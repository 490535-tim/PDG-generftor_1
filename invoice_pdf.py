"""Консольный генератор PDF-счетов из CSV/JSON данных и HTML-шаблонов.

Зависимости (установка через pip): jinja2, weasyprint.
Системные требования — библиотеки Pango/Cairo:
    Windows — GTK3-runtime либо MSYS2 (pacman -S mingw-w64-ucrt-x86_64-pango);
    macOS   — brew install pango cairo;
    Linux   — системные пакеты pango и cairo.
Актуальная инструкция по установке WeasyPrint:
    https://doc.courtbouillon.org/weasyprint/stable/install.html

Запуск:
    python invoice_pdf.py          — интерактивная генерация счетов;
    python invoice_pdf.py --check  — самопроверка (импорты, демо-файлы, шрифт, рендер);
    python invoice_pdf.py --all <файл данных> [шаблон] — пакетная генерация без меню.

Шаблоны: templates/{minimal,classic,modern}.html. Акцентный цвет и зебра
настраиваются CSS-переменными --accent / --zebra в блоке <style> шаблона.
Логотип: файл templates/logo.png (или logo.jpg / logo.svg) выводится
в шапке счёта; если файла нет, блок логотипа не отображается.
Блок «Примечания / условия оплаты» появляется, если в данных есть поле
notes / comment / Примечания / комментарий.
Последний выбор файлов запоминается в .last_choice.json: в меню
файлов достаточно нажать Enter.
"""

import csv
import json
import os
import platform
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Optional, Union

import jinja2
from jinja2 import ChainableUndefined, Environment, FileSystemLoader, select_autoescape

# --- Windows: UTF-8 для консоли и поиск GTK3 до импорта weasyprint -------------
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

if platform.system() == "Windows":
    for _gtk_bin in (Path("C:/Program Files/GTK3-Runtime Win64/bin"),
                     Path("C:/Program Files (x86)/GTK3-Runtime/bin")):
        if _gtk_bin.is_dir():
            os.environ.setdefault("WEASYPRINT_DLL_DIRECTORIES", str(_gtk_bin))
            break

try:
    import weasyprint  # noqa: E402
    from weasyprint import CSS, HTML  # noqa: E402
except (ImportError, OSError) as _import_error:
    print(f"Не удалось загрузить WeasyPrint: {_import_error}", file=sys.stderr)
    print(
        "Требуются системные библиотеки Pango/Cairo:\n"
        "  Windows: pacman -S mingw-w64-ucrt-x86_64-pango (MSYS2) или GTK3-runtime;\n"
        "  macOS:   brew install pango cairo;\n"
        "  Linux:   установить системные пакеты pango и cairo.\n"
        "Актуальная инструкция: "
        "https://doc.courtbouillon.org/weasyprint/stable/install.html",
        file=sys.stderr,
    )
    sys.exit(1)

# --- Пути -----------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
FONTS_DIR = BASE_DIR / "fonts"

# --- Синонимы --------------------------------------------------------------------
INVOICE_KEY_SYNONYMS = frozenset({
    "invoiceid", "invoice", "id", "invoiceno", "invoicenumber",
    "номер", "номерсчета", "счет", "idсчета", "код",
})
INVOICE_KEY_SUBSTRINGS = ("invoice", "счет")
INVOICE_KEY_EXCLUDE = ("date", "дата", "total", "сумма", "итого")
ITEM_LIST_KEYS = ("items", "positions", "строки", "товары")
COLUMN_SYNONYMS = {
    "qty": frozenset({"qty", "quantity", "количество", "колво", "шт"}),
    "price": frozenset({"price", "cost", "цена", "стоимость"}),
    "total": frozenset({"total", "итого", "сумма"}),
}
ENCODINGS_CSV = ("utf-8-sig", "cp1251")
LOGO_FILENAMES = ("logo.png", "logo.jpg", "logo.jpeg", "logo.svg", "logo.gif")
FONT_FALLBACKS = (
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    Path("/Library/Fonts/DejaVuSans.ttf"),
    Path("C:/Windows/Fonts/DejaVuSans.ttf"),
    Path("C:/Windows/Fonts/arial.ttf"),
    Path("C:/Windows/Fonts/segoeui.ttf"),
)

# --- Базовый CSS ------------------------------------------------------------------
CSS_BODY = """
@page { size: A4; margin: 15mm 18mm 20mm 18mm;
        @bottom-right { content: "Стр. " counter(page) " из " counter(pages);
                        font-size: 9pt; color: #6b7280;
                        font-family: 'InvoiceFont', 'DejaVu Sans', 'Segoe UI',
                                     Arial, sans-serif; } }
body { font-family: 'InvoiceFont', 'DejaVu Sans', 'Roboto', 'Segoe UI', Arial, sans-serif;
       font-size: 11pt; color: #1f2937; line-height: 1.45; }
h1 { font-size: 17pt; margin: 0 0 4mm; }
p { margin: 1mm 0; }
table { width: 100%; border-collapse: collapse; margin-top: 6mm; }
th, td { border: 0.4pt solid #9aa0a6; padding: 2mm 3mm; text-align: left; vertical-align: top; }
th { background: #eef1f4; }
thead { display: table-header-group; }
tr { page-break-inside: avoid; }
tbody tr:nth-child(even) td { background: var(--zebra, transparent); }
.num { text-align: right; white-space: nowrap; }
tfoot td { font-weight: bold; background: #f6f8fa; }
.muted { color: #6b7280; font-size: 9.5pt; }
img.logo { height: 12mm; }
"""

# Переменные, отсутствующие в данных, но использованные в шаблоне.
_MISSING_VARS: set = set()


# --- Утилиты ----------------------------------------------------------------------
def normalize_name(name: Any) -> str:
    """Нижний регистр, 'ё' -> 'е', только буквы и цифры."""
    lowered = str(name).lower().replace("ё", "е")
    return "".join(ch for ch in lowered if ch.isalnum())


def plural_ru(count: int, one: str, few: str, many: str) -> str:
    """Русская форма множественного числа: 1 позиция, 2 позиции, 5 позиций."""
    if count % 10 == 1 and count % 100 != 11:
        return one
    if 2 <= count % 10 <= 4 and not 12 <= count % 100 <= 14:
        return few
    return many


def safe_file_id(invoice_id: Any) -> str:
    """Безопасное имя файла: output/invoice_<safe_id>.pdf."""
    safe = re.sub(r"[^\w\-.]+", "_", str(invoice_id)).strip("_")
    return safe or "noid"


# --- Поиск ключа номера счёта -------------------------------------------------------
def find_key_candidate(keys: Iterable[Any]) -> Optional[str]:
    """Точное совпадение с синонимом; иначе подстрока invoice/счет (без полей даты и сумм)."""
    keys = list(keys)
    for key in keys:
        if normalize_name(key) in INVOICE_KEY_SYNONYMS:
            return key
    candidates = []
    for key in keys:
        norm = normalize_name(key)
        if not any(sub in norm for sub in INVOICE_KEY_SUBSTRINGS):
            continue
        if any(bad in norm for bad in INVOICE_KEY_EXCLUDE):
            continue
        candidates.append(key)
    if candidates:
        return min(candidates, key=len)
    return None


def find_invoice_key(keys: Iterable[Any]) -> Optional[str]:
    """Приоритет: точный синоним, подстрока, иначе первое поле."""
    keys = list(keys)
    candidate = find_key_candidate(keys)
    if candidate is not None:
        return candidate
    return keys[0] if keys else None


def find_json_invoice_key(header: dict, items: list) -> Optional[str]:
    """Для вложенного JSON ключ ищется сначала в полях родителя, затем в позициях."""
    key = find_key_candidate(header.keys())
    if key is None and items:
        key = find_key_candidate(items[0].keys())
    if key is None:
        key = next(iter(header), None)
    if key is None and items:
        key = next(iter(items[0]), None)
    return key


# --- Чтение данных -------------------------------------------------------------------
def read_csv_file(path: Path) -> list:
    """CSV: кодировки utf-8-sig -> cp1251; разделитель через csv.Sniffer (';,\t', иначе ';')."""
    raw = path.read_bytes()
    text = None
    last_error: Optional[UnicodeDecodeError] = None
    for encoding in ENCODINGS_CSV:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError as error:
            last_error = error
    if text is None:
        raise ValueError(f"не удалось декодировать файл ({last_error})")
    lines = text.splitlines()
    sample = "\n".join(lines[:10])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
        reader = csv.DictReader(lines, dialect=dialect)
    except csv.Error:
        reader = csv.DictReader(lines, delimiter=";")
    rows = []
    for row in reader:
        clean = {
            str(key): ("" if value is None else value)
            for key, value in row.items()
            if key is not None
        }
        rows.append(clean)
    return rows


def flatten_dict(source: dict, prefix: str = "") -> dict:
    """Разворачивает вложенные словари с префиксом: customer -> customer_name."""
    flat = {}
    for key, value in source.items():
        full_key = f"{prefix}{key}"
        if isinstance(value, dict):
            flat.update(flatten_dict(value, f"{full_key}_"))
        else:
            flat[full_key] = value
    return flat


def split_json_record(record: dict) -> tuple:
    """Разделяет запись JSON на шапку и плоские позиции (поле позиции побеждает)."""
    items_key = None
    for key in ITEM_LIST_KEYS:
        if isinstance(record.get(key), list):
            items_key = key
            break
    if items_key is None:
        for key, value in record.items():
            if isinstance(value, list) and value and all(
                isinstance(item, dict) for item in value
            ):
                items_key = key
                break
    raw_items = record.get(items_key, []) if items_key else []
    if not isinstance(raw_items, list):
        raise ValueError("список позиций счёта должен быть массивом JSON")
    flat_header = flatten_dict(
        {key: value for key, value in record.items() if key != items_key}
    )
    items = [
        {**flat_header, **flatten_dict(item)}
        for item in raw_items
        if isinstance(item, dict)
    ]
    return flat_header, items


def read_json_file(path: Path) -> list:
    """JSON: массив, одиночный объект или объект-обёртка с первым полем-списком."""
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(data, list):
        records = data
    elif isinstance(data, dict):
        records = None
        for value in data.values():
            if isinstance(value, list) and all(
                isinstance(item, dict) for item in value
            ):
                records = value
                break
        if records is None:
            records = [data]
    else:
        raise ValueError("ожидался массив или объект JSON")
    result = []
    for record in records:
        if not isinstance(record, dict):
            raise ValueError("каждый счёт должен быть объектом JSON")
        result.append(split_json_record(record))
    return result


# --- Группировка ----------------------------------------------------------------------
def new_group(invoice_id: str) -> dict:
    return {"invoice_id": invoice_id, "header": {}, "items": []}


def group_csv_rows(rows: list) -> list:
    """Группирует строки CSV по номеру счёта с сохранением порядка."""
    key = find_invoice_key(rows[0].keys()) if rows else None
    groups: dict = {}
    for row in rows:
        invoice_id = str(row.get(key, "")) if key is not None else ""
        group = groups.get(invoice_id)
        if group is None:
            group = new_group(invoice_id)
            group["header"] = dict(row)
            groups[invoice_id] = group
        group["items"].append(row)
    return list(groups.values())


def group_json_records(records: list) -> list:
    """Каждая запись JSON — счёт; записи с одинаковым номером объединяются."""
    groups: dict = {}
    for header, items in records:
        key = find_json_invoice_key(header, items)
        invoice_id = ""
        if key is not None:
            if key in header and header[key] is not None:
                invoice_id = str(header[key])
            elif items and key in items[0] and items[0][key] is not None:
                invoice_id = str(items[0][key])
        group = groups.get(invoice_id)
        if group is None:
            group = new_group(invoice_id)
            groups[invoice_id] = group
        if not group["header"]:
            group["header"] = header
        group["items"].extend(items)
    return list(groups.values())


def load_invoices(path: Path) -> list:
    """Читает CSV или JSON и возвращает список счетов."""
    if path.suffix.lower() == ".json":
        return group_json_records(read_json_file(path))
    return group_csv_rows(read_csv_file(path))


# --- Числа и суммы -----------------------------------------------------------------------
_SPACE_RE = re.compile(r"[\s\u00a0\u202f\u2007\u2009]+")
_CURRENCY_RE = re.compile(r"(?:₽|\$|€|£|руб|rub)", re.IGNORECASE)


def _resolve_single_separator(text: str, sep: str) -> str:
    """Разрешает одиночный разделитель: тысячи ('1.234') или десятичный ('12,5')."""
    parts = text.split(sep)
    if len(parts) == 2:
        head, tail = parts
        if len(tail) == 3 and head.lstrip("+-") not in ("", "0"):
            return head + tail  # тысячи: '1.234', '1,234'
        return head + "." + tail  # десятичные: '12,5', '0,125'
    if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
        return "".join(parts)  # тысячи: '1.234.567'
    return text


def _normalize_number_text(value: Any) -> str:
    """Подготовка строки числа: пробелы и валюта прочь, разделители приведены к '.'."""
    text = _CURRENCY_RE.sub("", _SPACE_RE.sub("", str(value))).strip()
    comma = text.rfind(",")
    dot = text.rfind(".")
    if comma != -1 and dot != -1:
        if comma > dot:  # '1.234,56'
            text = text.replace(".", "").replace(",", ".")
        else:  # '1,234.56'
            text = text.replace(",", "")
    elif comma != -1:
        text = _resolve_single_separator(text, ",")
    elif dot != -1:
        text = _resolve_single_separator(text, ".")
    return text.rstrip(".")


def parse_number(value: Any) -> Decimal:
    """Универсальный парсер числа: пусто -> 0; '1 234,56', '1.234,56', '1,234.56', валюта."""
    if isinstance(value, Decimal):
        return value
    if value is None:
        return Decimal("0")
    text = _normalize_number_text(value)
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def parse_number_strict(value: Any) -> Optional[Decimal]:
    """Decimal, если значение — корректное число; иначе None (пусто тоже None)."""
    if isinstance(value, Decimal):
        return value
    if value is None:
        return None
    text = _normalize_number_text(value)
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def money2(value: Decimal) -> Decimal:
    """Округление до 2 знаков по правилу HALF_UP."""
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def format_money(value: Any) -> str:
    """Вывод вида '1 234,56' (неразрывный пробел, запятая — десятичный разделитель)."""
    amount = money2(parse_number(value))
    return f"{amount:,}".replace(",", "\u00a0").replace(".", ",")


def find_column(keys: Iterable[Any], kind: str) -> Optional[str]:
    """Ищет колонку по общему словарю синонимов (qty / price / total)."""
    synonyms = COLUMN_SYNONYMS[kind]
    for key in keys:
        if normalize_name(key) in synonyms:
            return key
    return None


def with_amounts(items: list) -> list:
    """Добавляет в каждую позицию amount = qty * price (или итог колонки, или 0)."""
    result = []
    for item in items:
        row = dict(item)
        qty_key = find_column(row.keys(), "qty")
        price_key = find_column(row.keys(), "price")
        total_key = find_column(row.keys(), "total")
        if qty_key is not None and price_key is not None:
            amount = parse_number(row[qty_key]) * parse_number(row[price_key])
        elif total_key is not None:
            amount = parse_number(row[total_key])
        else:
            amount = Decimal("0")
        row["amount"] = money2(amount)
        result.append(row)
    return result


def invoices_total(items: list) -> Decimal:
    return money2(sum((row["amount"] for row in items), Decimal("0")))


def validate_invoices(invoices: list) -> list:
    """Предупреждения о подозрительных данных до генерации PDF."""
    warnings: list = []
    for invoice in invoices:
        invoice_id = invoice["invoice_id"] or "<без номера>"
        clients = set()
        for row in invoice["items"]:
            for key in ("Клиент", "customer_name"):
                value = str(row.get(key) or "").strip()
                if value:
                    clients.add(value)
        if len(clients) > 1:
            warnings.append(
                f"Счёт {invoice_id}: строки с разными клиентами "
                f"({', '.join(sorted(clients))}) — проверьте номер счёта"
            )
        for index, row in enumerate(invoice["items"], 1):
            name = (
                row.get("Наименование") or row.get("name") or row.get("товар")
                or f"позиция {index}"
            )
            qty_key = find_column(row.keys(), "qty")
            price_key = find_column(row.keys(), "price")
            if qty_key is not None:
                qty = parse_number_strict(row.get(qty_key))
                if qty is None or qty == 0:
                    warnings.append(
                        f"Счёт {invoice_id}, «{name}»: пустое или нечисловое "
                        f"количество ({row.get(qty_key)!r})"
                    )
                elif qty < 0:
                    warnings.append(
                        f"Счёт {invoice_id}, «{name}»: отрицательное количество"
                    )
            if price_key is not None:
                price = parse_number_strict(row.get(price_key))
                if price is None:
                    warnings.append(
                        f"Счёт {invoice_id}, «{name}»: нечисловая цена "
                        f"({row.get(price_key)!r})"
                    )
                elif price < 0:
                    warnings.append(
                        f"Счёт {invoice_id}, «{name}»: отрицательная цена"
                    )
    return warnings


def invoice_preview(invoice: dict) -> str:
    """Короткая сводка счёта: клиент, дата, позиции, итог."""
    header = invoice["header"]
    client = header.get("Клиент") or header.get("customer_name") or "клиент не указан"
    date = header.get("Дата") or header.get("date") or "дата не указана"
    items = with_amounts(invoice["items"])
    count = len(items)
    currency = header.get("Валюта") or header.get("currency") or "руб."
    return (
        f"{client}, {date}, {count} "
        f'{plural_ru(count, "позиция", "позиции", "позиций")}, '
        f"итого {format_money(invoices_total(items))} {currency}"
    )


# --- Шрифты и CSS ---------------------------------------------------------------------------
def find_logo(base_dir: Path = TEMPLATES_DIR) -> Optional[str]:
    """Имя файла логотипа в папке шаблона (resolved через base_url), иначе None."""
    for name in LOGO_FILENAMES:
        if (base_dir / name).is_file():
            return name
    return None


def find_font() -> Optional[Path]:
    """Первый найденный TTF: папка fonts/ рядом со скриптом, затем системные пути."""
    if FONTS_DIR.is_dir():
        for ttf in sorted(FONTS_DIR.glob("*.ttf")):
            return ttf
    for candidate in FONT_FALLBACKS:
        if candidate.is_file():
            return candidate
    return None


def describe_font(font_path: Optional[Path]) -> str:
    """Описание шрифта для --check: путь, поддержка кириллицы и знака №."""
    if font_path is None:
        return "TTF-шрифт не найден — кириллица может отобразиться неверно"
    try:
        from fontTools.ttLib import TTFont
    except ImportError:  # pragma: no cover - fontTools ставится вместе с weasyprint
        return f"{font_path.name} (проверка глифов недоступна: нет fontTools)"
    try:
        font = TTFont(str(font_path), fontNumber=0, lazy=True)
        cmap = font.getBestCmap() or {}
        cyrillic = "кириллица — OK" if any(
            0x0400 <= code_point <= 0x04FF for code_point in cmap
        ) else "КИРИЛЛИЦА НЕ ПОДДЕРЖИВАЕТСЯ"
        no_sign = "знак № — OK" if 0x2116 in cmap else "знак № ОТСУТСТВУЕТ"
        return f"{font_path.name}: {cyrillic}, {no_sign}"
    except Exception as error:  # noqa: BLE001 - любой битый шрифт не должен ломать --check
        return f"{font_path.name} (не удалось проверить: {error})"


def build_base_css() -> str:
    """Базовый CSS с @font-face для первого найденного шрифта."""
    font = find_font()
    if font is not None:
        font_face = (
            f"@font-face {{ font-family: 'InvoiceFont';"
            f" src: url('{font.as_uri()}'); }}\n"
        )
    else:
        print(
            "Предупреждение: TTF-шрифт не найден, "
            "кириллица может отобразиться некорректно."
        )
        font_face = ""
    return font_face + CSS_BODY


# --- Jinja2 ------------------------------------------------------------------------------------
class CollectingUndefined(ChainableUndefined):
    """Undefined, который не роняет рендер, но запоминает имя переменной."""

    def _remember(self) -> None:
        _MISSING_VARS.add(self._undefined_name)

    def __str__(self) -> str:
        self._remember()
        return ""

    def __iter__(self):
        self._remember()
        return iter(())

    def __len__(self) -> int:
        self._remember()
        return 0

    def __call__(self, *args: Any, **kwargs: Any) -> str:
        self._remember()
        return ""


def make_environment(loader_dir: Path = TEMPLATES_DIR) -> Environment:
    env = Environment(
        loader=FileSystemLoader(str(loader_dir)),
        autoescape=select_autoescape(),
        trim_blocks=True,
        lstrip_blocks=True,
        undefined=CollectingUndefined,
    )
    env.filters["num"] = parse_number
    env.filters["money"] = format_money
    return env


# --- Рендер PDF -----------------------------------------------------------------------------------
def render_invoice_pdf(
    env: Environment,
    template_name: str,
    invoice: dict,
    base_css: str,
    output_path: Optional[Path] = None,
    quiet: bool = False,
    base_dir: Path = TEMPLATES_DIR,
) -> tuple:
    """Рендерит счёт в PDF; возвращает (путь, список отсутствующих переменных)."""
    items = with_amounts(invoice["items"])
    total = invoices_total(items)
    invoice_id = str(invoice["invoice_id"])
    _MISSING_VARS.clear()
    template = env.get_template(template_name)
    html = template.render({
        "invoice_id": invoice_id,
        "invoice": invoice["header"],
        "items": items,
        "rows": items,
        "total": total,
        "now": datetime.now(),
        "base_css": base_css,
        "logo": find_logo(base_dir),
    })
    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / f"invoice_{safe_file_id(invoice_id)}.pdf"
    overwritten = output_path.exists()
    HTML(string=html, base_url=str(base_dir)).write_pdf(
        str(output_path), stylesheets=[CSS(string=base_css)]
    )
    if not quiet:
        if overwritten:
            print(f"Файл перезаписан: {output_path}")
        else:
            print(f"Создан файл: {output_path}")
    return output_path, sorted(_MISSING_VARS)


def warn_missing(missing: list) -> None:
    if missing:
        print("Предупреждение: в данных отсутствуют переменные: " + ", ".join(missing))


# --- Открытие PDF -----------------------------------------------------------------------------------
def open_pdf_file(path: Path) -> None:
    system = platform.system()
    if system == "Windows":
        os.startfile(str(path))
    elif system == "Darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


def open_generated_pdf(path: Path) -> None:
    """Автоматически открывает созданный PDF (без запроса)."""
    try:
        open_pdf_file(path)
    except OSError as error:
        print(f"Не удалось открыть PDF автоматически: {error}")


# --- Интерактивная сессия ------------------------------------------------------------------------------
LAST_CHOICE_PATH = BASE_DIR / ".last_choice.json"


def load_last_choice() -> dict:
    """Последний выбор файлов (данные + шаблон) из .last_choice.json."""
    try:
        data = json.loads(LAST_CHOICE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_last_choice(data_name: str, template_name: str) -> None:
    try:
        LAST_CHOICE_PATH.write_text(
            json.dumps({"data": data_name, "template": template_name},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass  # память выбора — не критичная функция


def ask(prompt: str) -> str:
    return input(prompt).strip()


def ask_yes_no(prompt: str, default: bool) -> bool:
    answer = ask(prompt).lower()
    if not answer:
        return default
    if answer[0] in ("y", "д"):
        return True
    if answer[0] in ("n", "н"):
        return False
    return default


def choose_file(title: str, paths: list, last_name: Optional[str] = None) -> Optional[Path]:
    """Меню файлов; Enter — прежний выбор, если он есть в списке."""
    last_path = None
    if last_name:
        for path in paths:
            if path.name == last_name:
                last_path = path
                break
    print(f"\n{title}:")
    for index, path in enumerate(paths, 1):
        mark = "  <- прежний выбор" if path is last_path else ""
        print(f"  {index}) {path.name} — {path.stat().st_size} байт{mark}")
    hint = " (Enter — прежний выбор)" if last_path else ""
    while True:
        answer = ask(f"Выберите номер файла (0 — выход){hint}: ")
        if answer in ("0", "q", "Q"):
            return None
        if not answer and last_path is not None:
            return last_path
        if answer.isdigit() and 1 <= int(answer) <= len(paths):
            return paths[int(answer) - 1]
        print("Неверный ввод. Введите номер из списка, 0 или q для выхода.")


def choose_invoice(invoices: list) -> Union[dict, str, None]:
    """Меню счетов с предпросмотром; 'a' — все, 0/q — выход."""
    print("\nСчета:")
    for index, invoice in enumerate(invoices, 1):
        label = invoice["invoice_id"] or "<без номера>"
        print(f"  {index}) Счёт {label} — {invoice_preview(invoice)}")
    print("  a) Сгенерировать все счета")
    while True:
        answer = ask("Выберите счёт (номер, a — все, 0/q — выход): ")
        if answer in ("0", "q", "Q"):
            return None
        if answer in ("a", "A"):
            return "all"
        if answer.isdigit() and 1 <= int(answer) <= len(invoices):
            return invoices[int(answer) - 1]
        print("Неверный ввод. Введите номер счёта, a или q.")


def interactive_session() -> int:
    data_files = sorted(
        [*DATA_DIR.glob("*.csv"), *DATA_DIR.glob("*.json")], key=lambda p: p.name
    )
    template_files = sorted(
        [*TEMPLATES_DIR.glob("*.html"), *TEMPLATES_DIR.glob("*.htm")],
        key=lambda p: p.name,
    )
    if not data_files:
        print(f"Не найдено файлов данных (*.csv, *.json) в папке {DATA_DIR}.")
        return 1
    if not template_files:
        print(f"Не найдено шаблонов (*.html, *.htm) в папке {TEMPLATES_DIR}.")
        return 1

    last_choice = load_last_choice()
    data_path = choose_file("Файлы данных", data_files, last_choice.get("data"))
    if data_path is None:
        return 0
    template_path = choose_file("Шаблоны", template_files, last_choice.get("template"))
    if template_path is None:
        return 0
    save_last_choice(data_path.name, template_path.name)

    try:
        invoices = load_invoices(data_path)
    except (OSError, ValueError, csv.Error) as error:
        print(f"Ошибка чтения {data_path.name}: {error}")
        return 1
    if not invoices:
        print("В файле не найдено ни одного счёта.")
        return 1
    for warning in validate_invoices(invoices):
        print(f"Предупреждение: {warning}")

    base_css = build_base_css()
    env = make_environment()
    template_name = template_path.name

    while True:
        choice = choose_invoice(invoices)
        if choice is None:
            return 0
        if choice == "all":
            created = []
            for invoice in invoices:
                try:
                    path, missing = render_invoice_pdf(
                        env, template_name, invoice, base_css
                    )
                except PermissionError:
                    print(
                        "Нет доступа к файлу: закройте открытый PDF-файл "
                        "и повторите генерацию."
                    )
                    continue
                except Exception as error:  # noqa: BLE001 - понятное сообщение вместо трейсбека
                    print(
                        f'Ошибка генерации счёта {invoice["invoice_id"]}: {error}'
                    )
                    continue
                warn_missing(missing)
                created.append(path)
                open_generated_pdf(path)
            print(f"\nСоздано файлов: {len(created)}")
            for path in created:
                print(f"  {path}")
            continue
        try:
            path, missing = render_invoice_pdf(env, template_name, choice, base_css)
        except PermissionError:
            print(
                "Нет доступа к файлу: закройте открытый PDF-файл "
                "и повторите генерацию."
            )
            continue
        except Exception as error:  # noqa: BLE001 - понятное сообщение вместо трейсбека
            print(f'Ошибка генерации счёта {choice["invoice_id"]}: {error}')
            continue
        warn_missing(missing)
        open_generated_pdf(path)
        if not ask_yes_no("Сгенерировать ещё один? (y/N) ", default=False):
            return 0


# --- Пакетный режим (--all) ------------------------------------------------------------------------------
def resolve_input_path(raw_name: str, base_dir: Path) -> Optional[Path]:
    """Путь как есть, иначе относительно папки проекта (data/ или templates/)."""
    path = Path(raw_name)
    if path.is_file():
        return path
    alt = base_dir / raw_name
    return alt if alt.is_file() else None


def run_batch(args: list) -> int:
    """--all <файл данных> [шаблон]: все счета без меню, PDF не открываются."""
    if not args or len(args) > 2:
        print("Использование: python invoice_pdf.py --all <файл данных> [шаблон]")
        return 1
    data_path = resolve_input_path(args[0], DATA_DIR)
    if data_path is None:
        print(f"Файл данных не найден: {args[0]}")
        return 1
    if len(args) == 2:
        template_path = resolve_input_path(args[1], TEMPLATES_DIR)
        if template_path is None:
            print(f"Шаблон не найден: {args[1]}")
            return 1
    else:
        last = load_last_choice().get("template")
        template_files = sorted(
            [*TEMPLATES_DIR.glob("*.html"), *TEMPLATES_DIR.glob("*.htm")],
            key=lambda p: p.name,
        )
        template_path = None
        if last:
            for path in template_files:
                if path.name == last:
                    template_path = path
                    break
        if template_path is None and template_files:
            template_path = template_files[0]
        if template_path is None:
            print(f"Не найдено шаблонов (*.html, *.htm) в папке {TEMPLATES_DIR}.")
            return 1
    save_last_choice(data_path.name, template_path.name)

    try:
        invoices = load_invoices(data_path)
    except (OSError, ValueError, csv.Error) as error:
        print(f"Ошибка чтения {data_path.name}: {error}")
        return 1
    if not invoices:
        print("В файле не найдено ни одного счёта.")
        return 1
    for warning in validate_invoices(invoices):
        print(f"Предупреждение: {warning}")

    base_css = build_base_css()
    env = make_environment(template_path.parent)
    created = []
    for invoice in invoices:
        print(f"Счёт {invoice['invoice_id'] or '<без номера>'}: {invoice_preview(invoice)}")
        try:
            path, missing = render_invoice_pdf(
                env, template_path.name, invoice, base_css,
                base_dir=template_path.parent,
            )
        except PermissionError:
            print("Нет доступа к файлу: закройте открытый PDF-файл и повторите.")
            continue
        except Exception as error:  # noqa: BLE001 - понятное сообщение вместо трейсбека
            print(f'Ошибка генерации счёта {invoice["invoice_id"]}: {error}')
            continue
        warn_missing(missing)
        created.append(path)
    print(f"\nСоздано файлов: {len(created)}")
    for path in created:
        print(f"  {path}")
    return 0 if created else 1


# --- Самопроверка (--check) ------------------------------------------------------------------------------
def run_check() -> int:
    print("1. Проверка импортов...")
    print(f"   jinja2 {jinja2.__version__} — OK")
    print(f"   weasyprint {weasyprint.__version__} — OK")
    print("2. Поиск демо-файлов...")
    demo_data = DATA_DIR / "invoices_demo.csv"
    demo_templates = [TEMPLATES_DIR / f"{name}.html"
                      for name in ("minimal", "classic", "modern")]
    for path in (demo_data, *demo_templates):
        if not path.is_file():
            print(f"   Не найден {path}")
            return 1
        print(f"   {path.name} — OK ({path.stat().st_size} байт)")
    print("3. Проверка шрифта...")
    print(f"   {describe_font(find_font())}")
    print("4. Тестовый рендер демо-счёта во всех шаблонах...")
    handle, temp_name = tempfile.mkstemp(suffix=".pdf")
    os.close(handle)
    temp_path = Path(temp_name)
    try:
        invoices = load_invoices(demo_data)
        if not invoices:
            print("   В демо-файле не найдено счетов.")
            return 1
        base_css = build_base_css()
        env = make_environment()
        for template in demo_templates:
            render_invoice_pdf(
                env, template.name, invoices[0], base_css,
                output_path=temp_path, quiet=True,
            )
            if not temp_path.read_bytes()[:1024].startswith(b"%PDF"):
                print(f"   {template.name}: файл не похож на PDF (нет %PDF).")
                return 1
            print(f"   {template.name} — рендер OK")
        print("Самопроверка пройдена.")
        return 0
    except Exception as error:  # noqa: BLE001 - понятное сообщение вместо трейсбека
        print(f"   Ошибка при самопроверке: {error}")
        return 1
    finally:
        temp_path.unlink(missing_ok=True)


# --- main ------------------------------------------------------------------------------------------------
def main() -> int:
    try:
        argv = sys.argv[1:]
        if "--check" in argv:
            return run_check()
        if argv and argv[0] == "--all":
            return run_batch(argv[1:])
        return interactive_session()
    except KeyboardInterrupt:
        print("\nРабота прервана.")
        return 130
    except EOFError:
        print("\nВвод недоступен, выход.")
        return 0


if __name__ == "__main__":
    sys.exit(main())