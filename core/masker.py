"""Inspect source shapes locally, and emit only approved headers and fake data."""

import argparse
from collections import Counter
from collections.abc import Mapping
from datetime import date
from decimal import Decimal
from itertools import islice
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
if __package__ in {None, ""}:
    sys.path.insert(0, str(ROOT))

from adapters.base_adapter import SOURCE_TYPES, SourceCell, SourceTable, blank, source_rows
from core.schema import DIRECTIONS, GL_FIELDS, TB_AMOUNT_FIELDS, TB_FIELDS, AuditError, date_value

SAMPLE_SIZE = 10
HEADER_HINTS = {
    "科目编码": "account_code", "科目代码": "account_code", "科目名称": "account_name",
    "科目级次": "level", "级次": "level", "余额方向": "direction", "借贷方向": "direction",
    "期初借方": "open_debit", "期初贷方": "open_credit", "本期借方": "period_debit", "本期贷方": "period_credit",
    "期末借方": "close_debit", "期末贷方": "close_credit",
    "借方金额": "debit", "贷方金额": "credit", "记账日期": "date", "凭证日期": "date",
    "凭证号": "voucher_no", "凭证编号": "voucher_no", "摘要": "summary", "摘要说明": "summary",
    "往来单位": "aux_name", "辅助核算": "aux_name", "客商": "aux_name", "客户名称": "aux_name",
    "往来单位名称": "aux_name", "辅助核算名称": "aux_name", "客户": "aux_name", "供应商": "aux_name",
    "金额": "amount", "余额": "amount", "本币金额": "amount",
}
HEADER_HINTS.update({field: field for field in (*TB_FIELDS, *GL_FIELDS)})
AMOUNT_ROLES = {*TB_AMOUNT_FIELDS, "debit", "credit", "amount"}
NUMBER_SHAPE = re.compile(r"(?:[0-9]+|[0-9]{1,3}(?:[,，][0-9]{3})+)(?:\.([0-9]+))?")


def safe_path(path, extensions=SOURCE_TYPES):
    path = path.absolute()
    destination = path.resolve()
    for label in ("raw", "mock"):
        folder = ROOT / "data" / label
        if folder.resolve() == folder and path.is_relative_to(folder) and destination.is_relative_to(folder):
            if destination.is_file() and destination.suffix.lower() in extensions:
                return destination, label
    raise AuditError("输入文件须位于项目 data/raw/ 或 data/mock/，并使用支持的格式。")


def number_shape(cell):
    value = cell.value
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        return None
    text = str(value).strip()
    parentheses = text.startswith("(") and text.endswith(")")
    if parentheses:
        text = text[1:-1]
    sign = text[0] if text.startswith(("+", "-")) else ""
    if sign:
        text = text[1:]
    if parentheses and sign:
        return None
    match = NUMBER_SHAPE.fullmatch(text)
    if match is None:
        return None
    decimals = len(match[1] or "")
    separator = "，" if "，" in text else "," if "," in text else ""
    if not isinstance(value, str) and cell.number_format not in {"", "General"}:
        parts = cell.number_format.split(";")
        section = parts[1] if sign == "-" and len(parts) > 1 else parts[0]
        fmt = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", section)
        places = re.search(r"\.([0#]+)", fmt)
        if places:
            decimals = max(places[1].count("0"), min(decimals, len(places[1])))
        if re.search(r"[0#],[0#]", fmt):
            separator = ","
        parentheses = sign == "-" and "(" in fmt and ")" in fmt
        if parentheses:
            sign = ""
    return sign, parentheses, separator, decimals


def cell_type(cell):
    if cell.kind in {"f", "e"}:
        return "formula" if cell.kind == "f" else "excel_error"
    if blank(cell):
        return "blank"
    if isinstance(cell.value, bool):
        return "boolean"
    if isinstance(cell.value, date):
        return "date_cell"
    if isinstance(cell.value, (int, float, Decimal)):
        return "number_cell"
    if isinstance(cell.value, str):
        try:
            date_value(cell.value)
            return "date_text"
        except ValueError:
            return "number_text" if number_shape(cell) else "text"
    return "other"


def fake_amount(shape):
    sign, parentheses, separator, decimals = shape
    number = f"1{separator}000" if separator else "100"
    if decimals:
        number += "." + "0" * decimals
    return f"({number})" if parentheses else sign + number


def profile(path, sheet=1, header_row=1, encoding="utf-8-sig", delimiter=",", approved_headers=()):
    spec = SourceTable(sheet, header_row, encoding, delimiter)
    with source_rows(path, spec, row_limit=header_row + SAMPLE_SIZE) as rows:
        headers = next(islice(rows, header_row - 1, None), ())
        sample = list(islice(rows, SAMPLE_SIZE))
    width = max((len(row) for row in (headers, *sample)), default=0)
    columns = []
    masked = [[] for _ in sample]
    customers = {}
    for index in range(width):
        header = headers[index] if index < len(headers) else SourceCell(None)
        label = header.value if header.kind not in {"f", "e"} and isinstance(header.value, str) else None
        role = HEADER_HINTS.get(label.strip()) if label else None
        if role is None and isinstance(approved_headers, Mapping):
            role = approved_headers.get(label)
        visible = label if role or label in approved_headers else None
        types = Counter()
        features = Counter(thousands_count=0, parenthesized_negative_count=0, blank_count=0,
                           decimal_count=0, invalid_amount_count=0)
        enums = set()
        for row, result in zip(sample, masked, strict=True):
            cell = row[index] if index < len(row) else SourceCell(None)
            kind = cell_type(cell)
            types[kind] += 1
            shape = number_shape(cell) if kind in {"number_cell", "number_text"} else None
            if shape:
                features["thousands_count"] += bool(shape[2])
                features["parenthesized_negative_count"] += bool(shape[1])
                features["decimal_count"] += bool(shape[3])
            features["blank_count"] += kind == "blank"
            if role in AMOUNT_ROLES and kind != "blank" and shape is None:
                features["invalid_amount_count"] += 1
            if role == "summary" or kind == "blank":
                value = None
            elif role == "aux_name" and kind not in {"formula", "excel_error"}:
                key = str(cell.value).strip()
                value = customers.setdefault(key, f"客商_{len(customers) + 1:03d}")
            elif role == "direction" and isinstance(cell.value, str) and cell.value.strip().upper() in DIRECTIONS:
                value = cell.value.strip()
                enums.add(value)
            elif shape:
                value = fake_amount(shape)
            elif kind in {"date_cell", "date_text"}:
                value = "2000-01-01"
            else:
                value = "[已脱敏]"
            result.append(value)
        columns.append({"column": index + 1, "header": visible, "role_hint": role, "types": dict(types),
                        "features": dict(features), "enum_values": sorted(enums)})
    return {"sheet_index": sheet, "header_row": header_row, "rows_sampled": len(sample),
            "columns": columns, "masked_sample": masked}


def approved_names(path):
    if path is None:
        return ()
    source, _ = safe_path(path, {".json"})
    with source.open(encoding="utf-8-sig") as file:
        names = json.load(file)
    if isinstance(names, list) and all(isinstance(name, str) and name.strip() for name in names):
        return frozenset(names)
    if isinstance(names, dict) and all(name.strip() and isinstance(role, str)
                                       and role in HEADER_HINTS.values() for name, role in names.items()):
        return names
    raise AuditError("公开表头清单须为列名字符串数组或列名到标准字段的映射。")


class ProbeParser(argparse.ArgumentParser):
    def error(self, message):
        self.exit(2, "探针参数无效。请使用 --help 查看用法。\n")


def main(argv=None):
    parser = ProbeParser(description="本地读取表头和随后十行，输出脱敏 JSON")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--sheet", type=int, default=1)
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--encoding", default="utf-8-sig", help="CSV 编码")
    parser.add_argument("--delimiter", default=",", help="CSV 单字符分隔符，制表符可写 tab")
    parser.add_argument("--approved-headers", type=Path, help="本地数据持有人审核过的公开列名 JSON")
    args = parser.parse_args(argv)
    try:
        delimiter = "\t" if args.delimiter == "tab" else args.delimiter
        if min(args.sheet, args.header_row) < 1 or len(delimiter) != 1:
            raise AuditError("工作表和表头行从 1 开始，分隔符须为单字符。")
        path, _ = safe_path(args.input)
        result = profile(path, args.sheet, args.header_row, args.encoding, delimiter,
                         approved_names(args.approved_headers))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except AuditError as error:
        print(str(error), file=sys.stderr)
    except Exception:
        print("探针执行失败。请在本地检查文件格式、编码与工作表位置。", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
