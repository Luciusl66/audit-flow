"""Passive local source readers and the adapter interface."""

from abc import ABC, abstractmethod
from contextlib import contextmanager
import csv
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from core.schema import AuditError, StandardData, TB_AMOUNT_FIELDS, ZERO, amount_value

SOURCE_TYPES = {".csv", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods"}


@dataclass(frozen=True)
class SourceCell:
    value: object
    kind: str = ""
    number_format: str = ""


@dataclass(frozen=True)
class SourceTable:
    sheet: int = 1
    header_row: int = 1
    encoding: str = "utf-8-sig"
    delimiter: str = ","


@contextmanager
def source_rows(path: Path, spec=SourceTable(), *, row_limit=None):
    extension = path.suffix.lower()
    if extension == ".csv":
        if spec.sheet != 1:
            raise AuditError("CSV 只有一个工作表。")
        with path.open(encoding=spec.encoding, newline="") as source:
            reader = csv.reader(source, strict=True, delimiter=spec.delimiter)
            yield (tuple(SourceCell(value) for value in row) for row in reader)
    elif extension in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        book = load_workbook(path, read_only=True, data_only=False, keep_links=False, keep_vba=False)
        try:
            if not 1 <= spec.sheet <= len(book.worksheets):
                raise AuditError("源表工作表序号超出范围。")
            tab = book.worksheets[spec.sheet - 1]
            tab.reset_dimensions()
            yield (
                tuple(SourceCell(cell.value, getattr(cell, "data_type", ""),
                                 getattr(cell, "number_format", "") or "") for cell in row)
                for row in tab.iter_rows()
            )
        finally:
            book.close()
    elif extension in {".xls", ".xlsb", ".ods"}:
        from python_calamine import CalamineWorkbook

        with CalamineWorkbook.from_path(path) as book:
            if not 1 <= spec.sheet <= len(book.sheet_names):
                raise AuditError("源表工作表序号超出范围。")
            rows = book.get_sheet_by_index(spec.sheet - 1).to_python(skip_empty_area=False, nrows=row_limit)
            yield (tuple(SourceCell(value) for value in row) for row in rows)
    else:
        raise AuditError("源文件须为 Excel 工作簿或 CSV。")


def blank(cell: SourceCell) -> bool:
    return cell.value is None or isinstance(cell.value, str) and not cell.value.strip()


def mapped_rows(rows, exact_names, label, header_row=1):
    for _ in range(header_row - 1):
        next(rows, None)
    headers = next(rows, ())
    positions = {}
    for field, name in exact_names.items():
        indexes = [
            i for i, cell in enumerate(headers)
            if cell.kind not in {"f", "e"} and cell.value == name
        ]
        if len(indexes) != 1:
            raise AuditError(
                f"{label} 字段 {field} 缺失或映射不唯一。"
                "请确认列名与脱敏探针输出逐字一致（含空格、全角/半角及标点）。"
            )
        positions[field] = indexes[0]
    if len(set(positions.values())) != len(positions):
        raise AuditError(f"{label} 多个标准字段映射到了同一源列。")
    for line, row in enumerate(rows, start=header_row + 1):
        if all(blank(cell) for cell in row):
            continue
        record = {}
        for field, index in positions.items():
            cell = row[index] if index < len(row) else SourceCell(None)
            if cell.kind in {"f", "e"}:
                raise AuditError(f"{label} 第 {line} 行的 {field} 包含公式或错误单元格。")
            record[field] = cell.value
        yield line, record


def convert_rows(rows, converters, label):
    for line, record in rows:
        for field, convert in converters.items():
            if field not in record:
                continue
            try:
                record[field] = convert(record[field])
            except ValueError:
                raise AuditError(f"{label} 第 {line} 行的 {field} 无法转换。") from None
        yield record


def source_totals(path, mapping, spec, blank_amounts_are_zero):
    """Read source controls separately from the adapter's transformed records."""
    try:
        names = {field: mapping[field] for field in TB_AMOUNT_FIELDS}
    except KeyError:
        raise AuditError("Adapter 必须指定全部六列原始 TB 金额。") from None
    converters = {name: lambda value: amount_value(value, optional=blank_amounts_are_zero) for name in names}
    totals = dict.fromkeys(TB_AMOUNT_FIELDS, ZERO)
    count = 0
    with source_rows(path, spec) as rows:
        mapped = mapped_rows(rows, names, "原始 TB", spec.header_row)
        for row in convert_rows(mapped, converters, "原始 TB"):
            count += 1
            for name, value in row.items():
                totals[name] += value
    return count, totals


class BaseAdapter(ABC):
    tb_columns: dict[str, str]
    ar_account: str
    cas_allowance_pairs: dict[str, tuple[str, ...]] | None = None
    tb_source = SourceTable()
    gl_source = SourceTable()
    blank_tb_amounts_are_zero = False

    @abstractmethod
    def transform(self, tb_path: Path, gl_path: Path) -> StandardData:
        """Normalize source records without suppressing conversion errors."""
