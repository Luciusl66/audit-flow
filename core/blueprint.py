"""Load and enforce the reviewed workbook structure."""

from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl.worksheet.cell_range import CellRange
from yaml import YAMLError, safe_load

from core.schema import AuditError

BLUEPRINT_PATH = Path(__file__).resolve().parents[1] / "templates" / "blueprint.yaml"
EXCEL_ROWS = 1_048_576


@dataclass(frozen=True)
class SheetBlueprint:
    name: str
    source: str
    headers: tuple[tuple[CellRange, tuple[str, ...]], ...]
    data: CellRange
    anchors: dict[str, str]
    protected_formulas: tuple[CellRange, ...]
    detail_formulas: CellRange | None


@dataclass(frozen=True)
class Blueprint:
    template: Path
    sheets: tuple[SheetBlueprint, ...]

    def sheet(self, name):
        for spec in self.sheets:
            if spec.name == name:
                return spec
        raise AuditError(f"底稿蓝图缺少工作表 {name}。")

    def validate(self, book, *, source=None, moved_anchors=None):
        specs = [spec for spec in self.sheets if source is None or spec.source == source]
        if book.sheetnames != [spec.name for spec in specs]:
            raise AuditError("工作表清单或顺序与底稿蓝图不一致。")
        for spec in specs:
            sheet = book[spec.name]
            for area, labels in spec.headers:
                rows = sheet.iter_rows(min_row=area.min_row, max_row=area.max_row,
                                       min_col=area.min_col, max_col=area.max_col)
                header = tuple(cell.value if cell.data_type != "f" else None for cell in next(rows, ()))
                if header != labels:
                    raise AuditError(f"工作表 {spec.name} 的固定表头与底稿蓝图不一致。")
            for name, original in spec.anchors.items():
                address = (moved_anchors or {}).get(name, original)
                try:
                    locations = [(title, cell.replace("$", ""))
                                 for title, cell in book.defined_names[name].destinations]
                except (AttributeError, IndexError, KeyError, TypeError, ValueError):
                    locations = []
                if locations != [(spec.name, address)]:
                    raise AuditError(f"模板锚点 {name} 与底稿蓝图不一致。")
            if spec.source == "report":
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is not None and (cell.data_type == "f" or cell.coordinate not in spec.data):
                            raise AuditError(f"工作表 {spec.name} 的报告写入超出底稿蓝图范围或包含公式。")

    def frozen_formulas(self, book):
        result = {}
        for spec in self.sheets:
            for region in spec.protected_formulas:
                for row in book[spec.name].iter_rows(min_row=region.min_row, max_row=region.max_row,
                                                     min_col=region.min_col, max_col=region.max_col):
                    for cell in row:
                        if cell.data_type != "f":
                            raise AuditError(f"工作表 {spec.name} 的蓝图保护区域缺少公式。")
                        result[spec.name, cell.coordinate] = formula_signature(cell.value)
        return result


def formula_signature(value):
    if isinstance(value, str):
        return value
    return type(value), dict(value), getattr(value, "text", None)


def cell_range(value):
    if not isinstance(value, str) or not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?", value):
        raise ValueError
    area = CellRange(value)
    if area.max_row > EXCEL_ROWS or area.max_col > 16_384:
        raise ValueError
    return area


def parse_sheet(config):
    name, source = config["name"], config["source"]
    if (not isinstance(name, str) or not name.strip() or len(name) > 31 or re.search(r"[\\/?*\[\]:]", name)
            or source not in ("template", "report")):
        raise ValueError
    if set(config) - {"name", "source", "headers", "data", "anchors", "protected_formulas", "detail_formulas"}:
        raise ValueError
    headers = []
    for address, labels in config["headers"].items():
        area = cell_range(address)
        if (area.min_row != area.max_row or not isinstance(labels, list)
                or len(labels) != area.max_col - area.min_col + 1
                or not all(isinstance(label, str) and label.strip() for label in labels)):
            raise ValueError
        headers.append((area, tuple(labels)))
    if not headers:
        raise ValueError
    anchors = config.get("anchors", {})
    protected = config.get("protected_formulas", [])
    dynamic = config.get("detail_formulas")
    if not isinstance(protected, list) or not isinstance(anchors, dict):
        raise ValueError
    for key, address in anchors.items():
        if not isinstance(key, str) or not key or ":" in cell_range(address).coord:
            raise ValueError
    if source == "report" and (anchors or protected or dynamic is not None):
        raise ValueError
    return SheetBlueprint(name, source, tuple(headers), cell_range(config["data"]), anchors,
                          tuple(cell_range(region) for region in protected),
                          cell_range(dynamic) if dynamic is not None else None)


def load_blueprint(path=BLUEPRINT_PATH):
    try:
        path = path.absolute()
        if path.resolve() != path:
            raise ValueError
        config = safe_load(path.read_text(encoding="utf-8"))
        if (set(config) != {"schema_version", "workpapers"} or type(config["schema_version"]) is not int
                or config["schema_version"] != 1):
            raise ValueError
        workpaper = config["workpapers"]["accounts_receivable"]
        filename = workpaper["template"]
        if (set(workpaper) != {"template", "sheets"} or not isinstance(filename, str)
                or not re.fullmatch(r"[^\\/:]+\.xlsx", filename)):
            raise ValueError
        template = path.parent / filename
        if template.resolve() != template or not template.is_file():
            raise ValueError
        sheets = tuple(parse_sheet(entry) for entry in workpaper["sheets"])
        if (not sheets or len({spec.name.casefold() for spec in sheets}) != len(sheets)
                or sum(len(spec.anchors) for spec in sheets) != len({key for spec in sheets for key in spec.anchors})):
            raise ValueError
        return Blueprint(template, sheets)
    except (OSError, YAMLError, AttributeError, KeyError, TypeError, ValueError, RecursionError):
        raise AuditError("底稿蓝图或模板路径无效。请检查 templates/blueprint.yaml。") from None
