"""Local Office recalculation, formula scanning, and cache preservation."""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
from posixpath import join, normpath
from shutil import which
import subprocess
import sys
from xml.etree.ElementTree import SubElement, register_namespace, tostring
from zipfile import BadZipFile, ZipFile

from defusedxml.ElementTree import fromstring

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ERRORS = frozenset(("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!",
                    "#SPILL!", "#CALC!", "#GETTING_DATA", "#BLOCKED!", "#CONNECT!", "#FIELD!"))
register_namespace("", MAIN)


def tag(name):
    return f"{{{MAIN}}}{name}"


@dataclass(frozen=True)
class Cache:
    kind: str
    value: str

    def number(self):
        if self.kind != "n":
            return None
        try:
            number = Decimal(self.value)
        except InvalidOperation:
            return None
        return number if number.is_finite() else None


@dataclass(frozen=True)
class Calculation:
    backend: str
    caches: dict[tuple[str, str], Cache]
    issues: tuple[tuple[str, str], ...]
    note: str = ""


def sheet_parts(archive):
    workbook = fromstring(archive.read("xl/workbook.xml"))
    links = fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {link.attrib["Id"]: link.attrib["Target"] for link in links}
    paths = {}
    for sheet in workbook.find(tag("sheets")):
        target = targets[sheet.attrib[f"{{{REL}}}id"]]
        paths[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else normpath(join("xl", target))
    return paths


def formulas(book):
    return {
        (sheet.title, cell.coordinate): cell.value
        for sheet in book.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.data_type == "f"
    }


def read_caches(path, expected):
    caches = {}
    with ZipFile(path) as archive:
        parts = sheet_parts(archive)
        for name in {key[0] for key in expected} & parts.keys():
            root = fromstring(archive.read(parts[name]))
            for cell in root.iter(tag("c")):
                key = (name, cell.get("r"))
                if key not in expected or cell.find(tag("f")) is None:
                    continue
                kind = cell.get("t", "n")
                value = cell.find(tag("v"))
                if value is not None and kind in {"n", "str", "b", "e"}:
                    if value.text is not None or kind == "str":
                        caches[key] = Cache(kind, value.text or "")
    return caches


def scan(expected, caches):
    issues = []
    for (sheet, coordinate), formula in expected.items():
        key = (sheet, coordinate)
        position = f"{sheet}!{coordinate}"
        cache = caches.get(key)
        if isinstance(formula, str) and "#REF!" in formula.upper():
            issues.append((position, "公式包含 #REF!"))
        if cache is None:
            issues.append((position, "缺少公式结果缓存或公式已丢失"))
        elif cache.kind == "e":
            issues.append((position, cache.value if cache.value in ERRORS else "Excel 公式错误"))
        elif cache.kind == "n" and cache.number() is None:
            issues.append((position, "数值缓存无效"))
    return tuple(issues)


def write_caches(source, target, caches):
    """Copy result caches into the original package without adopting Office rewrites."""
    with ZipFile(source) as original, ZipFile(target, "w") as result:
        by_name = sheet_parts(original)
        names = {by_name[sheet]: sheet for sheet, _ in caches}
        for entry in original.infolist():
            payload = original.read(entry.filename)
            if entry.filename in names:
                sheet = names[entry.filename]
                root = fromstring(payload)
                for cell in root.iter(tag("c")):
                    cache = caches.get((sheet, cell.get("r")))
                    if cache is None or cell.find(tag("f")) is None:
                        continue
                    cell.set("t", cache.kind)
                    value = cell.find(tag("v"))
                    if value is None:
                        value = SubElement(cell, tag("v"))
                    value.text = cache.value
                payload = tostring(root, encoding="utf-8")
            result.writestr(entry, payload)


def excel_worker(source, target):
    import xlwings as xw

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.enable_events = False
        app.api.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        book = app.books.open(str(source), update_links=False, read_only=True)
        try:
            app.api.CalculateFullRebuild()
            book.save(str(target))
        finally:
            book.close()


def soffice_path():
    binary = which("soffice") or which("libreoffice")
    if binary or os.name != "nt":
        return binary
    for variable in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(variable)
        if base:
            candidate = Path(base) / "LibreOffice" / "program" / "soffice.com"
            if candidate.is_file():
                return str(candidate)
    return None


def engine_command(backend, source, folder):
    if backend == "xlwings":
        return [sys.executable, str(Path(__file__).resolve()), "--excel-worker", str(source), str(folder / source.name)]
    binary = soffice_path()
    if binary is None:
        raise FileNotFoundError
    return [
        binary, f"-env:UserInstallation={(folder / 'profile').as_uri()}", "--headless",
        "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(folder), str(source),
    ]


def recalculate(source, expected, backend, timeout):
    order = ("xlwings", "libreoffice") if os.name == "nt" else ("libreoffice",)
    candidates = order if backend == "auto" else (backend,)
    if backend == "none":
        return Calculation("none", {}, scan(expected, {}), "本次执行跳过了公式重算。")
    for name in candidates:
        folder = source.parent / name
        folder.mkdir()
        try:
            env = {**os.environ, "TMP": str(folder), "TEMP": str(folder), "TMPDIR": str(folder)}
            subprocess.run(engine_command(name, source, folder), cwd=folder, env=env,
                           stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           check=True, timeout=timeout)
            caches = read_caches(folder / source.name, expected)
        except (OSError, subprocess.SubprocessError, BadZipFile, KeyError, SyntaxError):
            continue
        return Calculation(name, caches, scan(expected, caches))
    return Calculation("unavailable", {}, scan(expected, {}), "本地重算引擎不可用、重算失败或超时。")


if __name__ == "__main__" and len(sys.argv) == 4 and sys.argv[1] == "--excel-worker":
    excel_worker(Path(sys.argv[2]), Path(sys.argv[3]))
