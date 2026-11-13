"""Put VBA source files into an .xlsx / .xlsm and save it as .xlsm.

Usage:
    python3 inject_vba.py <input.xlsx|xlsm> <vba-dir> <output.xlsm>

- Standard modules: <vba-dir>/*.bas
- Class modules:    <vba-dir>/*.cls
- Document modules: <vba-dir>/ThisWorkbook.cls and <vba-dir>/sheets/<CodeName>.cls
  (optional; an empty document module is generated for every sheet otherwise)

Any existing vbaProject.bin in the input is replaced. Sheet contents and
formatting are kept as they are (the zip parts are copied byte for byte).
Python standard library only.
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vbaproject import (  # noqa: E402
    DOC_BASE_WORKBOOK,
    DOC_BASE_WORKSHEET,
    Module,
    build_vba_project,
    parse_source_file,
)

CT_MACRO_WORKBOOK = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def _sheet_parts(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Return [(sheet part path, sheet name)] in workbook order."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {
        m.group(1): m.group(2)
        for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)
    }
    targets.update(
        {
            m.group(2): m.group(1)
            for m in re.finditer(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)
        }
    )
    result = []
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        name = re.search(r'name="([^"]*)"', tag).group(1)  # type: ignore[union-attr]
        rid = re.search(r'r:id="([^"]*)"', tag).group(1)  # type: ignore[union-attr]
        target = targets[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        result.append((target, name))
    return result


def _set_sheet_codename(xml: str, code_name: str) -> str:
    if re.search(r"<sheetPr\b", xml):
        if "codeName=" in xml.split("<sheetPr", 1)[1].split(">", 1)[0]:
            return re.sub(r'(<sheetPr\b[^>]*codeName=")[^"]*(")', rf"\g<1>{code_name}\2", xml, count=1)
        return re.sub(r"<sheetPr\b", f'<sheetPr codeName="{code_name}"', xml, count=1)
    return re.sub(r"(<worksheet\b[^>]*>)", rf'\1<sheetPr codeName="{code_name}"/>', xml, count=1)


def _sheet_codename(xml: str) -> str | None:
    m = re.search(r'<sheetPr\b[^>]*codeName="([^"]+)"', xml)
    return m.group(1) if m else None


def inject(src: Path, vba_dir: Path, dst: Path, exclude: set[str] | None = None) -> None:
    with zipfile.ZipFile(src) as z:
        parts = {name: z.read(name) for name in z.namelist()}
        sheets = _sheet_parts(z)

    # Document modules: one for the workbook and one per sheet.
    modules: list[Module] = []
    wb_cls = vba_dir / "ThisWorkbook.cls"
    wb_code = parse_source_file("ThisWorkbook", _read_text(wb_cls), "document").code if wb_cls.exists() else ""
    modules.append(Module("ThisWorkbook", "document", wb_code, DOC_BASE_WORKBOOK))

    used = set()
    for i, (part, _name) in enumerate(sheets, start=1):
        xml = parts[part].decode("utf-8")
        code_name = _sheet_codename(xml) or f"Sheet{i}"
        while code_name in used:
            code_name += "_"
        used.add(code_name)
        parts[part] = _set_sheet_codename(xml, code_name).encode("utf-8")
        sheet_cls = vba_dir / "sheets" / f"{code_name}.cls"
        code = parse_source_file(code_name, _read_text(sheet_cls), "document").code if sheet_cls.exists() else ""
        modules.append(Module(code_name, "document", code, DOC_BASE_WORKSHEET))

    exclude = exclude or set()
    for path in sorted(vba_dir.glob("*.bas")):
        if path.stem in exclude:
            continue
        modules.append(parse_source_file(path.stem, _read_text(path), "standard"))
    for path in sorted(vba_dir.glob("*.cls")):
        if path.stem == "ThisWorkbook":
            continue
        modules.append(parse_source_file(path.stem, _read_text(path), "class"))

    for m in modules:
        if any(ord(c) > 127 for c in m.code):
            raise SystemExit(f"{m.name}: VBA source must be ASCII only (see docs/design.md)")
        if "\\" in m.code:
            # Japanese Mac Excel reads 0x5C in source code as a yen sign.
            raise SystemExit(f"{m.name}: do not write a backslash in VBA source; use ChrW(92)")

    parts["xl/vbaProject.bin"] = build_vba_project(modules)

    # workbook.xml: codeName
    wb = parts["xl/workbook.xml"].decode("utf-8")
    if "<workbookPr" in wb:
        if 'codeName="' not in wb.split("<workbookPr", 1)[1].split(">", 1)[0]:
            wb = wb.replace("<workbookPr", '<workbookPr codeName="ThisWorkbook"', 1)
    else:
        wb = re.sub(r"(<workbook\b[^>]*>)", r'\1<workbookPr codeName="ThisWorkbook"/>', wb, count=1)
    parts["xl/workbook.xml"] = wb.encode("utf-8")

    # [Content_Types].xml
    ct = parts["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(
        r'(<Override[^>]*PartName="/xl/workbook.xml"[^>]*ContentType=")[^"]*(")',
        rf"\g<1>{CT_MACRO_WORKBOOK}\2",
        ct,
    )
    if 'Extension="bin"' not in ct:
        ct = ct.replace("<Default ", f'<Default Extension="bin" ContentType="{CT_VBA}"/><Default ', 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    # workbook.xml.rels
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if REL_VBA not in rels:
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rIdVba1" Type="{REL_VBA}" Target="vbaProject.bin"/></Relationships>',
        )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        # [Content_Types].xml first, as Office writes it.
        out.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            out.writestr(name, data)
    print(f"wrote {dst} ({len(modules)} modules)")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(__doc__)
        raise SystemExit(2)
    inject(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))
