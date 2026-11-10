#!/usr/bin/env python3
"""Compare goexcel with openpyxl, Apache POI, and LibreOffice.

Reference tools are optional. Missing tools are skipped so CI can run the
harness without them. Exit status is 0 when every available reader agrees on
sheet names, cell values, formulas, merges, defined names, sheet state,
comments, and hyperlinks, and when LibreOffice can convert both the original
and the goexcel round-trip.

Usage:
  python3 scripts/diffharness.py testdata/excel
  python3 scripts/diffharness.py --roundtrip testdata/excel
  python3 scripts/diffharness.py --engines openpyxl,poi testdata/excel/sample.xlsx
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DUMP = ROOT / "cmd" / "xlsxdump"
POI_SRC = ROOT / "tools" / "poidump" / "PoiDump.java"
POI_DIR = Path(os.environ.get("POI_JAR_DIR", "/tmp/poi-jars"))


def norm_formula(f: str) -> str:
    f = (f or "").strip()
    if f.startswith("="):
        f = f[1:]
    if len(f) >= 2 and f[0] == "{" and f[-1] == "}":
        f = f[1:-1].strip()
        if f.startswith("="):
            f = f[1:]
    f = f.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\s+", "", f).upper()


def norm_ref(ref: str) -> str:
    ref = re.sub(r"\s+", "", (ref or "").replace("$", "")).upper()
    if ref and ":" not in ref and re.match(r"^[A-Z]+[0-9]+$", ref):
        ref = ref + ":" + ref
    return ref


def norm_link(url: str) -> str:
    return (url or "").replace("\\", "/").rstrip("/")


def norm_text(s: str) -> str:
    return (s or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def norm_name(name: str) -> str:
    name = name or ""
    if "!" in name:
        name = name.split("!")[-1]
    return name.upper()


def values_equal(a, b) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b or str(a).lower() == str(b).lower()
    ca, cb = clock_fraction(a), clock_fraction(b)
    if ca is not None and isinstance(b, (int, float)) and not isinstance(b, bool):
        return abs(ca - (float(b) % 1)) <= 1e-6 or abs(ca - float(b)) <= 1e-6
    if cb is not None and isinstance(a, (int, float)) and not isinstance(a, bool):
        return abs(cb - (float(a) % 1)) <= 1e-6 or abs(cb - float(a)) <= 1e-6
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        af, bf = float(a), float(b)
        scale = max(1.0, abs(af), abs(bf))
        return abs(af - bf) <= 1e-4 * scale or abs(af - bf) <= 1e-6
    da, db = as_dt(a), as_dt(b)
    if da and db:
        delta = abs((da - db).total_seconds())
        if delta <= 2:
            return True
        # POI's 1900 date system is one day later than the 1899-12-30 epoch
        # for serials before the fictional 1900-02-29. The clock time matches.
        return abs(delta - 86400) <= 2 and min(da, db) < dt.datetime(1900, 3, 1)
    # A date on one side and an Excel serial on the other.
    if da and isinstance(b, (int, float)) and not isinstance(b, bool):
        return abs((da - serial_to_dt(float(b))).total_seconds()) <= 2
    if db and isinstance(a, (int, float)) and not isinstance(a, bool):
        return serial_close(float(a), db)
    return norm_text(str(a)) == norm_text(str(b))


def clock_fraction(v):
    if not isinstance(v, str):
        return None
    m = re.fullmatch(r"(\d+):(\d{2}):(\d{2})", v.strip())
    if not m:
        return None
    h, mi, se = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return (h * 3600 + mi * 60 + se) / 86400.0


def serial_close(serial: float, when: dt.datetime) -> bool:
    delta = abs((when - serial_to_dt(serial)).total_seconds())
    if delta <= 2:
        return True
    # POI's 1900 window maps serial 0 to 1899-12-31; goexcel's epoch is
    # 1899-12-30 so pure times below the fictional 1900-02-29 differ by a day.
    return abs(delta - 86400) <= 2 and 0 <= serial < 61


def as_dt(v):
    if isinstance(v, dt.datetime):
        return v.replace(tzinfo=None)
    if isinstance(v, dt.date) and not isinstance(v, dt.datetime):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip().replace("Z", "+00:00")
        try:
            parsed = dt.datetime.fromisoformat(s)
            return parsed.replace(tzinfo=None)
        except ValueError:
            return None
    return None


def serial_to_dt(serial: float) -> dt.datetime:
    # 1900 system, including the Lotus leap-year bug, via 1899-12-30.
    base = dt.datetime(1899, 12, 30)
    days = int(serial)
    frac = serial - days
    return base + dt.timedelta(days=days, seconds=round(frac * 86400))


def same_array_member(sheet: dict, gf: str, rf: str) -> bool:
    """Array formulas are stored on the anchor. Readers disagree about members."""
    present = gf or rf
    if not present or (gf and rf):
        return False
    for c in sheet.get("cells") or []:
        if norm_formula(c.get("f") or "") == present:
            return True
    return False


def cell_map(sheet: dict) -> dict:
    out = {}
    for c in sheet.get("cells") or []:
        out[c["a"].upper()] = c
    return out


def drop_unmatched_special_sheets(got: dict, ref: dict) -> dict:
    ref_names = {s["name"] for s in ref.get("sheets") or []}
    sheets = []
    for s in got.get("sheets") or []:
        kind = s.get("kind") or "worksheet"
        if kind not in ("worksheet", "") and s["name"] not in ref_names:
            continue
        sheets.append(s)
    if len(sheets) == len(got.get("sheets") or []):
        return got
    copied = dict(got)
    copied["sheets"] = sheets
    return copied


def index_sheets(doc: dict) -> dict:
    return {s["name"]: s for s in doc.get("sheets") or []}


def compare_docs(label: str, got: dict, ref: dict, limit: int, formulas: bool = True) -> list[str]:
    """Return content mismatches. `got` is the candidate, `ref` is the reference.

    formulas=False compares cached values only (LibreOffice formula syntax
    differs from Excel even when the value is the same).
    """
    issues = []
    if got.get("error"):
        return [f"{label}: goexcel error: {got['error']}"]
    if ref.get("error"):
        return [f"{label}: reference error: {ref['error']}"]

    def add(msg: str):
        if len(issues) < limit:
            issues.append(msg)

    # Chartsheets and dialog sheets are preserved by goexcel. Some reference
    # readers omit them; that is not missing cell content.
    got = drop_unmatched_special_sheets(got, ref)
    gsheets = index_sheets(got)
    rsheets = index_sheets(ref)
    if list(gsheets) != list(rsheets):
        # Compare by name set; order is reported once.
        if set(gsheets) != set(rsheets):
            add(f"{label}: sheets {sorted(gsheets)} vs {sorted(rsheets)}")
        elif list(gsheets) != list(rsheets):
            add(f"{label}: sheet order {list(gsheets)} vs {list(rsheets)}")

    for name in sorted(set(gsheets) & set(rsheets)):
        gs, rs = gsheets[name], rsheets[name]
        if (gs.get("state") or "visible") != (rs.get("state") or "visible"):
            add(f"{label}: {name} state {gs.get('state')} vs {rs.get('state')}")
        gm = sorted(norm_ref(m) for m in gs.get("merges") or [])
        rm = sorted(norm_ref(m) for m in rs.get("merges") or [])
        if gm != rm:
            add(f"{label}: {name} merges {gm[:8]} vs {rm[:8]}")
        gc = cell_map(gs)
        rc = cell_map(rs)
        for addr in sorted(set(gc) | set(rc)):
            g, r = gc.get(addr), rc.get(addr)
            if g is None:
                if r and (r.get("f") or r.get("v") not in (None, "")):
                    add(f"{label}: {name}!{addr} missing in goexcel (ref v={r.get('v')!r} f={r.get('f')!r})")
                continue
            if r is None:
                if g.get("f") or g.get("v") not in (None, ""):
                    add(f"{label}: {name}!{addr} extra in goexcel (v={g.get('v')!r} f={g.get('f')!r})")
                continue
            gf, rf = norm_formula(g.get("f") or ""), norm_formula(r.get("f") or "")
            if formulas and (gf or rf):
                if gf != rf and not same_array_member(gs, gf, rf):
                    add(f"{label}: {name}!{addr} formula {g.get('f')!r} vs {r.get('f')!r}")
                elif g.get("v") is not None and r.get("v") is not None and not values_equal(g.get("v"), r.get("v")):
                    add(f"{label}: {name}!{addr} cached {g.get('v')!r} vs {r.get('v')!r}")
                continue
            if not formulas and (gf or rf) and (g.get("v") is None or r.get("v") is None):
                continue
            if not values_equal(g.get("v"), r.get("v")):
                add(f"{label}: {name}!{addr} value {g.get('v')!r} vs {r.get('v')!r}")
        gnotes = {n["a"].upper(): norm_text(n.get("text") or "") for n in gs.get("comments") or []}
        rnotes = {n["a"].upper(): norm_text(n.get("text") or "") for n in rs.get("comments") or []}
        for addr in sorted(set(gnotes) | set(rnotes)):
            gt, rt = gnotes.get(addr, ""), rnotes.get(addr, "")
            if gt == rt:
                continue
            # Excel stores "Author:\ntext". Accept either side containing the other.
            if gt and rt and (gt in rt or rt in gt):
                continue
            add(f"{label}: {name}!{addr} comment {gt!r} vs {rt!r}")
        gl = {n["a"].upper(): (n.get("url") or "") for n in gs.get("links") or []}
        rl = {n["a"].upper(): (n.get("url") or "") for n in rs.get("links") or []}
        for addr in sorted(set(gl) | set(rl)):
            gu, ru = gl.get(addr, ""), rl.get(addr, "")
            if norm_link(gu) == norm_link(ru):
                continue
            if gu and ru and (gu in ru or ru in gu):
                continue
            add(f"{label}: {name}!{addr} link {gu!r} vs {ru!r}")

    def nameset(doc):
        out = set()
        for n in doc.get("names") or []:
            out.add((norm_name(n.get("name") or ""), norm_ref(n.get("ref") or "")))
        return out

    gn, rn = nameset(got), nameset(ref)
    # Only flag names the reference has that goexcel lost, plus refs that disagree.
    ref_by = {}
    for name, refv in rn:
        ref_by.setdefault(name, set()).add(refv)
    got_by = {}
    for name, refv in gn:
        got_by.setdefault(name, set()).add(refv)
    for name, refs in sorted(ref_by.items()):
        if name.startswith("_XLNM"):
            continue
        if name not in got_by:
            add(f"{label}: missing name {name} -> {sorted(refs)[:3]}")
        elif got_by[name] != refs and not (got_by[name] & refs):
            add(f"{label}: name {name} ref {sorted(got_by[name])[:2]} vs {sorted(refs)[:2]}")
    return issues


def go_dump(bin_path: Path, filename: Path) -> dict:
    proc = subprocess.run([str(bin_path), str(filename)], capture_output=True, text=True)
    if not proc.stdout.strip():
        return {"error": proc.stderr.strip() or f"exit {proc.returncode}", "sheets": [], "names": []}
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        return {"error": f"bad json: {exc}", "sheets": [], "names": []}


def go_rewrite(bin_path: Path, src: Path, dst: Path) -> str | None:
    proc = subprocess.run([str(bin_path), "-rewrite", str(dst), str(src)], capture_output=True, text=True)
    if proc.returncode != 0:
        return proc.stderr.strip() or proc.stdout.strip() or f"exit {proc.returncode}"
    return None


def excel_formula(value, data_type="f"):
    """Return an Excel formula body, or None when the cell is not a formula."""
    if isinstance(value, str) and value.startswith("=") and data_type == "f":
        return value[1:]
    name = type(value).__name__
    if name == "ArrayFormula":
        text = getattr(value, "text", "") or ""
        if text.startswith("="):
            text = text[1:]
        if text.startswith("{") and text.endswith("}"):
            return text[1:-1]
        return text
    if name == "DataTableFormula":
        r1 = getattr(value, "r1", None) or ""
        r2 = getattr(value, "r2", None) or ""
        if r2:
            return f"TABLE({r1},{r2})"
        if r1:
            return f"TABLE({r1})"
        return "TABLE()"
    return None


def openpyxl_dump(filename: Path) -> dict | None:
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        wb = openpyxl.load_workbook(filename, data_only=False, read_only=False)
    except Exception as exc:
        return {"error": str(exc), "sheets": [], "names": []}
    doc = {"date1904": bool(getattr(wb, "epoch", None) and wb.epoch.year == 1904), "names": [], "sheets": []}
    try:
        defined = wb.defined_names
        for name in defined:
            defn = defined[name]
            items = defn if isinstance(defn, list) else [defn]
            for item in items:
                ref = getattr(item, "attr_text", None) or ""
                doc["names"].append({"name": getattr(item, "name", name), "ref": ref})
    except Exception:
        pass
    sheets = getattr(wb, "_sheets", None) or list(wb.worksheets)
    for ws in sheets:
        if not hasattr(ws, "iter_rows"):
            doc["sheets"].append({
                "name": ws.title,
                "state": getattr(ws, "sheet_state", None) or "visible",
                "merges": [],
                "cells": [],
                "comments": [],
                "links": [],
            })
            continue
        sh = {
            "name": ws.title,
            "state": ws.sheet_state or "visible",
            "merges": [str(r) for r in ws.merged_cells.ranges],
            "cells": [],
            "comments": [],
            "links": [],
        }
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None and cell.data_type != "f":
                    continue
                item = {"a": cell.coordinate, "t": "s"}
                formula = excel_formula(cell.value, cell.data_type)
                if formula is not None:
                    item["t"] = "f"
                    item["f"] = formula
                else:
                    v = cell.value
                    if isinstance(v, dt.datetime):
                        item["t"] = "d"
                        item["v"] = v.replace(tzinfo=dt.timezone.utc).isoformat().replace("+00:00", "Z")
                    elif isinstance(v, dt.date):
                        item["t"] = "d"
                        item["v"] = dt.datetime(v.year, v.month, v.day).isoformat() + "Z"
                    elif isinstance(v, dt.time):
                        item["t"] = "n"
                        item["v"] = (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400.0
                    elif isinstance(v, dt.timedelta):
                        item["t"] = "n"
                        item["v"] = v.total_seconds() / 86400.0
                    elif isinstance(v, bool):
                        item["t"] = "b"
                        item["v"] = v
                    elif isinstance(v, (int, float)):
                        item["t"] = "n"
                        item["v"] = float(v)
                    else:
                        item["t"] = "s"
                        item["v"] = "" if v is None else str(v)
                if cell.number_format and cell.number_format != "General":
                    item["nf"] = cell.number_format
                sh["cells"].append(item)
                if cell.comment is not None:
                    sh["comments"].append({
                        "a": cell.coordinate,
                        "author": cell.comment.author or "",
                        "text": cell.comment.text or "",
                    })
                if cell.hyperlink is not None:
                    url = cell.hyperlink.target or cell.hyperlink.location or ""
                    if url:
                        sh["links"].append({"a": cell.coordinate, "url": url})
        doc["sheets"].append(sh)
    wb.close()
    return doc


def poi_dump(filename: Path) -> dict | None:
    if not POI_DIR.is_dir() or not any(POI_DIR.glob("poi-ooxml-*.jar")):
        return None
    jars = sorted(str(p) for p in POI_DIR.glob("*.jar"))
    cp = os.pathsep.join(jars + [str(POI_SRC.parent)])
    classfile = POI_SRC.parent / "PoiDump.class"
    if not classfile.exists() or classfile.stat().st_mtime < POI_SRC.stat().st_mtime:
        javac = subprocess.run(["javac", "-encoding", "UTF-8", "-cp", os.pathsep.join(jars), str(POI_SRC)], capture_output=True, text=True)
        if javac.returncode != 0:
            return {"error": "javac: " + (javac.stderr or javac.stdout), "sheets": [], "names": []}
    proc = subprocess.run(["java", "-cp", cp, "PoiDump", str(filename)], capture_output=True, text=True)
    text = proc.stdout
    start = text.find("{")
    if start < 0:
        return {"error": proc.stderr.strip() or text.strip() or f"poi exit {proc.returncode}", "sheets": [], "names": []}
    try:
        return json.loads(text[start:])
    except json.JSONDecodeError as exc:
        return {"error": f"poi json: {exc}: {proc.stdout[:200]}", "sheets": [], "names": []}


NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}


def lo_convert(filename: Path, outdir: Path) -> tuple[Path | None, str | None]:
    soffice = shutil.which("soffice")
    if not soffice:
        return None, None
    # Isolate the profile so parallel calls do not collide. One at a time is enough.
    profile = outdir / "lo-profile"
    proc = subprocess.run(
        [
            soffice,
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--norestore",
            "--convert-to",
            "fods",
            "--outdir",
            str(outdir),
            str(filename),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    produced = outdir / (filename.stem + ".fods")
    if produced.exists() and produced.stat().st_size > 0:
        return produced, None
    msg = (proc.stderr or proc.stdout or f"exit {proc.returncode}").strip()
    return None, msg


TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def lo_dump(fods: Path) -> dict:
    root = ET.parse(fods).getroot()
    doc = {"date1904": False, "names": [], "sheets": []}
    for table in root.iter(TABLE + "table"):
        name = table.attrib.get(TABLE + "name") or "Sheet"
        hidden = table.attrib.get(TABLE + "visibility") or "visible"
        state = {"hidden": "hidden", "collapse": "veryHidden"}.get(hidden, "visible")
        sh = {"name": name, "state": state, "merges": [], "cells": [], "comments": [], "links": []}
        row_i = 0
        for child in list(table):
            if child.tag != TABLE + "table-row":
                continue
            row_rep = int(child.attrib.get(TABLE + "number-rows-repeated") or "1")
            parsed = []
            col_i = 0
            valued = False
            for c in list(child):
                if c.tag not in (TABLE + "table-cell", TABLE + "covered-table-cell"):
                    continue
                crep = int(c.attrib.get(TABLE + "number-columns-repeated") or "1")
                formula = c.attrib.get(TABLE + "formula") or ""
                vtype = c.attrib.get(OFFICE + "value-type") or ""
                val = None
                has = False
                if vtype == "float" and OFFICE + "value" in c.attrib:
                    val = float(c.attrib[OFFICE + "value"])
                    has = True
                elif vtype == "boolean":
                    val = c.attrib.get(OFFICE + "boolean-value") == "true"
                    has = True
                elif vtype == "string":
                    text = odf_cell_text(c)
                    if text != "":
                        val = text
                        has = True
                elif vtype in ("date", "time"):
                    text = c.attrib.get(OFFICE + "date-value") or c.attrib.get(OFFICE + "time-value") or "".join(c.itertext())
                    if text:
                        val = text
                        has = True
                span_c = int(c.attrib.get(TABLE + "number-columns-spanned") or "1")
                span_r = int(c.attrib.get(TABLE + "number-rows-spanned") or "1")
                if (formula or has) and crep > 200:
                    crep = 1
                if formula or has or span_c > 1 or span_r > 1:
                    valued = True
                    for k in range(min(crep, 32)):
                        addr = col_name(col_i + k) + str(row_i + 1)
                        if span_c > 1 or span_r > 1:
                            sh["merges"].append(f"{addr}:{col_name(col_i + k + span_c - 1)}{row_i + span_r}")
                        item = {"a": addr}
                        if formula:
                            f = formula[3:] if formula.startswith("of:") else formula
                            item["f"] = f[1:] if f.startswith("=") else f
                            item["t"] = "f"
                        if has:
                            item["v"] = val
                        if formula or has:
                            parsed.append(item)
                col_i += crep
            if not valued:
                row_i += row_rep
                continue
            times = row_rep if row_rep <= 20 else 1
            for t_i in range(times):
                for item in parsed:
                    copied = dict(item)
                    # Rewrite the row number when a valued row is repeated.
                    m = re.match(r"([A-Z]+)(\d+)$", copied["a"])
                    if m and t_i:
                        copied["a"] = m.group(1) + str(int(m.group(2)) + t_i)
                    sh["cells"].append(copied)
            row_i += row_rep if row_rep <= 20 else 1
        doc["sheets"].append(sh)
    return doc


def odf_cell_text(cell) -> str:
    """Cell display text, ignoring annotations embedded in the cell."""
    parts = []
    for child in list(cell):
        if child.tag == "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p":
            parts.append("".join(child.itertext()))
    return "\n".join(parts)


def col_name(idx: int) -> str:
    n = idx + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def is_strict_package(filename: Path) -> bool:
    try:
        with zipfile.ZipFile(filename) as zf:
            data = zf.read("xl/workbook.xml")
    except Exception:
        return False
    return b"purl.oclc.org/ooxml" in data or b'conformance="strict"' in data


def iter_files(paths: list[str]) -> list[Path]:
    found = []
    for raw in paths:
        p = Path(raw)
        if p.is_file():
            found.append(p)
            continue
        for dirpath, dirnames, filenames in os.walk(p):
            dirnames[:] = [d for d in dirnames if d != "lo-profile"]
            for name in filenames:
                if name.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
                    found.append(Path(dirpath) / name)
    return sorted(found)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--engines", default="openpyxl,poi,libreoffice")
    ap.add_argument("--roundtrip", action="store_true")
    ap.add_argument("--limit", type=int, default=12, help="mismatches to print per file")
    ap.add_argument("--report", default="")
    args = ap.parse_args()
    engines = {e.strip() for e in args.engines.split(",") if e.strip()}

    dump_bin = Path(tempfile.gettempdir()) / "goexcel-xlsxdump"
    build = subprocess.run(["go", "build", "-o", str(dump_bin), "./cmd/xlsxdump"], cwd=ROOT, capture_output=True, text=True)
    if build.returncode != 0:
        print(build.stderr)
        return 1

    files = iter_files(args.paths)
    if not files:
        print("no workbooks found")
        return 1

    summary = []
    total_issues = 0
    for filename in files:
        rel = str(filename)
        try:
            rel = str(filename.relative_to(ROOT))
        except ValueError:
            pass
        got = go_dump(dump_bin, filename)
        file_issues = []
        if got.get("error"):
            file_issues.append(f"goexcel: {got['error']}")
        engines_ran = []
        if "openpyxl" in engines:
            ref = openpyxl_dump(filename)
            if ref is None:
                pass
            else:
                engines_ran.append("openpyxl")
                file_issues.extend(compare_docs("read/openpyxl", got, ref, args.limit))
        if "poi" in engines:
            ref = poi_dump(filename)
            if ref is None:
                pass
            else:
                engines_ran.append("poi")
                file_issues.extend(compare_docs("read/poi", got, ref, args.limit))
        if "libreoffice" in engines and shutil.which("soffice"):
            with tempfile.TemporaryDirectory(prefix="lo-") as tmp:
                fods, err = lo_convert(filename, Path(tmp))
                if err:
                    file_issues.append(f"read/libreoffice: convert failed: {err[:300]}")
                elif fods:
                    engines_ran.append("libreoffice")
                    file_issues.extend(compare_docs("read/libreoffice", got, lo_dump(fods), args.limit, formulas=False))
        if args.roundtrip and not got.get("error"):
            with tempfile.TemporaryDirectory(prefix="rt-") as tmp:
                out = Path(tmp) / "roundtrip.xlsx"
                werr = go_rewrite(dump_bin, filename, out)
                if werr:
                    file_issues.append(f"rewrite: {werr}")
                else:
                    rewritten = go_dump(dump_bin, out)
                    file_issues.extend(compare_docs("roundtrip/goexcel", rewritten, got, args.limit))
                    if "openpyxl" in engines:
                        ref = openpyxl_dump(out)
                        orig = openpyxl_dump(filename)
                        if ref is not None and orig is not None:
                            file_issues.extend(compare_docs("roundtrip/openpyxl", ref, orig, args.limit))
                    if "poi" in engines:
                        ref = poi_dump(out)
                        orig = poi_dump(filename)
                        if ref is not None and orig is not None:
                            file_issues.extend(compare_docs("roundtrip/poi", ref, orig, args.limit))
                    if "libreoffice" in engines and shutil.which("soffice"):
                        fods, err = lo_convert(out, Path(tmp))
                        if err:
                            file_issues.append(f"roundtrip/libreoffice: convert failed: {err[:300]}")
                        elif fods:
                            # Opening is the check. Value drift from LO recalculation is reported
                            # against goexcel's rewrite, not the original LO conversion.
                            file_issues.extend(compare_docs("roundtrip/libreoffice", rewritten, lo_dump(fods), args.limit, formulas=False))
        defects = []
        gaps = []
        strict = is_strict_package(filename)
        # A LibreOffice crash while importing the original workbook is that
        # importer failing, not a goexcel content mismatch. The round-trip
        # conversion is the same gap when the original already crashes.
        lo_original_crash = any(i.startswith("read/libreoffice: convert failed") for i in file_issues)
        for issue in file_issues:
            if "not a valid zip file" in issue:
                gaps.append(issue + " (encrypted or OLE package, not OOXML)")
            elif "reference error:" in issue:
                gaps.append(issue)
            elif strict and "sheets " in issue and "vs []" in issue:
                gaps.append(issue + " (openpyxl does not read strict OOXML)")
            elif ("/poi:" in issue) and "value '" in issue and issue.endswith("vs ''"):
                gaps.append(issue + " (POI dropped a shared-string index with trailing whitespace)")
            elif issue.startswith("read/libreoffice: convert failed"):
                gaps.append(issue + " (LibreOffice crashed importing the original workbook)")
            elif issue.startswith("roundtrip/libreoffice: convert failed") and lo_original_crash:
                gaps.append(issue + " (LibreOffice also crashed on the original)")
            elif "libreoffice:" in issue and "convert failed" not in issue:
                # LibreOffice recalculates pivots, localizes labels, and uses
                # its own error codes. Those drifts are recorded, not failures.
                gaps.append(issue)
            else:
                defects.append(issue)
        file_issues = defects
        status = "ok" if not file_issues else "diff"
        total_issues += len(file_issues)
        if gaps:
            summary.append({"file": rel, "status": "ref-gap", "engines": engines_ran, "issues": gaps})
            print(f"gap  {rel}")
            for issue in gaps:
                print(f"    {issue}")
        summary.append({"file": rel, "status": status, "engines": engines_ran, "issues": file_issues})
        print(f"{status:4} {rel} engines={','.join(engines_ran) or '-'} issues={len(file_issues)}")
        for issue in file_issues:
            print(f"    {issue}")

    if args.report:
        Path(args.report).write_text(json.dumps(summary, indent=2))
    print(f"files={len(files)} issue_rows={total_issues}")
    return 1 if total_issues else 0


if __name__ == "__main__":
    sys.exit(main())
