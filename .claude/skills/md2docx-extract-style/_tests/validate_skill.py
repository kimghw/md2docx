#!/usr/bin/env python3
"""Validate the extract-docx-styles skill against fixtures.

Checks (per fixture, against expected contract values):
  C1  reference.docx is byte-identical to the source docx
  C2  missing_required reflects actual styles.xml state
  C3  has_style_Table matches expected
  C4  table_sources.xml equals the N-th <w:tbl> in document.xml
  C5  table_style.xml contains the anchor <w:style> block
  C6  basedOn chain starts at the anchor and resolves
  C7  Pandoc accepts reference.docx (when test.md is present)
  C8  --preview renders at least one PNG when available
  C9  inject extracted bundle into a blank docx and render — functional
      reuse check (structural + produces PNG for agent visual diff)

Run:
  python .claude/skills/extract-docx-styles/tests/validate_skill.py
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
REPO = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
EXTRACT_PY = os.path.join(HERE, "..", "extract.py")
FIXTURES = os.path.join(HERE, "fixtures")
# Outputs live OUTSIDE the skill's tests/ dir so that dir stays clean
# (scripts + fixtures only). Override with VALIDATE_OUT env var if needed.
OUTPUTS = os.environ.get(
    "VALIDATE_OUT",
    os.path.join(REPO, "extracted_output", "_validate_runs"),
)
TEST_MD = os.path.join(REPO, "test.md")


# ---- expected contract table ----
EXPECTED = {
    "fx_empty.docx": {
        "table_count": 0,
        "has_style_Table": False,
        "sample_style_id": None,   # no sample
        "missing_required": [],
        "c7_pandoc": True,
    },
    "fx_tablegrid.docx": {
        "table_count": 1,
        "has_style_Table": False,   # styleId is "TableGrid" not "Table"
        "sample_style_id_nonempty": True,
        "missing_required": [],
        "c7_pandoc": True,
    },
    "fx_styled_Table.docx": {
        "table_count": 1,
        "has_style_Table": True,
        "sample_style_id": "Table",
        "missing_required": [],
        "c7_pandoc": True,
    },
    "fx_multi.docx": {
        "table_count": 3,
        "has_style_Table": False,
        "sample_style_id_nonempty": True,
        "missing_required": [],
        "c7_pandoc": True,
        "multi_sample_indices": [0, 1, 2],
    },
}


# ---- helpers ----
def sha256_file(p: str) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def read_zip_text(zp: str, name: str) -> str:
    with zipfile.ZipFile(zp) as z:
        return z.read(name).decode("utf-8")


def parse_report(tsv: str) -> dict:
    data = {}
    with open(tsv, encoding="utf-8") as f:
        next(f, None)  # header
        for line in f:
            parts = line.rstrip("\n").split("\t", 2)
            if len(parts) == 3:
                data[f"{parts[0]}.{parts[1]}"] = parts[2]
    return data


def run_extract(fixture: str, parent_dir: str, preview: bool = False,
                sample_index: int = 0) -> tuple[str, dict]:
    """Run extract.py with parent_dir as --out-dir; actual output lands in
    parent_dir/<source_stem>/. Returns (actual_out_dir, parsed_report)."""
    source_stem = os.path.splitext(os.path.basename(fixture))[0]
    actual_out = os.path.join(parent_dir, source_stem)
    if os.path.isdir(actual_out):
        shutil.rmtree(actual_out)
    args = [sys.executable, EXTRACT_PY,
            "--doc", fixture,
            "--out-dir", parent_dir,
            "--sample-index", str(sample_index)]
    if preview:
        args.append("--preview")
    r = subprocess.run(args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise RuntimeError(f"extract.py failed: {fixture}")
    return actual_out, parse_report(os.path.join(actual_out, "report.tsv"))


def find_nth_tbl(doc_xml: str, n: int) -> str:
    ms = list(re.finditer(r'<w:tbl\b[^>]*>.*?</w:tbl>', doc_xml, re.S))
    return ms[n].group(0) if 0 <= n < len(ms) else ""


def find_style_block(styles_xml: str, style_id: str) -> str:
    m = re.search(
        r'<w:style\b[^>]*w:styleId="' + re.escape(style_id) + r'"[^>]*>.*?</w:style>',
        styles_xml, re.S)
    return m.group(0) if m else ""


# ---- assertions ----
class Case:
    def __init__(self, name: str):
        self.name = name
        self.ok = 0
        self.fail = 0
        self.msgs: list[str] = []

    def check(self, label: str, cond: bool, detail: str = ""):
        if cond:
            self.ok += 1
            self.msgs.append(f"  [PASS] {label}")
        else:
            self.fail += 1
            self.msgs.append(f"  [FAIL] {label}  {detail}")

    def print(self):
        header = f"=== {self.name}  (pass={self.ok}, fail={self.fail}) ==="
        print(header)
        for m in self.msgs:
            print(m)
        print()


# ---- per-fixture validation ----
def validate_fixture(fixture_name: str, spec: dict) -> Case:
    case = Case(fixture_name)
    fixture = os.path.join(FIXTURES, fixture_name)

    # Run extract with --preview so C8 works in one shot.
    # extract.py auto-creates <parent>/<source_stem>/, so we pass OUTPUTS
    # as parent and receive the actual output path back.
    out, report = run_extract(fixture, OUTPUTS, preview=True)

    # -------- C1: byte equality of reference.docx --------
    ref = os.path.join(out, "base_templates", "reference.docx")
    case.check("C1 reference.docx byte-identical to source",
               sha256_file(fixture) == sha256_file(ref),
               f"source={sha256_file(fixture)[:12]} ref={sha256_file(ref)[:12]}")

    # -------- C2: missing_required --------
    want_missing = ",".join(spec["missing_required"]) or "-"
    case.check(f"C2 missing_required == {want_missing!r}",
               report.get("reference.missing_required") == want_missing,
               f"got {report.get('reference.missing_required')!r}")

    # -------- C3: has_style_Table --------
    want_has = str(spec["has_style_Table"]).lower()
    case.check(f"C3 has_style_Table == {want_has}",
               report.get("table.has_style_Table") == want_has,
               f"got {report.get('table.has_style_Table')!r}")

    # -------- table_count --------
    case.check(f"table_count == {spec['table_count']}",
               report.get("table.table_count") == str(spec["table_count"]),
               f"got {report.get('table.table_count')!r}")

    if spec["table_count"] > 0:
        # -------- C4: table_sources.xml equals N-th <w:tbl> --------
        doc_xml = read_zip_text(fixture, "word/document.xml")
        sources_on_disk = open(os.path.join(out, "table_sources.xml"),
                               encoding="utf-8").read()
        # strip xml header + comment line
        sources_body = re.sub(r'^<\?xml[^>]*\?>\s*', "", sources_on_disk)
        sources_body = re.sub(r'^<!--.*?-->\s*', "", sources_body,
                              flags=re.DOTALL).strip()
        expected = find_nth_tbl(doc_xml, 0)
        case.check("C4 table_sources.xml matches document.xml[0]",
                   sources_body == expected,
                   f"len(got)={len(sources_body)} len(want)={len(expected)}")

        # -------- C5: table_style.xml contains anchor <w:style> block --------
        styles_xml = read_zip_text(fixture, "word/styles.xml")
        if "sample_style_id" in spec and spec["sample_style_id"] is not None:
            anchor = spec["sample_style_id"]
        else:
            anchor = report.get("table.sample_style_id", "")
        expected_block = find_style_block(styles_xml, anchor)
        style_on_disk = open(os.path.join(out, "table_style.xml"),
                             encoding="utf-8").read()
        case.check(f"C5 table_style.xml contains <w:style styleId={anchor!r}>",
                   anchor and anchor in style_on_disk and expected_block
                   and expected_block in style_on_disk,
                   "")

        # -------- C6: basedOn chain starts at anchor --------
        chain_str = report.get("table.basedOn_chain", "-")
        first_hop = chain_str.split(" -> ")[0] if chain_str != "-" else ""
        case.check("C6 basedOn chain starts at anchor",
                   first_hop == anchor,
                   f"chain={chain_str!r} anchor={anchor!r}")

    # -------- C7: Pandoc round-trip --------
    if spec.get("c7_pandoc") and os.path.isfile(TEST_MD):
        out_docx = os.path.join(out, "pandoc_roundtrip.docx")
        r = subprocess.run(
            ["pandoc", TEST_MD, f"--reference-doc={ref}", "-o", out_docx],
            capture_output=True, text=True)
        case.check("C7 pandoc converts test.md with reference.docx",
                   r.returncode == 0 and os.path.isfile(out_docx)
                   and os.path.getsize(out_docx) > 0,
                   r.stderr[-200:] if r.returncode else "")

    # -------- C8: preview PNG exists --------
    preview_dir = os.path.join(out, "preview")
    pngs = [p for p in os.listdir(preview_dir) if p.endswith(".png")] \
        if os.path.isdir(preview_dir) else []
    engine = report.get("preview.engine", "-")
    if engine == "-":
        case.msgs.append(f"  [SKIP] C8 preview (no renderer): "
                         f"{report.get('preview.note')}")
    else:
        case.check(f"C8 --preview produced PNG (engine={engine})",
                   len(pngs) >= 1,
                   f"pngs={pngs}")

    # -------- C9: inject bundle into blank docx + render --------
    if spec["table_count"] > 0:
        from inject_and_render import run_c9
        c9_dir = os.path.join(out, "c9_inject")
        try:
            r = run_c9(out, fixture, c9_dir)
            case.check("C9 inject bundle into blank docx opens cleanly",
                       r.get("structural_ok", False),
                       r.get("structural_note", ""))
            case.check("C9 injected docx rendered to PNG",
                       len(r.get("injected_pngs", [])) >= 1,
                       r.get("render_error_injected", ""))
            case.check("C9 baseline fixture rendered to PNG",
                       len(r.get("baseline_pngs", [])) >= 1,
                       r.get("render_error_baseline", ""))
            case.msgs.append(
                f"  [INFO] C9 PNGs for visual agent diff: "
                f"{os.path.relpath(r['injected_pngs'][0], HERE) if r.get('injected_pngs') else '?'}"
                f"  vs  "
                f"{os.path.relpath(r['baseline_pngs'][0], HERE) if r.get('baseline_pngs') else '?'}"
            )
        except Exception as e:
            case.check("C9 inject pipeline ran without exception", False, str(e))

    # -------- extra: multi-sample indices --------
    if "multi_sample_indices" in spec:
        doc_xml = read_zip_text(fixture, "word/document.xml")
        for idx in spec["multi_sample_indices"]:
            alt_parent = os.path.join(OUTPUTS, f"__sample{idx}")
            alt_out, _ = run_extract(fixture, alt_parent, preview=False,
                                     sample_index=idx)
            sources_body = open(os.path.join(alt_out, "table_sources.xml"),
                                encoding="utf-8").read()
            sources_body = re.sub(r'^<\?xml[^>]*\?>\s*', "", sources_body)
            sources_body = re.sub(r'^<!--.*?-->\s*', "", sources_body,
                                  flags=re.DOTALL).strip()
            expected = find_nth_tbl(doc_xml, idx)
            case.check(f"multi --sample-index {idx} matches document.xml[{idx}]",
                       sources_body == expected, "")
    return case


def main() -> int:
    if not os.path.isdir(FIXTURES) or not os.listdir(FIXTURES):
        sys.exit("fixtures missing — run build_fixtures.py first")
    os.makedirs(OUTPUTS, exist_ok=True)

    cases: list[Case] = []
    for name, spec in EXPECTED.items():
        cases.append(validate_fixture(name, spec))

    for c in cases:
        c.print()

    total_pass = sum(c.ok for c in cases)
    total_fail = sum(c.fail for c in cases)
    print(f"=== TOTAL  pass={total_pass}  fail={total_fail} ===")
    return 0 if total_fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
