#!/usr/bin/env python3
"""extract-docx-styles: 단일 Word 템플릿에서 두 산출을 만든다.

  1) reference.docx — Pandoc --reference-doc= 으로 그대로 넘길 사본
  2) 표 스타일 번들 — 다른 docx 에 이식 가능한 XML 스니펫 묶음

Usage:
  python extract.py --doc template.docx --out-dir extracted_output/_styles/
  python extract.py --doc template.dotx --out-dir extracted_output/_styles/
  python extract.py --doc template.docx --out-dir out/ --sample-index 1
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import zipfile
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, OSError):
    pass


# ---------- Pandoc expectation ----------
REQUIRED_NAMES = ["Normal", "heading 1"]
RECOMMENDED_NAMES = [
    "heading 2", "heading 3", "heading 4", "heading 5", "heading 6",
    "Title", "Subtitle", "Quote", "List Paragraph",
]


# ---------- zip helpers ----------
def read_zip_text(zpath: str, name: str) -> Optional[str]:
    """Return part contents or None when the part is missing."""
    with zipfile.ZipFile(zpath) as z:
        if name not in z.namelist():
            return None
        return z.read(name).decode("utf-8")


# ---------- source normalization ----------
def normalize_source(src: str, out_dir: str) -> str:
    """Copy src to <out_dir>/reference.docx. .dotx is copied as .docx.

    Returns the absolute path of the produced reference.docx.
    """
    ext = os.path.splitext(src)[1].lower()
    if ext not in (".docx", ".dotx"):
        raise SystemExit(f"Unsupported input extension: {ext} (expect .docx or .dotx)")
    dst = os.path.join(out_dir, "reference.docx")
    shutil.copy2(src, dst)
    if ext == ".dotx":
        print(f"  .dotx detected -> copied as {dst}")
    return dst


# ---------- style parsing (regex only, stdlib) ----------
_STYLE_BLOCK_RE = re.compile(
    r'<w:style\b[^>]*>.*?</w:style>', re.S
)


def iter_style_blocks(styles_xml: str):
    return _STYLE_BLOCK_RE.finditer(styles_xml)


def style_attr(block: str, attr: str) -> Optional[str]:
    m = re.search(r'\bw:' + re.escape(attr) + r'="([^"]*)"', block)
    return m.group(1) if m else None


def style_name(block: str) -> Optional[str]:
    m = re.search(r'<w:name\s+w:val="([^"]*)"', block)
    return m.group(1) if m else None


def style_basedon(block: str) -> Optional[str]:
    m = re.search(r'<w:basedOn\s+w:val="([^"]*)"', block)
    return m.group(1) if m else None


def find_style_block(styles_xml: str, style_id: str) -> Optional[str]:
    pat = re.compile(
        r'<w:style\b[^>]*w:styleId="' + re.escape(style_id) + r'"[^>]*>.*?</w:style>',
        re.S,
    )
    m = pat.search(styles_xml)
    return m.group(0) if m else None


def collect_style_names(styles_xml: str) -> set[str]:
    return {m.group(1) for m in re.finditer(r'<w:name\s+w:val="([^"]*)"', styles_xml)}


def walk_basedon_chain(styles_xml: str, start_id: str) -> list[str]:
    """Return [start_id, parent, grandparent, ...] up to a cycle / missing."""
    seen = set()
    chain = []
    cur = start_id
    while cur and cur not in seen:
        seen.add(cur)
        block = find_style_block(styles_xml, cur)
        if not block:
            break
        chain.append(cur)
        cur = style_basedon(block)
    return chain


# ---------- table extraction ----------
_TBL_RE = re.compile(r'<w:tbl\b[^>]*>.*?</w:tbl>', re.S)


def iter_tables(doc_xml: str):
    return list(_TBL_RE.finditer(doc_xml))


def table_style_id(tbl_xml: str) -> Optional[str]:
    m = re.search(r'<w:tblStyle\s+w:val="([^"]*)"', tbl_xml)
    return m.group(1) if m else None


# ---------- the two jobs ----------
def job_prepare_reference(ref_docx: str) -> dict:
    """Validate that reference.docx has the w:name set Pandoc expects."""
    styles = read_zip_text(ref_docx, "word/styles.xml") or ""
    names = collect_style_names(styles)
    missing_required = [n for n in REQUIRED_NAMES if n not in names]
    missing_recommended = [n for n in RECOMMENDED_NAMES if n not in names]
    return {
        "path": ref_docx,
        "style_name_count": len(names),
        "missing_required": missing_required,
        "missing_recommended": missing_recommended,
    }


def job_extract_table_style(ref_docx: str, out_dir: str, sample_index: int) -> dict:
    """Pull Table style definition + sample <w:tbl> out for re-use."""
    doc = read_zip_text(ref_docx, "word/document.xml") or ""
    styles = read_zip_text(ref_docx, "word/styles.xml") or ""

    tables = iter_tables(doc)
    result = {
        "table_count": len(tables),
        "sample_index": sample_index,
        "sample_style_id": None,
        "has_style_Table": False,
        "chain": [],
        "wrote_sample_table": False,
        "wrote_table_style": False,
        "wrote_bundle": False,
    }

    # Is there a styleId="Table" definition?
    table_block = find_style_block(styles, "Table")
    result["has_style_Table"] = table_block is not None

    if not tables:
        print("  [warn] no <w:tbl> in document.xml — table extraction skipped.")
        return result

    if sample_index < 0 or sample_index >= len(tables):
        raise SystemExit(
            f"--sample-index {sample_index} out of range (found {len(tables)} tables)"
        )

    sample = tables[sample_index].group(0)
    sid = table_style_id(sample)
    result["sample_style_id"] = sid

    # sample_table.xml
    sample_path = os.path.join(out_dir, "sample_table.xml")
    with open(sample_path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            '<!-- sample <w:tbl> extracted from ' + os.path.basename(ref_docx)
            + ' (index ' + str(sample_index) + ') -->\n'
        )
        f.write(sample)
    result["wrote_sample_table"] = True

    # Resolve style chain (prefer explicit "Table" then sample's own tblStyle).
    anchor = None
    if table_block is not None:
        anchor = "Table"
    elif sid:
        anchor = sid
    if anchor:
        result["chain"] = walk_basedon_chain(styles, anchor)

    # table_style.xml  — the anchor style block only
    if anchor:
        anchor_block = find_style_block(styles, anchor)
        if anchor_block:
            style_path = os.path.join(out_dir, "table_style.xml")
            with open(style_path, "w", encoding="utf-8") as f:
                f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
                f.write('<!-- w:style block for ' + anchor + ' -->\n')
                f.write(anchor_block)
            result["wrote_table_style"] = True

    # bundle/styles_excerpt.xml  — every style in the basedOn chain
    if result["chain"]:
        bundle_dir = os.path.join(out_dir, "table_style_bundle")
        os.makedirs(bundle_dir, exist_ok=True)

        pieces = []
        for sid_in_chain in result["chain"]:
            b = find_style_block(styles, sid_in_chain)
            if b:
                pieces.append(b)
        excerpt_path = os.path.join(bundle_dir, "styles_excerpt.xml")
        with open(excerpt_path, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write(
                '<styles-excerpt anchor="' + anchor + '" '
                'chain="' + ' -> '.join(result["chain"]) + '">\n'
            )
            for p in pieces:
                f.write(p + "\n")
            f.write('</styles-excerpt>\n')

        readme_path = os.path.join(bundle_dir, "README.txt")
        with open(readme_path, "w", encoding="utf-8") as f:
            f.write(
                "table_style_bundle — how to transplant\n"
                "======================================\n\n"
                "anchor styleId : " + anchor + "\n"
                "basedOn chain  : " + " -> ".join(result["chain"]) + "\n\n"
                "Option A (recommended): run md2docx/clone_table_props.py with\n"
                "  --template <this reference.docx> --target <other.docx>\n\n"
                "Option B (manual): open the target .docx's word/styles.xml and\n"
                "paste every <w:style> block from styles_excerpt.xml just before\n"
                "</w:styles>. Keep the styleId values exactly as-is so basedOn\n"
                "pointers still resolve. For the cell layout, copy <w:tblPr> and\n"
                "<w:tblGrid> from sample_table.xml into the target table.\n"
            )
        result["wrote_bundle"] = True

    return result


# ---------- reporting ----------
def write_report(out_dir: str, ref_res: dict, tbl_res: dict, src: str) -> str:
    path = os.path.join(out_dir, "report.tsv")
    lines = [
        "section\tkey\tvalue",
        f"input\tpath\t{src}",
        f"input\text\t{os.path.splitext(src)[1].lower()}",
        f"reference\tpath\t{ref_res['path']}",
        f"reference\tstyle_name_count\t{ref_res['style_name_count']}",
        f"reference\tmissing_required\t{','.join(ref_res['missing_required']) or '-'}",
        f"reference\tmissing_recommended\t{','.join(ref_res['missing_recommended']) or '-'}",
        f"table\ttable_count\t{tbl_res['table_count']}",
        f"table\tsample_index\t{tbl_res['sample_index']}",
        f"table\tsample_style_id\t{tbl_res['sample_style_id'] or '-'}",
        f"table\thas_style_Table\t{str(tbl_res['has_style_Table']).lower()}",
        f"table\tbasedOn_chain\t{' -> '.join(tbl_res['chain']) or '-'}",
        f"table\twrote_sample_table\t{str(tbl_res['wrote_sample_table']).lower()}",
        f"table\twrote_table_style\t{str(tbl_res['wrote_table_style']).lower()}",
        f"table\twrote_bundle\t{str(tbl_res['wrote_bundle']).lower()}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def print_summary(src: str, out_dir: str, ref_res: dict, tbl_res: dict) -> None:
    print()
    print(f"[input]  {src}")
    print(f"[out]    {out_dir}")
    print()
    print("[1/2] reference.docx — Pandoc --reference-doc readiness")
    print(f"      style w:name count : {ref_res['style_name_count']}")
    if ref_res["missing_required"]:
        print(f"      MISSING (required) : {', '.join(ref_res['missing_required'])}")
    else:
        print("      required w:name    : OK")
    if ref_res["missing_recommended"]:
        print(
            "      missing (recommended): "
            + ", ".join(ref_res["missing_recommended"])
        )
    print()
    print("[2/2] Table style extraction")
    print(f"      <w:tbl> in document : {tbl_res['table_count']}")
    if tbl_res["table_count"] == 0:
        print("      (no table — nothing to extract)")
        return
    print(f"      sample index        : {tbl_res['sample_index']}")
    print(f"      sample tblStyle     : {tbl_res['sample_style_id'] or '-'}")
    print(
        "      styleId='Table'     : "
        + ("present" if tbl_res["has_style_Table"] else "MISSING "
           "— Pandoc emits <w:tblStyle w:val=\"Table\"/> literally, "
           "so the target must define it")
    )
    if tbl_res["chain"]:
        print(f"      basedOn chain       : {' -> '.join(tbl_res['chain'])}")
    wrote = [
        ("sample_table.xml", tbl_res["wrote_sample_table"]),
        ("table_style.xml",  tbl_res["wrote_table_style"]),
        ("table_style_bundle/", tbl_res["wrote_bundle"]),
    ]
    for name, ok in wrote:
        print(f"      wrote {name:25s} {'yes' if ok else 'no'}")


# ---------- CLI ----------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--doc", required=True,
                    help="Input Word template (.docx or .dotx)")
    ap.add_argument("--out-dir", required=True,
                    help="Output directory; created if missing")
    ap.add_argument("--sample-index", type=int, default=0,
                    help="Which <w:tbl> to extract as sample_table.xml (default 0)")
    args = ap.parse_args()

    if not os.path.isfile(args.doc):
        raise SystemExit(f"input not found: {args.doc}")

    os.makedirs(args.out_dir, exist_ok=True)

    print("[0] Copy to reference.docx ...")
    ref_docx = normalize_source(args.doc, args.out_dir)

    ref_res = job_prepare_reference(ref_docx)
    tbl_res = job_extract_table_style(ref_docx, args.out_dir, args.sample_index)

    report_path = write_report(args.out_dir, ref_res, tbl_res, args.doc)
    print_summary(args.doc, args.out_dir, ref_res, tbl_res)
    print()
    print(f"[report] {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
