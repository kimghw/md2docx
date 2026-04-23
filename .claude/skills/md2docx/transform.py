#!/usr/bin/env python3
"""md2docx — Markdown → DOCX 변환 + 표 스타일 주입 (한 번에).

파이프라인 (3 단계):

  [1] Pandoc 변환
      pandoc <md> --reference-doc=<extract_out>/base_templates/reference.docx -o <out>
      - 제목/본문/리스트/코드/링크 스타일은 reference.docx 를 상속
      - 표는 구조(행·열·내용·정렬) 만 생성되고, Pandoc 은 <w:tblStyle w:val="Table"/>
        리터럴만 박아둠 → 이대로는 서식 미적용

  [2] 표 스타일 주입 (핵심)
      <extract_out>/table_style.xml 의 <w:style> 블록들을 출력 docx 의
      word/styles.xml 에 병합하고, Pandoc 이 박아둔 <w:tblStyle w:val="Table"/>
      을 앵커 styleId 로 치환. <extract_out>/table_sources.xml 의 <w:tblLook>
      을 복사해 firstRow/firstCol 조건부 서식을 활성화.

  [3] (선택) --preview 렌더
      최종 docx → PDF → 페이지별 PNG (docx2pdf + PyMuPDF)

입력:
  --md            .md 파일
  --extract-out   md2docx-extract-style 이 만든 디렉터리
                  (base_templates/reference.docx, table_style.xml, table_sources.xml 필수)
  --out           출력 .docx 경로

사용 예:
  python .claude/skills/md2docx/transform.py \
    --md          test.md \
    --extract-out extracted_output/reference_reg/ \
    --out         extracted_output/reference_reg/final.docx \
    --preview
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import zipfile
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, OSError):
    pass


# ---------- pandoc ----------
def run_pandoc(md: str, ref_docx: str, out_docx: str) -> None:
    r = subprocess.run(
        ["pandoc", md, f"--reference-doc={ref_docx}", "-o", out_docx],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise SystemExit("pandoc failed")


# ---------- XML helpers ----------
_STYLE_BLOCK = re.compile(r'<w:style\b[^>]*>.*?</w:style>', re.S)
_STYLEID_ATTR = re.compile(r'\bw:styleId="([^"]*)"')


def style_blocks_with_ids(xml_text: str) -> list[tuple[str, str]]:
    out = []
    for m in _STYLE_BLOCK.finditer(xml_text):
        block = m.group(0)
        sid_m = _STYLEID_ATTR.search(block)
        out.append((sid_m.group(1) if sid_m else "", block))
    return out


def read_anchor_and_tbllook(sample_xml: str) -> tuple[Optional[str], Optional[str]]:
    a = re.search(r'<w:tblStyle\s+w:val="([^"]*)"', sample_xml)
    t = re.search(r'<w:tblLook\b[^/]*/>', sample_xml)
    return (a.group(1) if a else None,
            t.group(0) if t else None)


def extract_tblpr_additions(sample_xml: str) -> dict:
    """Pull propagatable <w:tblPr> children from table_sources.xml.

    These bits live on each <w:tbl> (not in the style definition) and shape
    cell geometry — Pandoc doesn't emit them, so we have to copy from template.
    """
    out = {}
    tblpr_m = re.search(r'<w:tblPr\b[^>]*>(.*?)</w:tblPr>', sample_xml, re.S)
    if not tblpr_m:
        return out
    inner = tblpr_m.group(1)
    for tag in ("tblInd", "tblCellMar", "tblLayout"):
        m = re.search(
            rf'<w:{tag}\b[^>]*>.*?</w:{tag}>|<w:{tag}\b[^/]*/>',
            inner, re.S)
        if m:
            out[tag] = m.group(0)
    return out


def apply_tblpr_additions(tbl_body: str, additions: dict) -> str:
    """Ensure each <w:tblPr> in tbl_body carries the given child elements.

    If an element already exists, it's replaced; otherwise it's appended.
    """
    if not additions:
        return tbl_body
    tblpr_m = re.search(r'<w:tblPr\b[^>]*>(.*?)</w:tblPr>', tbl_body, re.S)
    if not tblpr_m:
        return tbl_body
    inner = tblpr_m.group(1)
    for tag, xml in additions.items():
        pat = rf'<w:{tag}\b[^>]*>.*?</w:{tag}>|<w:{tag}\b[^/]*/>'
        existing = re.search(pat, inner, re.S)
        if existing:
            inner = inner[:existing.start()] + xml + inner[existing.end():]
        else:
            inner = inner + xml
    return tbl_body[:tblpr_m.start(1)] + inner + tbl_body[tblpr_m.end(1):]


def extract_cell_formatting(sample_xml: str) -> dict:
    """Harvest representative cell-level formatting from table_sources.xml.

    Scans every <w:tc> in the sample and pulls the first non-trivial
    <w:pPr> / <w:rPr> found. The following children are **stripped** before
    deciding "non-trivial" because they are POSITION-DEPENDENT and must not
    be broadcast to arbitrary cells:
      - <w:pStyle>   — may point to a template-only style like "Compact"
      - <w:cnfStyle> — per-cell conditional formatting marker (firstRow/…) ;
                      Word auto-derives this from tblLook + actual position

    Returns a dict with 'pPr_inner' and/or 'rPr_inner' — may be absent when
    the sample cells had nothing reusable.
    """
    drop_ppr = re.compile(
        r'<w:pStyle\s+w:val="[^"]*"\s*/>|<w:cnfStyle\b[^/]*/>')
    out = {}
    tcs = re.findall(r'<w:tc\b[^>]*>.*?</w:tc>', sample_xml, re.S)
    for tc in tcs:
        if "pPr_inner" not in out:
            ppr = re.search(r'<w:pPr\b[^>]*>(.*?)</w:pPr>', tc, re.S)
            if ppr and ppr.group(1).strip():
                inner = drop_ppr.sub("", ppr.group(1)).strip()
                if inner:
                    out["pPr_inner"] = inner
        if "rPr_inner" not in out:
            rpr = re.search(r'<w:rPr\b[^>]*>(.*?)</w:rPr>', tc, re.S)
            if rpr and rpr.group(1).strip():
                out["rPr_inner"] = rpr.group(1).strip()
        if "pPr_inner" in out and "rPr_inner" in out:
            break
    return out


def overwrite_cell_formatting(tbl_body: str, cell_fmt: dict) -> tuple[str, int]:
    """Hard-overwrite every <w:tc>'s formatting with extracted sample.

    The user's intent: Pandoc's table formatting should be **erased entirely**
    and replaced with what the template's table_sources.xml provides.  So:

      - <w:tcPr>  → cleared to <w:tcPr/>            (table style's tcPr drives)
      - <w:pPr>   → replaced with sample pPr_inner  (or removed if none)
      - <w:rPr>   → replaced with sample rPr_inner  (or removed if none)
      - inline <w:b/>, <w:i/>, etc. from pandoc ARE stripped — any emphasis
        must come from the table style (e.g. firstRow), not inline.

    Returns (new_body, cells_overwritten).
    """
    cells_overwritten = 0
    sample_ppr = cell_fmt.get("pPr_inner") or ""
    sample_rpr = cell_fmt.get("rPr_inner") or ""
    new_ppr_xml = f"<w:pPr>{sample_ppr}</w:pPr>" if sample_ppr else ""
    new_rpr_xml = f"<w:rPr>{sample_rpr}</w:rPr>" if sample_rpr else ""

    def _rewrite_tc(m: re.Match) -> str:
        nonlocal cells_overwritten
        tc = m.group(0)
        # 1) nuke any <w:tcPr>…</w:tcPr> or self-closing — replace with empty
        tc = re.sub(r'<w:tcPr\b[^>]*>.*?</w:tcPr>', '<w:tcPr/>', tc, flags=re.S)
        tc = re.sub(r'<w:tcPr\b[^/]*/>', '<w:tcPr/>', tc)
        # If no tcPr at all, insert an empty one at the start of the cell
        if '<w:tcPr' not in tc:
            tc = re.sub(r'(<w:tc\b[^>]*>)', r'\1' + '<w:tcPr/>', tc, count=1)
        # 2) nuke every <w:pPr> and replace with sample pPr (or nothing)
        tc = re.sub(r'<w:pPr\b[^>]*>.*?</w:pPr>', '', tc, flags=re.S)
        tc = re.sub(r'<w:pPr\b[^/]*/>', '', tc)
        if new_ppr_xml:
            tc = re.sub(r'(<w:p\b[^>]*>)', r'\1' + new_ppr_xml, tc)
        # 3) nuke every <w:rPr> and replace with sample rPr (or nothing)
        tc = re.sub(r'<w:rPr\b[^>]*>.*?</w:rPr>', '', tc, flags=re.S)
        tc = re.sub(r'<w:rPr\b[^/]*/>', '', tc)
        if new_rpr_xml:
            tc = re.sub(r'(<w:r\b[^>]*>)', r'\1' + new_rpr_xml, tc)
        cells_overwritten += 1
        return tc

    new_body = re.sub(r'<w:tc\b[^>]*>.*?</w:tc>', _rewrite_tc, tbl_body,
                      flags=re.S)
    return new_body, cells_overwritten


def inject_styles(target_styles_xml: str,
                  donor_blocks: list[tuple[str, str]]) -> str:
    """Merge donor <w:style> blocks into target. Replace if styleId exists, else append."""
    out = target_styles_xml
    for sid, block in donor_blocks:
        if not sid:
            continue
        pat = r'<w:style\b[^>]*w:styleId="' + re.escape(sid) + r'"[^>]*>.*?</w:style>'
        existing = re.search(pat, out, re.S)
        if existing:
            out = out[:existing.start()] + block + out[existing.end():]
        else:
            out = out.replace("</w:styles>", block + "</w:styles>", 1)
    return out


def retype_pandoc_tables(doc_xml: str, anchor: str,
                         tbllook: Optional[str],
                         tblpr_extras: Optional[dict] = None,
                         cell_fmt: Optional[dict] = None
                         ) -> tuple[str, int, int]:
    """Rewrite Pandoc tables so the template's style wins completely.

    For every <w:tbl>:
      1) swap <w:tblStyle w:val="Table"/> → <w:tblStyle w:val="{anchor}"/>
      2) inject/refresh <w:tblLook> so firstRow/firstCol conditional formatting fires
      3) propagate tblPr children from table_sources.xml (tblInd, tblCellMar, tblLayout)
      4) **hard-overwrite** every <w:tc>:
         - strip <w:tcPr>/<w:pPr>/<w:rPr> (including pandoc's Compact pStyle)
         - replace <w:pPr> / <w:rPr> with sample's harvested contents if any
         - inline pandoc bold/italic is NOT preserved — emphasis must come
           from the table style (firstRow, etc.)

    Returns (new_xml, tables_touched, cells_overwritten).
    """
    touched = 0
    cells_overwritten = 0

    def _replace_tbl(m: re.Match) -> str:
        nonlocal touched, cells_overwritten
        open_tag, body, close_tag = m.group(1), m.group(2), m.group(3)
        # Swap tblStyle val — Pandoc emits '<w:tblStyle w:val="Table" />'
        # (with a space before />) so we tolerate optional whitespace.
        new_body, n = re.subn(
            r'<w:tblStyle\s+w:val="[^"]*"\s*/>',
            f'<w:tblStyle w:val="{anchor}"/>',
            body)
        # Tolerate the non-self-closing form too just in case.
        if n == 0:
            new_body, n = re.subn(
                r'<w:tblStyle\s+w:val="[^"]*"\s*>\s*</w:tblStyle>',
                f'<w:tblStyle w:val="{anchor}"/>',
                new_body)
        # If still no tblStyle, insert one inside <w:tblPr> (or create tblPr).
        if n == 0:
            tblpr_m = re.search(r'<w:tblPr\b[^>]*>(.*?)</w:tblPr>', new_body, re.S)
            if tblpr_m:
                inner = tblpr_m.group(1)
                if '<w:tblStyle' not in inner:
                    new_inner = f'<w:tblStyle w:val="{anchor}"/>' + inner
                    new_body = new_body[:tblpr_m.start(1)] + new_inner + new_body[tblpr_m.end(1):]
            else:
                new_body = f'<w:tblPr><w:tblStyle w:val="{anchor}"/></w:tblPr>' + new_body

        # Ensure tblLook is present
        if tbllook:
            tblpr_m = re.search(r'<w:tblPr\b[^>]*>(.*?)</w:tblPr>', new_body, re.S)
            if tblpr_m:
                inner = tblpr_m.group(1)
                if '<w:tblLook' in inner:
                    inner = re.sub(r'<w:tblLook\b[^/]*/>', tbllook, inner)
                else:
                    inner = inner + tbllook
                new_body = new_body[:tblpr_m.start(1)] + inner + new_body[tblpr_m.end(1):]
        # Propagate tblPr geometry bits (tblInd / tblCellMar / tblLayout)
        if tblpr_extras:
            new_body = apply_tblpr_additions(new_body, tblpr_extras)
        # Hard-overwrite every <w:tc>'s formatting with extracted sample.
        new_body, n_cells = overwrite_cell_formatting(new_body, cell_fmt or {})
        cells_overwritten += n_cells
        touched += 1
        return open_tag + new_body + close_tag

    new_xml = re.sub(r'(<w:tbl\b[^>]*>)(.*?)(</w:tbl>)', _replace_tbl,
                     doc_xml, flags=re.S)
    return new_xml, touched, cells_overwritten


# ---------- inject bundle into a docx ----------
def inject_bundle(docx_path: str, table_style_path: str,
                  table_sources_path: str) -> dict:
    excerpt = open(table_style_path, encoding="utf-8").read()
    sources = open(table_sources_path, encoding="utf-8").read()
    anchor, tbllook = read_anchor_and_tbllook(sources)
    if not anchor:
        raise SystemExit("no <w:tblStyle> in table_sources.xml — cannot determine anchor")

    with zipfile.ZipFile(docx_path) as z:
        members = {n: z.read(n) for n in z.namelist()}
    styles_xml = members["word/styles.xml"].decode("utf-8")
    doc_xml = members["word/document.xml"].decode("utf-8")

    donor = style_blocks_with_ids(excerpt)
    tblpr_extras = extract_tblpr_additions(sources)
    cell_fmt = extract_cell_formatting(sources)
    styles_xml = inject_styles(styles_xml, donor)
    doc_xml, touched, cells_overwritten = retype_pandoc_tables(
        doc_xml, anchor, tbllook, tblpr_extras, cell_fmt)

    members["word/styles.xml"] = styles_xml.encode("utf-8")
    members["word/document.xml"] = doc_xml.encode("utf-8")

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in members.items():
            z.writestr(name, data)

    return {
        "anchor": anchor,
        "donor_style_ids": [sid for sid, _ in donor if sid],
        "tbllook_copied": tbllook is not None,
        "tables_touched": touched,
        "tblpr_extras_copied": sorted(tblpr_extras.keys()),
        "cell_fmt_harvested": sorted(cell_fmt.keys()),
        "cells_overwritten": cells_overwritten,
    }


# ---------- preview render (optional) ----------
def render_preview(docx_path: str, out_dir: str) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    pdf = os.path.join(out_dir, f"{stem}.pdf")
    result = {"engine": None, "pdf": None, "pngs": [], "note": None}
    # try docx2pdf then soffice
    try:
        from docx2pdf import convert as _d2p
        _d2p(os.path.abspath(docx_path), os.path.abspath(pdf))
        if os.path.isfile(pdf):
            result["engine"] = "docx2pdf"
            result["pdf"] = pdf
    except Exception as e:
        result["note"] = f"docx2pdf failed: {e}"
    if not result["pdf"]:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", os.path.abspath(out_dir),
                 os.path.abspath(docx_path)],
                check=True, capture_output=True)
            generated = os.path.join(
                out_dir,
                os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
            if os.path.isfile(generated):
                shutil.move(generated, pdf)
                result["engine"] = "soffice"
                result["pdf"] = pdf

    if not result["pdf"]:
        return result

    try:
        import fitz  # PyMuPDF
    except ModuleNotFoundError:
        result["note"] = "PyMuPDF 미설치 (pip install pymupdf)"
        return result

    doc = fitz.open(pdf)
    try:
        zoom = 200 / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for i in range(doc.page_count):
            pix = doc.load_page(i).get_pixmap(matrix=matrix, alpha=False)
            p = os.path.join(out_dir, f"{stem}.page-{i + 1:02d}.png")
            pix.save(p)
            result["pngs"].append(p)
    finally:
        doc.close()
    return result


# ---------- CLI ----------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--md", required=True, help="Input .md file")
    ap.add_argument("--extract-out", required=True,
                    help="Directory produced by md2docx-extract-style")
    ap.add_argument("--out", required=True, help="Output .docx path")
    ap.add_argument("--preview", action="store_true",
                    help="Also render output to PDF + per-page PNG")
    args = ap.parse_args()

    ref_docx = os.path.join(args.extract_out, "base_templates", "reference.docx")
    table_style = os.path.join(args.extract_out, "table_style.xml")
    table_sources = os.path.join(args.extract_out, "table_sources.xml")

    for label, path in [("--md", args.md), ("reference.docx", ref_docx),
                        ("table_style.xml", table_style),
                        ("table_sources.xml", table_sources)]:
        if not os.path.isfile(path):
            raise SystemExit(f"missing {label}: {path}")

    out_dir_of_out = os.path.dirname(os.path.abspath(args.out)) or "."
    os.makedirs(out_dir_of_out, exist_ok=True)

    print(f"[1/3] pandoc  {args.md}  +  {ref_docx}")
    run_pandoc(args.md, ref_docx, args.out)
    print(f"      -> {args.out}")

    print(f"[2/3] inject table style bundle")
    meta = inject_bundle(args.out, table_style, table_sources)
    print(f"      anchor styleId      : {meta['anchor']}")
    print(f"      donor styleIds      : {', '.join(meta['donor_style_ids']) or '-'}")
    print(f"      tblLook copied      : {meta['tbllook_copied']}")
    print(f"      tblPr extras copied : {', '.join(meta['tblpr_extras_copied']) or '-'}")
    print(f"      cell fmt harvested  : {', '.join(meta['cell_fmt_harvested']) or '-'}")
    print(f"      <w:tbl> retyped     : {meta['tables_touched']}")
    print(f"      cells overwritten   : {meta['cells_overwritten']}")

    if args.preview:
        print(f"[3/3] render preview")
        preview_dir = os.path.join(out_dir_of_out, "preview")
        pr = render_preview(args.out, preview_dir)
        if pr["engine"]:
            print(f"      engine              : {pr['engine']}")
            print(f"      pdf                 : {pr['pdf']}")
            print(f"      pages               : {len(pr['pngs'])}")
        else:
            print(f"      engine              : FAILED — {pr['note']}")

    print(f"\n[done] {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
