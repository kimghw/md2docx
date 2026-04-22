table_style_bundle — how to transplant
======================================

anchor styleId : TableGrid
basedOn chain  : TableGrid -> TableNormal

Option A (recommended): run md2docx/clone_table_props.py with
  --template <this reference.docx> --target <other.docx>

Option B (manual): open the target .docx's word/styles.xml and
paste every <w:style> block from styles_excerpt.xml just before
</w:styles>. Keep the styleId values exactly as-is so basedOn
pointers still resolve. For the cell layout, copy <w:tblPr> and
<w:tblGrid> from sample_table.xml into the target table.
