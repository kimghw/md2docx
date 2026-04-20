# reference_template.docx 스타일 추출 결과

대상: [references/templates/reference_template.docx](../reference_template.docx)

## 산출물

| 파일 | 내용 |
|------|------|
| [reference_template.used_mapping.tsv](reference_template.used_mapping.tsv) | `document.xml`에서 실제 참조 중인 style ID와 그 w:name — 8건 |
| [reference_template.used_styles.xml](reference_template.used_styles.xml) | 참조된 styleId의 실제 `<w:style>` XML 블록 (테두리·폰트·basedOn 포함) |

## 실제 참조 중인 스타일

| reference | styleId | type | w:name | count |
|-----------|---------|------|--------|-------|
| pStyle | 1 | paragraph | heading 1 | 1 |
| pStyle | 2 | paragraph | heading 2 | 1 |
| pStyle | 3 | paragraph | heading 3 | 1 |
| pStyle | 4 | paragraph | heading 4 | 1 |
| pStyle | 5 | paragraph | heading 5 | 1 |
| pStyle | 6 | paragraph | heading 6 | 1 |
| tblStyle | 10 | table | Plain Table 1 | 1 |
| tblStyle | aa | table | Table Grid | 1 |

## 재추출 명령

```bash
# 1) document.xml에서 참조 중인 style ID
unzip -p references/templates/reference_template.docx word/document.xml \
  | grep -oE 'w:(pStyle|rStyle|tblStyle) w:val="[^"]+"' \
  | sort | uniq -c

# 2) 해당 ID의 type·w:name 조회 (styles.xml)
unzip -p references/templates/reference_template.docx word/styles.xml \
  | grep -oE '<w:style [^>]*w:styleId="<ID>"[^>]*>[^<]*<w:name w:val="[^"]+"'
```
