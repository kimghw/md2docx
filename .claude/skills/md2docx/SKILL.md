---
name: md2docx
description: Markdown 한 파일을 Word docx 한 파일로 변환한다. 입력은 (1) .md 파일, (2) md2docx-extract-style 이 만든 디렉터리 (base_templates/reference.docx + table_style.xml + table_sources.xml). Pandoc 으로 구조를 변환한 뒤, Pandoc 이 손대지 않은 **표 스타일을 후처리로 주입**해 최종 docx 를 만든다. `--preview` 로 PDF/PNG 프리뷰까지.
---

# md2docx — md + 추출 번들 → 스타일 적용된 docx

`md2docx-extract-style` 가 이미 **한 개의 Word 템플릿을 둘로 쪼개 놓은** 상태 (reference.docx + 표 스타일 번들) 를 전제로 한다. 이 스킬은 그 둘을 다시 합쳐서 md → 최종 docx 를 완성한다.

## 파이프라인 (3 단계)

```
  [입력] .md  +  extract_out/{base_templates/reference.docx, table_style.xml, table_sources.xml}
                         │
         ┌───────────────┘
         ▼
  [1] pandoc <md> --reference-doc=reference.docx -o <out>
         - 단락/제목/리스트/코드/링크/인용: reference.docx 상속  ✅
         - 표: 구조만 생성, <w:tblStyle w:val="Table"/> 리터럴만 박힘  ⚠️
         ▼
  [2] 표 스타일 주입
         - table_style.xml 의 <w:style> 를 <out>/word/styles.xml 에 병합
         - document.xml 의 <w:tblStyle w:val="Table"/> → <w:tblStyle w:val="{anchor}"/>
         - table_sources.xml 의 <w:tblLook> 을 복사 (firstRow/firstCol 활성화)
         ▼
  [3] (선택) --preview
         - 최종 docx → PDF → 페이지별 PNG  (extract-style 의 preview/ 폴더에 final.* 로 누적)
```

## 사용

```bash
python .claude/skills/md2docx/transform.py \
  --md          test.md \
  --extract-out extracted_output/reference_reg/ \
  --out         extracted_output/reference_reg/final.docx
```

PowerShell:

```powershell
python .claude\skills\md2docx\transform.py `
  --md          test.md `
  --extract-out extracted_output\reference_reg\ `
  --out         extracted_output\reference_reg\final.docx
```

프리뷰까지:

```bash
python .claude/skills/md2docx/transform.py \
  --md          test.md \
  --extract-out extracted_output/reference_reg/ \
  --out         extracted_output/reference_reg/final.docx \
  --preview
```

## 플래그

| 플래그 | 필수 | 의미 |
|---|---|---|
| `--md` | ✅ | 입력 Markdown 파일 |
| `--extract-out` | ✅ | `md2docx-extract-style` 의 출력 디렉터리 (보통 `<out-dir>/<원본파일이름>/`) |
| `--out` | ✅ | 출력 `.docx` 경로 (없는 상위 디렉터리는 자동 생성) |
| `--preview` | - | 최종 docx 를 PDF + PNG 로 렌더해 `<out 부모>/preview/` 에 `final.*` 접두어로 저장 (extract-style 의 `source.*` 와 같은 폴더) |

## 선행 조건 — `--extract-out` 이 가져야 할 파일

| 파일 | 출처 | 용도 |
|---|---|---|
| `base_templates/reference.docx` | extract-style ① | Pandoc `--reference-doc=` 입력 |
| `table_style.xml` | extract-style ② | `<w:style>` 블록 (앵커 + basedOn 체인, `<styles-excerpt>` 래퍼) |
| `table_sources.xml` | extract-style ② | 앵커 styleId 와 `<w:tblLook>` 추출 |

이 세 파일 중 하나라도 없으면 시작 전에 에러로 중단한다.

## 산출물

```
<out 부모>/                        ← 보통 extract-style 산출 폴더 = extract_out
├── <out>                          ← 주입까지 끝난 최종 docx
└── preview/                       ← --preview 시 (extract-style 의 preview/ 와 같은 위치)
    ├── final.pdf
    ├── final.page-01.png          ← 200 dpi
    ├── final.page-02.png
    └── ...                        ← 같은 폴더에 source.* (extract-style) 도 함께 누적
```

## 원리 핵심 3가지

1. **Pandoc 은 표의 구조만 만든다.** `<w:tbl>` 은 그리되 `<w:tblStyle w:val="Table"/>` 리터럴만 박아두고 그 styleId 가 reference.docx 에 존재하는지는 보지 않는다. 이 스킬은 그 리터럴을 템플릿의 실제 앵커 styleId (예: `aa`, `TableGrid`) 로 **사후 치환**한다.
2. **`<w:tblLook>` 이 꺼져 있으면 firstRow/firstCol 서식이 발현되지 않는다.** Pandoc 기본 출력이 이 속성을 꺼서 내보내는 경우가 있어, 템플릿 원본 표의 `<w:tblLook>` 을 복제한다.
3. **styleId 가 이미 존재하면 덮어쓰고 없으면 추가.** basedOn 체인 끝까지 가져와 주입하므로 어떤 docx 에도 이식 가능.

## 제한

- 표가 없는 md 도 OK — 치환 대상이 없으면 주입 단계가 무해하게 통과.
- Pandoc 이 이미 어떤 styleId 로 표를 내보내는지는 가정하지 않음 — **모든** `<w:tbl>` 의 `<w:tblStyle>` 을 앵커로 교체.
- 이 스킬은 `md2docx-extract-style` 을 호출하지 않는다 — 호출자가 먼저 extract 를 돌려야 함.
- `--preview` 는 Windows+Word 또는 LibreOffice 필요.

## 관련 스킬

- `md2docx-extract-style` — 이 스킬의 선행 단계. 단일 Word 템플릿을 reference.docx + 표 스타일 번들로 분리 추출.
