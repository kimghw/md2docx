---
name: extract-docx-styles
description: 단일 Word(.docx/.dotx) 템플릿을 받아 세 가지 산출물을 만든다. (1) Pandoc `--reference-doc=` 으로 그대로 넘길 수 있는 `reference.docx` — 변환 후 스타일 이름 검증 포함. (2) Pandoc 과 무관하게 원본 XML 에서 직접 잘라낸 표(Table) 스타일 스니펫 — 추출 후 basedOn 체인/스타일 ID 검증 포함. (3) `--preview` 옵션 시 source.pdf + 페이지별 PNG (docx2pdf 또는 LibreOffice) — 육안/에이전트 시각 검증용.
---

# extract-docx-styles — 1개 Word 템플릿에서 세 가지를 분리 추출

입력은 Word 템플릿 **한 개** (`.docx` 또는 `.dotx`). 이 한 파일에 대해 **독립된 작업**들을 수행한다. 각 작업은 서로 영향을 주지 않는다.

| 작업 | 목적 | 산출물 | Pandoc 관계 | 기본 ON? |
|---|---|---|---|---|
| **① reference.docx 준비** | Pandoc 의 `--reference-doc=` 에 바로 쓸 사본 확보 | `reference.docx` | **필요** (매칭 규칙 검증) | ✅ |
| **② 표 스타일 추출** | 원본 docx 의 `w:style` / `w:tbl` 을 그대로 떼어냄 | `table_style.xml`, `sample_table.xml`, `table_style_bundle/` | **무관** (순수 원본 XML) | ✅ |
| **③ 시각 프리뷰** | 사람/에이전트가 눈으로 스타일 확인 | `preview/source.pdf`, `preview/source.page-NN.png` | **무관** | ❌ `--preview` 시 |

---

## 사용 도구

| 도구 | 용도 |
|---|---|
| `Bash` | `python .claude/skills/extract-docx-styles/extract.py …` 실행 |
| `Read` / `Grep` | 산출물 검증 (reference.docx 의 `styles.xml`, `table_style.xml` 내용) |

> Pandoc 은 이 스킬이 직접 호출하지 않는다. 실제 md → docx 변환은 호출자가 별도로 수행한다.

---

## 실행

```bash
python .claude/skills/extract-docx-styles/extract.py \
  --doc     references/templates/<template>.docx \
  --out-dir extracted_output/<name>/
```

Windows PowerShell:

```powershell
python .claude\skills\extract-docx-styles\extract.py `
  --doc     references\templates\<template>.docx `
  --out-dir extracted_output\<name>\
```

옵션:

| 플래그 | 기본값 | 의미 |
|---|---|---|
| `--doc` | (필수) | 입력 `.docx` / `.dotx` |
| `--out-dir` | (필수) | 출력 디렉터리 (없으면 생성) |
| `--sample-index N` | `0` | 표가 여러 개일 때 `sample_table.xml` 로 뽑을 `<w:tbl>` 의 인덱스 |
| `--preview` | off | ③ 시각 프리뷰 생성 (MS Word 또는 LibreOffice 필요) |

---

## 산출물 구조

```
extracted_output/<name>/
├── reference.docx              ← ① Pandoc --reference-doc= 입력용
├── table_style.xml             ← ② 앵커 표 스타일 한 블록
├── sample_table.xml            ← ② 원본의 <w:tbl> 샘플 (tblPr/tblGrid 포함)
├── table_style_bundle/         ← ② basedOn 체인까지 포함한 이식용 번들
│   ├── styles_excerpt.xml         ← 표 관련 w:style 블록 전체
│   └── README.txt                 ← 이식 방법 요약
├── preview/                    ← ③ --preview 시에만
│   ├── source.pdf                 ← reference.docx → PDF
│   ├── source.page-01.png         ← 페이지별 PNG (200 dpi)
│   ├── source.page-02.png
│   └── preview_report.md          ← 페이지 인덱스
└── report.tsv                  ← 세 작업의 검증 결과 요약
```

- `.dotx` 입력은 자동으로 `.docx` 사본을 만들어 처리. 원본은 변경 안 됨.
- 표가 없는 템플릿이면 ② 산출물은 생성되지 않는다 (`report.tsv` 에 `table_count=0` 기록).

---

## ① reference.docx 준비 (Pandoc 용) — 변환 후 검증

### 1-1. 변환 절차

원본 docx/dotx 를 그대로 `reference.docx` 로 복사한다. 자동 수정은 하지 않는다. XML 내부 스타일 집합이 Pandoc 매칭 규칙을 충족하는지 확인만 한다.

### 1-2. 검증 기준 (`report.tsv` 에 기록)

| 항목 | 기준 | 행동 |
|---|---|---|
| `w:name = "Normal"` | **필수** | 누락 시 `missing_required` 에 기록, 경고 출력 |
| `w:name = "heading 1"` | **필수** | 위와 동일 |
| `w:name = "heading 2"~"heading 6"` | 권장 | `missing_recommended` 기록, 경고만 |
| `w:name = "Title"`, `"Subtitle"`, `"Quote"`, `"List Paragraph"` | 권장 | 위와 동일 |
| 총 스타일 이름 개수 | 참고 | `style_name_count` 기록 |

> **왜 `w:name` 만 보는가** — Pandoc 은 스타일을 `w:name` 리터럴로 매칭한다. `w:styleId` 는 `1`, `a`, `Heading1`, `제목1Char` 등 무엇이든 상관없다. 따라서 reference.docx 는 **내부 `w:name` 집합이 온전하면 된다**.

### 1-3. 검증 실패 시

- `missing_required` 가 있으면 **Pandoc 이 해당 스타일을 못 찾아 기본 서식으로 fallback** 된다.
- 해결: Word UI 에서 해당 스타일을 만들고 (예: `Normal`, `heading 1` 이라는 정확한 이름) 저장한 뒤 이 스킬을 재실행.
- 이 스킬은 XML 에 자동으로 스타일을 끼워 넣지 않는다 — Word 가 손상 파일로 판정하는 케이스 회피 목적.

### 1-4. 사용

```bash
pandoc input.md --reference-doc=extracted_output/<name>/reference.docx -o draft.docx
```

---

## ② 표 스타일 추출 (Pandoc 과 무관) — 추출 후 검증

### 2-1. 추출 절차 (순수 ZIP + XML 조작)

1. `word/document.xml` 에서 모든 `<w:tbl>...</w:tbl>` 블록을 찾는다.
2. `--sample-index` (기본 0) 에 해당하는 표 하나를 `sample_table.xml` 로 그대로 떠낸다.
3. 그 표의 `<w:tblStyle w:val="…">` 에서 앵커 styleId 를 얻는다.
4. `word/styles.xml` 에서 앵커 styleId 의 `<w:style>` 블록을 찾아 `table_style.xml` 로 저장한다.
5. `basedOn` 체인을 거슬러 올라가며 참조되는 모든 `<w:style>` 블록을 묶어 `table_style_bundle/styles_excerpt.xml` 로 저장한다.

> Pandoc 을 호출하거나 Pandoc 출력 문서를 읽지 않는다. 전부 **원본 템플릿의 XML** 만 읽고 자른다.

### 2-2. 검증 기준 (`report.tsv` 에 기록)

| 항목 | 기준 | 필드 |
|---|---|---|
| `<w:tbl>` 개수 | 최소 1 이상 (없으면 ② 전체 스킵) | `table_count` |
| 샘플 표의 `tblStyle` | 존재 여부 확인 | `sample_style_id` |
| `w:styleId="Table"` 정의 | 존재 여부 (Pandoc 용 리터럴 매칭) | `has_style_Table` |
| basedOn 체인 | 순환 없고, 체인 끝까지 styles.xml 안에서 해소됨 | `basedOn_chain` |
| 산출물 기록 | 각 파일이 실제 쓰였는지 | `wrote_sample_table`, `wrote_table_style`, `wrote_bundle` |

### 2-3. `has_style_Table=false` 일 때의 의미

- **Pandoc 으로 md → docx 할 때 표 서식이 적용되지 않는다.** Pandoc 은 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 리터럴을 그대로 쓰므로, 대상 템플릿에 `w:styleId="Table"` 이 없으면 매칭 실패.
- `w:name="Table"` 만 있어도 안 됨. `w:aliases="Table"` 은 무시됨.
- 단, ② 산출물 자체는 **앵커 styleId 가 다른 이름(예: `aa`) 이어도 그대로 추출**된다. 다른 docx 에 이식할 때는 styleId 까지 동일하게 유지하면 된다.

### 2-4. 추출된 표 스타일의 이식 (Pandoc 무관 경로)

**경로 — 수동 이식**

1. 대상 docx 의 `word/styles.xml` 을 열어 `</w:styles>` 직전에 `table_style_bundle/styles_excerpt.xml` 의 `<w:style>` 블록들을 **그대로** 붙여넣는다.
2. styleId 값을 바꾸지 말 것 (`basedOn` 포인터가 깨짐).
3. 대상 표의 셀 배치·테두리까지 복제하려면 `sample_table.xml` 의 `<w:tblPr>`·`<w:tblGrid>` 를 대상 표의 동일 요소에 덮어쓴다.

> 과거 버전에는 `md2docx/clone_table_props.py` 자동 이식 경로가 있었다. 현재 저장소에는 해당 스킬이 없다 — 필요하면 별도 스크립트로 구현.

---

## ③ 시각 프리뷰 (`--preview`) — 육안/에이전트 검증

기본 OFF. `--preview` 를 붙이면 `reference.docx` 를 **PDF → 페이지별 PNG** 로 렌더해 `out/preview/` 아래 쌓는다. Pandoc 과 무관하며, ① / ② 와는 완전히 독립.

### 3-1. 렌더링 엔진 우선순위

스킬이 자동으로 시도:

| 순위 | 엔진 | 조건 | 비고 |
|---|---|---|---|
| 1 | **Word COM (docx2pdf)** | Windows + MS Word 설치 | 실제 Word 렌더링 (최고 충실도) |
| 2 | **LibreOffice headless** | `soffice` / `libreoffice` 가 PATH 에 있음 | 크로스플랫폼 |
| 3 | (스킵) | 위 둘 다 없음 | `report.tsv` 의 `preview.note` 에 사유 기록, `preview.engine=-` |

PDF → PNG 변환은 **PyMuPDF (`fitz`)** 사용 (200 dpi).

### 3-2. 필요 패키지

```bash
pip install docx2pdf pymupdf
```

- `docx2pdf` — Word COM 래퍼 (Windows + Word 전제)
- `pymupdf` — 순수 바이너리 wheel, 시스템 의존성 없음
- LibreOffice 경로만 쓰려면 `docx2pdf` 는 생략 가능 (import 실패 시 자동으로 다음 엔진)

### 3-3. 용도 2가지 (같은 산출물 공유)

| 시나리오 | 누가 보는가 | 판단 |
|---|---|---|
| **실사용** — 추출 결과 육안 확인 | 사용자가 `preview/*.png` 를 IDE / 뷰어로 열어봄 | "내가 의도한 스타일 맞음" |
| **스킬 검증 (validation)** | 검증 스크립트가 PNG 를 에이전트에 전달 | 테두리 유무, firstRow 음영, firstCol 음영, 폰트 일치 여부 |

### 3-4. report.tsv 에 기록되는 항목

| 필드 | 의미 |
|---|---|
| `preview.attempted` | `--preview` 로 시도했는지 |
| `preview.engine` | `docx2pdf` / `soffice` / `-` (실패) |
| `preview.pdf_path` | 생성된 PDF 의 경로 |
| `preview.page_count` | 생성된 PNG 개수 |
| `preview.note` | 실패 사유·경고 메시지 |

### 3-5. 에이전트 검증 체크리스트 (권장)

검증 스크립트에서 PNG 를 Claude 멀티모달 에이전트에 전달할 때 질문할 기준:

1. 표 **모든 셀 경계**가 보이는가? 두께·색이 `table_style.xml` 의 `w:tblBorders` 와 일치?
2. **첫 행 음영** 색이 `tblStylePr w:type="firstRow"` 의 `w:shd w:fill` 값과 일치?
3. **첫 열 음영** 색이 `tblStylePr w:type="firstCol"` 의 값과 일치?
4. 제목 1·2·3 서체가 예상과 일치 (크기/굵기/정렬)?
5. 한국어 글자가 깨지지 않고 렌더링됐는가?

---

## report.tsv 예시

```
section    key                   value
input      path                  references/templates/reference_regu.docx
input      ext                   .docx
reference  path                  extracted_output/reference_regu/reference.docx
reference  style_name_count      39
reference  missing_required      -
reference  missing_recommended   -
table      table_count           2
table      sample_index          0
table      sample_style_id       aa
table      has_style_Table       false
table      basedOn_chain         aa -> a1
table      wrote_sample_table    true
table      wrote_table_style     true
table      wrote_bundle          true
```

- `missing_required = -` & `has_style_Table = true` → 양쪽 다 통과.
- `missing_required = -` & `has_style_Table = false` → ① 통과, ② 는 수동 이식 전용 (Pandoc 경로로는 표 서식 미적용).
- `missing_required = Normal` 등 → ① 실패, 원본 Word 수정 후 재실행.

---

## 원리 요약

- **Pandoc 의 단락/문자 스타일 매칭**: `w:name` 리터럴. styleId 는 자유.
- **Pandoc 의 표 스타일 매칭**: 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 를 리터럴로 emit. 따라서 템플릿에 **`w:styleId="Table"`** 정의가 반드시 있어야 서식이 적용된다. (실측: [references/guide_table_style.md](references/guide_table_style.md))
- **Pandoc 과 무관한 표 스타일 이식**: styleId 와 basedOn 체인만 유지하면 임의 docx 에 그대로 붙여 넣어도 동작. 이 스킬이 제공하는 ② 산출물이 그 재료.

---

## 제한

- 자동 수정 안 함. `w:name` 누락·`Table` styleId 부재는 보고만 하고 XML 을 건드리지 않는다.
- `.dotx` 는 `.docx` 사본만 만든다. 매크로·첨부는 보존되나 Pandoc 이 못 읽는 경우가 있다.
- 이 스킬은 Pandoc 을 실행하지 않는다 — `pandoc` 호출은 호출자 책임.

---

## 관련 자료 (스킬 내부 references/)

- [references/guide_table_style.md](references/guide_table_style.md) — Pandoc 표 스타일 매칭 규칙 실측
- [references/pandoc-reference-docx-guide.md](references/pandoc-reference-docx-guide.md) — reference.docx 운용 가이드
- [references/pandoc-manual.md](references/pandoc-manual.md) — Pandoc 매뉴얼 발췌
- [references/claude-skill-authoring-guidelines.md](references/claude-skill-authoring-guidelines.md) — 스킬 작성 가이드
- [../_extract-docx-styles_bak/SKILL.md](../_extract-docx-styles_bak/SKILL.md) — 구 버전 (전수검증/probe/roundtrip 포함, 참고용 백업)
