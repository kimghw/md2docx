---
name: md2docx-extract-style
description: 단일 Word(.docx/.dotx) 템플릿을 받아 세 가지 산출물을 만든다. 핵심 전제 — Pandoc 은 표의 **구조만 변환**하고 **표 스타일은 처리하지 않는다**. 그래서 ① 단락/문자 스타일 (Pandoc 이 맡음) 과 ② 표 스타일 (Pandoc 이 손대지 않음, 별도 주입 필요) 을 분리 추출한다. (1) `reference.docx` — Pandoc `--reference-doc=` 용, 변환 후 스타일 이름 검증 포함. (2) 원본 XML 에서 직접 잘라낸 표 스타일 스니펫 — basedOn 체인/styleId 검증 포함. (3) `--preview` 옵션 시 source.pdf + 페이지별 PNG (docx2pdf 또는 LibreOffice) — 육안/에이전트 시각 검증용.
---

# md2docx-extract-style — 1개 Word 템플릿에서 세 가지를 분리 추출

입력은 Word 템플릿 **한 개** (`.docx` 또는 `.dotx`). 이 한 파일에 대해 **독립된 작업**들을 수행한다. 각 작업은 서로 영향을 주지 않는다.

> **🔑 설계 전제 — Pandoc 책임 경계**
>
> Pandoc 의 md → docx 변환은 **표의 구조(행·열·셀·정렬)만 생성**하고 **표 스타일(테두리·음영·서체·음영 밴딩)은 건드리지 않는다.** 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 리터럴을 꽂아둘 뿐, 그 styleId 가 reference.docx 에 실제 존재하는지·어떤 속성을 갖는지는 검증하지 않는다.
>
> 그래서 이 스킬은 **① 단락/문자 스타일**(Pandoc 이 `w:name` 으로 매칭하는 부분) 과 **② 표 스타일**(Pandoc 이 손대지 않는 부분) 을 **분리**해 관리한다. ② 는 변환 이후 별도 단계에서 주입해야 한다 — Pandoc 한 번에 끝나지 않는다.

| 작업 | 목적 | 산출물 | Pandoc 관계 | 기본 ON? |
|---|---|---|---|---|
| **① reference.docx 준비** | Pandoc 의 `--reference-doc=` 에 바로 쓸 사본 확보 | `reference.docx` | **쓰임** — 단락/문자 스타일 매칭 규칙 검증 | ✅ |
| **② 표 스타일 추출** | 원본 docx 의 `w:style` / `w:tbl` 을 그대로 떼어냄 | `table_style.xml` (앵커 + basedOn 체인 정의), `table_sources.xml` (`<w:tbl>` 인스턴스 스냅샷) | **무관** — Pandoc 이 표 스타일을 처리하지 않아서 **따로 필요** | ✅ |
| **③ 시각 프리뷰** | 사람/에이전트가 눈으로 스타일 확인 | `preview/source.pdf`, `preview/source.page-NN.png` | **무관** | ❌ `--preview` 시 |

---

## 사용 도구

| 도구 | 용도 |
|---|---|
| `Bash` | `python .claude/skills/md2docx-extract-style/extract.py …` 실행 |
| `Read` / `Grep` | 산출물 검증 (reference.docx 의 `styles.xml`, `table_style.xml` 내용) |

> Pandoc 은 이 스킬이 직접 호출하지 않는다. 실제 md → docx 변환은 호출자가 별도로 수행한다.

---

## 실행

```bash
python .claude/skills/md2docx-extract-style/extract.py \
  --doc     references/templates/<template>.docx \
  --out-dir extracted_output/
```

Windows PowerShell:

```powershell
python .claude\skills\md2docx-extract-style\extract.py `
  --doc     references\templates\<template>.docx `
  --out-dir extracted_output\
```

산출물은 `--out-dir` 아래 **원본 파일명(확장자 제외)** 서브디렉터리에 들어간다.
예) `--doc references/templates/reference_reg.docx --out-dir extracted_output/`
→ `extracted_output/reference_reg/` 에 모든 산출물이 쌓인다.

옵션:

| 플래그 | 기본값 | 의미 |
|---|---|---|
| `--doc` | (필수) | 입력 `.docx` / `.dotx` |
| `--out-dir` | (필수) | **부모 디렉터리**. 실제 출력은 `<out-dir>/<원본파일이름>/` 에 생성 |
| `--sample-index N` | `0` | 표가 여러 개일 때 `table_sources.xml` 로 뽑을 `<w:tbl>` 의 인덱스 |
| `--preview` | off | ③ 시각 프리뷰 생성 (MS Word 또는 LibreOffice 필요) |

---

## 산출물 구조

```
<out-dir>/
└── <원본파일이름>/                ← extract.py 가 자동 생성
    ├── base_templates/             ← ① Pandoc --reference-doc= 입력용
    │   └── reference.docx             ← 원본 복사본
    ├── table_style.xml             ← ② 표 스타일 정의 (앵커 + basedOn 체인 전체를 <styles-excerpt> 로 묶음)
    ├── table_sources.xml           ← ② 원본의 <w:tbl> 인스턴스 스냅샷 (tblPr/tblGrid/tblLook/cells)
    ├── preview/                    ← ③ --preview 시에만 (md2docx 의 final.* 도 같은 폴더에 누적)
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
pandoc input.md --reference-doc=extracted_output/<원본파일이름>/base_templates/reference.docx -o draft.docx
```

---

## ② 표 스타일 추출 (Pandoc 과 무관) — 추출 후 검증

> **왜 ② 가 별도로 필요한가** — Pandoc 은 md 표를 변환할 때 **표의 구조만 본다**: 행 개수·열 개수·셀 내용·정렬. 표 스타일(테두리 두께·색·음영·첫 행 굵게·첫 열 강조 등) 은 전혀 고려하지 않고, 그저 `<w:tblStyle w:val="Table"/>` 이라는 **고정 리터럴**만 출력에 찍어둔다. 따라서 템플릿 원본이 아무리 예쁜 표 스타일을 정의해도, Pandoc 만 돌려서는 출력 docx 에 그 스타일이 연결되지 않는다. 이 섹션이 제공하는 스니펫을 **변환 후 별도 단계**에서 주입해야 서식이 살아난다.

### 2-1. 추출 절차 (순수 ZIP + XML 조작)

1. `word/document.xml` 에서 모든 `<w:tbl>...</w:tbl>` 블록을 찾는다.
2. `--sample-index` (기본 0) 에 해당하는 표 하나를 `table_sources.xml` 로 그대로 떠낸다.
3. 그 표의 `<w:tblStyle w:val="…">` 에서 앵커 styleId 를 얻는다.
4. `word/styles.xml` 에서 앵커 styleId 의 `<w:style>` 블록을 찾아 `table_style.xml` 로 저장한다.
5. `basedOn` 체인을 거슬러 올라가며 참조되는 모든 `<w:style>` 블록을 묶어 `table_style.xml` 에 `<styles-excerpt anchor="..." chain="...">` 래퍼로 저장한다.

> Pandoc 을 호출하거나 Pandoc 출력 문서를 읽지 않는다. 전부 **원본 템플릿의 XML** 만 읽고 자른다.

### 2-2. 검증 기준 (`report.tsv` 에 기록)

| 항목 | 기준 | 필드 |
|---|---|---|
| `<w:tbl>` 개수 | 최소 1 이상 (없으면 ② 전체 스킵) | `table_count` |
| 샘플 표의 `tblStyle` | 존재 여부 확인 | `sample_style_id` |
| `w:styleId="Table"` 정의 | 존재 여부 (Pandoc 용 리터럴 매칭) | `has_style_Table` |
| basedOn 체인 | 순환 없고, 체인 끝까지 styles.xml 안에서 해소됨 | `basedOn_chain` |
| 산출물 기록 | 각 파일이 실제 쓰였는지 | `wrote_table_sources`, `wrote_table_style` |

> `table_sources.xml` 은 원본 템플릿의 **실제 `<w:tbl>` XML 한 덩어리 그대로**를 담은 인스턴스 스냅샷이다 (Pandoc 출력이 아님). md2docx 의 주입 단계에서 이 파일의 `<w:tblLook>` / `<w:tblGrid>` / 셀 pPr·rPr 을 Pandoc 이 생성한 표에 덮어씌워 서식을 되살린다. 같은 표의 **서식 정의** 는 `table_style.xml` 에 별도로 들어 있다.

### 2-3. `has_style_Table=false` 일 때의 의미

- **Pandoc 으로 md → docx 할 때 표 서식이 적용되지 않는다.** Pandoc 은 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 리터럴을 그대로 쓰므로, 대상 템플릿에 `w:styleId="Table"` 이 없으면 매칭 실패.
- `w:name="Table"` 만 있어도 안 됨. `w:aliases="Table"` 은 무시됨.
- 단, ② 산출물 자체는 **앵커 styleId 가 다른 이름(예: `aa`) 이어도 그대로 추출**된다. 다른 docx 에 이식할 때는 styleId 까지 동일하게 유지하면 된다.

### 2-4. 추출된 표 스타일의 이식 (Pandoc 무관 경로)

**경로 — 수동 이식**

1. 대상 docx 의 `word/styles.xml` 을 열어 `</w:styles>` 직전에 `table_style.xml` 안쪽의 `<w:style>` 블록들을 (`<styles-excerpt>` 래퍼는 벗기고) **그대로** 붙여넣는다.
2. styleId 값을 바꾸지 말 것 (`basedOn` 포인터가 깨짐).
3. 대상 표의 셀 배치·테두리까지 복제하려면 `table_sources.xml` 의 `<w:tblPr>`·`<w:tblGrid>` 를 대상 표의 동일 요소에 덮어쓴다.

> 과거 버전에는 `md2docx/clone_table_props.py` 자동 이식 경로가 있었다. 현재 저장소에는 해당 스킬이 없다 — 필요하면 별도 스크립트로 구현.

### 2-5. 실제 재사용 가능성 검증 (C9)

정적 추출만으로는 **"잘라낸 조각이 실제로 Word 표에 적용되는가"** 를 증명할 수 없다. 이를 위해 `tests/inject_and_render.py` 가 **왕복 검증**을 수행한다:

```
[1] 빈 docx 생성 (python-docx 로 3×3 plain table)
[2] table_style.xml 의 <w:style> 블록들을 target 의 styles.xml 에 주입
    - 동일 styleId 가 있으면 교체, 없으면 </w:styles> 앞에 추가
[3] 빈 docx 의 <w:tbl> 의 <w:tblPr> 을 다시 씀:
    - <w:tblStyle w:val="{anchor}"/> 지정
    - table_sources.xml 의 <w:tblLook> 을 복사해 firstRow/firstCol 활성화
[4] injected.docx 를 python-docx 로 재오픈 (구조 sanity)
[5] docx2pdf → PNG 로 렌더 (injected + baseline 두 장)
[6] 에이전트(Claude vision) 가 두 PNG 의 서식 속성을 비교 판정
```

**C9 가 검증하는 것**:
- 추출 번들이 Word 가 열 수 있는 유효한 구조인가 (structural)
- 주입한 뒤 실제로 렌더에서 테두리·음영이 살아나는가 (visual)
- 앵커 styleId + basedOn 체인이 **원본과 다른 docx 에서도 해소**되는가 (reusability)

실행:

```bash
python .claude/skills/md2docx-extract-style/_tests/inject_and_render.py \
  --extract-out extracted_output/<원본파일이름>/ \
  --source      references/templates/<template>.docx \
  --out-dir     extracted_output/<원본파일이름>/c9_inject/
```

산출물:
- `c9_inject/injected.docx` — 빈 docx + 주입된 번들 + 앵커 적용된 표
- `c9_inject/injected.page-NN.png` — 렌더된 PNG
- `c9_inject/baseline.page-NN.png` — 원본 템플릿 렌더 (비교 기준)

> **왕복 검증 없이는** ② 산출물은 "XML 조각이 잘려 있다" 이상을 주장할 수 없다. 실제 Word 에 적용되는지는 C9 을 거쳐야 안다.

### 2-6. 검증 산출물 위치

`tests/` 디렉터리는 **스크립트 + fixture 만** 둔다. 실제 검증 실행 결과는 `extracted_output/_validate_runs/<fixture>/` 로 저장된다 (즉, 스킬 내부가 아닌 리포 루트의 생성물 디렉터리 하위). 환경변수 `VALIDATE_OUT` 으로 경로 재지정 가능.

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
reference  path                  extracted_output/reference_regu/base_templates/reference.docx
reference  style_name_count      39
reference  missing_required      -
reference  missing_recommended   -
table      table_count           2
table      sample_index          0
table      sample_style_id       aa
table      has_style_Table       false
table      basedOn_chain         aa -> a1
table      wrote_table_sources    true
table      wrote_table_style     true
```

- `missing_required = -` & `has_style_Table = true` → 양쪽 다 통과.
- `missing_required = -` & `has_style_Table = false` → ① 통과, ② 는 수동 이식 전용 (Pandoc 경로로는 표 서식 미적용).
- `missing_required = Normal` 등 → ① 실패, 원본 Word 수정 후 재실행.

---

## 원리 요약

- **Pandoc 책임 경계** — Pandoc 은 표의 **구조만 변환**하고 **표 스타일은 처리하지 않는다**. md 의 표 → docx 의 `<w:tbl>` (행·열·정렬·내용) 까지만. 테두리·음영·서체·밴딩 같은 서식 지정은 Pandoc 의 업무가 아니다.
- **Pandoc 의 단락/문자 스타일 매칭**: `w:name` 리터럴. styleId 는 자유. → ① 에서 검증.
- **Pandoc 의 표 스타일 "매칭" 은 사실상 고정 스텁**: 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 를 **조건 없이** emit. 따라서 reference.docx 에 **`w:styleId="Table"`** 정의가 **있든 없든** Pandoc 은 신경 쓰지 않는다. 서식이 살아나려면 변환 후 별도 주입이 필요하다. (실측: [references/table-style.md](references/table-style.md))
- **② 산출물의 역할** — Pandoc 이 맡지 않는 표 스타일을 **후처리 단계**에서 주입할 수 있도록, 원본 템플릿의 `w:style` / `w:tbl` 을 **Pandoc 을 거치지 않고** 그대로 잘라낸다. styleId 와 basedOn 체인만 유지하면 임의 docx 에 이식 가능.

---

## 제한

- 자동 수정 안 함. `w:name` 누락·`Table` styleId 부재는 보고만 하고 XML 을 건드리지 않는다.
- `.dotx` 는 `.docx` 사본만 만든다. 매크로·첨부는 보존되나 Pandoc 이 못 읽는 경우가 있다.
- 이 스킬은 Pandoc 을 실행하지 않는다 — `pandoc` 호출은 호출자 책임.

---

## 관련 자료 (스킬 내부 references/)

- [references/table-style.md](references/table-style.md) — OOXML 표 스타일 구조 + Pandoc 표 스타일 매칭 규칙 실측
- [references/pandoc.md](references/pandoc.md) — Pandoc 일반 사용법 + reference.docx 운용 가이드
- [references/claude-skill-authoring-guidelines.md](references/claude-skill-authoring-guidelines.md) — 스킬 작성 가이드
