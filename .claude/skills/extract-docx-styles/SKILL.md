---
name: extract-docx-styles
description: 단일 Word(.docx/.dotx) 템플릿을 받아 두 가지 산출물을 만든다. (1) Pandoc `--reference-doc=` 으로 그대로 넘길 수 있는 `reference.docx`, (2) 표(Table) 스타일 정보만 잘라내 다른 docx 에 이식 가능하도록 준비한 스니펫/번들. `md2docx` 나 수동 변환 파이프라인에 넘길 사전 준비 단계로 호출된다.
---

# extract-docx-styles — 1개 Word 템플릿 → reference.docx + 표 스타일 추출

입력은 Word 템플릿 **한 개** (`.docx` 또는 `.dotx`). 이 한 파일에 대해 **서로 다른 두 가지 작업**을 수행한다:

| 작업 | 산출물 | 용도 |
|---|---|---|
| **① reference.docx 준비** | `reference.docx` | Pandoc 에 `--reference-doc=` 로 그대로 넘겨 초안 변환에 사용. 단락/문자/제목 스타일 소스. |
| **② 표 스타일 추출** | `table_style.xml`, `sample_table.xml`, `table_style_bundle/` | 이 템플릿의 표 스타일만 잘라내 보관. 다른 docx 의 표에 이식하거나, `md2docx/clone_table_props.py` 가 참조할 소스로 사용. |

두 산출물은 독립적이다. reference.docx 로 Pandoc 변환 후, 표 스타일은 별도 단계에서 출력 docx 에 주입한다.

---

## 사용 도구

| 도구 | 용도 |
|---|---|
| `Bash` | `python .claude/skills/extract-docx-styles/extract.py …` 실행 |
| `Read` / `Grep` | 산출물 검증 (`reference.docx` 내부 styles.xml, `table_style.xml` 내용) |
| (외부 CLI) | Pandoc 은 이 스킬이 직접 호출하지 않는다. 실제 변환은 `md2docx` 가 담당. |

---

## 호출 절차

### 1. 단일 템플릿으로 실행

```bash
python .claude/skills/extract-docx-styles/extract.py \
  --doc     references/templates/style_template_table_name_rev3.docx \
  --out-dir extracted_output/_styles/
```

Windows PowerShell:

```powershell
python .claude\skills\extract-docx-styles\extract.py `
  --doc     references\templates\style_template_table_name_rev3.docx `
  --out-dir extracted_output\_styles\
```

### 2. 산출물 구조

```
extracted_output/_styles/
├── reference.docx              ← ① Pandoc --reference-doc= 로 그대로 사용
├── table_style.xml             ← ② Table 스타일 정의 스니펫 (w:style 블록)
├── sample_table.xml            ← ② 샘플 w:tbl 스니펫 (tblPr/tblGrid 포함)
├── table_style_bundle/         ← ② 이식용 번들 (basedOn 체인 포함)
│   ├── styles_excerpt.xml         ← 표 관련 w:style 블록 전체
│   └── README.txt                 ← 이식 방법 요약
└── report.tsv                  ← 두 작업의 검증 결과 요약
```

- `.dotx` 입력은 자동으로 임시 `.docx` 로 복사해 처리, 원본은 변경되지 않는다.
- `reference.docx` 는 원본의 내용/스타일을 그대로 가진 사본이다 (Pandoc 이 원하는 `w:name` 매칭만 검증).

### 3. reference.docx → Pandoc

```bash
pandoc input.md --reference-doc=extracted_output/_styles/reference.docx -o draft.docx
```

또는 `md2docx` 사용:

```bash
python .claude/skills/md2docx/transform.py \
  --md input.md \
  --ref extracted_output/_styles/reference.docx \
  --out output.docx
```

### 4. 표 스타일 → 다른 docx 에 이식

추출된 `table_style.xml` 을 대상 docx 에 적용하는 두 가지 경로:

**경로 A — `md2docx/clone_table_props.py` 재사용 (권장)**

`clone_table_props.py` 는 이미 "템플릿 docx → 대상 docx" 로 샘플 표 속성을 평탄화·주입한다. 이 스킬의 `reference.docx` 자체가 표 스타일도 가진 템플릿이므로 그대로 넣으면 된다:

```bash
python .claude/skills/md2docx/clone_table_props.py \
  --template extracted_output/_styles/reference.docx \
  --target   other_document.docx
```

**경로 B — 수동 이식**

`table_style_bundle/styles_excerpt.xml` 에 들어있는 `<w:style>` 블록을, 대상 docx 의 `word/styles.xml` 의 `</w:styles>` 직전에 추가. 표 셀의 `tblPr`/`tblGrid` 까지 그대로 복제하고 싶으면 `sample_table.xml` 을 참고해 대상 표의 해당 요소를 갱신.

---

## 검증 기준 (report.tsv 에 기록)

`extract.py` 는 자동으로 다음을 확인한다:

### ① reference.docx 측

| 항목 | 기준 |
|---|---|
| 필수 `w:name` | `Normal`, `heading 1` 존재 |
| 권장 `w:name` | `heading 2`~`heading 6`, `Title`, `Quote`, `List Paragraph` 등 |
| 누락 처리 | 자동 수정하지 않고 report.tsv 에 기록 (Word 에서 수동 편집 권장) |

### ② 표 스타일 측

| 항목 | 기준 |
|---|---|
| `w:styleId="Table"` | 존재 여부 (Pandoc 출력과 매칭되는 유일한 형태) |
| `<w:tbl>` 샘플 | `document.xml` 에 최소 1개 이상 존재 (표 없으면 ② 산출물 생성 생략) |
| basedOn 체인 | 순환 없고, 체인 끝까지 styles.xml 안에서 해소됨 |

---

## 원리 요약

- **Pandoc 의 단락/문자 스타일 매칭**: `w:name` 리터럴로 찾는다. styleId 는 `1`, `a`, `Heading1`, `제목1Char` 등 무엇이든 OK. 따라서 reference.docx 는 **내부 `w:name` 집합** 이 온전하면 된다.
- **Pandoc 의 표 스타일 매칭**: 출력 `document.xml` 에 `<w:tblStyle w:val="Table"/>` 리터럴을 그대로 emit 한다. 따라서 템플릿에 **`w:styleId="Table"`** 정의가 반드시 있어야 서식이 적용된다. `w:name="Table"` 만으로는 부족하고 `w:aliases="Table"` 은 무시된다. (실측: [references/guide_table_style.md](../../../references/guide_table_style.md))
- 표 스타일을 **다른 docx 에 이식**할 때는 styleId 까지 유지해야 한다. `w:styleId="Table"` 이면 그대로, 커스텀 ID 면 대상 docx 에서도 같은 ID 로 심거나 `basedOn` 체인째 이식.

---

## 제한

- **자동 수정 안 함.** 누락된 `w:name` 이나 `Table` 스타일을 이 스킬이 XML 로 삽입하지 않는다 — Word 가 손상 파일로 판정하는 케이스를 피하려는 보수적 정책. 누락이 보고되면 Word UI 에서 해당 스타일을 만들어 저장 후 재실행.
- **표가 여러 개**인 템플릿은 기본적으로 첫 번째 `<w:tbl>` 을 `sample_table.xml` 로 뽑는다. 다른 표를 쓰고 싶으면 `--sample-index N`.
- **.dotx** 입력은 `.docx` 사본만 만든다. 매크로·첨부는 보존되지만 Pandoc 이 읽지 못할 수 있다.
- **이 스킬은 Pandoc 을 실행하지 않는다.** 실제 변환은 호출자 (`md2docx` 또는 직접 `pandoc` 호출) 의 책임.

---

## 관련 자료

- [references/guide_table_style.md](../../../references/guide_table_style.md) — Pandoc 매칭 규칙 실측 결과
- [references/pandoc-reference-docx-guide.md](../../../references/pandoc-reference-docx-guide.md) — reference.docx 운용 가이드
- [../md2docx/SKILL.md](../md2docx/SKILL.md) — 변환 + 표 주입을 담당하는 후속 스킬
- [../_extract-docx-styles_bak/SKILL.md](../_extract-docx-styles_bak/SKILL.md) — 구 버전(전수검증/probe/roundtrip 포함), 참고용 백업
