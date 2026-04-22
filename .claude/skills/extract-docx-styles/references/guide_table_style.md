# Pandoc 표 스타일 매칭 규칙 (검증 결과)

Pandoc 3.9.0.2 기준, `--reference-doc`으로 넘긴 Word 템플릿의 스타일을
Pandoc이 어떻게 찾아 적용하는지 실변환으로 검증한 결과를 기록한다.

## 결론 요약

| 스타일 종류 | Pandoc 매칭 방식 |
|---|---|
| Paragraph / Character (Heading, Title, Normal, Quote, …) | **`w:name` 기준** — styleId는 임의의 값이어도 됨 |
| **Table** | **`w:styleId="Table"` 기준 (literal 하드코드)** — `w:name="Table"`만으론 부족, `w:aliases="Table"`은 무시됨 |

## 검증 방법 및 결과

### Test 1 — 템플릿 그대로 변환

`references/templates/` 의 rev3 / rev4 / rev5 를 각각 `--reference-doc`으로
지정해 동일한 MD(표 하나)를 변환하고, 출력 docx의
`word/document.xml` 안 `<w:tblStyle w:val="…">` 값과
`word/styles.xml`에 해당 styleId가 있는지 확인.

| 템플릿 | 출력 `<w:tblStyle w:val="…">` | 출력에 해당 styleId 존재? | 적용 결과 |
|---|---|:-:|---|
| rev3 | `Table` | ✅ (테두리 있음) | 정상 적용 |
| rev4 | `Table` | ✅ (테두리 없음) | 정상 적용(플레인) |
| rev5 | `Table` | ❌ (`w:aliases="Table"`만 있음) | **dangling reference → 서식 미적용** |

### Test 2 — Table 스타일: styleId vs w:name

rev3 복제본에서 커스텀 Table 스타일의 **`w:styleId="Table"` → `"MyTable"`** 로
변경, `w:name="Table"`은 그대로 둠.

| 확인 | 결과 |
|---|---|
| Pandoc이 emit한 값 | `<w:tblStyle w:val="Table"/>` |
| 출력 styles.xml에 `styleId="Table"` | ❌ |
| 출력 styles.xml에 `styleId="MyTable"` (name="Table") | ✅ |

→ Pandoc은 표에 대해 **literal 문자열 "Table"을 styleId로 하드코드 emit**한다.
`w:name`으로 lookup하여 치환하지 않는다.

### Test 3 — Paragraph 스타일: styleId vs w:name

default 복제본에서 paragraph 스타일들의 **`w:styleId`만 변경**, `w:name`은 유지.

| 원래 styleId | 변경 후 | w:name (유지) |
|---|---|---|
| `1` | `MyH1` | `heading 1` |
| `2` | `MyH2` | `heading 2` |
| `a` | `MyNormal` | `Normal` |
| `a3` | `MyTitle` | `Title` |
| `a5` | `MyQuote` | `Quote` |
| `a6` | `MyList` | `List Paragraph` |

변환 결과 Pandoc이 emit한 `<w:pStyle w:val="…">`:

| MD 요소 | emit된 값 |
|---|---|
| `# 제목1` | `MyH1` |
| `## 제목2` | `MyH2` |
| 본문 첫 문단 | `FirstParagraph` |
| `> 인용` | `BlockText` |
| `- 리스트` | `Compact` |

→ Pandoc은 paragraph 스타일에 대해 **`w:name`으로 조회하여 해당 스타일의
styleId를 찾아 emit**한다.

## 템플릿 평가

| 템플릿 | 문단 스타일 | 표 스타일 | 비고 |
|---|:-:|:-:|---|
| default | ✅ | ❌ | Table 스타일 자체가 없음 |
| rev1 | ✅ | ❌ | Table Grid만 있고 styleId="Table" 없음 |
| rev2 | ✅ | ❌ | custom styleId="10"(name="스타일1"), Table styleId 없음 |
| rev3 | ✅ | ✅ | styleId="Table" + 테두리 |
| rev4 | ✅ | ✅ | styleId="Table" + 테두리 없음 (플레인) |
| rev5 | ✅ | ❌ | `w:aliases="Table"` 방식은 Pandoc이 인식 못 함 |

## 템플릿 제작 시 체크리스트

1. 표 서식을 Pandoc으로 살리려면 템플릿에 반드시 아래 스타일이 존재해야 한다.
   ```xml
   <w:style w:type="table" w:customStyle="1" w:styleId="Table">
     <w:name w:val="Table"/>
     <w:basedOn w:val="..."/>   <!-- 기존 표 스타일 상속 -->
     ...서식 지정...
   </w:style>
   ```
   - **`w:styleId`는 반드시 `"Table"`**
   - `w:name`도 `"Table"`로 맞추는 것이 안전
   - `w:aliases`에 `"Table"`을 넣는 방식(rev5)은 **작동하지 않음**

2. 문단/문자 스타일(heading 1~9, Title, Subtitle, Quote, List Paragraph 등)은
   `w:name`만 Pandoc 기대값과 일치하면 styleId는 아무거나 상관없다.
   한글 Word가 자동 부여한 `1`, `a`, `a3` 같은 styleId도 그대로 OK.

## 테스트 환경

- Pandoc 3.9.0.2
- 입력 MD: 표 한 개 / 제목·본문·인용·리스트가 섞인 짧은 문서
- 레퍼런스 docx: `references/templates/style_template_*.docx`
