# 표(Table) 스타일 — OOXML 구조와 Pandoc 매칭 규칙

Word 문서의 표 서식이 OOXML 에서 어떻게 정의되는지,
그리고 Pandoc 이 표 스타일을 어떻게 매칭하는지 한 파일에 정리.
Pandoc 일반 사용법/단락·문자 스타일 매칭은 [pandoc.md](pandoc.md) 참고.

---

## 요약

- **Pandoc 은 표 스타일을 `w:styleId="Table"` 리터럴로 하드코드**한다.
  단락/문자 스타일처럼 `w:name` 으로 lookup 하지 않는다.
  → reference.docx 에 `styleId="Table"` 스타일이 없으면 표 서식이 전혀 적용되지 않는다.
- **`<w:style type="table">` 한 덩어리에 표 서식 대부분이 담긴다.**
  테이블/행/셀/셀 안 문단/셀 안 글자 속성을 하나의 스타일 안에서 모두 통제 가능.
- **표 본문(`<w:tbl>`)에만 담을 수 있는 정보**(열 폭, tblLook, 병합 정보 등)는 스타일에 못 담는다. → 샘플 표 XML 이 따로 필요한 이유.

---

## 1. Pandoc 의 표 스타일 매칭 규칙 (검증 결과)

Pandoc 3.9.0.2 기준, `--reference-doc` 으로 넘긴 Word 템플릿의 표 스타일을 어떻게 찾는지 실변환으로 검증한 결과.

| 스타일 종류 | Pandoc 매칭 방식 |
|---|---|
| Paragraph / Character (Heading, Title, Normal, Quote, …) | **`w:name` 기준** — styleId 는 임의의 값이어도 됨 |
| **Table** | **`w:styleId="Table"` 기준 (literal 하드코드)** — `w:name="Table"` 만으론 부족, `w:aliases="Table"` 은 무시됨 |

### 검증 1 — 템플릿 그대로 변환

`references/templates/` 의 rev3 / rev4 / rev5 를 각각 `--reference-doc` 으로 지정해 동일한 MD(표 하나)를 변환하고, 출력 docx 의 `word/document.xml` 안 `<w:tblStyle w:val="…">` 값과 `word/styles.xml` 에 해당 styleId 가 있는지 확인.

| 템플릿 | 출력 `<w:tblStyle w:val="…">` | 출력에 해당 styleId 존재? | 적용 결과 |
|---|---|:-:|---|
| rev3 | `Table` | ✅ (테두리 있음) | 정상 적용 |
| rev4 | `Table` | ✅ (테두리 없음) | 정상 적용(플레인) |
| rev5 | `Table` | ❌ (`w:aliases="Table"` 만 있음) | **dangling reference → 서식 미적용** |

### 검증 2 — Table 스타일: styleId vs w:name

rev3 복제본에서 커스텀 Table 스타일의 **`w:styleId="Table"` → `"MyTable"`** 로 변경, `w:name="Table"` 은 그대로 둠.

| 확인 | 결과 |
|---|---|
| Pandoc 이 emit 한 값 | `<w:tblStyle w:val="Table"/>` |
| 출력 styles.xml 에 `styleId="Table"` | ❌ |
| 출력 styles.xml 에 `styleId="MyTable"` (name="Table") | ✅ |

→ Pandoc 은 표에 대해 **literal 문자열 "Table" 을 styleId 로 하드코드 emit** 한다.
`w:name` 으로 lookup 하여 치환하지 않는다.

### 검증 3 — Paragraph 스타일 (대비)

default 복제본에서 paragraph 스타일들의 **`w:styleId` 만 변경**, `w:name` 은 유지.

| 원래 styleId | 변경 후 | w:name (유지) |
|---|---|---|
| `1` | `MyH1` | `heading 1` |
| `2` | `MyH2` | `heading 2` |
| `a` | `MyNormal` | `Normal` |
| `a3` | `MyTitle` | `Title` |
| `a5` | `MyQuote` | `Quote` |
| `a6` | `MyList` | `List Paragraph` |

변환 결과 Pandoc 이 emit 한 `<w:pStyle w:val="…">`:

| MD 요소 | emit 된 값 |
|---|---|
| `# 제목1` | `MyH1` |
| `## 제목2` | `MyH2` |
| 본문 첫 문단 | `FirstParagraph` |
| `> 인용` | `BlockText` |
| `- 리스트` | `Compact` |

→ Pandoc 은 paragraph 스타일에 대해서는 **`w:name` 으로 조회하여 해당 스타일의 styleId 를 찾아 emit** 한다. 표와 정반대.

### 템플릿 평가

| 템플릿 | 문단 스타일 | 표 스타일 | 비고 |
|---|:-:|:-:|---|
| default | ✅ | ❌ | Table 스타일 자체가 없음 |
| rev1 | ✅ | ❌ | Table Grid 만 있고 styleId="Table" 없음 |
| rev2 | ✅ | ❌ | custom styleId="10"(name="스타일1"), Table styleId 없음 |
| rev3 | ✅ | ✅ | styleId="Table" + 테두리 |
| rev4 | ✅ | ✅ | styleId="Table" + 테두리 없음 (플레인) |
| rev5 | ✅ | ❌ | `w:aliases="Table"` 방식은 Pandoc 이 인식 못 함 |

---

## 2. 템플릿에 표 스타일을 만드는 올바른 절차

표 스타일은 임의로 아무거나 갖다 쓰면 안 된다. **템플릿 안에 이미 있는 표가 실제로 어떤 스타일을 쓰는지** 먼저 확인하고, 그 스타일을 기반으로 `Table` 이름의 스타일을 만들어야 한다.

### 1단계: 템플릿의 표가 사용하는 스타일 ID 확인
```bash
unzip -p your-template.docx word/document.xml | grep -oE 'w:tblStyle w:val="[^"]+"' | sort | uniq -c
```
출력 예:
```
      3 w:tblStyle w:val="aa"        <- 주로 이걸 씀
      1 w:tblStyle w:val="10"
```
→ 템플릿의 표들이 실제로 사용 중인 style ID 를 파악

### 2단계: 해당 style ID 가 어떤 이름과 서식을 가진지 확인
```bash
unzip -p your-template.docx word/styles.xml | grep -oE '<w:style [^>]*w:styleId="aa"[^>]*>.{0,500}'
```
출력 예:
```xml
<w:style w:type="table" w:styleId="aa">
  <w:name w:val="Table Grid"/>
  <w:basedOn w:val="a1"/>
  <w:tblPr>
    <w:tblBorders>...</w:tblBorders>
  </w:tblPr>
</w:style>
```
→ 이 서식이 사용자가 실제로 원하는 표 모양

### 3단계: 그 스타일을 복제해서 styleId 를 `Table` 로 추가
`</w:styles>` 바로 앞에 삽입:
```xml
<w:style w:type="table" w:customStyle="1" w:styleId="Table">
  <w:name w:val="Table"/>
  <w:basedOn w:val="aa"/>   <!-- ⭐ 1단계에서 찾은 실제 사용 style ID -->
</w:style>
```
- `basedOn` 으로 기존 스타일을 상속 → 테두리, 배경, 폰트 등 모든 서식 자동 적용
- 일일이 `tblBorders` 복사할 필요 없음
- 템플릿 여러 곳에서 다른 표 스타일을 쓴다면, 가장 많이 쓰이는 것을 `basedOn` 에 지정

### 왜 이렇게 해야 하는가?
- ❌ 잘못된 방법: "한글 Word 는 보통 Table Grid 를 쓰니까 `aa` 갖다 쓰자" (추측)
- ✅ 올바른 방법: 템플릿 실제 내용을 보고 어떤 표 스타일이 쓰이는지 확인 후 그걸 기반으로 제작

템플릿마다 쓰는 표 스타일이 다르다. 회사 양식은 커스텀 `aa` 를 쓸 수도, 기본 `Plain Table 1` 을 쓸 수도, 아예 커스텀 스타일(`CompanyTable` 등)을 만들어 쓸 수도 있다. 추측하지 말고 실제 확인.

### Word UI 에서 하는 경우
1. 템플릿 Word 로 열기
2. 기존 표 하나 클릭 → 디자인 탭에서 현재 적용된 스타일 이름 확인
3. 스타일 갤러리에서 그 스타일 **우클릭 → 복제**
4. 복제본 이름을 정확히 **`Table`** 로 변경
5. 저장

### 템플릿 제작 체크리스트

1. 표 서식을 Pandoc 으로 살리려면 템플릿에 반드시 아래 스타일이 존재해야 한다.
   ```xml
   <w:style w:type="table" w:customStyle="1" w:styleId="Table">
     <w:name w:val="Table"/>
     <w:basedOn w:val="..."/>   <!-- 기존 표 스타일 상속 -->
     ...서식 지정...
   </w:style>
   ```
   - **`w:styleId` 는 반드시 `"Table"`**
   - `w:name` 도 `"Table"` 로 맞추는 것이 안전
   - `w:aliases` 에 `"Table"` 을 넣는 방식(rev5)은 **작동하지 않음**

2. 문단/문자 스타일(heading 1~9, Title, Subtitle, Quote, List Paragraph 등)은 `w:name` 만 Pandoc 기대값과 일치하면 styleId 는 아무거나 상관없다. 자세한 규칙은 [pandoc.md § 5](pandoc.md).

### 관련 표 관련 스타일 이름

| Pandoc 요소 | w:name |
|------------|--------|
| 표 본체 | `Table` (필수, 자주 누락) |
| 표 캡션 | `Table Caption` |
| 기본 테이블 | `Normal Table` |
| 테이블 그리드 | `Table Grid` |

---

## 3. OOXML 표 스타일 구조 — `<w:style type="table">` 완전 해부

Word 문서의 표(Table) 서식은 **하나의 `<w:style>` 블록** 에 전부 담을 수 있다. 테이블 레벨·셀 레벨·셀 안 문단 레벨·셀 안 글자 레벨 속성 **전부** 가 이 하나의 스타일 안에서 통제 가능하다.

> **제대로 정의된 테이블 스타일 한 덩어리만 있으면 표 서식 정보는 다 있다.**
> 추출도 이 스타일 하나만 올바르게 떠내면 재사용 가능. 셀 단위 XML 을 별도로 긁어올 필요는 **원칙적으로 없다** (스타일에 누락이 있는 템플릿일 때만 보충).

### 3-1. 스키마 — `<w:style type="table">` 의 자식 요소

OOXML 기준 ([ECMA-376 Part 1 §17.7.4.5](https://ecma-international.org/publications-and-standards/standards/ecma-376/)). 실제로 쓰이는 자식만 정리:

| 자식 요소 | 층위 | 역할 |
|---|---|---|
| `<w:name>` | 메타 | 스타일 이름 (Pandoc 표 매칭엔 안 쓰이지만 Word UI 표시용) |
| `<w:basedOn>` | 메타 | 부모 스타일 (체인 상속) |
| `<w:uiPriority>`, `<w:rsid>` | 메타 | 정렬·변경 추적 |
| **`<w:pPr>`** | **모든 셀의 기본 문단 속성** | 정렬·행간·들여쓰기 등 |
| **`<w:rPr>`** | **모든 셀의 기본 글자 속성** | 폰트·크기·색·굵기 |
| `<w:tblPr>` | 표 레벨 | 테두리·셀 margin·들여쓰기·tblLayout 등 |
| `<w:trPr>` | 모든 행의 기본 | 행 높이·헤더 반복 |
| `<w:tcPr>` | 모든 셀의 기본 | 셀 음영·세로 정렬·margin override |
| `<w:tblStylePr type="…">` | **조건부** | 특정 영역만 — 6종 타입 (아래) |

### 3-2. `<w:tblStylePr>` 의 조건부 타입

| `w:type` 값 | 적용 대상 |
|---|---|
| `firstRow` | 첫 번째 행 (예: 헤더) |
| `lastRow` | 마지막 행 (예: 합계) |
| `firstCol` | 첫 번째 열 |
| `lastCol` | 마지막 열 |
| `band1Horz` / `band2Horz` | 홀수/짝수 행 (행 밴딩) |
| `band1Vert` / `band2Vert` | 홀수/짝수 열 (열 밴딩) |
| `neCell` / `nwCell` / `seCell` / `swCell` | 네 모퉁이 셀 |

각 `<w:tblStylePr>` 은 다시 내부에 `<w:pPr>`·`<w:rPr>`·`<w:tblPr>`·`<w:trPr>`·`<w:tcPr>` 를 가질 수 있다. **즉 조건부 영역도 문단·글자·셀 속성을 따로 지정 가능**.

### 3-3. `<w:tblLook>` 과의 연동

테이블 본문(`<w:tbl>`) 의 `<w:tblPr>` 안에 있는 `<w:tblLook>` 이 이 조건부를 **켜고 끄는 스위치** 역할:

```xml
<w:tblLook w:val="04A0"
           w:firstRow="1"     ← firstRow tblStylePr 활성화
           w:firstColumn="1"  ← firstCol tblStylePr 활성화
           w:lastRow="0"
           w:lastColumn="0"
           w:noHBand="0" w:noVBand="1" />
```

스타일에 firstRow 정의가 있어도 `w:firstRow="0"` 이면 적용 안 됨. **tblLook 도 반드시 복제해야 한다.**

---

## 4. 계층 구조 (override 우선순위)

아래로 갈수록 더 구체적이고 이긴다:

```
1. 문서 기본값 (w:docDefaults in styles.xml)
2. 테이블 스타일 (w:style type="table")                    ← 이 문서의 주제
   ├─ 탑 레벨: pPr, rPr, tblPr, trPr, tcPr                  (모든 셀 기본값)
   └─ tblStylePr(firstRow, firstCol, …)                     (조건부 영역)
3. 표 본문의 tblPr / trPr / tcPr                            (개별 표·행·셀 override)
4. 셀 안 문단의 pPr
   └─ pStyle (다른 문단 스타일 참조)                         ← pandoc 이 Compact 로 박는 곳
5. 문단 인라인 pPr
6. 런의 rPr (bold, italic, 개별 글자 색 등)
```

**실무적 함의**:
- 4번 (`pStyle`) 이 박히면 2번 탑 레벨 pPr 이 **밀려남**
- Pandoc 은 모든 셀 문단에 `<w:pStyle w:val="Compact"/>` 를 박음 → 2번의 셀 기본 문단 서식이 먹지 않음
- 해결: 4번의 `pStyle` 제거 → 2번이 다시 살아남

---

## 5. 실제 예 — `reference_reg.docx` 의 `aa` 스타일

```xml
<w:style w:type="table" w:styleId="aa">
  <w:name w:val="Table Grid"/>
  <w:basedOn w:val="a1"/>                          ← Normal Table 기반
  <w:uiPriority w:val="39"/>

  <!-- 모든 셀의 기본 문단 속성 -->
  <w:pPr>
    <w:spacing w:after="0"/>                       ← 문단 뒤 간격 0
  </w:pPr>

  <!-- 표 전체 테두리 -->
  <w:tblPr>
    <w:tblBorders>
      <w:top    w:val="single" w:sz="18" w:color="000000"/>
      <w:left   w:val="single" w:sz="18" w:color="000000"/>
      <w:bottom w:val="single" w:sz="18" w:color="000000"/>
      <w:right  w:val="single" w:sz="18" w:color="000000"/>
      <w:insideH w:val="single" w:sz="18" w:color="000000"/>
      <w:insideV w:val="single" w:sz="18" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>

  <!-- 첫 행 조건부 -->
  <w:tblStylePr w:type="firstRow">
    <w:pPr><w:wordWrap/></w:pPr>                   ← 문단: 워드랩
    <w:rPr>                                        ← 글자: 폰트·굵게·크기·색
      <w:rFonts w:asciiTheme="majorHAnsi"
                w:eastAsiaTheme="majorEastAsia"
                w:hAnsiTheme="majorHAnsi"/>
      <w:b/><w:i w:val="0"/>
      <w:color w:val="000000" w:themeColor="text1"/>
      <w:sz w:val="24"/>
    </w:rPr>
    <w:tcPr>
      <w:shd w:val="clear" w:color="auto"
             w:fill="D9D9D9"/>                     ← 셀 음영
    </w:tcPr>
  </w:tblStylePr>

  <!-- 첫 열 조건부 -->
  <w:tblStylePr w:type="firstCol">
    <w:tblPr/>
    <w:tcPr>
      <w:shd w:val="clear" w:color="auto"
             w:fill="C1E4F5"/>                     ← 음영만, 글자는 Normal 상속
    </w:tcPr>
  </w:tblStylePr>
</w:style>
```

### 이 스타일이 정의한 것 / 안 한 것

| 영역 | 문단 | 글자 | 셀 음영 | 테두리 |
|---|---|---|---|---|
| **모든 셀 (기본)** | 문단 간격 0 | ❌ (없음, Normal 상속) | ❌ | 두꺼운 검정 |
| **첫 행** | wordWrap | 폰트·굵게·크기·검정 | D9D9D9 (회색) | 두꺼운 검정 |
| **첫 열** | 기본 상속 | Normal 상속 | C1E4F5 (연파랑) | 두꺼운 검정 |

→ **본문 셀 글씨체는 이 스타일이 정하지 않음**. Normal (문서 기본) 로 간다. 이게 reference_reg.docx 템플릿의 설계 선택.

---

## 6. "테이블 스타일에 정보가 다 있다" 는 말은 맞나?

**대체로 맞음, 조건부**. 정확히는:

| 상황 | 스타일 한 덩어리면 충분한가 |
|---|---|
| 템플릿 작성자가 스타일에 **전부** 담아 놓은 경우 | ✅ 충분 |
| 본문 셀 서식을 Normal 상속에 맡긴 경우 (위의 aa) | ✅ 충분 (Normal 도 함께 가져가면 됨) |
| 셀마다 **인라인 서식** (pStyle/rPr) 을 박아 놓은 경우 | ⚠️ 스타일만으로 부족, `table_sources.xml` 의 `<w:tc>` 에서 추가 추출 필요 |

현재 스킬 설계:
- `md2docx-extract-style` 이 **두 덩어리 모두 뽑음**:
  - `table_style.xml` — 스타일 계층 (앵커 + basedOn 체인, `<styles-excerpt>` 래퍼)
  - `table_sources.xml` — 실제 셀 XML 스냅샷 (보충용, 인라인 서식이 있을 때 대비)
- `md2docx` 가 사용할 때:
  - 스타일은 `styles.xml` 에 병합해서 **자동 적용**
  - pandoc 의 `pStyle="Compact"` 는 제거해서 스타일 계층이 먹도록 해제
  - (인라인 서식 케이스) sample 의 `<w:tc>` 속 pPr/rPr 을 추가 주입

---

## 7. 표 본문(`<w:tbl>`) 이 자체적으로 갖는 정보

스타일과 **별개로** 각 `<w:tbl>` 안에 박히는 요소들 (스타일로는 결정 못하는 것):

| 요소 | 의미 | 왜 스타일에 못 담나 |
|---|---|---|
| `<w:tblGrid>` | 각 열의 폭 | 표마다 열 개수가 다름 |
| `<w:tblLook>` | 조건부 스위치 | 표별로 끄고 켤 수 있어야 함 |
| `<w:tblW>` | 표 전체 폭 지정 방식 (dxa / pct / auto) | 표마다 다름 |
| `<w:gridSpan>` / `<w:vMerge>` | 셀 병합 정보 | 데이터 구조 |

→ 이들은 **표마다 고유**. 스타일에는 못 담음. `table_sources.xml` 이 유일한 재료.

---

## 8. 테스트 환경 / 참고

- 검증 Pandoc 버전: 3.9.0.2
- 입력 MD: 표 한 개 / 제목·본문·인용·리스트가 섞인 짧은 문서
- 레퍼런스 docx: `references/templates/style_template_*.docx` (rev1~rev5)
- [pandoc.md](pandoc.md) — Pandoc 일반 사용법 + 단락/문자 스타일의 `w:name` 매칭
- ECMA-376 Part 1 §17.7.4 — `<w:style>` 스키마 정식 정의
- ECMA-376 Part 1 §17.4.12 — `<w:tblStyle>` 런타임 매칭
