# OOXML 테이블 스타일 구조 — `w:style type="table"` 완전 해부

Word 문서의 표(Table) 서식은 **하나의 `<w:style>` 블록** 에 전부 담을 수 있다. 즉 테이블 레벨·셀 레벨·셀 안 문단 레벨·셀 안 글자 레벨 속성 **전부** 가 이 하나의 스타일 안에서 통제 가능하다.

이 문서의 결론:

> **제대로 정의된 테이블 스타일 한 덩어리만 있으면 표 서식 정보는 다 있다.**
> 추출도 이 스타일 하나만 올바르게 떠내면 재사용 가능. 셀 단위 XML 을 별도로 긁어올 필요는 **원칙적으로 없다** (스타일에 누락이 있는 템플릿일 때만 보충).

---

## 1. 스키마 — `<w:style type="table">` 가 가질 수 있는 자식 요소

OOXML 기준 ([ECMA-376 Part 1 §17.7.4.5](https://ecma-international.org/publications-and-standards/standards/ecma-376/)). 실제로 쓰이는 자식만 정리:

| 자식 요소 | 층위 | 역할 |
|---|---|---|
| `<w:name>` | 메타 | 스타일 이름 (Pandoc 매칭에 쓰이진 않지만 Word UI 표시용) |
| `<w:basedOn>` | 메타 | 부모 스타일 (체인 상속) |
| `<w:uiPriority>`, `<w:rsid>` | 메타 | 정렬·변경 추적 |
| **`<w:pPr>`** | **모든 셀의 기본 문단 속성** | 정렬·행간·들여쓰기 등 |
| **`<w:rPr>`** | **모든 셀의 기본 글자 속성** | 폰트·크기·색·굵기 |
| `<w:tblPr>` | 표 레벨 | 테두리·셀 margin·들여쓰기·tblLayout 등 |
| `<w:trPr>` | 모든 행의 기본 | 행 높이·헤더 반복 |
| `<w:tcPr>` | 모든 셀의 기본 | 셀 음영·세로 정렬·margin override |
| `<w:tblStylePr type="…">` | **조건부** | 특정 영역만 — 6종 타입 (아래) |

### 1-1. `<w:tblStylePr>` 의 6가지 타입

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

### 1-2. `<w:tblLook>` 과의 연동

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

## 2. 계층 구조 (override 우선순위)

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

## 3. 실제 예 — `reference_reg.docx` 의 `aa` 스타일

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

## 4. 그럼 "테이블 스타일에 정보가 다 있다" 는 말은 맞나?

**대체로 맞음, 조건부**. 정확히는:

| 상황 | 스타일 한 덩어리면 충분한가 |
|---|---|
| 템플릿 작성자가 스타일에 **전부** 담아 놓은 경우 | ✅ 충분 |
| 본문 셀 서식을 Normal 상속에 맡긴 경우 (위의 aa) | ✅ 충분 (Normal 도 함께 가져가면 됨) |
| 셀마다 **인라인 서식** (pStyle/rPr) 을 박아 놓은 경우 | ⚠️ 스타일만으로 부족, `sample_table.xml` 의 `<w:tc>` 에서 추가 추출 필요 |

현재 스킬 설계:
- `md2docx-extract-style` 이 **두 덩어리 모두 뽑음**:
  - `table_style_bundle/styles_excerpt.xml` — 스타일 계층 (주된 정보)
  - `sample_table.xml` — 실제 셀 XML 스냅샷 (보충용, 인라인 서식이 있을 때 대비)
- `md2docx` 가 사용할 때:
  - 스타일은 `styles.xml` 에 병합해서 **자동 적용**
  - pandoc 의 `pStyle="Compact"` 는 제거해서 스타일 계층이 먹도록 해제
  - (인라인 서식 케이스) sample 의 `<w:tc>` 속 pPr/rPr 을 추가 주입

---

## 5. 부수 요소 — 표 본문(`<w:tbl>`) 이 자체적으로 갖는 정보

스타일과 **별개로** 각 `<w:tbl>` 안에 박히는 요소들 (스타일로는 결정 못하는 것):

| 요소 | 의미 | 왜 스타일에 못 담나 |
|---|---|---|
| `<w:tblGrid>` | 각 열의 폭 | 표마다 열 개수가 다름 |
| `<w:tblLook>` | 조건부 스위치 | 표별로 끄고 켤 수 있어야 함 |
| `<w:tblW>` | 표 전체 폭 지정 방식 (dxa / pct / auto) | 표마다 다름 |
| `<w:gridSpan>` / `<w:vMerge>` | 셀 병합 정보 | 데이터 구조 |

→ 이들은 **표마다 고유**. 스타일에는 못 담음. `sample_table.xml` 이 유일한 재료.

---

## 6. 참고

- [guide_table_style.md](guide_table_style.md) — Pandoc 매칭 규칙 실측
- [pandoc-reference-docx-guide.md](pandoc-reference-docx-guide.md) — reference.docx 운용
- ECMA-376 Part 1 §17.7.4 — `<w:style>` 스키마 정식 정의
- ECMA-376 Part 1 §17.4.12 — `<w:tblStyle>` 런타임 매칭
