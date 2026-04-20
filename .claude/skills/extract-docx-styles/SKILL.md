---
name: extract-docx-styles
description: Pandoc `--reference-doc`용 Word(.docx) 템플릿의 스타일을 분석·검증·수정하여 Markdown→DOCX 변환 시 사용자 서식이 올바르게 적용되도록 돕는다. 한글 Word 템플릿의 style ID / w:name 매칭, 표(Table) 스타일 누락, reference.docx 준비가 필요한 상황에서 호출된다.
---

# extract-docx-styles — Pandoc reference.docx 스타일 추출·검증 가이드

Pandoc 변환 시 `--reference-doc` 옵션으로 Word 템플릿을 사용할 때, 사용자 서식이 올바르게 적용되도록 **템플릿의 스타일을 분석하고 필요한 경우 수정**한다.

---

## 핵심 원리

**Pandoc은 `w:styleId`가 아니라 `w:name`으로 스타일을 매칭한다.**

- ❌ 오해: Pandoc은 정해진 style ID(`Heading1`, `Normal`)가 필요하다
- ✅ 사실: Pandoc은 스타일의 `w:name` 속성을 본다

즉, 한글 Word가 자동 생성한 style ID(`1`, `a`, `a3` 등)여도 `w:name`이 `heading 1`, `Normal`, `Title` 등이면 정상 동작한다.

```xml
<w:style w:type="paragraph" w:styleId="1">
  <w:name w:val="heading 1"/>   <!-- Pandoc은 이 name을 봄 -->
</w:style>
```

---

## Pandoc이 찾는 `w:name` 목록

### 문단/텍스트 스타일

| Markdown 요소 | 필요한 w:name |
|--------------|--------------|
| 본문 | `Normal` |
| `#` ~ `######` | `heading 1` ~ `heading 6` |
| YAML `title:` | `Title` |
| YAML `subtitle:` | `Subtitle` |
| `>` 인용 | `Quote` |
| 블록 인용 | `Block Text` |
| 코드 블록 | `Source Code` 또는 `Verbatim Char` |
| 인라인 코드 | `Verbatim Char` |
| 링크 | `Hyperlink` |
| 이미지 캡션 | `Image Caption` |
| 표 캡션 | `Table Caption` |
| 각주 참조 | `Footnote Reference` |
| 각주 본문 | `Footnote Text` |
| 첫 문단 | `First Paragraph` |
| 리스트 간격 압축 | `Compact` |
| 이름 정의 | `Definition Term` |
| 내용 정의 | `Definition` |
| 저자 | `Author` |
| 날짜 | `Date` |
| 참고문헌 | `Bibliography` |
| 초록 | `Abstract` |
| 공통 캡션 | `Caption` |
| 그림 | `Figure` |

### 한글 Word UI 이름 → Pandoc w:name

| 한글 Word UI | Pandoc이 원하는 w:name |
|-------------|----------------------|
| 표준 | `Normal` |
| 제목 1 ~ 제목 6 | `heading 1` ~ `heading 6` |
| 제목 | `Title` |
| 부제 | `Subtitle` |
| 인용 | `Quote` |
| 강한 인용 | `Intense Quote` |
| 목록 단락 | `List Paragraph` |
| 머리글 | `header` |
| 바닥글 | `footer` |

> Word UI에 보이는 이름과 내부 `w:name`은 다를 수 있다. 한글 Word 기본 스타일은 내부적으로 영문 w:name을 유지하므로 보통 추가 작업 불필요.

---

## 작업 절차

### 1. 실제 참조 중인 스타일 추출

`styles.xml`에 정의만 되어 있고 본문에서 쓰지 않는 스타일은 매핑 검증에 필요 없다. **`document.xml`에서 실제로 참조되는 style ID만** 추출하여 (a) 매핑 TSV와 (b) 해당 ID의 실제 `<w:style>` XML 블록을 함께 저장한다. 매핑만 있으면 "어떤 이름인지"는 알지만 "실제 서식이 어떤지"는 모르므로 둘 다 필요하다.

출력 경로 (모두 `references/templates/extracted/` 아래):
- `<basename>.used_mapping.tsv` — 컬럼: `reference<TAB>styleId<TAB>type<TAB>w:name<TAB>count`
- `<basename>.used_styles.xml` — 참조된 styleId의 `<w:style>` 블록 모음 (테두리·폰트·basedOn 등 실 서식 포함)

```bash
TPL=your-template.docx
BASENAME=$(basename "$TPL" .docx)
OUTDIR=references/templates/extracted
MAP=$OUTDIR/$BASENAME.used_mapping.tsv
OUT_XML=$OUTDIR/$BASENAME.used_styles.xml
mkdir -p "$OUTDIR"

# --- (a) 매핑 TSV ---

# 1) document.xml에서 참조 중인 style ID와 건수
unzip -p "$TPL" word/document.xml \
  | grep -oE 'w:(pStyle|rStyle|tblStyle) w:val="[^"]+"' \
  | sed -E 's/w:(pStyle|rStyle|tblStyle) w:val="([^"]+)"/\1\t\2/' \
  | sort | uniq -c | awk '{print $2"\t"$3"\t"$1}' > /tmp/used_ids.tsv

# 2) 참조된 ID를 styles.xml과 조인하여 type·w:name 붙이기
unzip -p "$TPL" word/styles.xml \
  | grep -oE '<w:style [^>]*w:styleId="[^"]+"[^>]*>[^<]*<w:name w:val="[^"]+"' \
  | sed -E 's/.*w:type="([^"]+)".*w:styleId="([^"]+)".*<w:name w:val="([^"]+)".*/\2\t\1\t\3/' \
  | sort > /tmp/defined.tsv

printf 'reference\tstyleId\ttype\tw:name\tcount\n' > "$MAP"
awk -F'\t' 'NR==FNR{t[$1]=$2; n[$1]=$3; next} {printf "%s\t%s\t%s\t%s\t%s\n", $1, $2, t[$2], n[$2], $3}' \
  /tmp/defined.tsv /tmp/used_ids.tsv >> "$MAP"

# --- (b) 실제 스타일 XML 블록 ---
# 주의: grep -oP 는 UTF-8 로케일 필요 (한글 스타일 이름 때문). LC_ALL=C.UTF-8 필수.
STYLES=$(LC_ALL=C.UTF-8 unzip -p "$TPL" word/styles.xml)

{
  echo '<?xml version="1.0" encoding="UTF-8"?>'
  echo '<styles-extracted>'
  awk -F'\t' 'NR>1 {print $2}' "$MAP" | sort -u | while read -r ID; do
    echo "<!-- styleId=$ID -->"
    echo "$STYLES" | LC_ALL=C.UTF-8 grep -oP "<w:style [^>]*w:styleId=\"$ID\"[^>]*>.*?</w:style>"
    echo
  done
  echo '</styles-extracted>'
} > "$OUT_XML"

cat "$MAP"
```

두 파일을 통해 "어떤 스타일이 쓰였고(TSV), 그 스타일의 실제 서식이 무엇인지(XML)"를 모두 확보할 수 있다. `basedOn`이 가리키는 상위 스타일도 필요하면 TSV의 ID 뒤에 이어 붙여 같은 방식으로 추출한다. 정의된 전체 목록은 저장하지 않는다.

### 2. 템플릿의 w:name 목록 확인

```bash
unzip -p your-template.docx word/styles.xml | grep -oE '<w:name w:val="[^"]+"' | sort -u
```

출력에서 위 표의 Pandoc 필수 이름들이 존재하는지 확인 (저장은 하지 않음).

### 3. 누락된 w:name이 있으면 추가/수정

- 기존 스타일의 `w:name`을 Pandoc 규격으로 변경하거나
- 새 스타일을 추가할 때 `basedOn`으로 기존 스타일을 상속시키면 서식이 자동 적용됨

---

## ⚠️ 표(Table) 스타일 — 가장 흔한 누락 지점

Pandoc은 Markdown 표를 출력 docx에 다음처럼 기록한다:

```xml
<w:tblStyle w:val="Table"/>
```

→ 템플릿에 **`w:name="Table"`인 스타일이 없으면 표 서식이 전혀 적용되지 않는다**.

한글 Word 템플릿은 보통 `Normal Table`, `Table Grid`, 커스텀 ID(`aa` 등)만 있고 `Table`은 없다.

### 올바른 해결 절차

임의로 아무 표 스타일을 갖다 쓰지 말 것. **템플릿 안의 기존 표가 실제로 어떤 스타일을 쓰는지** 확인하고 그 스타일을 기반으로 `Table` 이름을 만든다.

#### 1단계: 템플릿 표가 실제로 사용하는 style ID 확인

```bash
unzip -p your-template.docx word/document.xml | grep -oE 'w:tblStyle w:val="[^"]+"' | sort | uniq -c
```

출력 예:
```
      3 w:tblStyle w:val="aa"        <- 주로 이걸 씀
      1 w:tblStyle w:val="10"
```

#### 2단계: 해당 style의 이름·서식 확인

```bash
unzip -p your-template.docx word/styles.xml | grep -oE '<w:style [^>]*w:styleId="aa"[^>]*>.{0,500}'
```

#### 3단계: 그 스타일을 복제하여 이름만 `Table`로

`</w:styles>` 바로 앞에 삽입:

```xml
<w:style w:type="table" w:styleId="Table">
  <w:name w:val="Table"/>
  <w:basedOn w:val="aa"/>   <!-- 1단계에서 찾은 실제 사용 style ID -->
</w:style>
```

- `basedOn`으로 기존 스타일을 상속 → 테두리·배경·폰트 자동 적용
- `tblBorders` 같은 서식을 일일이 복사할 필요 없음

#### Word UI에서 하는 경우

1. 템플릿을 Word로 열기
2. 기존 표 클릭 → 디자인 탭에서 현재 적용된 스타일 확인
3. 스타일 갤러리에서 우클릭 → **복제**
4. 복제본 이름을 정확히 **`Table`**로 변경 후 저장

### 관련 표 이름

| Pandoc 요소 | w:name |
|-----------|--------|
| 표 본체 | `Table` (필수, 자주 누락) |
| 표 캡션 | `Table Caption` |
| 기본 테이블 | `Normal Table` |
| 테이블 그리드 | `Table Grid` |

---

## 권장 워크플로우

### A. 사용자 Word 템플릿 → reference.docx (권장)

```bash
cp my-korean-template.docx reference.docx
pandoc input.md --reference-doc=reference.docx -o output.docx
```

- Word에서 스타일 갤러리 → `제목 1`, `제목 2`, `표준` 우클릭 → 수정으로 커스터마이즈
- Style ID를 변경할 필요 없음

### B. Pandoc 기본 reference.docx에서 시작

```bash
pandoc -o reference.docx --print-default-data-file reference.docx
```

- Word로 열어 `Heading 1`, `Normal` 등 편집
- 영문 style ID이므로 한글 Word 사용자에겐 덜 직관적

---

## 하지 말아야 할 것

### ❌ Pandoc으로 docx를 "정규화"하지 말 것

```bash
# 이렇게 하면 사용자 템플릿 서식이 Pandoc 기본값으로 덮어써짐
pandoc user-template.docx -o normalized.docx
```

`--reference-doc` 없이 docx → docx 하면 Heading 색상·폰트 등이 모두 사라진다.

### ❌ styles.xml을 sed로 직접 편집 금지

- style ID를 `1` → `Heading1`로 rename해도 변화 없음 (w:name으로 매칭되므로)
- `document.xml`, `numbering.xml`, `settings.xml` 참조가 어긋나면 Word가 "손상된 파일"로 판정

### ❌ PowerShell `Compress-Archive`로 docx 재패키징 금지

Windows 경로 구분자(`\`)가 들어가 Word가 열지 못함. 반드시 forward slash(`/`) 사용:

```powershell
[System.IO.Compression.ZipFile]::Open($dst, 'Create')
# CreateEntry에 relPath.Replace('\\', '/') 적용
```

---

## 한글 Word style ID 매핑 참고

수동 매핑이 필요할 때만 참고(일반적으로 **불필요**):

| 한글 Word style ID | w:name | Pandoc 표준 |
|-------------------|--------|-------------|
| `a` | Normal | Normal |
| `1` ~ `9` | heading 1~9 | Heading1~9 |
| `1Char` ~ `9Char` | 제목 1 Char~ | Heading1Char~ |
| `a3` | Title | Title |
| `a4` | Subtitle | Subtitle |
| `a5` | Quote | Quote |
| `a6` | List Paragraph | ListParagraph |
| `Char` | 제목 Char | TitleChar |
| `Char0` | 부제 Char | SubtitleChar |
| `Char1` | 인용 Char | QuoteChar |
| `ab`, `ac` | header, footer | Header, Footer |
| `aa` | Table Grid | TableGrid |

**Pandoc 사용 자체에는 ID 매핑이 불필요. w:name만 맞으면 됨.**

---

## Windows Pandoc PATH 문제

Winget 설치 경로:
```
C:\Users\<username>\AppData\Local\Microsoft\WinGet\Packages\JohnMacFarlane.Pandoc_*\pandoc-<version>\pandoc.exe
```

### PATH 미등록 시

1. PowerShell 재시작
2. 현재 세션에 수동 추가:
   ```powershell
   $env:Path += ";C:\Users\<username>\AppData\Local\Microsoft\WinGet\Packages\JohnMacFarlane.Pandoc_*\pandoc-<version>"
   ```
3. 영구 등록:
   ```powershell
   [Environment]::SetEnvironmentVariable("Path",
     [Environment]::GetEnvironmentVariable("Path", "User") + ";<pandoc_경로>",
     "User")
   ```

---

## 실전 변환 패턴

```powershell
# 기본
pandoc input.md -o output.docx

# 커스텀 스타일
pandoc input.md --reference-doc=reference.docx -o output.docx

# 옵션 조합
pandoc input.md `
  --reference-doc=reference.docx `
  --toc --toc-depth=3 `
  -N `
  -o output.docx
```

---

## 요약: 가장 효율적인 경로

1. Word에서 원하는 스타일이 적용된 빈 문서 작성 (또는 기존 양식 활용)
2. 그 파일을 `reference.docx`로 저장
3. **실제 참조 중인 스타일 추출** (`used_mapping.tsv`) — 수정 전/후 비교 기준
4. 필수 `w:name` 점검 (특히 `Table`)
5. `pandoc input.md --reference-doc=reference.docx -o output.docx`

style ID 변경·XML 수술·정규화 같은 복잡한 작업 필요 없음.
