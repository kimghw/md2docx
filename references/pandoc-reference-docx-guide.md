# Pandoc reference.docx 가이드 (한글 Word 템플릿 사용)

Pandoc 변환 시 `--reference-doc` 옵션으로 Word 스타일을 적용할 때의 핵심 원리와 주의사항.

---

## 핵심 원리: Pandoc은 `w:name`으로 스타일을 매칭한다

많은 사람이 오해하는 부분:
- ❌ 오해: Pandoc은 정해진 style ID(`Heading1`, `Normal` 등)가 있어야만 스타일을 적용한다
- ✅ 사실: Pandoc은 **스타일의 `w:name` 속성**을 보고 매칭한다

즉, 한글 Word가 자동 생성한 style ID(`1`, `a`, `a3` 등)여도 `w:name`이 `heading 1`, `Normal`, `Title` 등으로 되어 있으면 Pandoc이 정상 인식한다.

### 예시 (한글 Word에서 만든 docx)
```xml
<w:style w:type="paragraph" w:styleId="1">
  <w:name w:val="heading 1"/>   <!-- Pandoc은 이 name을 봄 -->
  ...
</w:style>
```
이 경우 Pandoc은 `w:pStyle w:val="1"`을 출력 docx에 사용한다.

---

## ⚠️ 단, `w:name`은 **반드시 Pandoc이 찾는 이름과 일치**해야 한다

Pandoc은 특정 이름을 기준으로 스타일을 찾는다. 이름이 다르면 아무리 서식이 잘 정의되어 있어도 적용되지 않는다.

### Pandoc이 찾는 `w:name` 목록 (필수)

| Markdown 요소 | Pandoc이 찾는 w:name |
|--------------|----------------------|
| 본문 | `Normal` |
| `#` (H1) | `heading 1` |
| `##` (H2) | `heading 2` |
| `###` ~ `######` | `heading 3` ~ `heading 6` |
| 제목 (YAML `title:`) | `Title` |
| 부제 (YAML `subtitle:`) | `Subtitle` |
| `>` 인용 | `Quote` |
| 블록 인용 | `Block Text` |
| 코드 블록 | `Source Code` 또는 `Verbatim Char` |
| 인라인 코드 | `Verbatim Char` |
| 링크 | `Hyperlink` |
| 이미지 캡션 | `Image Caption` |
| 표 | `Table` |
| 표 캡션 | `Table Caption` |
| 각주 참조 | `Footnote Reference` |
| 각주 본문 | `Footnote Text` |
| 첫 문단 | `First Paragraph` |
| 간격 압축 (리스트 내) | `Compact` |
| 이름 정의 | `Definition Term` |
| 내용 정의 | `Definition` |
| 저자 | `Author` |
| 날짜 | `Date` |
| 참고문헌 | `Bibliography` |
| 초록 | `Abstract` |
| 캡션 (공통) | `Caption` |
| 그림 | `Figure` |11

### 한글 Word 이름 → Pandoc 이름 변경이 필요한 경우

한글 Word는 기본 스타일 이름을 영문으로 쓰지만, 사용자가 한글로 변경했거나 커스텀 스타일을 만들었다면 매칭 실패.

| 한글 이름 (Word UI) | Pandoc이 원하는 w:name |
|---------------------|------------------------|
| 표준 | `Normal` |
| 제목 1 ~ 제목 6 | `heading 1` ~ `heading 6` |
| 제목 (Title) | `Title` |
| 부제 | `Subtitle` |
| 인용 | `Quote` |
| 강한 인용 | `Intense Quote` |
| 목록 단락 | `List Paragraph` |
| 머리글 | `header` |
| 바닥글 | `footer` |

**Word UI에 보이는 이름**과 **내부 `w:name`**은 다를 수 있다.
- Word UI: "제목 1" (한국어)
- 내부 w:name: "heading 1" (영문, Word가 자동 관리)

한글 Word의 기본 제공 스타일은 내부적으로 영문 w:name을 유지하므로 별도 작업 불필요.
커스텀 스타일을 추가했다면 w:name을 Pandoc 규격에 맞춰야 한다.

### 확인 방법
```bash
unzip -p your-template.docx word/styles.xml | grep -oE '<w:name w:val="[^"]+"' | sort -u
```
이 출력에 위 Pandoc 이름이 있으면 OK.

---

## ⚠️ 표(Table) 스타일은 특별히 주의 — `w:name="Table"` 필요

Pandoc은 Markdown 표를 변환할 때 출력 docx에 다음과 같이 표시한다:
```xml
<w:tblStyle w:val="Table"/>
```

### 문제 상황
한글 Word 템플릿은 보통 `Normal Table`, `Table Grid`, `Plain Table 1` 등의 스타일만 있고, **`w:name="Table"`** 스타일은 없다.
이 경우 Pandoc이 `Table` 스타일을 못 찾아 **사용자의 표 서식이 전혀 적용되지 않는다**.

### 올바른 해결 절차

표 스타일은 임의로 아무거나 갖다 쓰면 안 된다. **템플릿 안에 이미 있는 표가 실제로 어떤 스타일을 쓰는지** 먼저 확인하고, 그 스타일을 기반으로 `Table` 이름의 스타일을 만들어야 한다.

#### 1단계: 템플릿의 표가 사용하는 스타일 ID 확인
```bash
unzip -p your-template.docx word/document.xml | grep -oE 'w:tblStyle w:val="[^"]+"' | sort | uniq -c
```
출력 예:
```
      3 w:tblStyle w:val="aa"        <- 주로 이걸 씀
      1 w:tblStyle w:val="10"
```
→ 템플릿의 표들이 실제로 사용 중인 style ID를 파악

#### 2단계: 해당 style ID가 어떤 이름과 서식을 가진지 확인
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

#### 3단계: 그 스타일을 복제해서 이름만 `Table`로 변경하여 추가
`</w:styles>` 바로 앞에 삽입:
```xml
<w:style w:type="table" w:styleId="Table">
  <w:name w:val="Table"/>
  <w:basedOn w:val="aa"/>   <!-- ⭐ 1단계에서 찾은 실제 사용 style ID -->
</w:style>
```
- `basedOn`으로 기존 스타일을 상속 → 테두리, 배경, 폰트 등 모든 서식 자동 적용
- 일일이 `tblBorders` 복사할 필요 없음
- 템플릿 여러 곳에서 다른 표 스타일을 쓴다면, 가장 많이 쓰이는 것을 `basedOn`에 지정

### 왜 이렇게 해야 하는가?
- ❌ 잘못된 방법: "한글 Word는 보통 Table Grid를 쓰니까 `aa` 갖다 쓰자" (추측)
- ✅ 올바른 방법: 템플릿 실제 내용을 보고 어떤 표 스타일이 쓰이는지 확인 후 그걸 기반으로 제작

템플릿마다 쓰는 표 스타일이 다르다. 회사 양식은 커스텀 `aa`를 쓸 수도, 기본 `Plain Table 1`을 쓸 수도, 아예 커스텀 스타일(`CompanyTable` 등)을 만들어 쓸 수도 있다. 추측하지 말고 실제 확인.

### Word UI에서 하는 경우
1. 템플릿 Word로 열기
2. 기존 표 하나 클릭 → 디자인 탭에서 현재 적용된 스타일 이름 확인
3. 스타일 갤러리에서 그 스타일 **우클릭 → 복제**
4. 복제본 이름을 정확히 **`Table`**로 변경
5. 저장

### 관련 주요 표 이름
| Pandoc 요소 | w:name |
|------------|--------|
| 표 본체 | `Table` (필수, 자주 누락) |
| 표 캡션 | `Table Caption` |
| 기본 테이블 | `Normal Table` |
| 테이블 그리드 | `Table Grid` |

---

## 권장 워크플로우

### 1. Word 문서를 reference.docx로 사용하기 (가장 간단)
```bash
# 사용자가 만든 Word 템플릿을 그대로 reference.docx로 사용
cp my-korean-template.docx reference.docx

# 변환
pandoc input.md --reference-doc=reference.docx -o output.docx
```
- Word로 열고 스타일 갤러리에서 `제목 1`, `제목 2`, `표준` 등을 **우클릭 → 수정**으로 커스터마이징
- 저장 후 바로 reference.docx로 사용 가능
- Style ID를 바꿀 필요 없음

### 2. Pandoc 기본 reference.docx에서 시작
```bash
pandoc -o reference.docx --print-default-data-file reference.docx
```
- Word로 열어 `Heading 1`, `Normal` 등 수정
- 영문 style ID이므로 한글 Word 사용자에겐 직관적이지 않음

---

## 하지 말아야 할 것

### ❌ Pandoc으로 docx를 "정규화"하지 말 것
```bash
# 이렇게 하면 사용자 템플릿의 서식이 Pandoc 기본값으로 덮어써진다
pandoc user-template.docx -o normalized.docx
```
`--reference-doc` 없이 docx → docx 변환을 하면 Pandoc이 자체 기본 스타일로 출력한다. 원본의 Heading 색상, 폰트 등이 사라짐.

### ❌ styles.xml을 직접 sed로 편집하지 말 것
- style ID를 `1` → `Heading1`로 rename해도 변화 없음 (이미 w:name으로 매칭됨)
- document.xml, numbering.xml, settings.xml의 레퍼런스가 불일치하면 Word가 "손상된 파일"로 판정

### ❌ PowerShell의 `Compress-Archive`로 docx 재패키징 금지
Windows 경로 구분자(`\`)가 들어가 Word가 열지 못함. 반드시 forward slash(`/`) 사용:
```powershell
[System.IO.Compression.ZipFile]::Open($dst, 'Create')
# CreateEntry에 relPath.Replace('\\', '/') 적용
```

---

## 한글 Word 템플릿의 style ID 매핑 참고

한글 Word는 자동 생성 ID를 사용한다. 수동 매핑이 필요할 때 참고:

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

다시 강조: **Pandoc 사용에는 ID 매핑이 불필요**. w:name만 맞으면 됨.

---

## Windows에서 Pandoc 설치 후 PATH 문제

Winget으로 설치한 Pandoc 경로:
```
C:\Users\<username>\AppData\Local\Microsoft\WinGet\Packages\JohnMacFarlane.Pandoc_*\pandoc-<version>\pandoc.exe
```

### PATH가 자동 등록 안 될 때
1. PowerShell 재시작 (가장 먼저 시도)
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
# 기본 변환
pandoc input.md -o output.docx

# 커스텀 스타일 적용
pandoc input.md --reference-doc=reference.docx -o output.docx

# 여러 옵션 조합
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
3. `pandoc input.md --reference-doc=reference.docx -o output.docx`

끝. style ID 변경, XML 수술, 정규화 같은 복잡한 작업 필요 없음.
