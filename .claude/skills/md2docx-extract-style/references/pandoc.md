# Pandoc 가이드

Pandoc 일반 사용법과 `--reference-doc` 운용을 한 파일에 정리.
표(Table) 스타일 관련은 [table-style.md](table-style.md) 참고.

---

## 1. 설치

### Windows
```bash
# Chocolatey
choco install pandoc

# Winget
winget install --id JohnMacFarlane.Pandoc

# 또는 공식 인스톨러: https://pandoc.org/installing.html
```

### macOS
```bash
brew install pandoc
```

### Linux
```bash
sudo apt-get install pandoc          # Debian/Ubuntu
sudo dnf install pandoc              # Fedora
```

### 설치 확인
```bash
pandoc --version
```

### Windows PATH 문제

Winget 설치 경로:
```
C:\Users\<username>\AppData\Local\Microsoft\WinGet\Packages\JohnMacFarlane.Pandoc_*\pandoc-<version>\pandoc.exe
```

PATH 자동 등록이 안 될 때:
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

## 2. 기본 사용법

### 기본 문법
```bash
pandoc [옵션] [입력파일] -o [출력파일]
```

### 예시
```bash
pandoc input.md   -o output.docx   # Markdown → Word
pandoc input.md   -o output.pdf    # Markdown → PDF
pandoc input.md   -o output.html   # Markdown → HTML
pandoc input.docx -o output.md     # Word → Markdown
pandoc input.html -o output.md     # HTML → Markdown
```

### 입출력 포맷 명시
```bash
pandoc -f markdown -t docx input.md -o output.docx
# -f, --from : 입력 포맷
# -t, --to   : 출력 포맷
```

---

## 3. 주요 옵션

| 옵션 | 설명 |
|------|------|
| `-o FILE` | 출력 파일 지정 |
| `-f FORMAT` | 입력 포맷 (from) |
| `-t FORMAT` | 출력 포맷 (to) |
| `-s` | standalone (완전한 문서 생성) |
| `--toc` | 목차 자동 생성 |
| `--toc-depth=N` | 목차 깊이 지정 |
| `-N` | 섹션 자동 번호 |
| `-c FILE.css` | CSS 파일 적용 (HTML) |
| `--template=FILE` | 사용자 템플릿 적용 |
| `--reference-doc=FILE` | docx/odt 참조 스타일 |
| `--metadata KEY=VAL` | 메타데이터 설정 |
| `-V KEY=VAL` | 변수 설정 |
| `--highlight-style=STYLE` | 코드 하이라이팅 스타일 |
| `--pdf-engine=ENGINE` | PDF 엔진 지정 |
| `--resource-path=PATH` | 리소스 경로 추가 |
| `--extract-media=DIR` | 미디어 파일 추출 |

---

## 4. 지원 포맷

### 입력 포맷 (주요)
- markdown, commonmark, gfm (GitHub Flavored Markdown)
- html, latex, docx, odt, epub
- rst (reStructuredText), textile, mediawiki
- ipynb (Jupyter Notebook), org (Org-mode)

### 출력 포맷 (주요)
- 위 입력 포맷 전부 +
- pdf, pptx, beamer (프레젠테이션)
- man (manpage), docbook, jats
- revealjs, slidy, slideous (웹 프레젠테이션)

### 포맷 목록 확인
```bash
pandoc --list-input-formats
pandoc --list-output-formats
```

---

## 5. reference.docx — Word 스타일 적용의 핵심 원리

### 핵심: Pandoc은 `w:name` 으로 단락/문자 스타일을 매칭한다

많은 사람이 오해하는 부분:
- ❌ 오해: Pandoc은 정해진 style ID(`Heading1`, `Normal` 등)가 있어야만 스타일을 적용한다
- ✅ 사실: Pandoc은 **스타일의 `w:name` 속성**을 보고 매칭한다

즉, 한글 Word가 자동 생성한 style ID(`1`, `a`, `a3` 등)여도 `w:name`이 `heading 1`, `Normal`, `Title` 등으로 되어 있으면 Pandoc이 정상 인식한다.

```xml
<w:style w:type="paragraph" w:styleId="1">
  <w:name w:val="heading 1"/>   <!-- Pandoc은 이 name을 봄 -->
  ...
</w:style>
```
이 경우 Pandoc은 `w:pStyle w:val="1"` 을 출력 docx에 사용한다.

> ⚠️ 표(Table) 스타일은 예외다. Pandoc은 표에 대해서만 **styleId="Table" 을 리터럴로 하드코드** 한다. 자세한 내용과 검증 결과는 [table-style.md](table-style.md).

### Pandoc이 찾는 `w:name` 목록 (단락/문자 스타일)

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
| 그림 | `Figure` |

### 한글 Word 이름 → Pandoc 이름 매핑

한글 Word는 기본 스타일 이름을 영문으로 쓰지만, 사용자가 한글로 변경했거나 커스텀 스타일을 만들었다면 매칭 실패할 수 있다.

| 한글 이름 (Word UI) | Pandoc 이 원하는 w:name |
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

**Word UI 에 보이는 이름**과 **내부 `w:name`** 은 다를 수 있다.
- Word UI: "제목 1" (한국어)
- 내부 w:name: "heading 1" (영문, Word 가 자동 관리)

한글 Word 의 기본 제공 스타일은 내부적으로 영문 w:name 을 유지하므로 별도 작업 불필요.
커스텀 스타일을 추가했다면 w:name 을 Pandoc 규격에 맞춰야 한다.

### 확인 방법
```bash
unzip -p your-template.docx word/styles.xml | grep -oE '<w:name w:val="[^"]+"' | sort -u
```
이 출력에 위 Pandoc 이름이 있으면 OK.

### 한글 Word 템플릿의 style ID 참고

한글 Word 는 자동 생성 ID 를 사용한다 (참고용):

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

다시 강조: **Pandoc 사용에 ID 매핑은 불필요**. w:name 만 맞으면 된다.

---

## 6. 권장 워크플로우

### 1) Word 문서를 그대로 reference.docx 로 사용 (가장 간단)
```bash
cp my-korean-template.docx reference.docx
pandoc input.md --reference-doc=reference.docx -o output.docx
```
- Word 로 열고 스타일 갤러리에서 `제목 1`, `제목 2`, `표준` 등을 **우클릭 → 수정**으로 커스터마이징
- 저장 후 바로 reference.docx 로 사용 가능
- Style ID 를 바꿀 필요 없음

### 2) Pandoc 기본 reference.docx 에서 시작
```bash
pandoc -o reference.docx --print-default-data-file reference.docx
```
- Word 로 열어 `Heading 1`, `Normal` 등 수정
- 영문 style ID 이므로 한글 Word 사용자에겐 직관적이지 않음

---

## 7. 하지 말아야 할 것

### ❌ Pandoc 으로 docx 를 "정규화"하지 말 것
```bash
# 이렇게 하면 사용자 템플릿의 서식이 Pandoc 기본값으로 덮어써진다
pandoc user-template.docx -o normalized.docx
```
`--reference-doc` 없이 docx → docx 변환을 하면 Pandoc 이 자체 기본 스타일로 출력한다. 원본의 Heading 색상, 폰트 등이 사라진다.

### ❌ styles.xml 을 직접 sed 로 편집하지 말 것
- style ID 를 `1` → `Heading1` 로 rename 해도 변화 없음 (이미 w:name 으로 매칭됨)
- document.xml, numbering.xml, settings.xml 의 레퍼런스가 불일치하면 Word 가 "손상된 파일" 로 판정

### ❌ PowerShell 의 `Compress-Archive` 로 docx 재패키징 금지
Windows 경로 구분자(`\`) 가 들어가 Word 가 열지 못한다. 반드시 forward slash(`/`) 사용:
```powershell
[System.IO.Compression.ZipFile]::Open($dst, 'Create')
# CreateEntry 에 relPath.Replace('\\', '/') 적용
```

---

## 8. Markdown → Word(docx) 변환

### 기본 변환
```bash
pandoc input.md -o output.docx
```

### 참조 스타일 문서 사용
```bash
# 1) 참조 템플릿 생성
pandoc -o reference.docx --print-default-data-file reference.docx

# 2) reference.docx 를 Word 에서 열어 스타일(Heading 1, Normal 등) 수정 후 저장

# 3) 수정한 스타일로 변환
pandoc input.md --reference-doc=reference.docx -o output.docx
```

### 목차 포함 변환
```bash
pandoc input.md --toc --toc-depth=3 -o output.docx
```

### 여러 파일 병합 변환
```bash
pandoc chapter1.md chapter2.md chapter3.md -o book.docx
```

---

## 9. Markdown → PDF 변환

PDF 변환에는 LaTeX 엔진이 필요하다.

### 엔진 설치
- **Windows/Linux**: MiKTeX 또는 TeX Live
- **macOS**: MacTeX 또는 BasicTeX
- **경량 대안**: wkhtmltopdf, weasyprint

### 변환 예시
```bash
# 기본 (pdflatex)
pandoc input.md -o output.pdf

# 한글 포함 시 XeLaTeX
pandoc input.md --pdf-engine=xelatex -V mainfont="Malgun Gothic" -o output.pdf

# wkhtmltopdf 사용 (LaTeX 불필요)
pandoc input.md --pdf-engine=wkhtmltopdf -o output.pdf
```

### 한글 PDF 변환 옵션
```bash
pandoc input.md \
  --pdf-engine=xelatex \
  -V mainfont="Malgun Gothic" \
  -V geometry:margin=1in \
  -V fontsize=11pt \
  -o output.pdf
```

---

## 10. 메타데이터와 YAML 프론트매터

Markdown 파일 상단에 YAML 블록으로 메타데이터를 지정할 수 있다.

```markdown
---
title: "문서 제목"
author: "작성자"
date: "2026-04-20"
lang: ko
toc: true
toc-depth: 2
numbersections: true
---

# 본문 시작
```

또는 커맨드라인에서:
```bash
pandoc input.md \
  --metadata title="문서 제목" \
  --metadata author="작성자" \
  -o output.docx
```

---

## 11. 코드 하이라이팅

### 스타일 목록 확인
```bash
pandoc --list-highlight-styles
# pygments, tango, espresso, zenburn, kate, monochrome, breezedark, haddock
```

### 사용 예
```bash
pandoc input.md --highlight-style=zenburn -o output.html
```

### 하이라이팅 끄기
```bash
pandoc input.md --no-highlight -o output.html
```

---

## 12. 템플릿

### 기본 템플릿 확인
```bash
pandoc -D html  > template.html
pandoc -D latex > template.tex
```

### 사용자 템플릿 적용
```bash
pandoc input.md --template=my-template.html -o output.html
```

---

## 13. 필터(Filter)

### Lua 필터
```bash
pandoc input.md --lua-filter=my-filter.lua -o output.html
```

### pandoc-crossref (상호참조)
```bash
pandoc input.md --filter pandoc-crossref -o output.pdf
```

### 자주 쓰이는 필터
- `pandoc-crossref`: 그림/표/수식 상호참조
- `pandoc-citeproc`: 참고문헌 처리 (내장됨)
- `pandoc-include`: 외부 파일 포함

---

## 14. 참고문헌 (Citations)

### BibTeX 파일 사용
```bash
pandoc input.md --citeproc --bibliography=refs.bib -o output.docx
```

### CSL 스타일 지정
```bash
pandoc input.md \
  --citeproc \
  --bibliography=refs.bib \
  --csl=ieee.csl \
  -o output.pdf
```

### 본문 내 인용
```markdown
이전 연구 [@smith2020] 에 따르면...
```

---

## 15. 자주 쓰는 명령 모음

### 배치 변환 (여러 파일)
```bash
# Bash
for f in *.md; do pandoc "$f" -o "${f%.md}.docx"; done

# PowerShell
Get-ChildItem *.md | ForEach-Object { pandoc $_.Name -o ($_.BaseName + ".docx") }
```

### 슬라이드 생성 (reveal.js)
```bash
pandoc slides.md -t revealjs -s -o slides.html
```

### EPUB 전자책 생성
```bash
pandoc book.md -o book.epub --metadata title="책 제목"
```

### 목차 + 섹션 번호 + 스타일 모두 적용
```bash
pandoc input.md \
  -s \
  --toc \
  --toc-depth=3 \
  -N \
  --reference-doc=reference.docx \
  -o output.docx
```

---

## 16. Markdown 확장 문법

Pandoc 은 표준 Markdown 을 확장한 Pandoc Markdown 을 지원한다.

### 표(Table)
```markdown
| 헤더1 | 헤더2 | 헤더3 |
|-------|:-----:|------:|
| 왼쪽  | 중앙  | 오른쪽 |
```

### 수식(LaTeX)
```markdown
인라인: $E = mc^2$

블록:
$$
\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
$$
```

### 각주
```markdown
본문입니다[^1].

[^1]: 각주 내용
```

### 정의 목록
```markdown
용어
:   용어에 대한 설명
```

### 확장 기능 on/off
```bash
pandoc -f markdown+smart    -t html input.md
pandoc -f markdown-raw_html -t docx input.md
```

---

## 17. 트러블슈팅

### 한글 깨짐 (PDF)
- `--pdf-engine=xelatex` 와 `-V mainfont` 옵션 사용

### 이미지가 안 보임
- `--resource-path` 로 이미지 경로 추가
- 절대경로 또는 상대경로 확인

### 용량이 너무 큰 docx
- `--reference-doc` 으로 경량 템플릿 적용
- 내장 이미지 크기 축소

### LaTeX 에러
- 필요한 패키지 설치 (`tlmgr install <pkg>`)
- 에러 메시지의 누락 패키지명 확인

### 표 서식이 적용 안 됨
- → [table-style.md](table-style.md) 참고. Pandoc 은 표 스타일을 `styleId="Table"` 리터럴로 찾으며, 그 스타일이 없으면 별도 주입이 필요하다.

---

## 18. 참고 링크

- 공식 사이트: https://pandoc.org
- 사용자 가이드: https://pandoc.org/MANUAL.html
- 데모: https://pandoc.org/demos.html
- GitHub: https://github.com/jgm/pandoc

---

## 19. 빠른 참조 (Cheat Sheet)

```bash
# 가장 많이 쓰는 변환
pandoc in.md   -o out.docx                          # Markdown → Word
pandoc in.md   -o out.pdf                           # Markdown → PDF
pandoc in.md   -s -o out.html                       # Markdown → HTML
pandoc in.docx -o out.md                            # Word → Markdown

# 옵션 조합
pandoc in.md --toc -N -s -o out.html                # 목차+번호+완전문서
pandoc in.md --reference-doc=ref.docx -o out.docx   # 스타일 적용
pandoc in.md --pdf-engine=xelatex -V mainfont="Malgun Gothic" -o out.pdf

# 정보 확인
pandoc --version
pandoc --list-input-formats
pandoc --list-output-formats
pandoc --list-highlight-styles
pandoc -D FORMAT                                    # 기본 템플릿 출력
```

---

## 20. 요약: 가장 효율적인 경로

1. Word 에서 원하는 스타일이 적용된 빈 문서 작성 (또는 기존 양식 활용)
2. 그 파일을 `reference.docx` 로 저장
3. `pandoc input.md --reference-doc=reference.docx -o output.docx`
4. 표 서식까지 살리려면 → [table-style.md](table-style.md) 의 주입 절차 수행

끝. style ID 변경, XML 수술, 정규화 같은 복잡한 작업 필요 없음.
