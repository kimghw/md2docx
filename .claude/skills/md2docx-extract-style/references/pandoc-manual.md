# Pandoc 사용 매뉴얼

Pandoc은 다양한 문서 형식 간 변환을 지원하는 범용 문서 변환 도구입니다. Markdown, HTML, LaTeX, Word(docx), PDF, EPUB 등 40여 개 포맷을 상호 변환할 수 있습니다.

---

## 1. 설치

### Windows
```bash
# Chocolatey
choco install pandoc

# Winget
winget install --id JohnMacFarlane.Pandoc

# 또는 공식 인스톨러
# https://pandoc.org/installing.html
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

---

## 2. 기본 사용법

### 기본 문법
```bash
pandoc [옵션] [입력파일] -o [출력파일]
```

### 기본 예시
```bash
# Markdown → Word
pandoc input.md -o output.docx

# Markdown → PDF
pandoc input.md -o output.pdf

# Markdown → HTML
pandoc input.md -o output.html

# Word → Markdown
pandoc input.docx -o output.md

# HTML → Markdown
pandoc input.html -o output.md
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

### 지원 포맷 목록 확인
```bash
pandoc --list-input-formats
pandoc --list-output-formats
```

---

## 5. Markdown → Word(docx) 변환

### 기본 변환
```bash
pandoc input.md -o output.docx
```

### 참조 스타일 문서 사용
```bash
# 1) 참조 템플릿 생성
pandoc -o reference.docx --print-default-data-file reference.docx

# 2) reference.docx를 Word에서 열어 스타일(Heading 1, Normal 등) 수정 후 저장

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

## 6. Markdown → PDF 변환

PDF 변환에는 LaTeX 엔진이 필요합니다.

### 엔진 설치
- **Windows/Linux**: MiKTeX 또는 TeX Live
- **macOS**: MacTeX 또는 BasicTeX
- **경량 대안**: wkhtmltopdf, weasyprint

### 변환 예시
```bash
# 기본 (pdflatex)
pandoc input.md -o output.pdf

# 한글 포함 시 XeLaTeX 사용
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

## 7. 메타데이터와 YAML 프론트매터

Markdown 파일 상단에 YAML 블록으로 메타데이터를 지정할 수 있습니다.

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

## 8. 코드 하이라이팅

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

## 9. 템플릿 사용

### 기본 템플릿 확인
```bash
pandoc -D html > template.html
pandoc -D latex > template.tex
```

### 사용자 템플릿 적용
```bash
pandoc input.md --template=my-template.html -o output.html
```

---

## 10. 필터(Filter)

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

## 11. 참고문헌 (Citations)

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

## 12. 자주 쓰는 명령 모음

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

## 13. Markdown 확장 문법

Pandoc은 표준 Markdown을 확장한 Pandoc Markdown을 지원합니다.

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
pandoc -f markdown+smart -t html input.md
pandoc -f markdown-raw_html -t docx input.md
```

---

## 14. 트러블슈팅

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

---

## 15. 참고 링크

- 공식 사이트: https://pandoc.org
- 사용자 가이드: https://pandoc.org/MANUAL.html
- 데모: https://pandoc.org/demos.html
- GitHub: https://github.com/jgm/pandoc

---

## 16. 빠른 참조 (Cheat Sheet)

```bash
# 가장 많이 쓰는 변환
pandoc in.md -o out.docx                        # Markdown → Word
pandoc in.md -o out.pdf                         # Markdown → PDF
pandoc in.md -s -o out.html                     # Markdown → HTML
pandoc in.docx -o out.md                        # Word → Markdown

# 옵션 조합
pandoc in.md --toc -N -s -o out.html            # 목차+번호+완전문서
pandoc in.md --reference-doc=ref.docx -o out.docx  # 스타일 적용
pandoc in.md --pdf-engine=xelatex -V mainfont="Malgun Gothic" -o out.pdf

# 정보 확인
pandoc --version
pandoc --list-input-formats
pandoc --list-output-formats
pandoc --list-highlight-styles
pandoc -D FORMAT                                # 기본 템플릿 출력
```
