# Claude Agent Skill 작성 지침 (Anthropic 공식)

> 출처: [Claude Docs — Skill authoring best practices](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/best-practices), [Claude Code Docs — Extend Claude with skills](https://code.claude.com/docs/en/skills), [anthropics/skills GitHub](https://github.com/anthropics/skills)
> 갱신일: 2026-04-20

---

## 1. SKILL.md 기본 구조

모든 skill은 `SKILL.md` 파일 하나를 최소 요건으로 한다. 파일은 두 부분으로 구성된다:

1. **YAML frontmatter** (`---` 사이) — skill이 **언제** 사용되어야 하는지 Claude에게 알림
2. **Markdown 본문** — skill이 호출되었을 때 Claude가 따라야 할 지침

```markdown
---
name: my-skill-name
description: 이 skill이 무엇을 하는지와 언제 사용되어야 하는지를 3인칭으로 설명한다.
---

# 본문 제목

skill 사용 시 따라야 할 지침…
```

---

## 2. Frontmatter 필드 규칙

### `name` (필수)

| 항목 | 규칙 |
|------|------|
| 최대 길이 | **64자** |
| 허용 문자 | **소문자 영문 / 숫자 / 하이픈(`-`)만** |
| 금지 | 대문자, 공백, XML 태그, 예약어 |
| 권장 형태 | **Gerund (verb + -ing)** — skill이 제공하는 활동/능력을 명확히 표현 |
| 폴더명 일치 | skill 폴더명과 `name` 값은 **동일**해야 함 |

**좋은 예**: `extracting-docx-styles`, `reviewing-security`, `generating-reports`
**나쁜 예**: `extractStyle` (camelCase 불가), `Extract_Style` (대문자·언더스코어 불가)

### `description` (필수)

| 항목 | 규칙 |
|------|------|
| 최대 길이 | **1024자** |
| 내용 | **무엇을(What) + 언제(When)** 둘 다 포함 |
| 시점(POV) | **항상 3인칭** (system prompt에 주입되므로 일관성 중요) |
| 금지 | XML 태그, 빈 값 |

> description은 skill **discovery**의 핵심이다. "언제" 부분이 없으면 Claude가 호출 시점을 판단하지 못한다.

**좋은 예**:
```yaml
description: Pandoc `--reference-doc`용 Word 템플릿의 스타일을 분석·검증·수정한다. 한글 Word 템플릿의 w:name 매칭, 표 스타일 누락, reference.docx 준비가 필요한 경우 호출된다.
```

**나쁜 예**:
```yaml
description: Word 스타일을 추출해줘.  # 명령형, 2인칭, "언제" 없음
```

---

## 3. 본문 작성 가이드

### 길이
- **500줄 이하** 권장
- 초과 시 **Progressive Disclosure** 패턴으로 별도 파일 분리
- Claude가 SKILL.md를 로드한 순간 모든 토큰이 context window를 차지함 → 짧을수록 좋다

### Progressive Disclosure (점진적 공개)

SKILL.md에는 **핵심 지침**만 두고, 참조성 자료는 하위 파일로 분리해 필요할 때만 읽게 한다:

```
my-skill/
├── SKILL.md              # 핵심 지침 (항상 로드됨)
├── references/           # 참조 문서 (필요 시만 로드)
│   ├── detailed-spec.md
│   └── edge-cases.md
├── scripts/              # 실행 가능 코드
│   └── helper.py
└── assets/               # 템플릿·정적 파일
    └── template.docx
```

본문에서는 `자세한 엣지케이스는 [references/edge-cases.md](references/edge-cases.md) 참조` 식으로 포인터만 두면 된다.

---

## 4. 디렉토리 구조

| 폴더/파일 | 필수 여부 | 역할 |
|----------|----------|------|
| `SKILL.md` | 필수 | 메타데이터 + 핵심 지침 |
| `scripts/` | 선택 | 실행 가능 코드 (Python, Shell 등) |
| `references/` | 선택 | 필요 시 로드되는 참조 문서 |
| `assets/` | 선택 | 출력에 사용되는 템플릿/정적 파일 |

---

## 5. 작성 팁

- **Consistent naming**: 프로젝트 내 skill 이름 규칙(예: 모두 gerund)을 통일
- **Concrete triggers**: description에 구체적 키워드(파일 타입, 도구명, 문제 상황) 포함 → discovery 정확도 ↑
- **Actionable body**: 본문은 Claude가 실제로 실행할 수 있는 명령·절차 위주
- **No emojis (기본)**: Claude 기본 지침은 이모지 지양 — 사용자 skill에서 꼭 필요할 때만
- **반복 수정**: 실제 사용에서 Claude가 skill을 어떻게 쓰는지 관찰하며 개선 (예상 외 탐색 경로, 놓친 연결, 무시된 섹션 확인)

---

## 6. 안티패턴

- ❌ `name`에 camelCase/대문자 사용 → validation 실패
- ❌ description이 명령형("…해줘") → discovery 실패 유발
- ❌ description에 "언제" 빠짐 → Claude가 호출 시점 판단 불가
- ❌ 본문이 500줄 초과인데 분리 안 함 → context 낭비
- ❌ 폴더명과 `name` 필드 불일치
- ❌ 실행 불가능한 추상적 지침만 나열

---

## 7. 이 프로젝트(`c:\md2docx`) 내 skill 검토 기록

### `extract-docx-styles` (구 `extractStyle`)
**검토일**: 2026-04-20

**수정 전 문제점**:
- `name: extractStyle` — camelCase 사용으로 공식 규칙 위반
- description 마지막 문장이 명령형("…다룰 때 사용.")

**수정 내용**:
- `name` → `extract-docx-styles` (소문자·하이픈, gerund 형태)
- 폴더명 `.claude/skills/extractStyle/` → `.claude/skills/extract-docx-styles/` 로 리네임
- description을 3인칭 서술형으로 통일 ("…호출된다")

**유지한 점**:
- 본문 283줄 → 500줄 이내라 progressive disclosure 분리 미적용
- 다만 `Windows Pandoc PATH 문제`, `한글 Word style ID 매핑 참고` 섹션은 향후 본문이 늘어나면 `references/` 분리 후보
