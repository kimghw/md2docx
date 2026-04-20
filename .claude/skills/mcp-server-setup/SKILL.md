---
name: mcp-server-setup
description: Windows 환경에서 Claude Code에 외부 AI CLI(Codex, Gemini)를 MCP 서버로 등록·설정하는 가이드. Node.js 설치, Codex/Gemini CLI 설치, 인증, `claude mcp add` 등록, `%USERPROFILE%\.codex\config.toml` 및 `%USERPROFILE%\.claude\settings.json` 편집, 제공 도구(`codex`, `codex-reply`, `ask-gemini`, `brainstorm`, `fetch-chunk`)와 파라미터(model, reasoning_effort, approval-policy, sandbox 등) 설명이 필요할 때 호출된다.
---

# mcp-server-setup — Claude Code MCP 서버 설정 가이드 (Windows)

Windows(PowerShell 기준)에서 Claude Code에 **Codex CLI (OpenAI)** 와 **Gemini CLI (Google)** 를 MCP 서버로 등록하여 외부 AI 모델을 도구처럼 호출할 수 있게 한다.

> 아래 명령은 **PowerShell**을 기준으로 한다. `cmd.exe`를 쓴다면 `$env:VAR` → `%VAR%` 로 치환.
> 전역 npm 설치는 **관리자 권한 PowerShell**을 권장한다 (시작 메뉴 → "Windows PowerShell" 우클릭 → "관리자 권한으로 실행").

---

## 사전 준비

### Node.js 설치 (하나 선택)

```powershell
# winget (Windows 10/11 기본 제공)
winget install OpenJS.NodeJS.LTS

# 또는 Chocolatey
choco install nodejs-lts -y

# 또는 공식 설치 관리자
# https://nodejs.org/en/download 에서 "Windows Installer (.msi)" LTS 다운로드
```

- Node.js 20+ 필요
- 설치 후 새 PowerShell 창을 열어 `node -v`, `npm -v` 확인

### PATH 확인 및 새로고침 (설치 직후 필수 체크)

설치는 성공했는데 새 터미널에서 `node` / `npm` / `gemini` / `codex`가 `CommandNotFoundException`으로 뜨는 경우가 많다. Windows는 프로세스 시작 시점의 PATH를 고정하므로, **VS Code 등 부모 프로세스가 설치 이전에 떠 있었으면** 그 밑의 새 터미널도 구 PATH를 상속한다.

**1) 레지스트리상의 PATH에 node/npm 경로가 실제로 들어갔는지 확인**
```powershell
[Environment]::GetEnvironmentVariable('Path','Machine') -split ';' | Select-String 'nodejs'
[Environment]::GetEnvironmentVariable('Path','User')    -split ';' | Select-String 'npm'
```
기대 결과:
- Machine PATH: `C:\Program Files\nodejs\`
- User PATH:    `C:\Users\<USER>\AppData\Roaming\npm`

**2) 바이너리 실제 존재 여부**
```powershell
Test-Path "C:\Program Files\nodejs\node.exe"
Test-Path "$env:APPDATA\npm\gemini.cmd"
Test-Path "$env:APPDATA\npm\codex.cmd"
```

**3) 현재 세션 PATH 강제 새로고침** (터미널을 닫지 않고)
```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
node -v
```

**4) 그래도 안 되면 임시 주입**
```powershell
$env:Path = "C:\Program Files\nodejs;$env:APPDATA\npm;" + $env:Path
```

**5) 영구 해결**
- VS Code 통합 터미널에서 문제가 반복되면 **VS Code 자체를 완전히 종료 후 재시작** (새 터미널만 여는 것으로는 부족).
- 외부 PowerShell이라면 **로그아웃/재로그인** 또는 **재부팅** 한 번.

### Claude Code CLI 설치
```powershell
npm install -g @anthropic-ai/claude-code
```

---

## 1. Codex CLI (OpenAI)

### 설치
```powershell
npm install -g @openai/codex
```

### 인증 설정

구독 (ChatGPT Pro/Plus/Team)으로 로그인:

```powershell
codex login --device-auth
```

→ 출력된 URL을 브라우저에서 열고 디바이스 코드 입력.

### MCP 서버 등록
```powershell
claude mcp add -s user --transport stdio codex -- codex mcp-server
```

`-c` 플래그로 config 값을 인라인 오버라이드할 수 있다:
```powershell
claude mcp add -s user --transport stdio codex -- codex mcp-server -c reasoning_effort="xhigh" -c model="gpt-5.4"
```

### Codex 설정 파일 (`%USERPROFILE%\.codex\config.toml`)

MCP 등록 시 `-c`로 매번 지정하는 대신, `%USERPROFILE%\.codex\config.toml`에서 기본값을 설정할 수 있다.
파일/폴더가 없으면 생성:

```powershell
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.codex" | Out-Null
notepad "$env:USERPROFILE\.codex\config.toml"
```

내용 예시:
```toml
model = "gpt-5.4"
reasoning_effort = "xhigh"      # low | medium | high | xhigh
```

| 키 | 설명 | 기본값 |
|---|---|---|
| `model` | 사용할 모델 | `o3` |
| `reasoning_effort` | 추론 수준 | `medium` |
| `sandbox_permissions` | 샌드박스 권한 | `[]` |
| `shell_environment_policy.inherit` | 셸 환경변수 상속 | - |

> `-c` 플래그는 config.toml 값을 오버라이드한다. config.toml에 이미 설정된 값은 별도 지정 없이 MCP 서버에 자동 적용된다.

### 제공 도구

| 도구 | 설명 |
|---|---|
| `codex` | 새 코딩 세션 시작 (prompt 전달) |
| `codex-reply` | 기존 세션에 후속 지시 (threadId로 이어서 대화) |

### 주요 파라미터

| 파라미터 | 설명 | 예시 |
|---|---|---|
| `model` | 모델 선택 | `gpt-5.2`, `gpt-5.2-codex` |
| `prompt` | 초기 프롬프트 (필수) | 자유 텍스트 |
| `approval-policy` | 명령어 실행 승인 정책 | `untrusted`, `on-failure`, `on-request`, `never` |
| `sandbox` | 샌드박스 모드 | `read-only`, `workspace-write`, `danger-full-access` |
| `cwd` | 작업 디렉토리 (Windows 경로, 예: `C:\work\repo`) | 경로 |

---

## 2. Gemini CLI (Google)

### 설치
```powershell
npm install -g @google/gemini-cli
```

### 인증 (최초 1회)
```powershell
gemini
```

### MCP 서버 등록

Gemini CLI는 자체 MCP 서버 모드가 없으므로 서드파티 래퍼 사용:

```powershell
npm install -g gemini-mcp-tool
claude mcp add -s user --transport stdio gemini -- gemini-mcp
```

> 패키지명은 `gemini-mcp-tool`이지만 실행 바이너리는 `gemini-mcp`이다.
> Windows에서는 `gemini-mcp.cmd` 로 설치된다 (`where.exe gemini-mcp`로 확인).
> Codex와 달리 등록 시 모델을 고정하지 않고, 도구 호출 시 `model` 파라미터로 선택한다.

### 제공 도구

| 도구 | 설명 |
|---|---|
| `ask-gemini` | Gemini에 질문/작업 요청 (`model` 파라미터로 모델 선택) |
| `brainstorm` | 브레인스토밍 (창의적 프레임워크 자동 적용) |
| `fetch-chunk` | 대용량 콘텐츠 청크 단위 조회 |

### 주요 파라미터 (`ask-gemini`)

| 파라미터 | 설명 | 예시 |
|---|---|---|
| `prompt` | 분석 요청 (필수). `@file` 구문으로 파일 포함 가능 | `@main.py 이 코드 설명해줘` |
| `model` | 모델 선택 (기본: `gemini-2.5-pro`) | **`gemini-3.1-pro`**, `gemini-2.5-flash` |
| `sandbox` | 샌드박스 모드에서 코드 실행 | `true` / `false` |
| `changeMode` | 구조화된 편집 제안 반환 | `true` / `false` |

---

## 참고: 설정 파일 직접 편집

`%USERPROFILE%\.claude\settings.json`을 직접 수정하는 방법:

```powershell
notepad "$env:USERPROFILE\.claude\settings.json"
```

```json
{
  "mcpServers": {
    "codex": {
      "command": "codex",
      "args": ["mcp-server", "-c", "model=\"gpt-5.4\"", "-c", "reasoning_effort=\"xhigh\""]
    },
    "gemini": {
      "command": "gemini-mcp",
      "args": [],
      "env": {
        "GEMINI_MODEL": "gemini-3.1-pro"
      }
    }
  }
}
```

> Windows에서 `command` 값이 인식되지 않으면 `.cmd` 확장자를 명시하거나 전체 경로를 지정:
> ```json
> "command": "C:\\Users\\<USER>\\AppData\\Roaming\\npm\\codex.cmd"
> ```
> JSON에서 백슬래시는 `\\`로 이스케이프해야 한다.

---

## 트러블슈팅

- **`npm install -g` 실행 시 EACCES/권한 오류**: PowerShell을 **관리자 권한**으로 다시 실행. 또는 npm prefix를 사용자 경로로 변경 — `npm config set prefix "$env:APPDATA\npm"`.
- **`node` / `claude` / `codex` / `gemini` / `gemini-mcp` 명령을 찾지 못함 (`'xxx' is not recognized...` / `CommandNotFoundException`)**: 상단 "PATH 확인 및 새로고침" 섹션 참조. 핵심은 (1) 레지스트리 PATH에 실제 등록됐는지 확인, (2) `$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')` 로 현재 세션 갱신, (3) 반복되면 VS Code/셸 부모 프로세스 재시작.
- **`claude mcp add` 실행 시 권한/스코프 이슈**: `-s user`는 사용자 전역(`%USERPROFILE%\.claude\settings.json`), `-s project`는 현재 프로젝트의 `.mcp.json`에 등록된다. 스코프를 의도에 맞게 선택한다.
- **Gemini `gemini-mcp` 명령을 찾지 못함**: 패키지명(`gemini-mcp-tool`)과 바이너리명(`gemini-mcp`)이 다르다. 글로벌 설치 후 `where.exe gemini-mcp`로 확인 (`where`는 PowerShell alias와 충돌하므로 `.exe` 명시).
- **Codex 모델/추론 수준이 적용되지 않음**: `-c` 플래그가 config.toml을 오버라이드하므로 MCP 등록 명령에서 `-c` 누락 시 config.toml 값이 쓰인다. 예상 동작과 다르면 `%USERPROFILE%\.codex\config.toml`을 확인.
- **MCP 서버가 시작되지 않음 (Windows 특유)**: `settings.json`의 `command`에 `.cmd` 확장자를 명시하거나 절대경로로 교체. 예: `"command": "codex.cmd"`.
- **PowerShell 실행 정책 오류 (`running scripts is disabled` / `...ps1 파일을 로드할 수 없습니다`)**: npm이 만든 `.ps1` 래퍼가 차단된 경우. 두 가지 해결책 —
  - (권장) 사용자 범위 정책 완화 한 번만: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` → `Y`
  - (즉시) `.cmd` 확장자로 직접 호출: `gemini.cmd`, `codex.cmd`, `claude.cmd`
