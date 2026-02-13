# Installation Guide

## Install as Excel Add-In
1. Close all running Excel instances.
2. Copy `AI_Assistant.xlam` to:
- `C:\Users\<YourUser>\AppData\Roaming\Microsoft\AddIns\`
3. Open Excel.
4. Go to:
- `File -> Options -> Add-ins`
5. In `Manage`, choose `Excel Add-ins`, then click `Go...`.
6. Click `Browse...`, select `AI_Assistant.xlam`, and enable it.
7. Restart Excel.

## First-Run Setup
1. Open the assistant UI.
2. Open `Settings`.
3. Configure one or more providers:
- DeepSeek key.
- OpenRouter key.
- OpenAI key.
- LM Studio host/port/model (optional local mode).
4. Save settings.

## Trust Center Requirement
For full functionality and dev tooling compatibility, ensure:
- `Trust access to the VBA project object model` is enabled.

## Verification Checklist
- Assistant form opens from Excel.
- You can send a prompt in Cloud mode.
- If using Codex CLI mode:
- `codex --version` works in terminal.
- `codex login status` confirms authentication.
