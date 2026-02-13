# AI Assistant for Excel (English Fork)

AI-powered Excel assistant implemented in VBA, packaged as an Excel add-in (`.xlam`).

This fork is focused on English-first UX, development reproducibility, and practical integration with multiple model providers.

## What It Does
- Opens a chat form inside Excel.
- Builds context from the active workbook and selected range.
- Sends tasks to an AI model (cloud, local, or Codex CLI).
- Extracts an executable `commands` block from the AI response.
- Validates commands before execution.
- Executes valid commands directly in Excel.

## Core Features
- Conversational assistant UI (`frmChat`) with:
- Cloud/Local mode switch.
- Model selection.
- Optional selected-data inclusion.
- Optional command preview before execution.
- Image attachment support for compatible routes.
- Settings UI (`frmSettings`) with:
- DeepSeek API key.
- OpenRouter API key.
- OpenAI API key.
- LM Studio host/port/model.
- Response language selector (English, Russian, Ukrainian, Czech, Spanish, German).
- Strict command validation prior to execution.
- Retry/backoff for transient API/network failures.

## Model and Provider Support
Cloud routes:
- DeepSeek direct (`deepseek-chat`).
- OpenRouter routes:
- Claude Sonnet.
- GPT route via OpenRouter.
- Gemini Pro / Gemini Flash routes via OpenRouter.
- OpenAI direct routes:
- GPT-5.2 direct.
- GPT-5.2 Codex direct.

Local routes:
- LM Studio (OpenAI-compatible local endpoint).
- Codex CLI route (`Codex CLI (ChatGPT Plus)`), executed locally from VBA.

## Repository Layout
- `vba_unpacked_utf8/`: editable VBA source of truth.
- `rebuild_editable_fullforms.ps1`: rebuilds a fully editable `.xlsm` from source files.
- `unpack_vba.py`: extracts VBA modules from `vbaProject.bin`.
- `AI_Assistant.xlam`: add-in artifact used for Excel AddIns deployment.
- `AI_Assistant_latest_dev.xlsm`: latest workbook build artifact.
- `CONTEXT.MD`: deep technical context for session restoration.
- `IMPROVEMENT.MD`: prioritized roadmap.

## Quick Start (Developers)
1. Edit VBA source files in `vba_unpacked_utf8/`.
2. Rebuild:
```powershell
.\rebuild_editable_fullforms.ps1 -InputDir .\vba_unpacked_utf8 -Track dev
```
3. Sync add-in from latest workbook:
```powershell
.\scripts\sync_addin.ps1
```
4. Install `AI_Assistant.xlam` in Excel AddIns.

See:
- `INSTALLATION.md` for end-user add-in setup.
- `DEVELOPMENT.md` for development workflow and troubleshooting.
- `RELEASE.md` for release/publish checklist.

## Requirements
- Windows + desktop Excel with VBA enabled.
- Excel Trust Center: `Trust access to the VBA project object model`.
- `7z` available in `PATH` (used by tooling scripts).
- Python 3.x for `unpack_vba.py` (standard-library only).
- Network access for cloud providers.
- Optional: Codex CLI installed and authenticated for Codex CLI mode.

## Project Status
This fork is actively maintained for practical Excel automation workflows and iterative model-provider integration.
