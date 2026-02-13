# AI Assistant for Excel (VBA)

AI-powered Microsoft Excel add-in (`.xlam`) with support for cloud and local LLM providers.

This fork (`AI-Assistant-Excel-En`) is maintained in English and focused on practical automation workflows, reproducible builds, and safe command execution.

## Overview
The assistant runs directly inside Excel and can:
- collect workbook and selection context automatically,
- send tasks to an AI model,
- parse a structured `commands` block from the response,
- validate commands before execution,
- apply actions to the active workbook.

## Screenshots
Recommended screenshots for this README:
- Chat form (`frmChat`) with model selection and command preview toggle.
- Settings form (`frmSettings`) with provider keys and response language options.
- Example image-to-data extraction flow.

## Features
- Chat assistant with Excel-aware context.
- Cloud/local model routing.
- Optional selected-data injection into prompt.
- Optional command preview before execution.
- Image attachment support for compatible routes.
- Strict command validation before execution.
- Retry/backoff for transient API and network failures.

## Supported Models and Providers
Cloud routes:
- DeepSeek direct (`deepseek-chat`).
- OpenRouter routes:
- Claude Sonnet 4.5.
- GPT-5.2 via OpenRouter.
- Gemini 3 Pro via OpenRouter.
- Gemini 3 Flash via OpenRouter.
- OpenAI direct routes:
- GPT-5.2 (Direct OpenAI).
- GPT-5.2 Codex (Direct OpenAI).

Local routes:
- LM Studio (OpenAI-compatible local API endpoint).
- Codex CLI route (`Codex CLI (ChatGPT Plus)`) executed locally from VBA.

## Command Coverage
The assistant supports a large set of executable Excel commands across:

| Category | Example commands |
|---|---|
| Cells and formulas | `SET_VALUE`, `SET_FORMULA`, `FILL_DOWN`, `COPY`, `PASTE_VALUES` |
| Formatting | `BOLD`, `FONT_COLOR`, `FILL_COLOR`, `BORDER`, `MERGE` |
| Rows and columns | `INSERT_ROW`, `DELETE_COLUMN`, `HIDE_ROW`, `GROUP_ROWS` |
| Sort and filter | `SORT`, `FILTER`, `REMOVE_DUPLICATES`, `AUTOFILTER` |
| Charts | `CREATE_CHART`, `CHART_TITLE`, `CHART_LEGEND` |
| Pivot tables | `CREATE_PIVOT`, `PIVOT_ADD_ROW`, `PIVOT_ADD_VALUE` |
| Sheets | `ADD_SHEET`, `RENAME_SHEET`, `PROTECT_SHEET` |
| Conditional formatting | `COND_HIGHLIGHT`, `DATA_BARS`, `COLOR_SCALE` |
| Data validation | `VALIDATION_LIST`, `VALIDATION_NUMBER` |

## Installation
1. Copy `AI_Assistant.xlam` to:
`C:\Users\<YourUser>\AppData\Roaming\Microsoft\AddIns\`
2. In Excel:
- `File -> Options -> Add-ins`
- `Manage: Excel Add-ins -> Go...`
- `Browse...` and select `AI_Assistant.xlam`
- Enable checkbox and confirm
3. Restart Excel.

Detailed guide: `INSTALLATION.md`

## API Setup
Open the assistant and click `Settings`, then configure provider keys.

| Route | Where to get credentials |
|---|---|
| DeepSeek | https://platform.deepseek.com |
| OpenRouter (Claude/GPT/Gemini routes) | https://openrouter.ai |
| OpenAI Direct routes | https://platform.openai.com |
| Codex CLI route | `codex login` (local CLI authentication) |

For local LM Studio mode:
- install and run LM Studio,
- configure IP/Port in settings (default `127.0.0.1:1234`),
- refresh model list in the Settings form.

## Usage
1. Open a workbook with data.
2. Open AI Assistant.
3. Choose `Cloud` or `Local` mode.
4. Optionally enable `Include Data`.
5. Enter task in natural language.

Example prompts:
- `Add a new column with B*C totals and format as currency.`
- `Sort this table by date descending and highlight negative values.`
- `Build a monthly sales chart from the selected range.`
- `Create a pivot table by region with sum of revenue.`
- `Extract table values from the attached image and place them into the sheet.`

## Repository Layout
- `vba_unpacked_utf8/` - editable VBA source of truth.
- `rebuild_editable_fullforms.ps1` - rebuilds a fully editable `.xlsm` from source.
- `scripts/sync_addin.ps1` - updates `AI_Assistant.xlam` from latest workbook build.
- `unpack_vba.py` - extracts VBA modules from `vbaProject.bin`.
- `AI_Assistant.xlam` - add-in artifact.
- `AI_Assistant_latest_dev.xlsm` - latest workbook build artifact.
- `CONTEXT.MD` - deep technical context for session recovery.
- `IMPROVEMENT.MD` - prioritized roadmap.

## Development Quick Start
1. Edit files in `vba_unpacked_utf8/`.
2. Rebuild:
```powershell
.\rebuild_editable_fullforms.ps1 -InputDir .\vba_unpacked_utf8 -Track dev
```
3. Sync add-in:
```powershell
.\scripts\sync_addin.ps1
```

See also:
- `DEVELOPMENT.md`
- `RELEASE.md`
- `CONTRIBUTING.md`

## Requirements
- Windows and desktop Excel with VBA enabled.
- Trust Center setting: `Trust access to the VBA project object model`.
- `7z` in `PATH` for packaging scripts.
- Python 3.x for `unpack_vba.py`.
- Internet access for cloud routes.
- Optional: Codex CLI installed/authenticated for Codex CLI mode.

## Security Notes
- API keys are stored in Windows Registry:
`HKEY_CURRENT_USER\Software\ExcelAIAssistant\`
- Data is sent only to the selected provider route.
- With LM Studio local mode, traffic stays on local endpoint.

## Disclaimer
Use at your own risk. Always keep backups of important workbooks before automatic command execution.

## License
MIT License. See `LICENSE`.
