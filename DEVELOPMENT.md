# Development Guide

## Source of Truth
Always edit VBA code in:
- `vba_unpacked_utf8/modAIHelper.bas`
- `vba_unpacked_utf8/modExcelHelper.bas`
- `vba_unpacked_utf8/modMain.bas`
- `vba_unpacked_utf8/frmChat.frm`
- `vba_unpacked_utf8/frmSettings.frm`
- `vba_unpacked_utf8/ThisWorkbook.cls`
- `vba_unpacked_utf8/Sheet1.cls`

Generated artifacts (`.xlsm`, `.xlam`) are build outputs.

## Build Workflow
Rebuild editable workbook:
```powershell
.\rebuild_editable_fullforms.ps1 -InputDir .\vba_unpacked_utf8 -Track dev
```

Default outputs:
- Versioned workbook: `AI_Assistant_<track>_yyyyMMdd_HHmm_bNN.xlsm`
- Latest alias: `AI_Assistant_latest_<track>.xlsm`

## Add-In Sync Workflow
After rebuilding, update add-in artifact:
```powershell
.\scripts\sync_addin.ps1
```

Default behavior:
- Reads `AI_Assistant_latest_dev.xlsm`.
- Reuses `AI_Assistant.xlam` as add-in container template.
- Replaces `xl/vbaProject.bin`.
- Produces updated `AI_Assistant.xlam`.

## Install Local Add-In for Testing
1. Close all Excel windows.
2. Copy `AI_Assistant.xlam` to:
- `C:\Users\<User>\AppData\Roaming\Microsoft\AddIns\`
3. Excel -> `File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go...`
4. Browse and enable `AI_Assistant.xlam`.

## Debug and Troubleshooting
### VBA project automation issues
- Symptom: build script cannot inject modules/forms.
- Check: Excel Trust Center setting `Trust access to the VBA project object model`.

### Codex CLI issues
- Ensure CLI is installed and logged in.
- Validate from terminal:
```powershell
codex --version
codex login status
```
- If process was interrupted, app can return exit `-1073741510`.

### Local LM Studio issues
- Verify endpoint is reachable:
- `http://127.0.0.1:1234/v1/models` (or configured host/port).

### Encoding issues
- Use `unpack_vba.py` with explicit output encoding when required:
```powershell
python .\unpack_vba.py vbaProject.bin --out vba_unpacked_utf8 --encoding utf-8-sig
```

## Branching and PR Strategy (Recommended)
1. Create feature branch from `main`.
2. Keep commits scoped (one functional concern per commit).
3. Rebuild and smoke-test `AI_Assistant_latest_dev.xlsm`.
4. Sync `AI_Assistant.xlam` only when needed for release/testing.
5. Open PR with:
- change summary,
- user-visible impact,
- testing notes.
