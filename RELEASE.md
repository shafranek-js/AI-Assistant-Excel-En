# Release Guide

## Goal
Publish a stable add-in artifact (`AI_Assistant.xlam`) and updated docs.

## Release Steps
1. Update source in `vba_unpacked_utf8/`.
2. Rebuild workbook:
```powershell
.\rebuild_editable_fullforms.ps1 -InputDir .\vba_unpacked_utf8 -Track dev
```
3. Sync add-in:
```powershell
.\scripts\sync_addin.ps1
```
4. Smoke test:
- Open `AI_Assistant_latest_dev.xlsm`.
- Test at least one cloud model route.
- Test command extraction/execution flow.
- Test settings save/load.
- If enabled, test Codex CLI route.
5. Verify docs:
- `README.md`
- `INSTALLATION.md`
- `DEVELOPMENT.md`
- `CONTEXT.MD` (if major behavior changed)
6. Commit and push to release branch.
7. Tag release and publish notes.

## Release Artifact Naming
- Add-in artifact name should remain:
- `AI_Assistant.xlam`

## Minimal Release Notes Template
1. Added
2. Changed
3. Fixed
4. Known limitations
