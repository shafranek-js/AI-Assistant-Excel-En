# Contributing

## Language and Documentation
- Keep user-facing docs in English.
- Keep commit messages and PR descriptions clear and specific.

## Development Rules
1. Edit source files in `vba_unpacked_utf8/`.
2. Rebuild with `rebuild_editable_fullforms.ps1`.
3. Verify behavior in `AI_Assistant_latest_dev.xlsm`.
4. Sync `AI_Assistant.xlam` only when a release/test artifact is required.

## Pull Request Checklist
- Feature/fix is scoped and documented.
- No unrelated changes mixed in.
- Build script still works.
- Core chat flow tested.
- If model routing changed, tested relevant route(s).
- If command parsing/execution changed, validated command safety path.

## Suggested PR Template
1. Summary
2. Motivation
3. Technical changes
4. User-visible changes
5. Testing performed
6. Risks and rollback notes
