# AI Assistant for Excel - Improvement Roadmap

## Objective
Increase reliability, safety, maintainability, and usability of the VBA assistant without breaking current workflows.

## Prioritization Method
- Impact: how much user value/risk reduction the change brings.
- Effort: estimated implementation complexity/time.
- Risk: probability of regressions or support burden.

Scoring principle: do high-impact, low-risk, low/medium-effort items first.

## Priority Levels
- `P0` Critical: stability/safety issues and blockers.
- `P1` Important: major UX and reliability improvements.
- `P2` Strategic: architecture/security/test maturity.

## Roadmap

### Phase 1 (P0) - Stabilization and Safe Defaults (1-3 days)
1. Add response language setting (`English/Russian`) in Settings form instead of hardcoded prompt language.
2. Add `Preview commands before execute` toggle.
3. Add strict command validation before execution (address/range/argument checks).
4. Add retry with backoff for API/network errors (timeout, 429, 5xx).
5. Add startup diagnostics check (macros enabled, Trust Center access, API key presence).

Success criteria:
- No destructive command executes without valid syntax.
- User can switch default response language from UI.
- Temporary API failures recover automatically in most cases.

### Phase 2 (P1) - Reliability and User Control (4-7 days)
1. Add `Dry Run` mode (assistant generates commands but does not apply them).
2. Add action log sheet (`AI_Log`) with timestamp, request, model, commands, result, error.
3. Add `Safe Mode` for confirmation on risky commands (`DELETE_*`, `CLEAR_ALL`, etc.).
4. Add persistent chat history (optional, user-toggle).
5. Improve forms layout and DPI behavior for 125%-200% scaling.

Success criteria:
- Users can audit all assistant actions.
- Risky operations require explicit confirmation.
- UI remains usable on common DPI settings.

### Phase 3 (P2) - Security and Maintainability (1-2 weeks)
1. Refactor `modAIHelper` into separate modules: `ApiClient`, `PromptBuilder`, `CommandParser`, `CommandExecutor`.
2. Store API keys securely (DPAPI or Windows Credential Manager wrapper).
3. Add workbook-level smoke tests for key macros and parser behavior.
4. Add import/export settings (JSON) for easy migration.
5. Add VBA project signing process for deployment stability.

Success criteria:
- Reduced regression risk when changing one subsystem.
- API keys are not stored in plaintext.
- Release checklist includes smoke tests and signature status.

## Candidate Backlog (After P2)
1. Undo/rollback strategy for batch command execution.
2. Model capability matrix (image support, context limits, tools support).
3. Caching for repeated context extraction on large sheets.
4. Guided setup wizard for first launch.

## Execution Plan by Week
1. Week 1: Complete all `P0` items and release `vNext-safe`.
2. Week 2: Complete `P1` items and release `vNext-controlled`.
3. Weeks 3-4: Complete `P2` core refactor/security/tests and release `vNext-stable`.

## Risks and Mitigations
1. Regression risk in parser/executor changes.
Mitigation: add smoke tests before enabling strict mode by default.
2. UX friction from too many confirmations.
Mitigation: keep `Safe Mode` toggleable and remember user preference.
3. VBA/Excel environment variability across user machines.
Mitigation: startup diagnostics plus clear remediation messages.

## First Implementation Slice (Recommended)
1. Language setting in UI + prompt binding.
2. Command preview toggle.
3. Strict parser validation.
4. API retry/backoff.

This slice delivers the biggest safety and usability gains with minimal architecture change.
