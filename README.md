# Claude Formula Compiler for Excel

A formula *compiler* (not a chatbot): you type a shorthand verb into a cell, Claude
resolves it **once at authoring time**, and a native Excel formula is written into
the cell. After that the workbook recalculates at native speed forever, with no API
dependency.

See `CLAUDE_EXCEL_TOOL_SPEC.md` (provided separately) for the full design.

## Modules (`vba/`)

| File | Build step | Purpose |
|------|-----------|---------|
| `Sandbox.bas` | 1 | Filesystem path guard. Restricts all disk access to `C:\Users\colet\Documents\INDEX`. Includes `Test_Sandbox` self-test. |

## Build progress

- [x] Gate 0 — API credentials verified
- [x] Step 1 — Path guard + unit tests
- [ ] Step 2 — Undo snapshot + restore
- [ ] Step 3 — API plumbing (cache + retry)
- [ ] Step 4 — Catalog indexer
- [ ] Step 5 — `IMM` thunk + `CLAUDE()` macro + validator
- [ ] Step 6 — Resolution cache
- [ ] Step 7 — `LOOK`, `TIER`
- [ ] Step 8 — Notes, `EXPLAIN`, remaining verbs

## How to test a module

1. In Excel press `Alt`+`F11` (VBA editor), then `Ctrl`+`G` (Immediate Window).
2. `Insert` > `Module`, paste the module code.
3. Press `F5`, choose the module's `Test_*` sub, read the Immediate Window output.
