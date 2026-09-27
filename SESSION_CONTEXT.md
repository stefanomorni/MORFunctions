# Session Context: MORFunctions

- **Date**: 2026-09-27
- **Branch**: `main` at `0be02a1`
- **State**: Framework-aligned, git initialized, VBA at parity, handbook Deep

## Pickup

- Full project state → `PROJECT_PROGRESS.md`
- Architecture + module catalog → `PROJECT_HANDBOOK.md`
- Git repo live at `stefanomorni/MORFunctions`

## Last session (2026-09-27)

- Migrated `run-addin-publish.bat` to `publish_addin.py` (auto-unload/reload)
- Initialized git repo, rebased onto remote, pushed `0be02a1`
- Verified VBA tri-surface parity (38/38 components)
- Fixed `Informational.bas` cp1252 mojibake (disk → workbook Safe-Sync)
- Extended `.gitignore` (logs, staging, secrets)

## Open

- [ ] Feature work against `-dev` + watcher
- [ ] Decide UDF "supported vs experimental" surface
- [ ] Investigate hard VBA reference issue on publish/reload
- [ ] Ribbon: keep boilerplate or design real CustomUI?
