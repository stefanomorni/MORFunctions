# PROJECT_PROGRESS — MORFunctions

> Lossless resume packet. Prefer this over chat history.

- **Updated**: 2026-09-27
- **Phase**: Framework-aligned, git initialized, VBA at parity

## Current truth

| Item | Value |
|------|--------|
| Project root | `%PROJECTS_ROOT%\MORFunctions\` |
| Canonical add-in | `MORFunctions.xlam` symlink → `%APPDATA%\Microsoft\AddIns\` |
| Dev workbook | `MORFunctions-dev.xlsm` (open for editing) |
| VBA on disk | 38 modules in `vba-files\` (parity verified 2026-09-27) |
| Python env | `xlpy` (`vba-config.toml`) |
| Watcher | `run-vba-watcher.bat` → `MORFunctions-dev.xlsm` |
| Publish | `run-addin-publish.bat` → `publish_addin.py` (auto-unload/reload) |
| Ribbon | No production CustomUI; boilerplate `ribbon\customUI14.xml` only |
| Git | `main` at `0be02a1` → `stefanomorni/MORFunctions` |
| Handbook | **Deep** — `PROJECT_HANDBOOK.md` |

## Do / Don't

- **Do** edit via `MORFunctions-dev.xlsm` + `vba-files\` + watcher
- **Don't** import/watch the loaded `.xlam` (IsAddin lock / corruption risk)
- **Publish** only when shipping: `run-addin-publish.bat`, then reload add-in

## Accomplishments — 2026-09-27

1. Migrated `run-addin-publish.bat` from inline `python -c` to `publish_addin.py`
2. Fixed `run-pq-watcher.bat` (removed stray `watch` subcommand)
3. Created `publish_addin.py` + `Resolve-OfficePython.ps1` symlinks via `Repair-ProjectSymlinks.ps1`
4. Extended `.gitignore` (logs, staging, secrets)
5. Initialized git repo, rebased onto existing remote (`064fe3f`), pushed `0be02a1`
6. Verified VBA tri-surface parity (38/38 components identical)
7. Fixed `Informational.bas` cp1252 mojibake (disk → workbook Safe-Sync)
8. Deepened `PROJECT_HANDBOOK.md` (Starter → Deep) with full module catalog

## Prior — 2026-09-21

1. Scaffolded project via `Initialize-OfficeProject.ps1`
2. Created `MORFunctions-dev.xlsm` from add-in copy
3. Exported 38 VBA components to `vba-files\`
4. Retargeted launchers to `-dev`; added `run-addin-publish.bat`

## Next actions

1. Feature work (user-driven)
2. Decide UDF "supported vs experimental" surface
3. Investigate hard VBA reference issue on publish/reload
4. Optional: real ribbon if needed (icons from `%INFRASTRUCTURE_ROOT%\Assets\Office-Ribbon-Icons\`)
