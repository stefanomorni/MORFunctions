# PROJECT_PROGRESS — MORFunctions

> Lossless resume packet. Prefer this over chat history.

- **Updated**: 2026-09-21
- **Phase**: Scaffold complete — ready for feature work on `-dev`

## Current truth

| Item | Value |
|------|--------|
| Project root | `%PROJECTS_ROOT%\MORFunctions\` |
| Canonical add-in | `MORFunctions.xlam` symlink → `%APPDATA%\Microsoft\AddIns\` |
| Dev workbook | `MORFunctions-dev.xlsm` (open for editing) |
| VBA on disk | 38 modules in `vba-files\` |
| Python env | `xlpy` (`vba-config.toml`) |
| Watcher | `run-vba-watcher.bat` → `MORFunctions-dev.xlsm` |
| Publish | `run-addin-publish.bat` |
| Ribbon | No prior CustomUI; boilerplate `ribbon\customUI14.xml` only |
| Git | Not initialized yet |
| Handbook | **Starter** — `PROJECT_HANDBOOK.md` |

## Do / Don't

- **Do** edit via `MORFunctions-dev.xlsm` + `vba-files\` + watcher
- **Don't** import/watch the loaded `.xlam` (IsAddin lock / corruption risk)
- **Publish** only when shipping: `run-addin-publish.bat`, then reload add-in

## Next actions

1. Feature work (user-driven)
2. Optional: init git + remote
3. Optional: real ribbon if needed (icons from `%INFRASTRUCTURE_ROOT%\Assets\Office-Ribbon-Icons\`)
4. Deepen handbook when UDF catalog / active sites are clarified
