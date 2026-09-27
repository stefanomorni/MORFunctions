# Session Context: MORFunctions

- **Date**: 2026-09-21
- **Branch**: _(no git repo yet)_
- **State**: Scaffolded + first VBA export; publish-out pipeline ready

## What This Project Is

`MORFunctions.xlam` is a **universal Excel add-in** (UDFs / helpers: web scrape, JSON, charts, LUKB, worksheet functions, etc.). It is typically loaded in every Excel session alongside `MORProcedures.xlam`.

## Architecture

- **Canonical add-in**: `MORFunctions.xlam` — symlink → `%APPDATA%\Microsoft\AddIns\MORFunctions.xlam` (via Office Junctions)
- **Dev workbook**: `MORFunctions-dev.xlsm` — visible `.xlsm` for active development (created 2026-09-21 from a copy of the add-in)
- **Scripts**: symlinks → `Office-Automation-Framework\scripts\` (single-writer SSoT)
- **Launchers**: root `run-*.bat` — all target `MORFunctions-dev.xlsm`
- **Python**: `vba-config.toml` → `python_env = xlpy`

> **Single-writer rule**: script edits go only to `Office-Automation-Framework/scripts/`.
> See `Office-Automation-Framework/docs/SINGLE-WRITER-RULE.md`.

## Add-In Exception — Publish-Out Pipeline

**Never run the VBA watcher or import into `MORFunctions.xlam` while it is loaded.**
Loaded add-ins have `IsAddin=True`; direct VBA import can corrupt the project.

### Development workflow
1. Open `MORFunctions-dev.xlsm` in Excel
2. Run `run-vba-watcher.bat` — live sync with `vba-files\`
3. Edit modules on disk or in VBE; watcher keeps parity
4. When ready to ship: `run-addin-publish.bat`
   - Exports VBA → `vba-files\`
   - SaveAs `xlAddIn` → `_stage_MORFunctions.xlam`
   - Copies stage → `MORFunctions.xlam` (canonical)
   - Reload add-in in Excel (uncheck/recheck or restart)

## Launchers

| File | Purpose |
|---|---|
| `run-vba-export.bat` | Export VBA from dev `.xlsm` → `vba-files\` |
| `run-vba-import.bat` | Import `vba-files\` → dev `.xlsm` |
| `run-vba-watcher.bat` | Live bidirectional watcher (dev `.xlsm` only) |
| `run-pq-*` | Provisioned; add-ins usually have no PQ |
| `run-ribbon-export.bat` / `run-ribbon-import.bat` | Fluent Ribbon ↔ `ribbon\` |
| `run-addin-publish.bat` | Publish: dev `.xlsm` → staged → `MORFunctions.xlam` |

## Accomplishments — 2026-09-21
1. Symlink to canonical add-in already present
2. `Initialize-OfficeProject.ps1` scaffold (folders, script symlinks, `xlpy`)
3. Created `MORFunctions-dev.xlsm` via copy+SaveAs (installed add-in left loaded)
4. Exported **38** VBA components → `vba-files\`
5. Retargeted launchers/tasks to `-dev`; added `run-addin-publish.bat`

## Open / Next
- [x] Ribbon check — no production CustomUI; boilerplate only
- [x] Starter `PROJECT_HANDBOOK.md` written 2026-09-21
- [ ] Init git repo + GitHub remote (`stefanomorni/MORFunctions`) when ready
- [ ] First feature work against `-dev` + watcher
