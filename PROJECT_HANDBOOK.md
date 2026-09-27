---
project: MORFunctions
type: project_handbook
handbook_maturity: starter
habits: on
created: 2026-09-21
updated: 2026-09-21
source_code:
  - "%PROJECTS_ROOT%\\MORFunctions"
---

# Project Handbook: MORFunctions

> **Habits: ON** · **Handbook: Starter**  
> Durable project handbook (intent, constraints, architecture). Not a session log.  
> Session deltas belong in `SESSION_CONTEXT.md`. Gaps marked `Unknown — needs Stefano`.

## 1. Purpose (plain language)

- **What this is for**: Ship and maintain `MORFunctions.xlam` — a **universal Excel add-in** that exposes shared VBA helpers and worksheet-callable utilities (UDFs) across Stefano’s Excel sessions (web/JSON tooling, time/file helpers, charts, LUKB-related helpers, site/credential config, etc.).
- **Who it serves**: Stefano (personal/professional Excel automation), often alongside workbooks such as PMS / GP workflows that may call into the add-in.
- **Analogy**: A toolbox that stays loaded in Excel; sister add-in `MORProcedures.xlam` is the ribbon/masking toolkit, while this one is the **function library**.

## 2. Use cases & workflows

- **Must support**:
  - Keep a git-friendly text copy of all VBA (`vba-files\`) in sync with a visible **dev** workbook.
  - Publish a tested build into the canonical installed add-in (`%APPDATA%\Microsoft\AddIns\MORFunctions.xlam`) without corrupting a loaded `IsAddin=True` project.
  - Zero-install end-user runtime: consumers need only Excel + the `.xlam` (no Python on end machines).
- **Nice to have**:
  - Fluent Ribbon UI (none in production package yet — boilerplate only).
  - Formal module inventory / public UDF catalog for agents and docs.
- **Out of scope / non-goals**:
  - Replacing `MORProcedures` (mask/unmask ribbon and related utilities).
  - Live-watching or importing into the **loaded** `.xlam`.
  - Power Query pipelines (provisioned folders/launchers only; add-in is VBA-centric).

## 3. Constraints & restrictions

- **Hard rules**:
  - **Publish-out only** for the installed add-in: develop in `MORFunctions-dev.xlsm`; ship via `run-addin-publish.bat` (stage-then-copy). See `Office-Automation-Framework/docs/SINGLE-WRITER-RULE.md`.
  - **Never** `watch` / `import-all` against `MORFunctions.xlam` while installed/loaded (`IsAddin` lock → corruption risk).
  - VBA sources on disk **must use CRLF**; LF-only breaks `VBComponents.Import`.
  - Framework scripts live only in `Office-Automation-Framework\scripts\`; this project’s `scripts\` entries are **NTFS symlinks** (GoodSync: Copy as is).
  - Binary Office files are not source of truth for git (`*.xlam` / `*.xlsm` gitignored); text in `vba-files\`, `ribbon\`.
  - Canonical binary path: `%APPDATA%\Microsoft\AddIns\MORFunctions.xlam` (project root holds a symlink via Office Junctions).
- **Soft preferences**: Prefer editing modules as text under agent/watcher; keep production add-in reload explicit after publish.
- **Forbidden approaches**: Direct COM import into the loaded `.xlam`; duplicating framework scripts inside this repo; inventing ribbon icons / random `imageMso`.
- **Office ribbon icons** (if/when CustomUI is real): `%INFRASTRUCTURE_ROOT%\Assets\Office-Ribbon-Icons\` only (Pure Flat; see README). No decorative Business-pack icons.

## 3b. Runtime / tooling

> End-user Excel needs **no** Python. Python is for **developer** sync/publish only.

- **Python / mamba env**: `xlpy`
- **How to run**: project `run-*.bat` (via `scripts/Resolve-OfficePython.ps1`) or `mamba run -n xlpy python …`
- **Config SSoT**: `vba-config.toml` → `python_env = "xlpy"`; `file = 'MORFunctions-dev.xlsm'`
- **Resolution order for agents/scripts**:
  1. `vba-config.toml` `python_env`
  2. else `OFFICE_PYTHON_ENV`
  3. else framework fallback `xlpy`

## 4. Architecture now

- **Where code lives**: `%PROJECTS_ROOT%\MORFunctions\`
  - `vba-files\` — 38 exported components (`.bas` / `.cls`)
  - `MORFunctions-dev.xlsm` — visible edit/test surface
  - `MORFunctions.xlam` — symlink to installed add-in
  - `ribbon\customUI14.xml` — boilerplate only (no production CustomUI found on first export)
  - `scripts\` — symlinks → Office-Automation-Framework
  - Launchers: `run-vba-*`, `run-ribbon-*`, `run-pq-*`, `run-addin-publish.bat`
- **Where docs live**: `PROJECT_HANDBOOK.md`, `PROJECT_PROGRESS.md`, `SESSION_CONTEXT.md` (this folder). No separate vault project note yet.
- **Key component clusters** (from export; not a full API map):
  - Web / JSON: `WebScraping`, `WebDownloader`, `WebAuth`, `JsonConverter`, `JsonHelper`, `Cookies`, `SiteConfig`, `SiteManagement`
  - Time / files: `modTimeConversionFunctions`, `modGetSetFileTimes`, `modWorksheetFunctions`, `Tempo`
  - Charts / app events: `Grafici`, `ChartEvents`, `CAppEvent`, `CChartEvent`
  - Domain helpers: `lukb`, `Investing`, `Calcoli`, `Commissioni`, `Informational`, …
  - Infra: `Shell`, `Registro`, `ManageDLL`, `Ambiente`, `Aggiornamenti`
- **External systems**: Excel COM; optional web sites behind `SiteConfig` / credentials — **Unknown — needs Stefano** (which sites are still active). Sister: `MORProcedures.xlam`.

## 5. Decision log (append-only)

| Date | Decision | Why | Still valid? |
|------|----------|-----|--------------|
| 2026-09-21 | Develop in `MORFunctions-dev.xlsm`; publish-out to `.xlam` | Same sanctioned add-in pattern as MORProcedures; avoids IsAddin corruption | Yes |
| 2026-09-21 | `python_env = xlpy` in `vba-config.toml` | Matches Office COM/watcher tooling used across OAF children | Yes |
| 2026-09-21 | Create `-dev` via file copy of installed `.xlam` + SaveAs `.xlsm` | Installed add-in was referenced/locked; copy avoids uninstalling mid-session | Yes |

## 6. Open questions & risks

- Exact public UDF / API surface Stefano considers “supported” vs experimental/test modules (`modTest`, `TestSites`, …) — **Unknown — needs Stefano**.
- Whether any workbook holds a hard VBA **reference** to `MORFunctions` that complicates publish/reload — observed “referenced by another workbook” when reopening the installed file.
- Git repo + GitHub remote not created yet.
- Ribbon: keep boilerplate unused, or design a real CustomUI?

## 7. Validated commands / recipes

```text
# From %PROJECTS_ROOT%\MORFunctions\ with MORFunctions-dev.xlsm open:
run-vba-watcher.bat          # live sync <-> vba-files\
run-vba-export.bat           # force export
run-vba-import.bat           # force import into -dev only
run-addin-publish.bat        # -dev -> stage -> MORFunctions.xlam
run-ribbon-export.bat        # CustomUI <-> ribbon\ (boilerplate today)
```

## 8. Related docs

- USER_GUIDE: `Office-Automation-Framework/docs/USER_GUIDE.md` (framework)
- SINGLE-WRITER / add-in exception: `Office-Automation-Framework/docs/SINGLE-WRITER-RULE.md`
- PROJECT_PROGRESS: `PROJECT_PROGRESS.md`
- SESSION_CONTEXT: `SESSION_CONTEXT.md`
- Sister project: `%PROJECTS_ROOT%\MORProcedures\`
- LEARN / KB: `LEARN-VBA-Import-CRLF-And-AddIn-Architecture.md` (CRLF + add-in symlink architecture)
