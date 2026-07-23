# Sarwar PCMS — Excel Add-in

A custom Excel-DNA add-in (VB.NET) that adds:
- A **Sarwar PCMS** ribbon tab (task pane toggle, Refresh, Validate, Export
  Findings, About)
- 13 custom **EVM/schedule formulas** usable directly in any cell:
  `PCMS_SPI`, `PCMS_CPI`, `PCMS_SV`, `PCMS_CV`, `PCMS_EAC`, `PCMS_EAC_COMPOSITE`,
  `PCMS_VAC`, `PCMS_ETC`, `PCMS_TCPI`, `PCMS_TOTAL_FLOAT`, `PCMS_FREE_FLOAT`,
  `PCMS_IS_CRITICAL`, `PCMS_FORECAST_FINISH`, `PCMS_PCT_COMPLETE`
- A **DCMA-style validation engine** with 7 checks (dates, cost variance,
  EV-vs-BAC, negative float, missing predecessor/successor logic, zero
  duration, % complete range) — each check auto-skips if the sheet doesn't
  have the matching column, so it works with whatever layout you have
- A **task pane** with a color-coded results grid (red = error, yellow =
  warning) instead of a plain text dump
- **Export Findings** writes every issue to a new `PCMS_Validation` sheet
  (Row / Check / Severity / Message) so you can filter, sort, or share it
  like normal data
- **Local SQLite history** (matches Elwaha's architecture): "Save Snapshot"
  writes the active sheet's schedule/EVM rows plus current validation
  findings into a per-user SQLite database at
  `%USERPROFILE%\Documents\Sarwar_PCMS\pcms_data.sqlite`. "Load Last
  Snapshot" pulls the most recent save back into a new sheet for
  comparison against live data.
- Everything else still works **directly on the open workbook** — the
  database is additive history/versioning, not a requirement for day-to-day
  formulas or validation
- **EVM Dashboard**: one click builds a `PCMS_Dashboard` sheet with
  cumulative PV/EV/AC totals, headline SPI/CPI, and an S-curve line chart
- **Critical Path Highlighter**: one click colors every row on the active
  sheet by Total Float — red for critical (float ≤ 0), yellow for
  near-critical (float ≤ 5 days) — using plain cell fill, no data changes

## How to build (Windows + Visual Studio)

1. Install **Visual Studio 2022 Community** (free) with the **.NET desktop
   development** workload.
2. Open `Sarwar_PCMS.sln`.
3. Build in **Release** mode, platform **x64** (or **x86** to match your
   Excel bitness — check via Excel > File > Account > About Excel).
4. NuGet will automatically restore `ExcelDna.AddIn` and the Interop package
   on first build.
5. Output appears in `Sarwar_PCMS\bin\Release\net48\`:
   - `Sarwar_PCMS-AddIn64.xll` (or `-AddIn.xll` for x86) — **rename this to
     `Sarwar_PCMS.xll`** to match the installer scripts, or edit
     `ADDIN_FILE` in `Install.bat`/`Uninstall.bat` to match the real name.
   - `Sarwar_PCMS.dll`
   - Any Interop/support DLLs it pulls in

## Packaging for distribution

Create this folder layout next to `Install.bat` / `Uninstall.bat`:

```
Sarwar_PCMS_Package/
├── Install.bat
├── Uninstall.bat
└── DLL Files/
    ├── Sarwar_PCMS.xll
    ├── Sarwar_PCMS.dll
    └── (any other DLLs from the build output)
```

Zip the whole `Sarwar_PCMS_Package` folder and send it to whoever needs the
tool. They just unzip and double-click `Install.bat`.

**Important — SQLite native binaries:** since PCMS now uses
`System.Data.SQLite.Core`, the build output includes a `runtimes\` folder
(e.g. `runtimes\win-x64\native\SQLite.Interop.dll`,
`runtimes\win-x86\native\SQLite.Interop.dll`). Copy that whole `runtimes`
folder into `DLL Files\` too — `Install.bat` already has a step that copies
it automatically if present:

```
Sarwar_PCMS_Package/
├── Install.bat
├── Uninstall.bat
└── DLL Files/
    ├── Sarwar_PCMS.xll
    ├── Sarwar_PCMS.dll
    ├── System.Data.SQLite.dll
    └── runtimes/
        ├── win-x64/native/SQLite.Interop.dll
        └── win-x86/native/SQLite.Interop.dll
```

## Column headers the validator recognizes

Put these as header text in row 1 of your sheet — any subset is fine, each
check silently skips if its columns are missing:

`Start Date`, `Finish Date`, `EV`, `PV`, `AC`, `BAC`, `Total Float`,
`Predecessor`, `Successor`, `Duration`, `% Complete`

## Extending it

- **New formulas**: add a `<ExcelFunction>` method to `PcmsFunctions` in
  `PcmsEngine.vb` — it's live in Excel the next time you build, no registration
  needed.
- **New ribbon buttons**: add an entry to `Ribbon.xml`, then a matching
  `Public Sub OnXxx(control As IRibbonControl)` in `RibbonController.vb`.
- **Task pane UI**: edit `PcmsTaskPaneControl.vb` (plain WinForms — add
  labels, grids, buttons as needed).
- **Sheet logic**: `PcmsEngine.RefreshSheet` / `ValidateSheet` in
  `PcmsEngine.vb` read/write the active `Worksheet` directly via the Excel
  Interop object model — extend these for your real EVM/schedule rules
  (rollups, DCMA-style checks, etc.).

## Notes

- Uses `net48` (classic .NET Framework) because Excel-DNA add-ins target the
  Excel VSTO/COM interop model, not .NET Core/5+.
- The installer scripts register the add-in per-user (`HKCU`), add a Trusted
  Location, and optionally a Defender exclusion if run as admin — same
  pattern as your original tool, just renamed.
