# Large file process isolation

## Scope

- Keep the corefix5 + CRAfix1 roll-forward, CRA, SAD, and workbook formatting behavior unchanged.
- Run each selected subject in a separate spawned process so openpyxl work does not block the GUI process.
- Keep pause and stop behavior at subject boundaries. The active subject is never force-killed.
- Continue later subjects after one child process fails.
- Emit a processing heartbeat every 15 seconds and release child-process memory after each subject.

## Locked files

The following files have no diff from tag `corefix5-crafix1-20260716`:

- `roll_forward_core.py`
- `cra_support.py`
- `subjects_config.json`

## Validation

- CRA regression: 31 passes, 0 failures.
- Existing roll-forward regression: 82 passes, 0 failures.
- Real UexpVCVD prior workbook, 5.77 MB: success; GUI event loop remained responsive.
- Real J1 prior workbook, 8.01 MB: success; GUI event loop remained responsive.
- Real K1 prior workbook, 12.97 MB: success; GUI event loop remained responsive.
- Generated workbook ZIP/XML package checks: no failures; all three outputs load successfully.
- Direct versus isolated J1 output: only run timestamps and output path differ; business cells and styles match.
- Real N prior workbook, 25.92 MB: GUI/process isolation remained healthy, but the unchanged core did not finish within the 20-minute test command limit. This is retained as a performance boundary rather than changing validated roll logic.

## Build

`dist/AuditRollForward_latestui_corefix5_crafix1_largefile1.exe`
