# Corefix4 Baseline

This directory is the source baseline for `AuditRollForward_latestui_corefix4.exe`.
Future roll-forward changes should start here while preserving the current UI and project framework.

## Included Source

- `main_gui.py`: current UI and workflow
- `roll_forward_core.py`: roll-forward processing core
- `cra_support.py`: CRA parsing, applicability, and workbook writing
- `subjects_config.json`: subject-specific processing rules
- `dialog_helper.py`: isolated file dialog helper
- `llm_enhancement.py`: optional LLM review support
- `build.py`: packaging entry point
- `goal_regression_corefix3.py`: integrated regression checks for this baseline

## Verified State

- Integrated regression: 53 checks passed, 0 failed.
- Excel automation opened the generated Q1, VCVD, L2, and Uexp workbooks without repair mode.
- The suffix test executable passed a basic startup check.
- The default `dist/AuditRollForward.exe` was not replaced by the corefix4 build.

## Repository Safety

Customer samples, generated workbooks, build caches, Python caches, and executable artifacts are intentionally excluded from this source snapshot.
