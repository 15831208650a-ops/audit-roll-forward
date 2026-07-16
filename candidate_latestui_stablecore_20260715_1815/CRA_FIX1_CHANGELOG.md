# CRA Fix1 Change Record

Date: 2026-07-16
Base: `corefix5-20260716` (`a1eb3514f8a8f96b086e15ad859e390cbb5a3cf1`)

## Scope

- Support both 8-column and 9-column CRA/Canvas pasted layouts.
- Map C cash records to `C` without using loose single-letter substring matches.
- Map VC, VD, and VC&VD records to `UexpVCVD` while retaining the sales or
  administrative expense account name.
- Preserve the confirmed applicability rule, including N and N/A writing N/A.
- Mark saved previews from older parser versions as stale and prevent them from
  being written until the user parses the current source text again.
- Keep all parsed records visible and editable while adding numeric sorting,
  search, subject/status filters, and an exception-only filter.

## Verification

- Dedicated CRA regression: 31 passed, 0 failed.
- Existing roll-forward regression: 82 passed, 0 failed.
- Five locally saved real CRA source texts: no C or UexpVCVD mapping errors.
- Packaged executable startup: passed.
- The packaged `roll_forward_core` bytecode matches corefix5 exactly.

## Package

- `dist/AuditRollForward_latestui_corefix5_crafix1.exe`
- SHA256: `D8AB10F2570C4CEC76FC354A4EE583ABBEAA110FF057E4D6E3716BF2C476FDF9`
- Large-workbook process isolation is intentionally deferred to the next,
  separate delivery.
