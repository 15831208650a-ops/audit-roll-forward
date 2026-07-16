# Corefix5 Change Record

Date: 2026-07-16
Base: `corefix4-baseline-20260716` (`ef1a8aca8e8beb5c599020c48f120969a6deebf7`)

## Scope

- Suppress the optional PMTE Level/RP missing warnings.
- Always roll the six C.00 BKD account identity fields by dynamic headers.
- Roll the labeled C.03 cutoff period and rationale only when wording is enabled.
- Roll the labeled J.03 selection rationale and Note only when wording is enabled.
- Extend L2.01.1 rows through the real AM business boundary, including formulas,
  array formulas, totals, downstream references, styles, and data validations.
- Roll all L2.01.1 Note responses by occurrence.
- Roll populated L2.00 Lead adjustment content while retaining template labels.

## Verification

- Real prior input: `常规样本/吉安`
- Fresh output: `tmp_live_integrated_corefix5_final`
- Regression: 82 passed, 0 failed.
- ZIP/XML/relationship checks passed for the high-risk output workbooks.
- Excel read-only automation opened C, J1, K1, L1, L2, M, Q1, Uexp, and
  UexpVCVD successfully.
- N still fails Excel COM open; the same failure was reproduced from the untouched
  corefix4 baseline and is not introduced by this change set.

## Package

- `dist/AuditRollForward_latestui_corefix5.exe`
- SHA256: `309B5BE31A1B0B6F2F76D19C4A3EA22F625A537E00FA0889568EE9348932C47A`
- The default exe and corefix4 suffix exe were not overwritten.
