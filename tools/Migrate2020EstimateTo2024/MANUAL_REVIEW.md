# Manual Review Column

When `reports/missing-resources.xlsx` exists, `Precheck` and `Apply` read column L as the manual review decision.

- `1`: use `best_candidate_code`.
- `0`: keep this row as a supplement material or supplement machine.
- Any other numeric value: use that value as the target 2024 resource code.

The generated `reports/manual-overrides.csv` records every row replaced by the manual review column, including the source name and the target 2024 name for review.
