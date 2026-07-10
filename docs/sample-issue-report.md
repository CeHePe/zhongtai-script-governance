# Sample GitHub-Safe Issue Report

## Task

- Report: `<sanitized report name>`
- Metric: `<sanitized metric name>`
- Period: `YYYY-MM`
- Dimension: `<dimension>`

## Result

- Status: validated
- Local data files used: described by category only, not by raw filename
- Script: `scripts/validation/<validation_script>.py`

## Detail

| Dimension | Report Value | Expected Value | Difference | Result |
| --- | ---: | ---: | ---: | --- |
| P000001 xiang-mu-a | 100.00 | 100.00 | 0.00 | pass |
| P000002 xiang-mu-b | 88.00 | 90.00 | -2.00 | fail |

## Notes

- Project rows show project code plus project-name pinyin only.
- Professional-company rows must show pinyin only.
- Region, line, space, and other higher dimensions may show original names only when they are not customer-identifying.
- Do not paste full indicator-list field descriptions, raw calculation logic, source workbook filenames, screenshots, or source workbook contents.
