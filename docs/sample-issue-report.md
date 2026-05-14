# Sample GitHub-Safe Issue Report

## Task

- Report: 半收付归母净利润202512项目
- Metric: 示例指标_项目
- Period: 2025-12
- Dimension: 项目

## Result

- Status: validated
- Local data files used by name only: `JKS_数据中台二期_指标清单.xlsx`, `项目查询.xlsx`, `<source report>.xlsx`
- Script: `scripts/validation/<validation_script>.py`

## Detail

| Dimension | Report Value | Expected Value | Difference | Result |
| --- | ---: | ---: | ---: | --- |
| P000001 xiang-mu-a | 100.00 | 100.00 | 0.00 | pass |
| P000002 xiang-mu-b | 88.00 | 90.00 | -2.00 | fail |

## Notes

- Project rows show project code plus project-name pinyin only.
- Professional-company rows must show pinyin only.
- Region, line, space, and other higher dimensions may show original names.
- Do not paste full indicator-list field descriptions or attach source workbooks.
