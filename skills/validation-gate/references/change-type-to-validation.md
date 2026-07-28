# Change Type To Validation

| Change Type | Minimum Validation |
| --- | --- |
| Docs-only design note update | Level 0 or 1 |
| Reference rename or class rename | Level 1 |
| Public class/property/method signature change | Level 2 unless the user explicitly skips |
| Function behavior change | Level 2 or 3 |
| Objective-level workflow behavior | Level 3 minimum when tests exist |
| Harness/model/Test Manager interaction | Level 4 |
| Broad refactor or release candidate | Level 5 when available |

If the user says not to test, do not run validation. Report the selected level that would normally apply and that it was skipped by request.
