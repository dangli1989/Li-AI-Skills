# Validation Levels

Select the smallest level that gives useful confidence for the change.

| Level | Name | Use |
| --- | --- | --- |
| 0 | Skipped | User explicitly requests no validation, or the change is docs-only and review is enough. |
| 1 | Static/reference scan | Renames, reference updates, doc/API table checks, or low-risk refactors. |
| 2 | MATLAB static/unit check | MATLAB class/function changes with focused behavior or syntax risk. |
| 3 | MATLAB workflow test | Package workflow changes that can be exercised without full Simulink compile. |
| 4 | Simulink compile/harness test | Changes that affect Simulink models, harnesses, ports, or Test Manager behavior. |
| 5 | Full regression workflow | Broad refactors, behavior redesign, or release-candidate validation. |

For repos constrained to a specific MATLAB version, use that version only.
