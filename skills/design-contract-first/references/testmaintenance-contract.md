# TestMaintenance Contract

For the TestMaintenance package, use these repo-local artifacts as the design authority:

- `reports/TestMaintenancePlannedApiDiscussion.md` for class, property, method, and config API.
- `reports/TestCaseFirstWorkflowRequirementAndDesign.md` for workflow behavior.

## Branch Rule

Other branches can provide implementation evidence, Excel examples, API experiments, or regression comparisons. They are not API authority unless the design table in the current branch explicitly accepts the idea.

## Scope Rule

This contract applies to MATLAB package work under `src/+TestMaintenance`, especially class folders, public methods, issue objects, workflow checks, config behavior, and docs tied to the public API.

Use Simulink-specific skills for model construction, harness creation, model compile, and model project setup unless the task also changes the MATLAB package API.
