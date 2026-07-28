---
name: validation-gate
description: Validation discipline for MATLAB package/tool development. Use when changing MATLAB code, classes, public APIs, workflow behavior, refactors, or objective-level implementation so the validation level is selected, run, skipped explicitly, and reported.
---

# Validation Gate

## Goal

Make validation explicit for MATLAB package/tool changes, scaled to the risk and blast radius.

## Scope

Use this skill for MATLAB package and tool development, including objective-level workflows and function/class implementation. It applies to public API changes, behavior changes, refactors, renames, and config changes.

For Simulink model construction or harness authoring, use Simulink-specific validation skills as the primary workflow. Use this skill only for the MATLAB package/code side of that work.

## Workflow

1. Classify the change type and blast radius.
2. Select a validation level before or with implementation.
3. If the user explicitly says not to test, record validation as skipped and state residual risk.
4. Run the selected validation after implementation when allowed.
5. Report the validation level, commands/tools used, result, skipped checks, and residual risk.

## References

- Read `references/validation-levels.md` to select the validation level.
- Read `references/change-type-to-validation.md` when mapping a change type to a minimum validation level.
- Read `references/validation-report-template.md` before the final response for a validated or intentionally unvalidated change.
