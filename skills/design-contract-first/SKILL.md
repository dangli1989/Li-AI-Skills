---
name: design-contract-first
description: Guardrail for MATLAB package/tool development. Use when changing MATLAB classes, public APIs, workflow behavior, configuration fields, or objective-level implementation plans so code stays traceable to an accepted design contract and single source of truth.
---

# Design Contract First

## Goal

Keep MATLAB package/tool work traceable to an accepted design contract before implementation starts.

## Scope

Use this skill for MATLAB package and tool development, including objective-level workflow changes and function/class implementation. Apply it strongly to public classes, properties, methods, config fields, package behavior, renames, and refactors.

Do not use it as the primary workflow for Simulink model construction, presentation work, exploratory analysis, or one-off generated artifacts unless those changes also affect the MATLAB package API or behavior contract.

## Workflow

1. Identify whether the request is objective-level work or function/class-level work.
2. Identify the authoritative requirement, workflow note, class/property/method table, or config table.
3. If no accepted contract exists, update the design note before implementing behavior.
4. State the contract being implemented: accepted rows, workflow steps, affected classes/files, public API changes, config changes, and out-of-scope items.
5. Implement only the accepted contract. If code, docs, tests, and chat disagree, stop and resolve the design contract first.
6. Update the design note when a public API, config field, or workflow behavior changes.

## References

- Read `references/source-of-truth.md` when there is any conflict between chat, docs, tests, code, or another branch.
- Read `references/change-contract-template.md` before nontrivial package/class/API changes.
- Read `references/testmaintenance-contract.md` for TestMaintenance package work.
