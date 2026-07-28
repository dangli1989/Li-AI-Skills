# Adapter Contract

Adapters convert source skills into a target agent runtime format.

## Rules

- Read from `skills/`.
- Select skills from `profiles/*.yaml`.
- Write generated output only under `generated/<adapter>/` or the target runtime path.
- Do not read or copy vendor-managed skill sets.
- Do not modify `%USERPROFILE%\.matlab\agentic-toolkits`.
- Do not modify `%USERPROFILE%\.codex\skills\.system`.

## Adding A New Adapter

Create:

```text
adapters/<tool>/
  adapter.yaml
  install.ps1
  sync-from-runtime.ps1
  validate-runtime.ps1
```

The adapter should preserve each skill folder shape unless the target tool requires conversion.
