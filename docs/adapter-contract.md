# Adapter Contract

Adapters convert source skills into a target agent runtime format.

## Rules

- Read from `skills/`.
- Select skills from `profiles/*.yaml`.
- Write generated output only under `generated/<adapter>/` or the target runtime path.
- Prefer runtime registrations that point to source skills when the target agent supports them. Use copied runtime folders only when the target agent requires copies.
- Do not read or copy vendor-managed skill sets.
- Do not modify vendor-managed MATLAB Agentic Toolkit installs.
- Do not modify Codex `.system` skills.

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

Codex should register selected skills with junctions or symbolic links instead of copied folders, so the repo remains the only maintained copy.
