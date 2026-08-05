# AI Maintainer Notes

This repo is the source of truth for Li's reusable AI-agent skills.

## Rules

- Use the `do-it-with-a-plan` workflow for repo maintenance. Plan first and wait before file writes, install actions, runtime registration changes, or multi-step investigations.
- Edit source skills only under `skills/`.
- Do not edit generated output as source.
- Do not copy or vendor Codex `.system` skills.
- Do not copy or vendor MathWorks toolkit skills.
- Do not hard-code machine paths, user names, or checkout locations. Resolve paths from script location, `$env:CODEX_HOME`, `$HOME`, or explicit parameters.
- Keep `skill.yaml` portable and tool-neutral.
- Update `profiles/*.yaml` when adding, removing, or changing which skills should be active for a workflow.
- For Codex, register skills by junction or symbolic link from the runtime skill folder to `skills/<skill-id>`. Do not maintain copied Codex runtime skill folders.
- Run `tools/validate-repo.ps1` after changes.

## Source Model

```text
skills/<skill-id>/         real source copy
<codex-home>/skills/<id>   runtime registration pointing to the source copy
generated/                 disposable backups and generated artifacts
```

If a runtime registration needs to change, update the adapter script or profile first, then run the adapter intentionally.
