# Li's AI Skills

Portable source-of-truth repo for reusable AI-agent skills.

This repo is intentionally separate from vendor or system-managed skill sets:

- Do not copy Codex system skills from `%USERPROFILE%\.codex\skills\.system`.
- Do not copy vendor-managed MATLAB Agentic Toolkit skills from `%USERPROFILE%\.matlab\agentic-toolkits`.
- Keep only user-authored reusable skills under `skills/`.

## Skill System Map

```mermaid
flowchart LR
  subgraph Source["Source of truth"]
    Skills["skills/*"]
    Metadata["skill.yaml"]
    References["references / scripts / assets"]
  end

  subgraph Profiles["Selectable profiles"]
    Default["default.yaml"]
    Coding["coding.yaml"]
    Presentation["presentation.yaml"]
  end

  subgraph Adapters["Tool adapters"]
    Codex["adapters/codex"]
    Claude["adapters/claude-code"]
  end

  subgraph Runtime["Runtime registrations or generated copies"]
    CodexRuntime["Codex skills runtime"]
    ClaudeRuntime["Claude skills runtime"]
    Generated["generated/*"]
  end

  Skills --> Metadata
  Skills --> References
  Skills --> Profiles
  Profiles --> Codex
  Profiles --> Claude
  Codex --> CodexRuntime
  Claude --> ClaudeRuntime
  Claude --> Generated
```

The important rule is that `skills/` is edited by humans; adapters create tool-specific runtime registrations or generated copies.

For Codex, runtime entries are junctions or symbolic links back to `skills/<skill-id>`. This keeps one real copy of each skill in the repo. Updating the repo updates the registered skill contents without maintaining duplicate runtime folders.

## Layout

```text
Li-AI-Skills/
  skills/                 # Source of truth, hand-edited
  adapters/               # Tool-specific install/sync logic
  profiles/               # Which skills are enabled for each workflow
  generated/              # Disposable generated bundles and backups
  tools/                  # Repo-level helper scripts
  docs/                   # Maintenance notes
  tests/                  # Validation and sample prompts
```

## Current Runtime Support

- Codex CLI via `adapters/codex`.
- Claude Code via `adapters/claude-code`.

The repo is adapter-based so more agents can be added later without changing the source skills.

## Common Commands

Validate the repo:

```powershell
.\tools\validate-repo.ps1
```

List skills:

```powershell
.\tools\list-skills.ps1
```

Register enabled skills into Codex:

```powershell
.\adapters\codex\install.ps1
```

Export/install enabled skills for Claude Code:

```powershell
.\adapters\claude-code\install.ps1
```

## Source-of-Truth Rule

Edit skills in this repo first. Then run the adapter install script for the target agent.

For Codex, runtime skills are registrations that point to this repo. Editing a Codex runtime skill edits this repo directly.

For copy-based adapters, if a skill was edited directly in a runtime folder, run that adapter's `sync-from-runtime.ps1` to bring the change back into this repo before continuing.
