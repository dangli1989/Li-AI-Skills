# Claude Code Adapter

Exports selected source skills into:

```text
generated\claude-code\skills
```

By default it also installs to:

```text
%USERPROFILE%\.claude\skills
```

Use `-RuntimeSkillsPath` if your Claude Code setup uses a different location.

## Commands

```powershell
.\install.ps1
.\install.ps1 -ExportOnly
.\validate-runtime.ps1
.\sync-from-runtime.ps1
```
