# Codex Adapter

Registers selected source skills into the Codex skill runtime.

The runtime path is resolved from `CODEX_HOME` when set, otherwise from the current user's home directory.

Codex registrations are junctions or symbolic links back to this repo's `skills/<skill-id>` folders. The runtime folder should not contain copied Codex skill folders for this repo.

This adapter never installs or syncs `.system` skills.

## Commands

```powershell
.\install.ps1
.\install.ps1 -Profile coding
.\validate-runtime.ps1
.\sync-from-runtime.ps1
```

Use `-Profile coding` or another profile name to register a subset.

Restart Codex after changing skill registrations.

`sync-from-runtime.ps1` exists for legacy copied installs. In junction mode, editing a runtime skill edits this repo directly, so sync is usually unnecessary.
