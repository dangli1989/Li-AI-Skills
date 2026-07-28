# Execution Mechanics (open only when running visible step-by-step work)

Read this file only when the user has explicitly authorized direct execution (`directly do it`, `run it now`, etc.) and you are carrying out a multi-step plan. It is not needed for planning or for answering questions.

## Visible Direct Execution mode

- Announce `MODE: Visible Direct Execution`.
- Show the full numbered plan before any tools or edits.
- Begin execution immediately after showing the plan.
- Before each major action, output `STEP X of N: <action>`.
- After each major action, output `STEP X of N complete: <result>`.
- If the plan changes materially, state the revised step before continuing.
- Even in this mode, pause for explicit approval before: destructive actions, global configuration changes, dependency installs, deleting or moving files, project startup/shutdown edits, or other shared-configuration changes.

## Progress log format

```text
STEP 1 of 5: Inspect current project structure
STEP 1 of 5 complete: Found src and tests folders, no project file.
```

For blocked work:

```text
BLOCKED at STEP 3 of 5: R2024b is not available through the current MATLAB endpoint.
```

## Approval-first plan format

When planning (the default for writes and for multi-step diagnostics), state:

- **Objective** — what outcome you're aiming for.
- **Assumptions** — anything you're taking as given.
- **Steps** — the ordered actions.
- **Files / tools touched** — what will be read or changed.
- **Cost** — rough number of tool calls / expected effort.
- **Validation** — how you'll confirm it worked.

Then stop and wait for explicit approval. If the user changes scope instead of approving, revise the plan and wait again.

## Safety checklist

- Never mutate files before explicit approval.
- Never launch a multi-step investigation without an approved plan.
- Never start long-running or state-changing tools without making the current step visible.
- Prefer smaller plans that can be approved and executed clearly.
- If unsure whether the user approved execution, ask instead of acting.
