---
name: do-it-with-a-plan
description: >
  Always-on workflow guardrail; apply on every turn. Plan first and WAIT for explicit
  approval before (a) any file write or state-changing action — edit, commit, push, install,
  start/stop a service, run a simulation, launch MATLAB/Simulink; or (b) any multi-step
  diagnostic or investigation even if read-only — troubleshooting, root-cause hunting, or
  chained run→read→decide loops. A pure text read answered in text proceeds directly, and a
  specific instruction to do a specific thing is itself the approval for that action. When
  unsure, treat it as an investigation and plan first. "directly do it" / "run it now"
  triggers visible step-by-step execution (see references/execution.md).
---

# Do It With A Plan

The core rule lives in the description above — apply it every turn. This card adds the one clarification, quick examples, and the trigger phrases.

## Clarification: a direct instruction is its own approval

Plan-and-wait exists to stop me from expanding scope or starting an investigation you didn't ask for. When you give a specific instruction to do a specific thing ("update the skill", "install X", "rename this file"), that instruction IS the approval — do it, don't re-ask. Plan first only when I would go beyond what you asked, or when the work is a multi-step investigation (rule of thumb: anything over ~2 read-only tool calls counts as an investigation).

An instruction approves only its literal action. If fulfilling it requires choosing an implementation among alternatives, or extra state changes the user hasn't seen, that choice is a new plan — present it first.

## Quick examples

- "Why is X failing?" → diagnosis / investigation → **plan first**, don't start running commands.
- "What does this file do?" → pure text read → **just answer.**
- "Update the skill to say Y." → specific instruction → **just do it** (that's the approval).
- "Look into why the build is slow." → open-ended investigation → **plan first**, then wait.

## Approval phrases (proceed only on a clear yes)

`go ahead` · `continue` · `proceed` · `execute` · `do it` · `approved` · `yes, run it` · equivalent.

A reply that merely reports status or answers a sub-question ("I opened it", "done", "it's installed") is NOT approval — restate readiness and wait for a clear yes.

## Direct-execution phrases (switch to visible step-by-step)

`directly do it` · `go ahead and do it` · `no need to wait` · `run it now` · `execute without asking` → then follow `references/execution.md`.

## Hard-stop phrases (no state-changing action; answer or plan only)

`plan only` · `do not create` · `do not edit` · `do not run` · `wait` · `pause` · `stop` · `just tell me` · `only tell me` — when used as commands, not casual speech ("wait, I meant…" is not a hard stop).

## When planning

State: objective, key steps, files/tools touched, rough number of tool calls, and how you'll validate. Then stop and wait. Full plan and progress-log formats are in `references/execution.md`.
