# Ground Truth

- Title: MBD Establishment Workshop
- Subtitle: A reusable workshop story for moving from hand code to model-based design
- Use case: training-workshop
- Audience: System and software engineers evaluating how to establish an MBD workflow from existing embedded C code.
- Purpose: Reframe the source workshop into a clear training flow that shows why generic conversion steps are insufficient, what artifacts must be inspected, and how a repeatable MBD establishment path works.
- Template: public
- Output: C:\DevGit\Li-AI-Skills\generated\style-intake\li-mathworks-presentation\examples\story-first-test\MBD Establishment Workshop - Story First Li Style.pptx
- Status: generated from JSON spec; update this ledger when content, assets, or slide order changes.

## Source Facts
- The source deck is a 61-slide MBD establishment workshop created from a customer-safe example.
- The original story starts with generic hand-code-to-MBD steps, then shows why those steps need real project inputs.
- The source workshop repeatedly reviews the path: define scope, analyze workflow and software architecture, convert typical software components, convert architecture, test, and plan next steps.
- Source evidence includes C code excerpts, type/data definitions, software component boundaries, workflow diagrams, and a final next-step discussion.

## Progress Labels
- Scope
- Workflow
- Code artifacts
- Architecture
- Testing
- Next steps

## Slide Ledger
| Slide | Type | Title | Intent | Assets / gaps |
| --- | --- | --- | --- | --- |
| 1 | cover | MBD Establishment Workshop | Establish title, ownership, and delivery context. | Template cover. |
| 2 | agenda | The workshop has one through-line: turn code evidence into model decisions | Set the story around artifact-driven decisions instead of a list of topics. | No open gap recorded. |
| 3 | what-to-expect | Workshop setup: every concept is tied to a concrete artifact | Define the training rhythm without declaring a deck-wide color legend that is not needed. | No open gap recorded. |
| 4 | v-model-tool-map | The target workflow links decisions, tools, and verification | Create the high-level mental model before showing detailed code evidence. | No open gap recorded. |
| 5 | process-state-diagram | Generic conversion steps fail when project inputs are missing | Turn the early source-slide question into a visual risk argument. | No open gap recorded. |
| 6 | process-state-diagram | The workshop path resolves the missing-input problem | Show the replacement process that makes conversion actionable. | No open gap recorded. |
| 7 | code-review-excerpt | Source code is not just implementation; it exposes modeling decisions | Use a concrete code artifact so the workshop does not remain abstract. | No open gap recorded. |
| 8 | concept-artifact | Data definitions become the interface contract | Demonstrate the concept-plus-artifact pattern from the MBSE/training corpus. | Placeholder used because the real customer-safe screenshot was not provided.<br>- Add a sanitized DataBase.h or data dictionary screenshot when available. |
| 9 | artifact-map | A useful conversion review connects artifacts to decisions | Provide a repeatable review scaffold that can be reused in future customer workshops. | No open gap recorded. |
| 10 | v-model-tool-map | System Composer captures architecture after the evidence is understood | Teach best practice for sequencing System Composer work in an MBD establishment service. | No open gap recorded. |
| 11 | comparison-table | The workshop leaves decisions, not just observations | Replace a plain text table with a status-driven decision table. | No open gap recorded. |
| 12 | demo-exercise | Exercise: convert one component using the artifact map | Turn the workshop story into an action the audience can practice. | No open gap recorded. |
| 13 | recap | Reusable MBD establishment pattern | Close the training loop with a memorable reusable method. | No open gap recorded. |
| 14 | results-table | Where are we now, and what is next? | Preserve the source deck's closing intent while making ownership and next action readable. | No open gap recorded. |

## Open Inputs
- Replace generated placeholders with project-safe screenshots or exported model views if this becomes a customer-ready deck.
