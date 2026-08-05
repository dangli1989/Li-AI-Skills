# Speech

## Slide 1 - MBD Establishment Workshop
This workshop is about turning an existing hand-code workflow into a repeatable model-based design establishment process. The deck is organized around decisions and artifacts, not around copying every original slide.

## Slide 2 - The workshop has one through-line: turn code evidence into model decisions
We start with the problem: the usual conversion checklist sounds reasonable, but it does not tell us what to do with real project constraints. The rest of the workshop shows how source code, architecture, workflow, and test evidence drive the conversion decisions.

Build sequence:
- Reveal the four agenda points in order.

## Slide 3 - Workshop setup: every concept is tied to a concrete artifact
The workshop works because every idea is connected to a project artifact. When we introduce a method, we immediately show the evidence that justifies it and the model action that follows.

Build sequence:
- Show method first.
- Reveal evidence, modeling action, and practice as the instructional rhythm.

## Slide 4 - The target workflow links decisions, tools, and verification
This is the target pattern we want the audience to remember. Scope decides what is in or out. Architecture turns boundaries into System Composer structure. Behavior becomes Simulink implementation. Requirements and tests make the result reviewable.

Build sequence:
- Reveal the path from scope to tests.
- Reveal tool labels with each phase.
- Reveal the artifact and repeatability callouts.

## Slide 5 - Generic conversion steps fail when project inputs are missing
The first teaching point is that a checklist does not produce a useful model by itself. Each step depends on specific input evidence. If the evidence is missing, the model is either blocked or built at the wrong level of abstraction.

Build sequence:
- Reveal the generic steps.
- Add state labels to show why each step is not actionable yet.

## Slide 6 - The workshop path resolves the missing-input problem
The workshop solves the problem by changing the order. We do not model first. We define the scope, analyze the workflow, inspect code artifacts, then create architecture and tests from evidence.

Build sequence:
- Reveal scope and workflow.
- Reveal code artifacts as the bridge to modeling.
- Reveal architecture, testing, and next steps.

## Slide 7 - Source code is not just implementation; it exposes modeling decisions
This code excerpt is useful because it shows several modeling decisions at once. The function sequence suggests behavior boundaries. The RTDB access suggests interfaces. The status bit assignment suggests requirements and test observations.

Build sequence:
- Show the code excerpt first.
- Reveal each callout as a modeling decision.

## Slide 8 - Data definitions become the interface contract
The data-definition slide turns source code into an interface contract. The point is not to paste every typedef. The point is to decide what the model owns, what it exposes, and what tests need to observe.

Build sequence:
- Reveal the three concept cards.
- Reveal the artifact placeholder and the own/expose/verify callouts.

## Slide 9 - A useful conversion review connects artifacts to decisions
This map is the reusable working view. Every artifact has to answer a decision question. If an artifact does not answer a decision, it belongs in backup material, not the main workflow.

Build sequence:
- Reveal the first three source-side artifacts.
- Reveal architecture, tests, and gaps as the target-side decisions.

## Slide 10 - System Composer captures architecture after the evidence is understood
System Composer is most effective after the team understands the source evidence. We use code and workflow analysis to decide components and interfaces, then move behavior into Simulink and verification into tests.

Build sequence:
- Reveal source functions and System Composer components.
- Reveal interfaces, behavior, and test evidence.

## Slide 11 - The workshop leaves decisions, not just observations
This table is intentionally not just a list. It shows the decision, the evidence supporting it, whether the decision is validated, and the next action. That makes the slide useful even when someone reads it later.

Build sequence:
- Reveal the decision table.
- Call attention to status and next-action columns.

## Slide 12 - Exercise: convert one component using the artifact map
The exercise uses one representative component. The audience starts with source evidence, drafts the modeling boundary, and ends by defining what would prove the conversion is correct.

Build sequence:
- Reveal Start, Build, and Check as a three-step exercise.

## Slide 13 - Reusable MBD establishment pattern
The reusable pattern is evidence first, architecture second, behavior third, verification always. Missing information is not ignored; it is tracked as a project gap.

Build sequence:
- Reveal the four takeaways in order.

## Slide 14 - Where are we now, and what is next?
The close should make ownership clear. MathWorks packages the reusable method and demo assets. The customer team provides project-safe evidence. Together, the next step is to run one component through the full workflow.

Build sequence:
- Reveal owner rows.
- End on the joint next step.
