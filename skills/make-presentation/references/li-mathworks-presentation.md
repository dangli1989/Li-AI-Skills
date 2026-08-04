# Li's MathWorks Presentation

Use this reference only when the user asks for "Li's MathWorks Presentation", "MathWorks style", or a deck that should match Li's prior MathWorks customer, internal, or workshop presentation style.

## Required Questions

Before building:

- Ask whether the output should be `PPTX` or `HTML` if the user did not specify.
- Ask whether the deck is confidential. Use the confidential template only when the user says it is confidential.
- Ask or infer the use case: internal project sharing, customer wrap-up or technical deep-dive, training/workshop, or general technical presentation.

## Generation Contract

Build Li's MathWorks Presentation as a story-first deliverable, not as a styling pass over slides.

- For normal generation, draft the outline in chat and wait for approval before creating slides.
- If the user explicitly asks for fast/direct generation, proceed without waiting for outline approval, but still create the required notes and track assumptions.
- Always create `ground-truth.md` and `speech.md`, even for fast/direct generation.
- Use `ground-truth.md` as the source-of-truth ledger: purpose, audience, use case, source facts, working or approved outline, slide intent, asset inventory, missing inputs, placeholders, progress labels, decisions, work completed, and final output path.
- Use `speech.md` as the presenter-facing talk track mapped to the final slide order. The final delivery is `deck.pptx + speech.md` or `deck.html + speech.md`.
- Put missing screenshots, data, project facts, customer-safe names, and asset gaps in `ground-truth.md`. Also mention unresolved gaps in the final response.
- Decide content, story, assets, slide intent, and speech before applying MathWorks template styling, Li's color/shape conventions, animation, or transitions.
- Prefer diagrams, screenshots, model/code excerpts, evidence tables with emphasis, and staged builds over text-only slides. A text-only slide needs a clear story reason recorded in `ground-truth.md`.
- Apply Li's style because it supports the use case and message. Do not copy a visual pattern only because it appeared in an example deck.

## Official Template Assets

For PPTX output, start from the official MathWorks template assets:

- Non-confidential: `assets/li-mathworks-presentation/pptx/public.pptx`
- Confidential: `assets/li-mathworks-presentation/pptx/confidential.pptx`

Do not recreate the PPTX base from scratch. Preserve the template master, layouts, theme colors, template font usage, slide size, and confidential markings. Delete resource/instruction slides from the copied template before final delivery.

For PPTX, use this workflow:

1. Run `scripts/inspect_li_mathworks_pptx.ps1` on the chosen official template.
2. Create slides from official layouts only: `Title Slide`, `Agenda`, `Section Header`, `Title and Content`, `Two Content`, `Feature`, or `Title Only`.
3. Use existing placeholders first. Do not draw a replacement title, subtitle, cover, footer, line, logo, or section background.
4. For the cover, use the official `Title Slide` slide or layout and replace placeholder text only.
5. For section dividers, use the official `Section Header` layout. Do not create full-blue or custom decorative divider slides.
6. For new PPTX generation, prefer `scripts/new_li_mathworks_pptx.ps1` with a JSON slide spec.
7. Export previews for every slide by default and run `scripts/review_li_mathworks_pptx.ps1` before claiming the deck is done. Use sampled previews only for quick debugging and label them as a sample.

Run PowerPoint automation scripts serially. Do not run generation, review, or export scripts in parallel; PowerPoint COM automation can leave hidden processes or false failures when multiple jobs run at the same time.

For HTML output, use `assets/li-mathworks-presentation/html/template.html` and `assets/li-mathworks-presentation/html/styles.css` as the starting point, then adapt the content. The HTML version should resemble the MathWorks template and Li's slide construction style, but it does not need to duplicate PowerPoint internals.

## Style Source Corpus

Li's MathWorks Presentation style is a combined style system learned from the full example corpus, not from one deck. When local examples are available, consider all relevant decks before generating:

- Official template examples: `public.pptx` and `confidential.pptx`.
- Internal/project workshop examples: `MBD Establishment Workshop.pptx`, `MBD Establishment Workshop - Reorganized.pptx`, and `X_MBD_Workshop.pptx`.
- Customer technical deep-dive/wrap-up examples: `P13744 2022-09-15 deep-dive.pptx`, `P13744 Pile Driver Path Planning.pptx`, and related customer result decks.
- MBSE/System Composer training examples: `Session_1.pptx`, `Session_2.pptx`, `Session_3 - Final.pptx`, and `Session_4 - Outline.pptx` from the MBSE source repo when accessible.

Use the corpus this way:

- Use the official MathWorks template only as the base chrome, placeholder system, confidential marking, logo, font/theme, and layout source.
- Use Li's examples to choose content patterns, page roles, color coding, diagram density, animation pacing, and artifact framing.
- Choose slide patterns because they support the story: training instruction, customer evidence, internal progress review, technical decision-making, or a specific artifact walkthrough.
- Do not copy source slides unless preserving a customer-approved diagram, complex animation, exact screenshot/evidence artifact, or reviewed technical content. If copied slides visually clash with the new deck, recreate the idea using the official template.
- Do not let one example dominate the output. Combine patterns by use case and by the user's requested audience.

## Template Facts

The available PPTX templates use:

- 16:9 widescreen slide size.
- Arial as the primary font.
- MathWorks theme colors: blue `#0076A8`, light blue `#ECF5F8`, orange `#D78824`, deep blue `#004B87`, cyan `#00A9E0`, yellow `#F2A900`, red `#B7302C`, green `#48A23F`.
- Standard layouts: Title Slide, Title and Content, Title Only, Blank, Feature, Section Header, Two Content, Agenda.
- Confidential template: built-in `CONFIDENTIAL` marking in the master/title layout.

## Use-Case Patterns

Choose the setup by delivery context before choosing colors or visual motifs:

| Use case | Primary story need | Prefer | Avoid |
| --- | --- | --- | --- |
| Training/workshop | Teach a repeatable method and let attendees practice it. | Agenda, what-to-expect, concept-plus-artifact pages, demo/exercise pages, recap pages, progressive object builds. | Customer wrap-up density without exercises; color-code declarations that are not used throughout. |
| Internal project sharing | Make status, workflow, tradeoffs, and next decisions reviewable. | Persistent progress rails, process-state diagrams, repeated scaffolds with visible X/check/clock/status changes, dense evidence tables. | Sparse section pages that do not add new state; decorative cards without decisions. |
| Customer wrap-up/deep-dive | Show what was done, what evidence supports it, and what should happen next. | Large screenshots, results/evidence tables, technical comparison pages, static or lightly animated workflow diagrams. | Training color bands, excessive animations, or copied workshop scaffolds that do not support the customer story. |
| General technical presentation | Explain a technical decision or method clearly for the audience. | The nearest pattern above based on audience and purpose. | Mixing patterns from all use cases just to resemble prior decks. |

Training/workshop:

- Use agenda, "what to expect", chapter dividers, concepts, demos, exercises, recap, and best-practice slides.
- Use progressive object reveals to pace instruction.
- Prefer diagrams, screenshots, tables, and step-by-step visual builds over long bullet lists.
- Use explicit page-type color coding when teaching: best practice, extra note, why, how, what to do, demo/exercise, and discussion pages can use consistent colored bands or framed blocks.
- Training decks may use structured dense detail when it is part of an exercise, recap, table, or screenshot explanation. Do not force every label above 14 pt when the source pattern is a readable dense technical page.
- Use concrete System Composer, Simulink, requirements, test, project, or architecture screenshots when available. Recreate key labels as text if screenshots are too small.
- Require concrete technical artifacts in every major chapter: source code excerpt, architecture/data map, component boundary, wrapper/model view, harness, requirement/test evidence, or screenshot.
- Derive agenda items from actual section/anchor slides; do not list chapters that are not represented in the deck.
- Do not put generator/build instructions in visible slide text. Put presenter/build intent in speaker notes.

Customer wrap-up or technical deep-dive:

- Use a polished story: project statement, scope, workflow, implementation, evidence/results, recommendations, and next steps.
- Keep most slides static or lightly animated.
- Make screenshots and results easy to inspect. Recreate labels if source screenshots are too small.
- It is acceptable for customer technical slides to be denser than workshop teaching slides when the density is evidence: project constraints, implementation options, comparisons, results, or customer-facing decisions.
- Prefer title/content placeholders, readable bullets, result screenshots, tables, and model/test evidence over repeated pastel process-card slides.

Internal project sharing:

- Use denser setup/context slides, workflow diagrams, grouped shapes, tables, screenshots, and process views.
- Show tradeoffs, current state, open issues, and decisions.
- Use progressive reveals where they help explain a workflow or architecture.
- Use persistent progress devices when walking through a process repeatedly: left rail, current-step highlight, repeated scaffold, added X/check/status marks, clocks, arrows, and outcome labels.
- Process-review decks can use repeated slide states, but each repeated slide should add visible information or a build state. Do not create multiple nearly identical horizontal card slides.

General technical presentation:

- Use the customer wrap-up structure for external audiences and the internal sharing structure for engineering audiences.
- Keep titles concrete and message-bearing.

## Visual Style

- Use clean MathWorks corporate layouts with restrained technical density.
- Use theme blues, orange, yellow, green, red, light blue, white, and neutral grays. Avoid purple AI/SaaS gradients.
- Use the template font through placeholders. Do not force random fonts. Typical generated-slide sizes should skew larger than dense source slides: 14-16 pt minimum for supporting text, 18-22 pt for primary bullets, and template title placeholders for titles.
- Use rectangles, rounded rectangles, connectors, lines, right/down arrows, chevrons, flowchart shapes, and callouts.
- Prefer rounded rectangles, chevrons, subtle callouts, connector lines, and grouped diagram regions over plain sharp square boxes.
- Group related shapes heavily when building process diagrams, architecture views, and workshop explanations.
- Use native PowerPoint diagram primitives for the recurring Li patterns: concept-plus-artifact teaching frames, V/process diagrams with MathWorks tool labels, persistent progress rails, process-state diagrams with X/check/clock/status marks, and screenshot/evidence callout regions.
- In concept-plus-artifact pages, put the concept or method on one side and the concrete artifact, screenshot, code/model excerpt, table, or evidence on the other side. Use a thin teaching frame and callouts instead of a plain bullet page.
- In V/process diagrams, use low-saturation model/workflow regions, orange-outlined tool labels, and side benefits or evidence lists. Reveal the path in order for workshops.
- Use low-saturation color coding on top of the MathWorks template. Prefer light blue, light cyan, pale yellow, pale green, light red, and neutral gray fills with theme-colored outlines.
- Use semantic color coding: gray = current/legacy, blue = model/workflow, orange = conversion/action, green = validated/pass, red = risk/blocker, yellow = review highlight.
- Do not declare a deck-specific color code unless the deck uses that code consistently across the whole presentation. For short or mixed-purpose decks, use semantic colors locally without a color-code declaration slide.
- For training color-code pages, the source corpus also uses stronger teaching bands: purple for best-practice concepts, light purple for extra best-practice notes, orange for "why", green for "how", blue for "what to do", cyan for demo/exercise, and light red for discussion or risk. Use these only when the deck is explicitly training/workshop and the page type benefits from it.
- If a fill would match the slide background, use transparent/no fill instead of a same-color fill.
- Prefer aligned boxes, connector routes, and consistent shape spacing over decorative illustration.
- Keep slides presenter-led: bigger text, fewer words, and more diagram/image/icon/animation support. Leave details to the speaker.
- Do not add fake footers, bottom stripes, logo areas, gray title bands, or other decorations already handled by the MathWorks template.
- Avoid an entire deck of generic horizontal process cards. Alternate between process maps, evidence screenshots, tables, code/model excerpts, left-rail step reviews, recap pages, and decision pages.

## Animation And Transition Style

- Most slide transitions should be none. Use a simple push only when it supports section movement.
- Training/workshop decks can use many object animations for progressive reveals.
- Customer wrap-up/deep-dive decks should use few animations unless a build sequence improves clarity.
- Common object effects: fade, wipe from left/down/up/right, random horizontal bars, occasional wheel/barn/strips.
- Common timing: about 500 ms for normal reveals. Very short timings may be used for instant setup/state changes.
- Use click, after, and with effects intentionally to create staged explanations. Avoid flashy motion or constant movement.
- When reorganizing workshop slides, inspect how animation sequences teach the flow. Recreate the intended reveal sequence instead of copying slides blindly.
- If importing an existing slide, write down why it is being imported: preserve approved content, preserve complex animation, preserve a screenshot/evidence artifact, or preserve a customer-reviewed diagram. If there is no reason, recreate the idea in the template style instead.

## Review Checklist

Before delivery:

- Confirm the deck started from the correct official MathWorks template.
- Confirm the cover uses the official `Title Slide` layout/placeholders.
- Confirm cover title, subtitle, author/presenter line, and date are not merged into the wrong placeholder.
- Confirm new slides use official layouts and placeholders.
- Confirm confidential/non-confidential handling is correct.
- Confirm `ground-truth.md` and `speech.md` exist, are current, and match the final slide order.
- Confirm unresolved missing assets/placeholders are tracked in `ground-truth.md`.
- Remove template resource or instruction slides.
- Check agenda, section headers, demos, recaps, and next steps match the use case.
- Inspect dense slides for clipped text, small labels, unaligned groups, and overpacked screenshots.
- Verify animations are purposeful and not distracting.
- Verify speaker notes capture slide reason, talk track, and build sequence for generated workshop slides.
- Verify training/workshop decks contain concrete technical artifact slides, not only outline/process slides.
- Verify no generated shape covers template lines, logos, master text, or title elements.
- Verify no custom bottom stripe, fake footer, or decorative block was added where the template already provides structure.
- Verify generated slides visually match adjacent imported slides; otherwise recreate or restyle the slide.
- Verify the chosen page patterns match the use case. A customer wrap-up can be dense and static; a training deck should have what-to-expect, demo/exercise, recap, and artifact pages; an internal process-sharing deck should have progress/review scaffolds and visible build states.
- Verify section/progress locators are generated from one source of truth. Do not hand-enter repeated labels that can drift from agenda items.
- For HTML, check the generated deck visually in a browser and keep it offline-ready.
