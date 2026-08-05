# Review Checklist

Run the sense checks first. A deck that fails them is not fixed by passing the mechanical checks.

## Sense Checks (first, before visual/mechanical review)

- Does the deck answer the audience's core question for its use case (see `storytelling.md`)?
- Read the slide titles in order, alone: do they form a coherent argument?
- Does each slide answer a question a previous slide raised? Flag any slide with no setup.
- Does every fact, number, name, and result on every slide trace to an entry in `ground-truth.md`? Flag anything that looks invented.
- Is any slide pure filler — generic statements that would be true of any project? Cut or ground it.
- Does each slide title state the slide's actual message rather than a topic label?
- Does the agenda match the real final slide order?
- Does the text/image balance deliver the message efficiently — could a diagram, screenshot, or emphasized table replace a text block?
- Are unresolved gaps visible as placeholders and listed in `ground-truth.md`, rather than papered over?

## Human-Signature Checks (does it read as machine-made?)

- Do card counts, bullet counts, and table rows all land on 3? Vary them.
- Does any item exist only to complete a set? Delete it.
- Do two adjacent slides use the same pattern with the same shape counts? Change one.
- Are the two or three slides that carry the argument visibly the most substantial?
- Does every title carry a number and an assertion, with no plain label anywhere?
- Is the same string rendered twice on one slide?
- Any ASCII hyphens or straight apostrophes where a person typing would get em dashes and curly quotes?
- Does the register match the room (internal shorthand internally, formal for customers) using only names and terms the source provides?
- Does any status label describe the enum rather than what actually happened?

## Process Checks

- Was `ground-truth.md` created before outlining, kept current, and updated with the final output path?
- Was `speech.md` created and mapped to the final slide order?
- For normal generation, was the outline shown in chat and approved before deck creation?
- For fast/direct generation, were assumptions recorded in `ground-truth.md` even though approval was skipped?
- Are missing assets, screenshots, data, placeholders, and progress labels tracked?
- Is the final deliverable `pptx + speech.md` or `html + speech.md`?

## Visual/Mechanical Checks

- Can a room reader understand labels and captions?
- Are any important labels clipped, hidden, or too small?
- Are animations helpful and not distracting?
- Does the deck use visual evidence, diagrams, screenshots, staged builds, or emphasized tables where possible instead of defaulting to text?
- Do buttons have reliable fallbacks?
- Are local image/script/CSS assets present?
- Do slide counters match slide count?
- Does JavaScript pass syntax check?
- Were important slides visually inspected from screenshots?

## Li's MathWorks PPTX — Template-Override Audit

- Does slide 1 use the official Title Slide layout, with cover title, subtitle, author/presenter line, and date in the right template placeholders instead of merged or redrawn?
- Do new slides use official layouts/placeholders rather than fake decorations?
- Is every custom shape either a documented preset pattern (tagged) or justified in `ground-truth.md`? An untagged, unjustified custom shape is a defect.
- Does any generated shape cover or substitute for template lines, logos, master text, title placeholders, or footer elements? Any overlap is a defect.
- Are all colors from the documented palette and all fonts from the template?
- Are any slides on the `Blank` layout? Each needs a recorded justification.
- Are template resource slides removed?
- Are PowerPoint automation scripts run serially?
- Were previews exported for every slide, unless clearly marked as a sampled debug preview?
- Do exported previews show no overlapped template text, logo conflicts, or custom bottom stripes?
- Are empty generated text boxes deleted and unused placeholders cleared?
- Are color-code declaration slides used only when the code is applied consistently across the deck?
- Do tables communicate status/outcome/risk visually rather than remaining plain text grids?
