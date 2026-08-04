# Review Checklist

Before finalizing:

- Was `ground-truth.md` created, kept current, and updated with the final output path?
- Was `speech.md` created and mapped to the final slide order?
- For normal generation, was the outline shown in chat and approved before deck creation?
- For fast/direct generation, were assumptions recorded in `ground-truth.md` even though approval was skipped?
- Are missing assets, screenshots, data, placeholders, and progress labels tracked?
- Is the final deliverable `pptx + speech.md` or `html + speech.md`?
- Does each slide title match the actual message?
- Does the agenda match the real order?
- Can a room reader understand labels and captions?
- Are any important labels clipped, hidden, or too small?
- Are animations helpful and not distracting?
- Does the deck use visual evidence, diagrams, screenshots, staged builds, or emphasized tables where possible instead of defaulting to text?
- Do buttons have reliable fallbacks?
- Are local image/script/CSS assets present?
- Do slide counters match slide count?
- Does JavaScript pass syntax check?
- Were important slides visually inspected from screenshots?

For Li's MathWorks PPTX:

- Does slide 1 use the official Title Slide layout?
- Are cover title, subtitle, author/presenter line, and date placed in the right template placeholders instead of merged?
- Do new slides use official layouts/placeholders rather than fake decorations?
- Are template resource slides removed?
- Are PowerPoint automation scripts run serially?
- Were previews exported for every slide, unless clearly marked as a sampled debug preview?
- Do exported previews show no overlapped template text, logo conflicts, or custom bottom stripes?
- Are empty generated text boxes deleted and unused placeholders cleared?
- Are color-code declaration slides used only when the code is applied consistently across the deck?
- Do tables communicate status/outcome/risk visually rather than remaining plain text grids?
