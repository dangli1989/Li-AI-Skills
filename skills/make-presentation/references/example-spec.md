# Worked Example: JSON Spec, Ground Truth, and Speech

One complete small example for `new_li_mathworks_pptx.ps1` / `new_li_mathworks_pptx_mac.py`. Use it as the anchor for structure and tone; scale up for real decks. Note what makes it Li-style: every fact is concrete and sourced, every slide has a `reason` and `talkTrack`, titles state messages, and evidence outweighs prose.

## JSON slide spec (`spec.json`)

```json
{
  "title": "Battery Harness Repair: Wrap-Up",
  "subtitle": "Code-to-model conversion results for the P2119 legacy harness",
  "authorLine": "Li Dang, MathWorks Consulting",
  "dateLine": "2026-08-05",
  "useCase": "customer-wrap-up",
  "audience": "P2119 customer engineering team and their project lead",
  "purpose": "Show what was converted, the evidence it works, and what the customer should do next.",
  "coverTalkTrack": "Thank the team, restate the goal in their words: retire the hand-maintained harness scripts without losing test coverage.",
  "sourceFacts": [
    "Legacy harness: 4,200 lines of hand-written MATLAB across 12 scripts (customer repo, commit 8f31c02).",
    "Converted to Simulink Test harnesses: 9 of 12 scripts fully migrated, 2 partially, 1 intentionally retired.",
    "Regression suite: 214 of 218 legacy checks reproduced; 4 differences traced to a legacy rounding bug (issue #47).",
    "Runtime: nightly regression dropped from 3.1 h to 55 min on the customer's CI runner."
  ],
  "missingInputs": [
    "Customer-approved screenshot of the CI dashboard (placeholder on the results slide until provided)."
  ],
  "slides": [
    {
      "type": "content",
      "title": "Goal: retire the hand-maintained harness without losing coverage",
      "reason": "Restate the project statement in the customer's terms before showing work.",
      "talkTrack": "This is the sentence we agreed on in the kickoff. Everything in this deck is evidence against this sentence.",
      "bullets": [
        "12 legacy scripts, 4,200 lines, one owner who is retiring",
        "218 regression checks that must keep passing",
        "Target: Simulink Test harnesses your team can maintain"
      ]
    },
    {
      "type": "v-model-tool-map",
      "title": "Conversion workflow: analyze, wrap, migrate, verify",
      "reason": "One diagram showing how the work was done, before diving into results.",
      "talkTrack": "Walk the path left to right: static analysis of the scripts, wrapper isolation of legacy I/O, harness migration, then back-to-back verification.",
      "builds": ["Reveal analyze and wrap first", "Then migrate", "Then verify with the equivalence arrow"]
    },
    {
      "type": "results-table",
      "title": "214 of 218 checks reproduced; the 4 gaps are a legacy bug",
      "reason": "The core evidence slide; the deck stands or falls here.",
      "talkTrack": "Row by row: green rows are reproduced coverage. The four red checks differ because the legacy code rounds before saturation - issue 47, confirmed by your team as a bug.",
      "headers": ["Result", "Evidence", "Status", "Next action"],
      "rows": [
        "Scripts migrated | 9 of 12 full, 2 partial, 1 retired | Complete | Remaining 2 in phase 2",
        "Checks reproduced | 214 / 218 | Pass | See issue #47 for the 4 diffs",
        "Nightly runtime | 3.1 h to 55 min | Pass | Same CI runner",
        "Legacy rounding bug | 4 checks affected | Risk | Customer decision: fix or preserve"
      ]
    },
    {
      "type": "image-evidence",
      "title": "CI dashboard: harness suite running nightly since July 20",
      "reason": "Independent visual proof the migration is live, not a demo artifact.",
      "talkTrack": "This is your own dashboard, not our lab. Point out the streak since July 20 and the runtime column.",
      "images": [{ "path": "assets/ci-dashboard.png", "placeholder": true }],
      "assetStatus": "Waiting for customer-approved screenshot",
      "missingInputs": ["Customer-approved CI dashboard screenshot"]
    },
    {
      "type": "decision",
      "title": "Next step: decide on the rounding bug, then phase 2",
      "reason": "End on the decision the customer must make, with a recommendation.",
      "talkTrack": "Two options for issue 47: fix the legacy bug and accept 4 changed baselines, or preserve bug-compatibility in the model. We recommend fixing; here is why.",
      "items": [
        "Fix rounding bug - 4 baselines change (recommended)",
        "Preserve bug-compatible behavior in model",
        "Phase 2: migrate remaining 2 scripts (est. 3 weeks)"
      ]
    }
  ]
}
```

## Per-type content fields (cheat sheet)

Every slide needs `type`, `title`, `reason`, `talkTrack`; optional everywhere: `builds`, `intro`, `assetStatus`, `placeholderStatus`, `missingInputs`, `activeProgressLabel`/`locator`. Content fields by type:

- `agenda`/`learning-path`: `items` (+ optional `intro`)
- `content`/`chapter-objectives`: `bullets`
- `two-content`: `leftTitle`, `leftBullets`, `rightTitle`, `rightBullets`
- `section`/`chapter-divider`: `outcome`, locator via `activeProgressLabel`
- `what-to-expect`: `sections` [{label, detail, role}] — roles: best-practice, extra-note, why, how, what, demo, discussion
- `process`: `steps` [string or {label, detail}]
- `progress-sidebar`/`process-build-series`: `activeProgressLabel` (rail), `mainTitle`, `cards` [{label, detail, role}]
- `process-state-diagram`: `steps` [{label, detail, status}]
- `decision`: `items` (2-4 short options)
- `recap`/`recap-bridge`: `items` or `bullets` (+ optional `subtitle`)
- `artifact-map`/`artifact-review`/`architecture-layer`/`team-context`: `artifacts` [{label, detail, role}]
- `concept-artifact`: `concepts` [{label, detail, role}], `images` [{path, caption, placeholder}], `artifactLabel`, `callouts`
- `v-model-tool-map`: `phases` [{label, tool}] (max 5), `callouts`
- `code-review-excerpt`/`code-to-model-review`: `code` (short excerpt), `callouts`
- `image-evidence`/`screenshot-evidence`/`screenshot-callout`/`model-screenshot`: `images` [{path, caption, placeholder}], `callouts` [{label, role}]
- `comparison-table`/`comparison-evidence-table`/`results-table`: `headers`, `rows` — each row is a `"a | b | c"` pipe-string or `{cells: [...]}`; optional `callout`
- `demo-exercise`/`exercise-demo`: `panels` [{label, detail, role}] or `start`/`task`/`output`

Semantic roles for colors: `legacy/current` gray, `model/workflow` blue, `conversion/action` orange, `verification/validated/pass` green, `risk/blocker` red, `review/highlight/pending` yellow.

## What the generator writes from it

`ground-truth.md` — the ledger: title/audience/purpose header, the `sourceFacts` list verbatim, a slide table (number, type, title, intent, open gaps), and `missingInputs` as Open Inputs. Keep updating it by hand as facts, assets, or slide order change; it is the source of truth the review checks trace against.

`speech.md` — one `## Slide N - Title` section per slide with the `talkTrack` text and any `builds` sequence. Deliverable is always `deck.pptx + speech.md`.

## Anti-example (what NOT to do)

A spec whose facts read "significantly improved performance", "modern workflow", "robust verification" with no numbers, no file names, and no sources is filler. If the real numbers are not in hand, the fact list must instead contain the questions ("What was the legacy runtime? Ask customer") and the slides must show placeholders — never invented specifics, and never generic praise.
