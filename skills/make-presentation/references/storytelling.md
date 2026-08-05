# Storytelling Guide

Pick the arc that matches the use case, then apply the coherence rules. Do not mix arcs: a training deck built like a status review, or a customer wrap-up built like a workshop, is the most common way a deck stops making sense.

## Coherence Rules (all use cases)

- Every slide answers a question the previous slide raised. If no earlier slide makes the audience ask the question this slide answers, move it, add the setup, or cut it.
- Every slide title states the slide's actual message ("Wrapper isolates legacy I/O from the model" — not "Architecture").
- One message per slide. Supporting evidence lives on the slide; supporting nuance lives in `speech.md`.
- The agenda is generated from the final slide order, never written first and left stale.
- The deck must work as a story when only titles are read in sequence. Read the title list out loud as a final check: it should sound like a coherent argument.
- Facts on slides come from `ground-truth.md`. A slide that needs a fact the ground truth does not have gets a visible placeholder and a question to the user — not an invented number.

## Slide Economy Rules (all use cases)

- **Every slide must carry a comparison, a quantified change, a concrete artifact, or a decision.** A slide that carries none of these merges into a neighbour or moves to backup. Three short bullets restating the title is not a slide.
- **Cap any single fact at two appearances** across the deck. Repeating the same number on four slides reads as padding.
- **Suppress the agenda** for decks under ~18 slides whose titles are already assertions, and for internal recurring reviews with a small, familiar audience. Keep it for workshops and first-time customer audiences.
- **A supporting callout must add a fact the title does not already state.** If the chips beside a figure re-parse the title, delete them.
- **Do not derive facts the source does not state.** Never compute a ratio, subset, or set relationship between two source facts ("4 of 21 public functions") unless the source states the relation. Never assert a future commitment ("tracked for next review") that is not in the source.
- **Layout arity follows the facts.** Never synthesise a third card to fill a three-card pattern.
- **If the talk track explains a relationship** — replaces, flows into, before/after, caused — that relationship must be visible in the graphic through arrows, columns, or a legend. It may not live only in narration or in colour.
- **When the arc calls for a section the facts cannot support**, do not invent it and do not silently drop it: emit the slide with a visible placeholder and an open question. Record it in `ground-truth.md`.

## The Decision Slide

If the deck exists to get a decision, the decision slide is the **most substantive slide in the deck**, and it is the **last** slide — open items go before it or into an appendix, so the room is not left staring at what it cannot decide.

Every decision item carries: what is being decided, who proposes, who approves, when it takes effect, and what happens if the answer is no. Unknown fields render as visible placeholders rather than being omitted. Items that are not decisions do not appear on the decision slide.

If a stated deadline is close to the deck date, put the remaining time on the slide — urgency the audience cannot see is urgency that does not exist.

## Placeholders Are Dependencies

Every visible placeholder must produce a matching action item — artifact, owner, due date — on the open-items slide and, if it blocks the ask, on the decision slide. A placeholder that never becomes a request is an unfulfilled dependency the presenter forgot to chase.

## Arc: Training / Workshop

The audience's question: "What will I be able to do, and how do I practice it?"

1. **Why this matters** — the pain or risk the method removes, one concrete example.
2. **What to expect** — format, page-type color coding, when attendees act vs. listen.
3. **Agenda / learning path** — chapters mapped to outcomes, generated from real sections.
4. Per chapter, repeat this teaching unit:
   a. **Concept** — the idea, minimal text.
   b. **Concrete artifact** — real code, model, screenshot, or table showing the idea (every chapter must have at least one).
   c. **Demo or exercise** — task, starting artifact, expected action, expected output.
   d. **Recap** — what was learned, common mistakes, bridge to the next chapter.
5. **Final recap + best practices** — the repeatable method on one slide.
6. **Close** — what attendees should do in their own project next week.

## Arc: Customer Wrap-Up / Technical Deep-Dive

The audience's question: "What did you do, does the evidence hold, and what should we do next?"

1. **Project statement** — customer's goal and constraints in the customer's terms.
2. **Scope** — what was in, what was out, what was assumed.
3. **Workflow** — how the work was done, one diagram.
4. **Implementation highlights** — the 2-4 technical decisions that mattered, each with its reason and its artifact.
5. **Evidence / results** — screenshots, result tables, verification status. This is the heart of the deck; make it inspectable and mostly static.
6. **Limitations and open items** — named honestly; customers trust decks that show edges.
7. **Recommendations / next steps** — concrete, owned, time-bound.

## Arc: Internal Project Sharing

The audience's question: "Where does this stand, what changed, and what do you need from me?"

1. **Context reset** — one slide: what the project is, for anyone new in the room.
2. **Current state** — the workflow/process scaffold with visible status marks (done, blocked, in progress).
3. **What changed since last time** — same scaffold, state deltas highlighted; each repeated slide must add visible new state.
4. **Tradeoffs and open issues** — the decisions on the table, with options and evidence.
5. **Decisions needed** — exactly what the audience must decide or unblock, one slide.
6. **Next steps** — owners and dates.

## Arc: General Technical Presentation

Use the customer wrap-up arc for external audiences and the internal sharing arc for engineering audiences. Keep titles concrete and message-bearing.

## Timing

- 15-minute talk: roughly 8-12 slides. 30 minutes: 15-20. 60-minute workshop: 25-40 with exercises.
- Fewer words per slide than feels natural — the speaker carries nuance; `speech.md` carries the speaker.
- Implementation details go in speaker notes or backup slides, not the main flow.

## Reordering Rule

If a slide visually explains the whole talk, move it near the beginning. Then regenerate the agenda immediately.
