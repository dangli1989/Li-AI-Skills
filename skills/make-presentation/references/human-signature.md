# Human Signature

Goal: the finished deck should read as something Li built by hand in PowerPoint over a couple of evenings — not as something a machine emitted. Decks that fail this are not "wrong" slide by slide; they fail because every slide is *equally* correct.

## The hard boundary

Vary **form** freely. Never invent **facts**.

Everything in this file is about shape, rhythm, density, and wording of things that are already true. None of it authorises inventing a name, number, repo path, ticket ID, date, or anecdote. If a humanising detail would require a fact you do not have, ask for it or leave it out. A fabricated "(rig 1 not retested — no time)" is a far worse failure than a deck that reads as machine-made.

## What gives a generated deck away

Reviewers who correctly identified generated decks named these, in order of how loudly they signalled:

1. **A fixed body anchor.** Every slide's content starts at the same y, whether the title runs one line or two, and ends early — so the bottom third is empty on slide after slide. Nothing is ever denser because it mattered more.
2. **Everything counts to three.** Three bullets, three cards, three flow boxes, three table rows, over and over — including a padded third item that adds nothing just to complete the set.
3. **Identical geometry on consecutive slides.** Two slides instantiate the same template with the same box bounds; the audience sees the same picture twice.
4. **Uniform card sizing.** Equal widths and equal gaps regardless of how much text each card holds, and a grid rule that orphans the 4th card onto its own row.
5. **Zero lazy titles.** Every single title is a full assertion with a number in it. A real deck has at least one "Agenda", "Open items", or "July numbers".
6. **Duplicated strings.** The same sentence rendered in two places on one slide.
7. **ASCII punctuation.** Hyphens where a person typing in PowerPoint would get an em dash; straight apostrophes where autocorrect gives curly.
8. **Over-formal register for the room.** "D. Mercer" in a deck presented to the team that calls him Dave; no shorthand, no team abbreviations.
9. **Mechanical taxonomy.** A fixed status vocabulary applied literally — labelling an event that already happened as "risk" because that is a valid enum value.

## Rules to apply

**Rhythm and counts**
- Do not let card counts, bullet counts, and row counts all land on 3. Across a deck, deliberately mix: a 2-card row, a 4-card row, a 5-bullet list, one list with a single short fragment in it.
- If an item exists only to complete a set of three, delete it. Arity follows the facts, never the layout.

**Density variation**
- Decide, before building, which two or three slides carry the argument. Those get more content, more evidence, and more vertical space. Let genuinely thin slides be visibly short rather than padding them to match.
- Never leave the same amount of dead space on consecutive slides. If two adjacent slides would look alike, merge them, cut one, or change the pattern.

**Geometry**
- Consecutive slides must not use the same preset with the same shape counts. Alternate: evidence, then table, then a two-column comparison, then a diagram.
- Size cards to their content: a card with two lines of text may be shorter than its neighbour. Perfectly equal boxes are the loudest tell after dead space.

**Titles**
- Assertion titles are Li's habit and should dominate — but include one or two plain labels where a plain label is honest ("Agenda", "Open items", "Backlog"). Not every title needs a number.
- Keep titles to roughly one line (~55-58 characters). If a title must wrap, break at a clause so the second line carries at least three words, never a single orphan word.

**Language**
- Use the register of the actual room. Internal team decks use the shorthand the team uses — first names, tool names, ticket prefixes — **only where the source material provides them**. Customer decks stay formal.
- Vary sentence shape between slides. If three slides in a row end with a provenance stamp ("Measured on rig 2, week of 2026-07-06"), rewrite two of them.
- Use em dashes and curly apostrophes; the generator's `typographic()` helper does this, so do not fight it by hand-writing ASCII.

**Status and colour**
- Apply status labels that describe what actually happened, not the nearest enum value. An incident that was found and fixed is "fixed", not "risk".
- Reserve one accent colour for the single most important thing in the deck. If four pastel fills encode nothing, the colour is decoration and should be removed.

## Where the residue should come from

The most convincing human signals are not decorations added at the end — they are consequences of having actually thought about the talk:

- One slide is denser than the rest because that evidence is what the audience will argue about.
- One slide is nearly empty because the point is a single number.
- A blunt phrase survives because it is the honest phrase ("Bus factor is 1", "Two approvals sit outside this room").
- The agenda is a plain list because the presenter has given this monthly review eleven times.

Get those from the ground truth and the story, and most of this file takes care of itself.
