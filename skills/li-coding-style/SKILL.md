---
name: li-coding-style
description: Use in WFHBridge whenever writing or editing code, in any language. This is Li's personal coding style — naming, structure, comment density, safety habits — not a MATLAB-only convention. Currently derived from the +HarnessRepair MATLAB package; extend the language-specific section as we see Li's style in other languages.
---

# Li's Coding Style

Baselines:
- `01-active/2026-06-04-mbd-test-case-maintenance-class/src/+HarnessRepair` (28 files, 2,989 lines, analyzed 2026-07-06).
- `C:\DevGit\temp\CreateParameterList.m` (single large MATLAB extraction/export workflow, analyzed 2026-07-22).

This is still MATLAB-heavy evidence, but it now includes both package/class code and a long procedural System Composer/data-dictionary extraction script. The principles below are split into what's genuinely cross-language (apply in any language WFHBridge code gets written in) and what's MATLAB syntax detail (apply only to `.m` files). When code shows up in a new language, extend this file with what's language-specific for it rather than assuming the MATLAB section generalizes.

Apply this to new code. Do not rewrite existing code purely to match it — only apply when touching a file for a real reason.

## Universal principles (any language)

**Naming**
- Boolean variables get a clear boolean marker prefix (`f_` in MATLAB — the equivalent idiomatic marker in another language, e.g. `is_`/`has_` in Python, is the same instinct, not a contradiction). Applied consistently, not just sometimes.
- Loop indices get a short base name plus a context suffix once more than one is in scope (`idxTC`, `idxTS`) — not just `i`, `j`, `k` once ambiguity is possible.
- Public/object properties use UpperCamelCase. Class names, method names, function names, and local variables use lowerCamelCase. Function or method input arguments/parameters use UpperCamelCase.
- Follow each language's own casing convention for that construct (camelCase vars in MATLAB/JS, snake_case in Python, etc.) rather than importing MATLAB casing wholesale — the naming *discipline* (clear prefixes, context suffixes, no abbreviations without prior establishment) is what's personal style, not the specific casing.

**Function/method design**
- Success-flag-first returns where a language supports multiple returns: report success/failure as the primary signal, initialize outputs to failure defaults, set them on the success path.
- Guard clauses with early return over deep nested conditionals.
- Do not create a helper function when the helper body would be fewer than 5 lines and it is used only once. Keep the logic inline at the call site unless Li explicitly asks for it to be a function.
- Helper extraction rule: if a helper-worthy code block is more than 10 lines and makes the main function/method harder to read, first try to explain the block with one clear local comment or section comment. If one simple comment is not enough, the block is more than 20 lines, and it makes the main flow hard to track, create a helper function in the same `.m` file. If the helper grows beyond 40 lines, class/method code may move it into the class `private/` folder. For standalone function development, keep helpers as subfunctions in the same file unless Li explicitly asks for a separate private helper.
- Recursion is the right tool for graph/tree-shaped problems — prefer it over manual stack/queue bookkeeping when the problem is naturally recursive.
- Order destructive/repair actions least-destructive-first (e.g. reroute > rename > recreate) rather than jumping straight to the most direct/destructive fix.

**Architecture**
- One function/method per file when it's substantial; small helpers can stay inline.
- Internal/helper functions live separately from the public API (a `private/`-equivalent), not mixed into the public surface.
- Access control should be narrow and explicit (friend-class/module-level access) rather than defaulting to fully public or fully private — open exactly as much as a specific caller needs.
- Objects that need to reach back into an owner hold an explicit back-reference, rather than passing owner state around as loose arguments.

**Comments & logging**
- Target roughly **15% comment density** (comment lines ÷ non-blank lines) as a calibration point — not a floor to pad toward or a ceiling to trim to. Measured baseline: 396/2,700 ≈ 14.7% in the MATLAB source.
- Comments explain **why**, not what the code already says. The `CreateParameterList` sample confirms three preferred comment categories: dated change notes for meaningful behavior/history, `%%` phase headers for long workflows, and local why/rule comments near non-obvious business logic or API behavior.
- Keep comments brief and calm. If several consecutive lines do similar setup, validation, assignment, or extraction work, use one short section comment for the group instead of commenting each line or property. Avoid making the code look busy.
- Don't strip existing commented-out code when editing nearby lines — kept deliberately as history.
- Long-running operations get progress output with clear phase markers; warnings/errors say what to do next, not just what went wrong.
- Use section-header comments to break up long functions into logical phases. In long MATLAB workflows, Li commonly uses sections such as `%% Prepare and validate work environment`, `%% Initialize output data`, `%% Parse Input`, `%% Extract Lists`, and helper group headings like `%% Major Helper Functions - Nested`.

**Safety habits** (when writing code that mutates external state — files, models, records)
- Back up before destructive edits.
- Check/record state before touching something shared, restore it afterward.
- Refuse to operate on something with unsaved/uncommitted changes rather than silently overwriting.
- Prefer marking-for-removal over hard delete when touching user-authored content.

## MATLAB-specific (syntax detail, `.m` files only)

- Package namespaces (`+Pkg`) containing `@`-folder handle classes.
- `properties (SetAccess = {?Pkg.classA, ?classB})` for the narrow-access-control principle above.
- Property validation blocks: `mustBeMember` for enumerated states, size specs (`(1,:) char`, `(1,1)`), typed object arrays (`(:,1) Pkg.issue`).
- `arguments` blocks for input validation and defaults, with name-value options as `opt.*` inside them — this is the current direction, not legacy style to avoid.
- Nested functions for subroutines sharing context with their parent — don't hoist them out just to avoid nesting.
- `switch`/`case` dispatch on type strings (`BlockType`, `origSrc`/`origSink`) is the dominant control-flow pattern for branching on kind.
- Include the offending object's full path (`getfullname(...)`) in warnings/errors.
- Wrap noisy calls in `warning off` / `warning on` pairs, scoped as tightly as possible.
- For long procedural MATLAB utilities, a top `%% Script Change Note:` block is acceptable when it records real dated behavior changes, migration notes, or compatibility fixes. Keep it factual and specific; do not use it as a generic changelog for tiny edits.
- Prefer `%%` phase comments in large single-file workflows. The phase name should describe the workflow stage or artifact boundary, not the next statement.
- Nested helper functions are idiomatic when they share the parent workflow state, especially for parsing/export helpers that need access to accumulated lists, dictionaries, model handles, or configuration constants.
- It is acceptable for externally-defined artifact names to drive casing and field names. Preserve CSV column names, stereotype property names, requirement/report column names, and customer schema identifiers exactly even when they do not match normal MATLAB variable style.
- For generated/exported records, initialize struct arrays with explicit field lists so the schema is visible at the top of the workflow.
- Use dictionaries/maps for duplicate detection and model-vs-input reconciliation when keys are business identifiers and values are record numbers or source rows.
- Warning text should include the artifact being skipped and the action/result, for example that a component will not be added to the component list.

**Commenting calibration from `CreateParameterList`**
- Do not rely only on function headers. Long workflows should have top-level phase comments and targeted local comments that explain business rules, external API constraints, or historical fixes.
- Good comments include why a CSV parser uses `readcell` instead of `readtable`, why model data is treated as ground truth over input CSV rows, why line endings follow the input/export contract, or why a port trace needs a special model-reference path.
- Avoid comments that merely restate code, such as "loop over rows" or "set flag true". Comments should preserve context that would be hard to recover from the statement itself.
- Keep dated comments when they explain customer-visible behavior changes or fragile API workarounds. Prefer concise bullet notes over prose paragraphs.
- Sparse or placeholder function headers like "Summary of this function goes here" should be replaced when touching the file; a thin header is worse than no useful header.

**Don't carry forward into new MATLAB code** (legacy idioms present in the baseline — don't replicate, but don't churn existing lines to fix them unless that's the actual task):
- `getfield`/`setfield` — use dynamic fieldnames (`s.(name)`) instead.
- `eval` for parsing dimension/value expressions — prefer validated parsing.
- Swallowing exceptions with no handling — always at least warn/rethrow with context.
- Broad `warning off` without a message ID.
