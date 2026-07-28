# Single Source Of Truth

Before changing MATLAB package/tool code, identify the authoritative artifact for the task.

## Authority Order

1. Requirement or workflow note.
2. Class/property/method table.
3. Accepted config table or behavior contract.
4. Tests and examples as validation evidence, not design authority.
5. Chat as temporary discussion only until written into the design note.

## Rules

- If a decision is made in chat, update the design note before or with the code change.
- Do not implement public behavior that exists only in chat.
- Do not import API concepts from another branch unless the design note explicitly accepts them.
- If implementation and design disagree, decide whether to change the design or the code before continuing.
- Every new public class, property, method, config field, or workflow behavior must map to a design-table row or workflow step.
