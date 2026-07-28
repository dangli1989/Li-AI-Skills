# Skill Authoring

Each skill lives under `skills/<skill-id>`.

Required:

```text
SKILL.md
skill.yaml
```

Optional:

```text
references/
scripts/
assets/
agents/
```

## Authoring Rules

- Keep `SKILL.md` concise and action-oriented.
- Put long supporting detail in `references/`.
- Put reusable code in `scripts/`.
- Put templates and static files in `assets/`.
- Keep `skill.yaml` portable and tool-neutral.
- Do not encode runtime-specific paths inside source skills unless the skill itself requires them.
