# Auto Version Control Rules - Claude AI

You are a senior software developer. These rules override your default behavior. Follow them on every action without being asked.

## Trigger Prompt

When the user says **"run auto version control"** (or any close variation like "run avc", "auto version control", "start version control"), immediately run the full assessment:

1. Run `git status`, `git branch`, and `git log --oneline -10`
2. Report the current state: branch, uncommitted changes, recent commits, version tags
3. Flag any issues: working on main, uncommitted changes, missing .gitignore, no tags
4. Recommend next actions

This is how the user explicitly asks you to check in on the project.

## Rule 1: Git Is Mandatory

- If the project is not a git repository, run `git init` and create an initial commit before doing anything else.
- Never work directly on `main`, `master`,`stable` . Always create a feature branch first then merge into `stable` and `PSM stable`.
- Branch naming: `feat/description`, `fix/description`, `refactor/description`, `docs/description`, `chore/description`.
- If you are on `main` when you start, create and switch to a feature branch immediately.

## Rule 2: Conventional Commits

Every commit message must follow this format:

```
type: short description (imperative, lowercase, no period)
```

Valid types: `feat`, `fix`, `refactor`, `docs`, `test`, `style`, `perf`, `chore`, `ci`, `build`.

Examples:
- `feat: add user authentication endpoint`
- `fix: prevent null pointer in payment handler`
- `refactor: extract validation logic into shared module`
- `docs: add API usage examples to README`

Rules:
- One logical change per commit. Do not bundle unrelated changes.
- Commit after every meaningful change, not at the end of a long session.
- If a commit touches more than 3 unrelated things, you are bundling too much. Split it.

## Rule 3: Semantic Versioning

update git hub releases on minor version changes of `stable`

Tag releases using `vMAJOR.MINOR.PATCH`:
- **MAJOR** -- breaking changes (removed features, changed APIs, incompatible updates)
- **MINOR** -- new features that do not break existing functionality
- **PATCH** -- bug fixes, typo corrections, minor improvements
