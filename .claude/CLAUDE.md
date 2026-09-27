# Auto Version Control Rules - Claude AI

You are a senior software developer. These rules override your default behavior. Follow them on every action without being asked.

## Rule 0: Always Read First

Before taking any action on this project — including edits, commits, or file creation:

1. Read `.claude/CLAUDE.md` and `.claude/CODING_NOTES.md`.
2. Run `gh pr list` — if a PR exists for the current branch, run `gh pr view <number> --comments` and read all CodeRabbit comments before proceeding.
   - If `gh` is unavailable or unauthenticated, check the PR manually in the GitHub web UI and acknowledge any open CodeRabbit findings before proceeding.
3. Do not make any edits until outstanding CR findings are addressed or acknowledged.

No exceptions.

### Checking PR review status

`.claude/CODING_NOTES.md` is a standards and practices reference — a log of coding patterns and past findings, grouped by topic. It is **not** the source of truth for PR review status.

- To check if a PR review is complete or paused: **always use `gh pr view <number> --comments`**.
- CodeRabbit may auto-pause reviews after rapid commits — check for `review paused` in the summary comment.
- If paused, trigger a new run with: `gh pr comment <number> --body "@coderabbitai review"`
- If CR hits a rate limit (`Rate limit exceeded`), run `date -u` to get the current UTC time, calculate the UTC timestamp when the window clears, and state it explicitly (e.g. "clears at 05:04 UTC"). Re-trigger on the first user interaction at least 5 minutes after that time to allow for clock drift.
- **Sequential PR workflow:** Open one PR, wait for CR to finish and address all findings, merge, then open the next. Do not trigger multiple concurrent CodeRabbit reviews.

## Trigger Prompt

When the user says **"run auto version control"** (or any close variation like "run avc", "auto version control", "start version control"), immediately run the full assessment:

1. Run `git status`, `git branch`, and `git log --oneline -10`
2. Report the current state: branch, uncommitted changes, recent commits, version tags
3. Flag any issues: working on main, uncommitted changes, missing .gitignore, no tags
4. Recommend next actions

This is how the user explicitly asks you to check in on the project.

## Rule 1: Git Is Mandatory

- If the project is not a git repository, run `git init` and create an initial commit before doing anything else.
- Never work directly on `main` or `master`. Always create a feature branch first, then merge into `master`.
- Branch naming: `feat/description`, `fix/description`, `refactor/description`, `docs/description`, `chore/description`.
- If you are on `main` or `master` when you start, create and switch to a feature branch immediately.

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
- If a new feature is added or changed, update the top-level README.md before committing.
- After every commit, check if a PR exists for the current branch (`gh pr list --head <branch>`). If none exists, open one immediately via `gh pr create`. Never leave a commit on a feature branch without an open PR.

## Rule 3: Semantic Versioning

Update GitHub releases on minor version changes of `master`.

Tag releases using `vMAJOR.MINOR.PATCH`:
- **MAJOR** — breaking changes (removed features, changed APIs, incompatible updates)
- **MINOR** — new features that do not break existing functionality
- **PATCH** — bug fixes, typo corrections, minor improvements

**To cut a release:**
```bash
git tag v1.2.3
git push origin v1.2.3
```

**Note:** Only tag from `master`.

### Rehearse before releasing

Any workflow that publishes something (a release, a package, a deploy) must have a **rehearsal mode** that does everything except the final publish step. For GitHub Actions that means a `workflow_dispatch` trigger with an input for the version or tag, and a publish job that is skipped when it is a rehearsal. A rehearsal needs no tag. GitHub only lets you dispatch a workflow when the copy on the default branch already has a `workflow_dispatch` trigger, but the run uses the workflow file from the `--ref` you pick, so a change to a publish workflow that already has that trigger can be rehearsed from its own branch before it merges (`gh workflow run build-release.yml --ref <branch> -f version=vX.Y.Z`). A brand-new publish workflow, or a change that adds `workflow_dispatch` to one, has to merge first and is then rehearsed from the default branch before the first tag. Before that merge, make sure its publish job runs only for the real release event (for example a `v*` tag push) and is skipped on dispatch, so neither the merge nor the rehearsal publishes anything.

Rehearse before the first real tag, and again after any change to the release pipeline:

1. Run the rehearsal and read its job summary (the release notes it would publish, the artifacts it built).
2. **Download the rehearsal's artifacts and use them the way a user will**: install the package, import the plugin, run the binary on a clean machine or in a clean container. A green rehearsal proves the pipeline ran. It does not prove the artifact works where it is consumed: a package that built and passed its rehearsal can still be refused by the tool that imports it (for example, because its target-system label does not match the computer).
3. Only then tag. The human sign-off before tagging still applies (Rule 6 in repos that have it), and the rehearsal result belongs in the summary you give the human.

Do not rehearse by tagging a throwaway version: a pushed tag is public and hard to take back. If a repo has no publish workflow yet, this rule waits until one exists.

## Rule 4: CodeRabbit Pull Request Reviews

When a pull request is open or being prepared:

- Always open PRs via `gh pr create` — never merge directly to `master` without a PR.
- After CodeRabbit submits its review, read the review comments before making any further changes.
- For each finding (whether from CodeRabbit or a Claude code review):
  1. If it matches an existing `.claude/CODING_NOTES.md` entry — fix it immediately and reference the note's topic in the commit message.
  2. If it is a new pattern — fix it, then add or amend a note under the relevant topic in `.claude/CODING_NOTES.md` before committing, following that file's style rule (clear, ≤300 characters, grouped by topic).
- Do not dismiss or ignore nitpicks — log them to `.claude/CODING_NOTES.md` even if not immediately actionable.
- Only merge a PR after all blocking CodeRabbit comments are resolved.
