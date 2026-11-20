# Claude Code — Agent Operating Instructions

You are working inside the **ae-proposal-engine** repository (or a proposal
repository scaffolded from it). This file is loaded automatically at the start
of every Claude Code session in this repo. Follow the rules below for every
session, without exception.

---

## 1. Working directory discipline

- Your working directory is the **repo root** of this project. All file
  creation, edits, scripts, and git operations must operate on paths **inside
  this repo**.
- Before any `Write`, `Edit`, or `Bash` call that creates or modifies a file,
  resolve the target path and verify it is inside the repo root. If it is
  outside the repo, **stop and ask the user for explicit confirmation** —
  do not proceed on assumption.
- Never write to `~/Desktop`, `~/Documents`, `~/Downloads`, `/tmp`, or any
  other path outside the repo for "scratch," "preview," or "temporary" files.
  If you need a scratch location, use a path inside the repo (e.g. a
  gitignored `_scratch/` folder) — not the user's home directory.
- If `cwd` drifts during a session (e.g. after a `cd` in a Bash call), `cd`
  back to the repo root before the next write.
- When a user request is ambiguous about location ("save this script," "make
  a note file"), **default to a path inside the repo** and name the path
  explicitly in your response so the user can redirect before you act.

**Why this rule exists:** On a prior pursuit, a session wrote files to the
user's Desktop instead of the proposal folder ("Why were you creating files
on my desktop?"). With multiple concurrent proposals in flight, cross-pursuit
contamination or silent data loss is a catastrophic failure mode. This rule is
non-negotiable.

**How it's enforced (programmatically):** A `PreToolUse` hook configured in
`.claude/settings.json` runs `python scripts/hooks/cwd_guard_hook.py` before
every `Write` / `Edit` / `NotebookEdit` tool call. The hook reads the tool
input, determines the active git repo root via `git rev-parse --show-toplevel`,
and rejects (exit 2) any `file_path` outside that root. The hook is repo-
relative — it works the same way in the engine repo and in any scaffolded
proposal repo. Shell scripts have a parallel guard via `scripts/lib/guard_cwd.sh`
(sourced by every entry-point script); Python scripts use
`scripts/lib/guard_cwd.py`. The prose rule above remains your first line
of defense; the hook is the safety net.

**One legitimate exception:** the hook allowlists writes under
`~/.claude/projects/` because that's where Claude Code's auto-memory storage
lives. Memory is meta-conversation state, not project artifacts, and is
managed by the agent by design. All other paths outside the repo are
rejected.

---

## 2. Repository context

- **Engine repo** (`ae-proposal-engine`): the template and toolchain. Identified
  by `scripts/scaffold/new_proposal.sh` at the root. Scripts here are the
  template scripts that scaffolded proposals will inherit.
- **Scaffolded proposal repo**: produced by `new_proposal.sh`. Identified by
  `00_admin/proposal_config.yml` at the root. Contains live content for a
  single pursuit.
- If you are unsure which you are in, check for those marker files before
  making structural changes.

---

## 3. Scripts and configuration

- All shell scripts in `scripts/` derive `REPO_ROOT` from their own location
  (`SCRIPT_DIR`/`BASH_SOURCE`) — they do **not** rely on the caller's `cwd`.
  Preserve this pattern when adding or editing scripts.
- A shared guard helper (`scripts/lib/guard_cwd.sh`) is sourced by every
  script to verify it is running in a valid repo root. If you add a new
  script, source and call the guard: see existing scripts for the pattern.
- `00_admin/proposal_config.yml` is the **single source of truth** for
  proposal metadata. Never hardcode values that belong in the config.

## 4. Internal notes vs. published content

When writing or editing proposal markdown, distinguish between content the
client will see and notes that must not leak into the delivered `.docx`.

**Use HTML comments for invisible internal notes.** Pandoc strips them
during the docx build, so they appear in source but never in output:

```markdown
<!-- INTERNAL: confirm pricing with Bill before final submission -->
The contractor shall provide [SCOPE] services per Section 3.2.
```

**Avoid these visible-marker patterns** in committed source — the publish
lint gate (`scripts/publish/check_output_pdf.sh`) blocks submission when
any of them appear in the compiled `.docx`:

`PLACEHOLDER`, `[VERIFY]`, `[NEEDS_VERIFICATION]`, `[FIXME]`, `[TODO]`,
`[INTERNAL]`, `[DRAFT]`, `[INSERT…]`, `[CONFIRM…]`, `[PENDING…]`,
`[NUMBER]`, `[TBD]`, `[XX]`, `[XXX]`, `*End of draft*`.

Use these markers freely **during drafting** to flag work-in-progress;
just resolve every one before invoking `combine_volumes.sh` (the gate is
hard-blocking there). For a working draft, use `<!-- … -->` instead — it
won't leak even if you forget to remove it.

---

## 5. Workflow conventions

See the "Branching, Commits, and Versioning" section of `README.md`:

- Trunk-based development on `main`, PRs required, squash-merge.
- Conventional Commits (`feat:`, `fix:`, `docs:`, `chore:`, `ci:`,
  `refactor:`).
- Semantic Versioning; pre-1.0 (`0.x.y`) allows breaking changes between
  minors.
- Every release updates `package.json`, moves `## [Unreleased]` entries in
  `CHANGELOG.md` under a new dated version heading, and tags annotated.

---

## 6. Output publishing rule — save state and merge to `main` at every publish

This rule is non-negotiable and applies to every session in this repo.

- **Locations.** All pursuit outputs go in `output/`:
  - `.md` → `output/artifacts/`
  - `.docx`, `.xlsx`, `.ics` and other deliverables → `output/published/`
    (calendar files in `output/published/calendar/`).
- **Never overwrite.** Every file is versioned `<base>_vNN.<ext>`. When a
  newer version is published, the current one is **moved** to
  `output/<area>/_superseded/vNN/`. Nothing in `output/` is deleted or
  replaced in place.
- **Save state and merge to `main` at every publish.** Each publish commits
  the whole working tree on a short-lived `publish/<timestamp>` branch,
  fast-forward merges it into `main`, deletes the branch and pushes `main`
  to `origin`. This overrides the PR-required convention in §5 for output
  publishes only.
- **How.** Stage files under `_scratch/stage/{artifacts,published}/`, then run
  `python scripts/publish/publish_outputs.py --stage _scratch/stage -m
  "docs(output): <what changed>" --push`. The script enforces versioning,
  supersession and the git steps. Do not hand-copy files into `output/`.
- **Generators.** Pursuit-specific data and builders live in
  `scripts/pursuit/` (e.g. `build_outputs.py`; `--docx` converts staged
  `.md` to `.docx` via pandoc with `templates/docx_reference/reference.docx`).
  Edit the data there and regenerate rather than editing published files.
