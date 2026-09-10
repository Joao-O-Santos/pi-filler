# Agent instructions

This repository is a small Pi extension. Keep work proportionate to that
scale.

## Before changing code

Read `PROJECT.md`, `ARCHITECTURE.md`, `STYLE.md`, and `PLAN.md` while
the plan exists. Inspect the relevant code and tests before proposing
new structure. Treat the current repository state as more authoritative
than an old plan or comment.

## Design boundary

Prefer the smallest implementation that is easy to inspect. Use Node.js
and Pi primitives directly when they are sufficient. Prefer deleting
code to moving it behind a new abstraction.

Do not create frameworks, service layers, dependency injection,
compatibility shims, plugin systems, generic registries, or shared
utility packages without a concrete repeated need. A little duplication
is cheaper than a premature abstraction in these projects.

Keep the package understandable in one sitting. Split files when that
improves navigation, not to imitate a larger application architecture.

## Upstream policy

Track current upstream releases. Do not pin dependencies, actions, or
tool versions merely to preserve an old environment. Do not commit
`package-lock.json`.

When upstream breaks the project, update our implementation, tests, and
docs to work with current upstream behavior. Add backward-compatibility
code only when the user explicitly requires an older supported version.

## Style

Follow `STYLE.md`. Use tabs for code indentation where syntax permits
and keep Biome's line width at 100 characters. Hard-wrap ordinary
Markdown prose at 72 characters with the repository formatter.

Do not reformat generated files or syntax whose format forbids tabs.
YAML uses two-space indentation.

## Tests and verification

For behavior changes, add the smallest deterministic regression test
that would have caught the problem. Run `make verify` before reporting
completion. Run `make site` when documentation or site infrastructure
changes.

Do not claim checks passed unless they were actually run. Report
unavailable checks and their reason.

## Project state

Update `PROJECT.md` when accepted scope or direction changes. Update
`ARCHITECTURE.md` only after the implementation it describes is true.
Update `SYNC.json` after durable project artifacts are current, not
before.

`PLAN.md` is temporary implementation guidance. Remove it after its
accepted decisions have been implemented and transferred into durable
project files where needed.

## Commits and releases

Commit subjects are imperative, begin with a capitalized verb, and stay
within 50 characters. Use `Add ...`, `Remove ...`, `Simplify ...`, or
similar present forms. Do not use `Added ...` and do not use prefixes
such as `fix:`, `feat:`, `chore:`, or `docs:`.

Work in small verified commits. Do not tag, publish, or release without
explicit instruction.

GitLab is the sole release authority. GitHub is a mirror for
verification and Pages only. Do not add GitHub release or npm-publishing
workflows.
