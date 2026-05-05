# Upstream sync catalogue — python-docx

Governed by ADR 004 in the
[ooxml-reference-corpus](https://github.com/loadfix/ooxml-reference-corpus/blob/master/docs/adr/004-upstream-sync.md)
repo. See that ADR for cadence, tooling, and disposition vocabulary.

## Tracking

- **Upstream project.** `python-openxml/python-docx`
  (`https://github.com/python-openxml/python-docx.git`).
  Note: the original upstream `scanny/python-docx` has been deleted;
  `python-openxml/python-docx` is the surviving canonical repo and
  carries the `v1.2.0` tag.
- **Fork baseline tag.** `v1.2.0`
- **Fork baseline SHA.** `e45454602b53e8e572b179ccf1c91093ec9f4ed7`
  (upstream `master` HEAD as of 2026-05-05; the release commit subject
  is `release: prepare v1.2.0 release`, dated 2025-06-16).
- **Fork divergent commit count.** 246 (fork additions on top of the
  baseline, as of 2026-05-05).
- **Upstream remote** is **not** configured in a default clone.
  Maintainers add it on demand during a sweep:
  ```bash
  git remote add upstream https://github.com/python-openxml/python-docx.git
  git fetch upstream
  git log --no-merges --oneline e4545460..upstream/master
  ```

## Sync status (sweep: 2026-05-05)

As of this sweep `upstream/master` is **at the baseline SHA** — no
new commits have landed upstream since v1.2.0. There is nothing to
evaluate and nothing to pull.

Upstream activity since baseline:

| short-sha  | date       | subject                               | tier   | disposition | rationale |
|------------|------------|---------------------------------------|--------|-------------|-----------|
| *(none)*   | —          | —                                     | —      | —           | `upstream/master` == fork baseline SHA |

Non-master upstream branches observed (not pulled):

- `feature/bookmarks` — experimental; fork has its own bookmark
  authoring surface (Phase D). Disposition: `blocked-by-fork-divergence`.
- `feature/header` — experimental; fork ships header/footer APIs.
  Disposition: `blocked-by-fork-divergence`.

## Next sweep due

**2026-08-03** (first Monday of August 2026).

If the upstream repository continues to move away from `scanny` in
this period, re-verify the canonical URL at the start of the sweep
(GitHub redirects are not stable for long-dead forks).

## History

- 2026-05-05 — initial catalogue created by Wave 11-D. Baseline
  confirmed; zero upstream divergence.
