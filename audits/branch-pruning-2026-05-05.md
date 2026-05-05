# Branch pruning — 2026-05-05

Remote branches deleted from `origin` as part of the Wave-C audit cleanup.
Every branch listed here was strictly merged into `origin/master` at time
of deletion (verified with `git branch -r --merged origin/master`), so no
work was lost. The commit history is preserved via the merge commits
already on master.

## Pruned branches (7)

- `origin/feat/w10-a-smartart-authoring`  — merged in `0c873592`
- `origin/feat/w10-b-bibliography-authoring`  — merged in `19221e52`
- `origin/feat/w10-f-field-eval`  — merged in `100c9449`
- `origin/fix/w11-a-indexing-perf`  — merged in `9773d977`
- `origin/fix/w8-a-part-drop-narrowing`  — merged in `46f92e0f`
- `origin/fix/w8-b-reproducible-fixes`  — merged in `30e16052`
- `origin/fix/w8-e-api-gaps`  — merged in `71355efb`

## Intentionally retained

- `origin/feat/w11-d-upstream-sync`  — retained per audit policy (handled by Wave-D).
- `origin/feat/w1-e-conformance-ci`  — parked per policy (conformance CI still in design).
- `origin/agent/issue-28`, `origin/develop`, `origin/fix/overnight-n4-section-valign`
  — merged, but outside the Wave-C safe-to-prune patterns (`feat/w10-*`,
  `fix/w8-*`, `fix/w11-a-*`, `chore/overnight-*`, `worktree-agent-*`).
  Left alone for a future hygiene pass.
