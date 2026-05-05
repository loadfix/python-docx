# Release automation — python-docx (fork)

This directory holds the fork's release tooling. `.github/workflows/` is
deliberately absent on this fork (removed 2026-05-02, commit `2aa5e4c`),
so the PyPI-publish workflow lives here as a parked template rather than
an active workflow.

## Files

- `release.yml.parked` — GitHub Actions workflow that, when moved to
  `.github/workflows/release.yml`, builds `sdist` + `wheel` on every
  `v*` tag push and publishes to PyPI via trusted publishing (OIDC, no
  stored API tokens). See the header comment in the file for the exact
  activation steps.
- The matching `scripts/prepare_release.py` (one level up) is the
  *local* half of the flow: bump `__version__`, insert a `HISTORY.rst`
  entry, validate its formatting, and run `pytest` + `pyright`. That
  script runs today without any GitHub Actions configuration.

## End-to-end release flow (once activated)

1. On a clean tree:
   `python scripts/prepare_release.py 2026.MM.N --title "Short title"`
   — bumps version, prompts for the HISTORY entry, validates format,
   runs `pytest` + `pyright`.
2. Review: `git diff`, then commit:
   `git commit -am 'chore(release): bump to 2026.MM.N'`.
3. Tag and push:
   `git tag -a v2026.MM.N -m '2026.MM.N'`
   `git push origin master --follow-tags`.
4. The tag trigger fires `release.yml` (once activated) → build →
   publish to PyPI.

See also `docs/adr/005-release-process.md` in
`loadfix/ooxml-reference-corpus` for the cross-library rationale.
