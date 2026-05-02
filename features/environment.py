"""Set testing environment before and after behave acceptance test runs.

Helpers for step implementations live in ``features/steps/helpers.py``:

- ``test_docx(name)``     — absolute path to a fixture ``.docx`` in ``features/steps/test_files/``
- ``test_file(name)``     — absolute path to any fixture file in ``features/steps/test_files/``
- ``saved_docx_path``     — canonical scratch output path (``features/_scratch/test_out.docx``)

Tag conventions (extend as needed, no enforcement today):

- ``@slow``          — scenarios that take noticeably longer than the ~ms average
- ``@fixture-heavy`` — scenarios that exercise large fixtures or LibreOffice round-trips

Example invocations::

    behave features/                         # run everything
    behave --tags "not @slow" features/      # skip slow scenarios
    behave --tags "@fixture-heavy" features/ # run only fixture-heavy scenarios
"""

from __future__ import annotations

import os
import time

scratch_dir = os.path.abspath(os.path.join(os.path.split(__file__)[0], "_scratch"))


def before_all(context):
    if not os.path.isdir(scratch_dir):
        os.mkdir(scratch_dir)


def before_scenario(context, scenario):
    """Record a baseline mtime so ``after_scenario`` can scrub fresh scratch files."""
    # -- ensure the scratch dir exists even if before_all was skipped (e.g. --dry-run) --
    if not os.path.isdir(scratch_dir):
        os.mkdir(scratch_dir)
    # -- a tiny epsilon avoids races where a file written in the same tick as the
    # -- baseline is treated as pre-existing. time.time() has sub-second resolution. --
    context._scratch_baseline_time = time.time()


def after_scenario(context, scenario):
    """Remove scratch files created during this scenario.

    Only files whose mtime is newer than ``context._scratch_baseline_time`` are
    removed; pre-existing fixtures stay put. Sub-directories and their contents
    are left alone — scenarios that want to manage a tree should clean it up
    themselves.
    """
    baseline = getattr(context, "_scratch_baseline_time", None)
    if baseline is None or not os.path.isdir(scratch_dir):
        return
    for name in os.listdir(scratch_dir):
        path = os.path.join(scratch_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            continue
        if mtime >= baseline:
            try:
                os.remove(path)
            except OSError:
                # -- best-effort cleanup; don't fail the scenario over a stray file --
                pass
