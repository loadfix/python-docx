"""Property-based tests for python-docx using Hypothesis.

This package contains property-based tests that exercise python-docx's
read/write contracts with randomly-generated input. Tests live here
rather than under ``tests/unit/`` so they can be opted-in or skipped
independently when ``hypothesis`` is not installed.

Pattern
-------

1. Define a Hypothesis strategy (``@composite`` or ``st.*``) that
   produces well-formed input — text, run formatting, table shapes,
   etc. The strategies blacklist control characters and lone
   surrogates that the OOXML spec forbids.
2. Drive ``docx.Document`` authoring with the generated input, save
   to an in-memory ``BytesIO``, reload, and assert that the round
   trip preserves the input.

Each property test uses Hypothesis's default settings
(``max_examples=100``); per-test ``@settings`` overrides are applied
where the cost of one example is high (e.g. table authoring).

Run
---

    pytest tests/property/ -q

The dependency on ``hypothesis`` is dev-only; production users do
not need it installed.
"""

from __future__ import annotations
