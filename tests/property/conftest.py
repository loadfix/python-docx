"""Local pytest config for the property test package.

The repository-wide ``filterwarnings = ["error"]`` in ``pyproject.toml``
escalates Hypothesis's ``UserWarning`` about its ``.hypothesis``
example database directory into a collection error. Hypothesis emits
this warning whenever a project sets a custom ``norecursedirs`` (we
do, to skip ``features/`` and friends) because that replaces — rather
than extends — pytest's defaults. Silence it locally so the property
suite collects cleanly under the strict filter.
"""

from __future__ import annotations

import warnings

warnings.filterwarnings(
    "ignore",
    message="Skipping collection of '.hypothesis' directory.*",
    category=UserWarning,
)
