"""pytest fixtures that are shared across test modules."""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from docx.package import Package
from tests.unitutil.mock import instance_mock

if TYPE_CHECKING:
    from docx import types as t
    from docx.parts.story import StoryPart
    from tests.unitutil.mock import FixtureRequest, Mock


@pytest.fixture
def fake_parent() -> t.ProvidesStoryPart:
    class ProvidesStoryPart:
        @property
        def part(self) -> StoryPart:
            raise NotImplementedError

    return ProvidesStoryPart()


@pytest.fixture
def package_(request: FixtureRequest) -> Mock:
    """Mock `docx.package.Package` instance, shared across test modules."""
    return instance_mock(request, Package)
