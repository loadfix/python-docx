"""pytest fixtures that are shared across test modules."""

from __future__ import annotations

from typing import TYPE_CHECKING
from unittest.mock import MagicMock

import pytest

if TYPE_CHECKING:
    from docx import types as t
    from docx.opc.package import OpcPackage
    from docx.parts.story import StoryPart


@pytest.fixture
def fake_parent() -> t.ProvidesStoryPart:
    class ProvidesStoryPart:
        @property
        def part(self) -> StoryPart:
            raise NotImplementedError

    return ProvidesStoryPart()


@pytest.fixture
def fake_package() -> OpcPackage:
    """A mock OpcPackage suitable for constructing parts in tests."""
    return MagicMock(name="package")
