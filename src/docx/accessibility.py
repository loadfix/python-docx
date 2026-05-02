"""Accessibility helpers for python-docx documents.

This module provides utilities to detect common accessibility issues in a document's
heading structure, such as skipped outline levels, multiple top-level headings, or
empty heading paragraphs. Screen readers and assistive technologies rely on a clean
heading hierarchy to navigate a document, so flagging these issues helps authors
produce more accessible content.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph


# -- kind constants for HeadingIssue.kind --
SKIPPED_LEVEL = "skipped_level"
MULTIPLE_H1 = "multiple_h1"
EMPTY_HEADING = "empty_heading"
NO_H1 = "no_h1"


_HEADING_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)


@dataclass(frozen=True)
class HeadingIssue:
    """A single heading-structure issue reported by :func:`validate_heading_structure`.

    `paragraph` is the |Paragraph| that exhibits the issue. `kind` is a short string
    identifier (one of ``"skipped_level"``, ``"multiple_h1"``, ``"empty_heading"``, or
    ``"no_h1"``). `message` is a human-readable description of the problem suitable for
    display to the author.

    .. versionadded:: 2026.05.0
    """

    paragraph: Paragraph
    kind: str
    message: str


def _heading_level(paragraph: Paragraph) -> int | None:
    """Return the integer heading level for `paragraph`, or |None| if not a heading.

    A paragraph is considered a heading when its style name matches "Heading N" where
    N is 1-9, case-insensitively (so "heading 2" and "HEADING 2" also match).
    """
    style = paragraph.style
    if style is None:
        return None
    name = style.name
    if name is None:
        return None
    match = _HEADING_RE.match(name.strip())
    if match is None:
        return None
    return int(match.group(1))


def validate_heading_structure(
    paragraphs: Iterable[Paragraph],
) -> list[HeadingIssue]:
    """Return a list of |HeadingIssue| objects describing heading-structure problems.

    The following accessibility issues are detected:

    * ``"skipped_level"`` — a heading skips one or more outline levels (e.g. a
      "Heading 3" that directly follows a "Heading 1" without an intervening
      "Heading 2").
    * ``"multiple_h1"`` — the document contains more than one top-level heading
      ("Heading 1"). Only the *second* and later H1s are flagged.
    * ``"empty_heading"`` — a heading paragraph has no visible text (after
      stripping whitespace).
    * ``"no_h1"`` — the first heading in the document is below "Heading 1"
      (e.g. starts at "Heading 2"). The offending heading paragraph is flagged.

    Non-heading paragraphs are ignored. Issues are returned in document order.

    .. versionadded:: 2026.05.0
    """
    issues: list[HeadingIssue] = []
    previous_level: int | None = None
    h1_count = 0
    first_heading_seen = False

    for paragraph in paragraphs:
        level = _heading_level(paragraph)
        if level is None:
            continue

        # -- empty-heading check --
        if not paragraph.text.strip():
            issues.append(
                HeadingIssue(
                    paragraph=paragraph,
                    kind=EMPTY_HEADING,
                    message=f"Heading {level} paragraph is empty",
                )
            )

        # -- no-H1 check: first heading is not H1 --
        if not first_heading_seen and level > 1:
            issues.append(
                HeadingIssue(
                    paragraph=paragraph,
                    kind=NO_H1,
                    message=(
                        f"First heading is Heading {level}; document is missing "
                        f"a top-level Heading 1"
                    ),
                )
            )

        # -- multiple-H1 check --
        if level == 1:
            h1_count += 1
            if h1_count > 1:
                issues.append(
                    HeadingIssue(
                        paragraph=paragraph,
                        kind=MULTIPLE_H1,
                        message=(
                            "Document contains more than one Heading 1; "
                            "exactly one top-level heading is recommended"
                        ),
                    )
                )

        # -- skipped-level check: heading level jumps by more than 1 --
        if previous_level is not None and level > previous_level + 1:
            missing = previous_level + 1
            issues.append(
                HeadingIssue(
                    paragraph=paragraph,
                    kind=SKIPPED_LEVEL,
                    message=(
                        f"Heading {level} follows Heading {previous_level}; "
                        f"Heading {missing} is missing"
                    ),
                )
            )

        first_heading_seen = True
        previous_level = level

    return issues
