# pyright: reportPrivateUsage=false

"""Regression tests for the one-shot ``huge_tree=True`` warning.

``Document(..., huge_tree=True)`` opts the caller into lxml's
``huge_tree`` parser mode, which disables libxml2's built-in
XML-bomb safeguards (the 10 MB per-AttValue cap and the 256-deep
nesting limit). Because the flag is trivially discoverable by
developers ingesting untrusted documents, python-docx emits a
``UserWarning`` the first time the relaxed parser actually runs.

The warning is one-shot per process (guarded by the module-level
``_huge_tree_warned`` flag) — subsequent calls stay quiet.
"""

from __future__ import annotations

import warnings

from docx.oxml import parser as parser_mod
from docx.oxml.parser import huge_tree_mode, parse_xml


def _reset_warned_flag():
    parser_mod._huge_tree_warned = False


def it_emits_one_warning_on_first_huge_tree_parse():
    _reset_warned_flag()
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        with huge_tree_mode():
            parse_xml(b"<root/>")

    user_warnings = [
        w for w in captured if issubclass(w.category, UserWarning)
    ]
    assert len(user_warnings) == 1
    assert "huge_tree" in str(user_warnings[0].message)


def it_only_warns_once_across_repeated_calls():
    _reset_warned_flag()
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        with huge_tree_mode():
            parse_xml(b"<root/>")
            parse_xml(b"<root/>")
            parse_xml(b"<root/>")

    user_warnings = [
        w for w in captured if issubclass(w.category, UserWarning)
    ]
    assert len(user_warnings) == 1


def it_does_not_warn_when_huge_tree_flag_is_off():
    _reset_warned_flag()
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        parse_xml(b"<root/>")

    user_warnings = [
        w
        for w in captured
        if issubclass(w.category, UserWarning)
        and "huge_tree" in str(w.message)
    ]
    assert user_warnings == []
