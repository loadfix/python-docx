# pyright: reportPrivateUsage=false

"""Focused regression tests for the hardened ``=``-field formula evaluator.

``docx.fields._evaluate_formula`` previously called :func:`eval` on input
that had passed a character-class whitelist. The whitelist accepted two
adjacent ``*`` characters, letting ``**`` (exponentiation) through: a
field ``= 9**9**9**9`` would trigger an effectively-unbounded bignum
computation and pin CPU.

After remediation the evaluator parses the expression with :mod:`ast`
and walks it with a node allow-list that deliberately omits
``ast.Pow``, so exponentiation is rejected regardless of what the
character filter accepts.
"""

from __future__ import annotations

import pytest

from docx.fields import _evaluate_formula


@pytest.mark.parametrize(
    "expr,expected",
    [
        ("= 1+2", "3"),
        ("= 2+3*4", "14"),
        ("= (1+2)*(3+4)", "21"),
        ("= -5 + 10", "5"),
        ("= +3", "3"),
        ("= 7/2", "3.5"),
        ("= 10 % 3", "1"),
        ("= 100 - 58", "42"),
    ],
)
def it_evaluates_basic_arithmetic(expr, expected):
    assert _evaluate_formula(expr, {}) == expected


@pytest.mark.parametrize(
    "expr",
    [
        "= 2**8",
        "= 9**9",
        "= 9**9**9**9",
        "= 2 ** 3",
    ],
)
def it_rejects_power_operator(expr):
    assert _evaluate_formula(expr, {}) is None


def it_rejects_deeply_nested_power_chains_quickly():
    # -- pin: the evaluator must *not* attempt to compute this. The AST
    # -- check rejects ``ast.Pow`` before any numeric work happens, so
    # -- this returns |None| instantly instead of hanging. --
    import time

    start = time.monotonic()
    result = _evaluate_formula("= 9**9**9**9", {})
    elapsed = time.monotonic() - start

    assert result is None
    assert elapsed < 1.0  # -- generous ceiling; actual is microseconds --


@pytest.mark.parametrize(
    "expr",
    [
        "= __import__('os')",
        "= foo.bar",
        "= 1;2",
        "= [1,2,3]",
        "= (x for x in range(10))",
    ],
)
def it_rejects_non_arithmetic_constructs(expr):
    assert _evaluate_formula(expr, {}) is None


def it_returns_none_on_syntax_errors():
    assert _evaluate_formula("= 1+", {}) is None
    assert _evaluate_formula("=", {}) is None


def it_returns_none_when_not_a_formula_prefix():
    assert _evaluate_formula("TOC \\o", {}) is None
