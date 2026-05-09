"""OOXML-validate conformance harness for python-docx.

Companion to :mod:`tests.conformance.test_round_trip`. Where the sibling
harness enforces the **byte-round-trip** contract (the saved package
must be byte-identical to the input), this module enforces an
orthogonal **schema-validity** contract:

    python-docx's save output must validate cleanly against the
    Microsoft Open XML SDK validator.

The two contracts catch different classes of bug:

- Byte-round-trip catches silent drops (parts lost, attributes
  reordered, elements elided) on fixtures Word itself wrote.
- ``ooxml-validate`` catches spec-invalid output on *any* save —
  including saves Microsoft Word would still open, because Word is
  famously lenient about its own schema. Something that round-trips
  fine but fails validation means python-docx is emitting OOXML that
  happens to survive Word's leniency but that strict consumers (other
  Office clients, Open XML SDK tooling, LibreOffice in strict mode)
  may reject.

Scope — pure instrumentation, infrastructure-not-mounted friendly
----------------------------------------------------------------

- If :mod:`ooxml_validate` is not importable (not installed, or the
  optional ``conformance`` extra not added), the whole module skips.
- If the .NET runtime is missing at call time
  (:exc:`ooxml_validate.OoxmlValidateToolNotFound`), the whole module
  skips — bundled validator needs ``dotnet`` on PATH. Install with
  ``sudo apt-get install dotnet-runtime-8.0`` or point
  ``OOXML_VALIDATE_DOTNET`` at an explicit executable.
- Fixtures the reader refused (surfaced by ``test_round_trip`` as an
  empty ``saved_bytes`` slot) are skipped here too — can't validate
  bytes that were never produced.
- Fixtures whose byte-round-trip *fails* are still validated —
  validation is about the saved output's spec conformance, not its
  fidelity to the original.

When a fixture fails, the failure gets its own ticket; this harness
stays frozen as a surface, not a fixer. See the docx ``CLAUDE.md``
section "Running the conformance harness" for the developer-facing
operator docs.

Opt-in install
--------------

``ooxml-validate`` is listed under the ``conformance`` extra; install
with::

    pip install -e '.[conformance]'

CI runs this lane only when the extra is present; default ``test``
extras skip it via the import-guard below.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from tests.conformance.conftest import ALL_FIXTURE_IDS, ALL_FIXTURES

# Module-level marker so ``pytest -m conformance`` selects this
# harness alongside ``test_round_trip``. All tests inherit it.
pytestmark = pytest.mark.conformance


# ---------------------------------------------------------------------------
# Import guard — infrastructure-not-mounted pattern
# ---------------------------------------------------------------------------
#
# ``ooxml-validate`` is an optional conformance-extra dep. On a default
# dev install (``pip install -e '.[dev]'``) the package isn't present
# and importing it raises ModuleNotFoundError. Skip the whole module in
# that case so the default CI lane stays green without the validator
# mounted.

ooxml_validate = pytest.importorskip(
    "ooxml_validate",
    reason=(
        "ooxml-validate not installed — run "
        "`pip install -e '.[conformance]'` to enable the schema-"
        "validation harness."
    ),
)


# ---------------------------------------------------------------------------
# Module-cached load + save
# ---------------------------------------------------------------------------
#
# The ``test_round_trip`` sibling also computes a ``{name: (orig, saved)}``
# map, but its fixture is module-scoped to that module and not visible
# across test files. Rather than refactor the byte-equality contract's
# harness (explicit non-goal of this instrumentation pass — see
# module docstring "Scope") we recompute the round-trip here; cached
# once per module so the parametrised tests share it.


@pytest.fixture(scope="module")
def _saved_bytes_by_fixture() -> dict[str, bytes]:
    """Return ``{fixture_name: saved_bytes}``.

    Fixtures the reader rejects outright are omitted; the per-fixture
    test skips on that basis with :exc:`KeyError` handling. We
    intentionally don't surface the underlying exception here — the
    sibling ``test_round_trip`` already does, and surfacing it twice
    just clutters output with duplicate failure context.
    """
    from docx import Document  # lazy — keep importable without editable install

    out: dict[str, bytes] = {}
    for path in ALL_FIXTURES:
        try:
            document = Document(str(path))
            buf = io.BytesIO()
            document.save(buf)
            out[path.name] = buf.getvalue()
        except Exception:  # noqa: BLE001 — reader refusals handled via absent key
            continue
    return out


# ---------------------------------------------------------------------------
# Per-fixture validation
# ---------------------------------------------------------------------------


@pytest.mark.skipif(
    not ALL_FIXTURES,
    reason=(
        "no docx fixtures available — land *.office.docx files under "
        "~/code/ooxml-reference-corpus/fixtures/docx/ (Office-authored "
        "via Microsoft Word desktop) or point DOCX_CORPUS_ROOT at an "
        "alternative corpus"
    ),
)
@pytest.mark.parametrize(
    "fixture_path",
    ALL_FIXTURES,
    ids=ALL_FIXTURE_IDS,
)
def it_saves_spec_valid_ooxml(
    fixture_path: Path,
    _saved_bytes_by_fixture: dict[str, bytes],
    tmp_path: Path,
) -> None:
    """The Open XML SDK validator must report zero issues on save output.

    The validator only accepts a filesystem path, not bytes, so we
    stage the saved bytes into ``tmp_path`` and call ``validate()`` on
    that copy.

    A non-empty issue list fails the test with a length-capped
    rendering (first 20 issues at most) so the report stays readable
    when a single regression produces hundreds of findings.
    """
    name = fixture_path.name

    if name not in _saved_bytes_by_fixture:
        # Reader refused this fixture — ``test_round_trip`` already
        # surfaces the underlying exception. Nothing to validate here.
        pytest.skip(f"{name}: reader refused the fixture; nothing to validate.")

    saved_bytes = _saved_bytes_by_fixture[name]
    staged = tmp_path / name
    staged.write_bytes(saved_bytes)

    try:
        issues = ooxml_validate.validate(staged)
    except ooxml_validate.OoxmlValidateToolNotFound as exc:
        # The .NET runtime is the infrastructure leg of this harness.
        # Treat a missing ``dotnet`` the same as a missing validator
        # package — skip cleanly with an actionable message.
        pytest.skip(
            f"ooxml-validate bundled .NET CLI could not run: {exc} — "
            "install the .NET 8+ runtime "
            "(`sudo apt-get install dotnet-runtime-8.0`) or set "
            "OOXML_VALIDATE_DOTNET to point at a dotnet executable."
        )
    except ooxml_validate.OoxmlValidateError as exc:
        # Validator ran but produced unparseable output. Surface the
        # error rather than masking it as a pass — something is wrong
        # with the validator install, not necessarily the fixture.
        pytest.fail(
            f"{name}: ooxml-validate subprocess failed: "
            f"{type(exc).__name__}: {exc}"
        )

    if not issues:
        return  # status='pass' — empty issue list per the validator's contract.

    preview = issues[:20]
    extra = len(issues) - len(preview)
    lines = [
        f"{name}: ooxml-validate reported {len(issues)} issue(s) on the",
        "saved package. python-docx is emitting OOXML the Open XML SDK",
        "validator considers invalid — Word may still open it, but the",
        "spec does not. First findings:",
        "",
    ]
    lines.extend(f"  {issue}" for issue in preview)
    if extra > 0:
        lines.append(f"  ... ({extra} further issue(s) truncated)")
    pytest.fail("\n".join(lines))
