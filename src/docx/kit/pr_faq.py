"""Amazon-style press-release / FAQ ("PR/FAQ") template family.

Closes #295.

Amazon's internal product-development practice is to write the press
release *first* — before any code is written — and pair it with a
"frequently asked questions" addendum that surfaces the difficult
questions the team will have to answer before launch. Working
backwards from the customer-facing announcement forces clarity about
who the customer is, what problem the product solves, and what the
launch will actually feel like to a reader.

This module exposes three composition helpers::

    from docx import Document
    from docx.kit import pr_faq

    doc = Document()
    pr_faq.press_release(
        doc,
        headline="Acme launches FrobnitzPro",
        subheadline="The fastest frobnitz on the market",
        location="Seattle, WA",
        date="2026-05-29",
        summary="One-paragraph summary of the launch...",
        problem="Customers struggle with...",
        solution="FrobnitzPro solves this by...",
        quote_speaker="Jane Doe, VP Product",
        quote_text='"FrobnitzPro is the most exciting product we have launched in a decade."',
        customer_quote_speaker="John Smith, ACME Corp",
        customer_quote_text='"It has revolutionised our workflow."',
        call_to_action="Visit acme.com/frobnitz to learn more.",
    )
    pr_faq.faq(
        doc,
        items=[
            ("What is FrobnitzPro?", "It is a frobnitz that..."),
            ("How much does it cost?", "$99/month..."),
            ("When is it available?", "Today."),
        ],
    )
    doc.save("pr_faq.docx")

The shape of an Amazon press release is canonical: headline,
optional subheadline, dateline (``LOCATION — DATE``), one-paragraph
summary, the problem, the solution, an internal-spokesperson quote,
an optional customer quote, and a call to action. The FAQ is a
straightforward question / answer list. :func:`press_release` and
:func:`faq` each *append* to the end of an existing |Document| and
return the list of newly-appended paragraphs in document order;
:func:`pr_faq_doc` is a one-line convenience that builds a fresh
document with both sections and saves it (or returns it).

The helpers prefer Word's conventional built-in styles
(``Title``, ``Subtitle``, ``Heading 1``, ``Heading 2``, ``Quote``)
and fall back to ``Normal`` when a custom template lacks a style —
the spirit of a *kit* is "works out of the box, customise as you
like".

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Sequence, Tuple, Union

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the kit reaches for. The default python-docx
# -- template ships ``Normal``, ``Title``, ``Subtitle``, the
# -- ``Heading 1..9`` family, and ``Quote``. When a caller-supplied
# -- template is missing a particular style we fall back to ``Normal``
# -- rather than raise — same policy as ``front_matter`` /
# -- ``back_matter``.
_STYLE_TITLE = "Title"
_STYLE_SUBTITLE = "Subtitle"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_QUOTE = "Quote"
_STYLE_NORMAL = "Normal"


def _has_style(document: "DocumentCls", style_name: str) -> bool:
    """Return |True| when `document` defines a paragraph style named `style_name`."""
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve_style(document: "DocumentCls", preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _add_quote_block(
    document: "DocumentCls",
    speaker: str,
    text: str,
) -> List["Paragraph"]:
    """Append a two-paragraph quote block (text + attribution) and return them.

    The quote text is rendered in the ``Quote`` style (or ``Normal``
    fallback). The attribution sits on its own paragraph in italics so
    the reader can scan the speaker without losing the visual rhythm
    of the surrounding press release.
    """
    paragraphs: List["Paragraph"] = []
    quote_style = _resolve_style(document, _STYLE_QUOTE)

    quote_para = document.add_paragraph(text, style=quote_style)
    paragraphs.append(quote_para)

    attribution = document.add_paragraph()
    attribution_run = attribution.add_run(f"— {speaker}")
    attribution_run.italic = True
    paragraphs.append(attribution)

    return paragraphs


# -- Press release -------------------------------------------------------


def press_release(
    document: "DocumentCls",
    *,
    headline: str,
    subheadline: Optional[str] = None,
    location: str,
    date: str,
    summary: str,
    problem: str,
    solution: str,
    quote_speaker: str,
    quote_text: str,
    customer_quote_speaker: Optional[str] = None,
    customer_quote_text: Optional[str] = None,
    call_to_action: str,
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append an Amazon-style press release to `document` and return the new paragraphs.

    Renders (in order):

    1. A centred ``Title``-styled headline paragraph.
    2. An optional centred ``Subtitle``-styled subheadline paragraph.
    3. A bold dateline of the form ``"LOCATION — DATE — "`` whose
       trailing run is the `summary` body. Word press-release house
       style emits the location and date as the first run of the
       opening summary paragraph; the kit follows that convention.
    4. A ``Heading 2`` "The Problem" paragraph followed by `problem`.
    5. A ``Heading 2`` "The Solution" paragraph followed by `solution`.
    6. A spokesperson quote (``Quote`` style) attributed to
       `quote_speaker`.
    7. An optional customer quote attributed to
       `customer_quote_speaker` (only emitted when both `customer_quote_*`
       arguments are supplied).
    8. A bold "Call to action:" paragraph whose body is `call_to_action`.
    9. A trailing page break, when ``page_break`` is true (the default).

    All required arguments must be non-empty strings. The optional
    customer-quote pair is only honoured when both halves are
    supplied; supplying just one raises ``ValueError``.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true.

    Parameters
    ----------
    document
        The target |Document|. Required.
    headline
        Required headline text.
    subheadline
        Optional subheadline rendered under the headline.
    location, date
        Required dateline pieces (e.g. ``"Seattle, WA"`` and
        ``"2026-05-29"``).
    summary
        Required one-paragraph launch summary.
    problem, solution
        Required problem and solution paragraphs.
    quote_speaker, quote_text
        Required spokesperson attribution and quotation.
    customer_quote_speaker, customer_quote_text
        Optional customer-testimonial pair. Both must be supplied
        together.
    call_to_action
        Required call-to-action line.
    page_break
        When |True| (the default), a trailing page break is appended
        so the FAQ that typically follows starts on a fresh page.

    Raises
    ------
    ValueError
        When any required argument is empty or only one half of the
        customer-quote pair is supplied.

    .. versionadded:: 2026.05.29
    """
    _required = {
        "headline": headline,
        "location": location,
        "date": date,
        "summary": summary,
        "problem": problem,
        "solution": solution,
        "quote_speaker": quote_speaker,
        "quote_text": quote_text,
        "call_to_action": call_to_action,
    }
    for name, value in _required.items():
        if not value or not str(value).strip():
            raise ValueError(f"{name} must be a non-empty string")

    customer_supplied = (
        bool(customer_quote_speaker and str(customer_quote_speaker).strip()),
        bool(customer_quote_text and str(customer_quote_text).strip()),
    )
    if any(customer_supplied) and not all(customer_supplied):
        raise ValueError(
            "customer_quote_speaker and customer_quote_text must be supplied "
            "together (both or neither)"
        )

    paragraphs: List["Paragraph"] = []
    title_style = _resolve_style(document, _STYLE_TITLE)
    subtitle_style = _resolve_style(document, _STYLE_SUBTITLE)
    heading_2_style = _resolve_style(document, _STYLE_HEADING_2)

    # -- Headline --
    headline_para = document.add_paragraph(headline, style=title_style)
    headline_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraphs.append(headline_para)

    # -- Subheadline --
    if subheadline:
        sub_para = document.add_paragraph(subheadline, style=subtitle_style)
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(sub_para)

    # -- Dateline + summary (single paragraph, bold dateline run, body run) --
    dateline_para = document.add_paragraph()
    dateline_run = dateline_para.add_run(f"{location} — {date} — ")
    dateline_run.bold = True
    dateline_para.add_run(summary)
    paragraphs.append(dateline_para)

    # -- Problem --
    problem_heading = document.add_paragraph("The Problem", style=heading_2_style)
    paragraphs.append(problem_heading)
    paragraphs.append(document.add_paragraph(problem))

    # -- Solution --
    solution_heading = document.add_paragraph("The Solution", style=heading_2_style)
    paragraphs.append(solution_heading)
    paragraphs.append(document.add_paragraph(solution))

    # -- Spokesperson quote --
    paragraphs.extend(_add_quote_block(document, quote_speaker, quote_text))

    # -- Customer quote (optional pair) --
    if all(customer_supplied):
        # -- type-narrowing: both halves are non-empty by the validation
        # -- above; cast for the type checker.
        assert customer_quote_speaker is not None
        assert customer_quote_text is not None
        paragraphs.extend(
            _add_quote_block(document, customer_quote_speaker, customer_quote_text)
        )

    # -- Call to action --
    cta_para = document.add_paragraph()
    cta_label = cta_para.add_run("Call to action: ")
    cta_label.bold = True
    cta_para.add_run(call_to_action)
    paragraphs.append(cta_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


# -- FAQ -----------------------------------------------------------------


def faq(
    document: "DocumentCls",
    *,
    items: Sequence[Tuple[str, str]],
    title: str = "Frequently Asked Questions",
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append a Q&A list to `document` and return the new paragraphs.

    Renders (in order):

    1. A ``Heading 1`` paragraph holding `title` (default
       ``"Frequently Asked Questions"``). Pass ``title=""`` to
       suppress the heading and emit only the Q&A pairs.
    2. For each ``(question, answer)`` tuple in `items`, a paragraph
       whose first run is ``"Q: "`` (bold) followed by the question
       text, then a paragraph whose first run is ``"A: "`` (bold)
       followed by the answer text.
    3. A trailing page break, when ``page_break`` is true (the default).

    `items` is a sequence of ``(question, answer)`` 2-tuples. Both
    members of every tuple must be non-empty strings.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true.

    Raises
    ------
    ValueError
        When `items` is empty or contains a tuple whose question or
        answer is empty / missing.

    .. versionadded:: 2026.05.29
    """
    if not items:
        raise ValueError("items must contain at least one (question, answer) tuple")

    coerced: List[Tuple[str, str]] = []
    for index, item in enumerate(items):
        if not isinstance(item, tuple) or len(item) != 2:
            raise ValueError(
                "items[%d] must be a (question, answer) 2-tuple" % index
            )
        question, answer = item
        if not question or not str(question).strip():
            raise ValueError("items[%d] question must be a non-empty string" % index)
        if not answer or not str(answer).strip():
            raise ValueError("items[%d] answer must be a non-empty string" % index)
        coerced.append((str(question), str(answer)))

    paragraphs: List["Paragraph"] = []
    heading_1_style = _resolve_style(document, _STYLE_HEADING_1)

    if title:
        heading_para = document.add_paragraph(title, style=heading_1_style)
        paragraphs.append(heading_para)

    for question, answer in coerced:
        q_para = document.add_paragraph()
        q_label = q_para.add_run("Q: ")
        q_label.bold = True
        q_para.add_run(question)
        paragraphs.append(q_para)

        a_para = document.add_paragraph()
        a_label = a_para.add_run("A: ")
        a_label.bold = True
        a_para.add_run(answer)
        paragraphs.append(a_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


# -- Convenience builder -------------------------------------------------


def pr_faq_doc(
    *,
    press_release_kwargs: dict,
    faq_items: Sequence[Tuple[str, str]],
    output_path: Optional[str] = None,
) -> "DocumentCls":
    """Build a fresh PR/FAQ |Document|, optionally save it, and return it.

    Convenience wrapper around :func:`press_release` and :func:`faq` that
    creates a fresh |Document|, appends the press release using the
    keyword arguments in `press_release_kwargs`, then appends the FAQ
    using `faq_items`. When `output_path` is supplied the document is
    saved at that path; the |Document| is returned in either case so
    callers can perform further mutation before re-saving.

    Parameters
    ----------
    press_release_kwargs
        A mapping of keyword arguments to forward to
        :func:`press_release`. ``page_break`` is honoured if present;
        when omitted, the default (``True``) is used so the FAQ
        starts on a fresh page.
    faq_items
        The list of ``(question, answer)`` tuples to forward to
        :func:`faq`.
    output_path
        Optional file path. When supplied, the document is saved
        there before being returned.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with
        :meth:`Document.save` if `output_path` was not supplied.

    Raises
    ------
    ValueError
        When the underlying :func:`press_release` or :func:`faq` call
        rejects its arguments.

    .. versionadded:: 2026.05.29
    """
    if not isinstance(press_release_kwargs, dict):
        raise ValueError("press_release_kwargs must be a dict")

    document = Document()
    press_release(document, **press_release_kwargs)
    faq(document, items=faq_items)

    if output_path is not None:
        document.save(output_path)

    return document


__all__ = [
    "press_release",
    "faq",
    "pr_faq_doc",
]
