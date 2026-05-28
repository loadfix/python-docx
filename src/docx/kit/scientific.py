"""Scientific paper template family — IEEE / ACM / APA / Nature.

Closes #82.

This module exposes four template factories that build entire
scientific-paper drafts in one call::

    from docx.kit.scientific import (
        ieee_paper,
        acm_paper,
        apa_paper,
        nature_paper,
    )

    doc = ieee_paper(
        title='A Distributed Consensus Algorithm',
        authors=[
            {'name': 'Alice', 'affiliation': 'Acme Corp', 'email': 'alice@acme.com'},
            {'name': 'Bob',   'affiliation': 'Beta Labs', 'email': 'bob@beta.io'},
        ],
        abstract='We present...',
        keywords=['consensus', 'distributed systems'],
        sections=[
            {'heading': 'Introduction',  'body': '...'},
            {'heading': 'Related Work',  'body': '...'},
            {'heading': 'Algorithm',     'body': '...'},
            {'heading': 'Evaluation',    'body': '...'},
            {'heading': 'Conclusion',    'body': '...'},
        ],
        references=[
            {'authors': 'Lamport, L.',
             'title':   'The Part-Time Parliament',
             'venue':   'TOCS',
             'year':    1998},
        ],
    )
    doc.save('paper.docx')

The four factories — :func:`ieee_paper`, :func:`acm_paper`,
:func:`apa_paper`, :func:`nature_paper` — each return a fresh
|Document| pre-populated with the conventional sections and the
typesetting style that publication expects. The shapes are inspired
by each venue's official author kit (IEEE Conference template,
ACM ``acmart`` ``sigconf`` mode, APA 7th-edition manuscript style,
Nature standard article style); the output is a *starting point*
that captures the venue's structural skeleton — final camera-ready
formatting still requires the venue's own LaTeX / Word stylesheet,
but for early drafting and iterative authoring the kit removes the
boilerplate of building each template by hand.

Common conventions across the four factories:

- **Title block** — title is centred in the ``Title`` style. Author
  names + affiliations + emails follow in the venue's preferred
  layout (centred for IEEE / ACM / Nature, flush-left double-spaced
  for APA).
- **Abstract** — preceded by a bold ``Abstract`` label. APA renders
  the label as its own centred ``Heading 1`` per the manual; the
  others render the label inline at the start of the abstract
  paragraph (the IEEE / ACM / Nature house style).
- **Keywords / Index Terms** — IEEE renders ``Index Terms—`` in
  italics, ACM uses ``CCS Concepts`` plus ``Keywords``, APA uses
  ``Keywords:`` italicised, Nature omits keywords entirely (per the
  Nature style guide).
- **Body sections** — caller supplies an arbitrary ``sections`` list;
  each entry is rendered as a ``Heading 1`` plus body paragraphs.
- **References** — caller supplies an arbitrary ``references`` list;
  each entry is rendered into the venue's citation format
  (numbered ``[1] Authors, "Title," Venue, Year.`` for IEEE; numbered
  ``[1] Authors. Year. Title. Venue.`` for ACM; ``Authors (Year).
  Title. Venue.`` for APA; numbered superscript-style ``1. Authors.
  Title. Venue Year.`` for Nature).
- **Layout** — IEEE and Nature switch the document body to a
  two-column layout via :meth:`Section.set_columns` (with the title
  block sitting in a leading single-column continuous section so the
  banner stays full-width). APA keeps a single column and applies
  double line-spacing to every body paragraph. ACM stays
  single-column at draft time (the ``sigconf`` two-column rendering
  is handled by the ACM stylesheet at typesetting).
- **Style fallback to ``Normal``.** Helpers prefer Word's built-in
  ``Title`` / ``Heading 1`` / ``Heading 2`` styles; when the loaded
  template lacks one, fall back to ``Normal`` rather than raise. The
  spirit of a kit is "works out of the box".
- **No XML reach-down** — the kit composes only public python-docx
  API (``Document.add_paragraph``, ``Document.add_heading``,
  ``Document.add_section``, ``Section.set_columns``, …).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Optional, Sequence, Union

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.text.paragraph import Paragraph


# -- Type aliases ---------------------------------------------------------

AuthorDict = Mapping[str, str]
SectionDict = Mapping[str, Union[str, Sequence[str]]]
ReferenceDict = Mapping[str, Union[str, int]]


# -- Shared helpers -------------------------------------------------------


def _add_styled(
    document: DocumentCls,
    text: str,
    style: str,
    fallback: str = "Normal",
) -> Paragraph:
    """Append ``text`` in ``style`` (falling back to ``fallback``)."""
    try:
        document.styles[style]
    except KeyError:
        style = fallback
    return document.add_paragraph(text, style=style)


def _add_title(
    document: DocumentCls,
    title: str,
    align: int = WD_ALIGN_PARAGRAPH.CENTER,
) -> Paragraph:
    """Append a centred title in the ``Title`` style (or fallback)."""
    para = _add_styled(document, title, "Title")
    para.alignment = align
    return para


def _add_centered(document: DocumentCls, text: str) -> Paragraph:
    """Append a single centred paragraph in ``Normal`` style."""
    para = document.add_paragraph(text)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_section(
    document: DocumentCls,
    heading: str,
    body: Union[str, Sequence[str]],
    level: int = 1,
) -> List[Paragraph]:
    """Append a ``Heading N`` plus body paragraphs and return them."""
    paragraphs: List[Paragraph] = [document.add_heading(heading, level=level)]
    chunks: Sequence[str] = [body] if isinstance(body, str) else list(body)
    for chunk in chunks:
        if not chunk:
            continue
        paragraphs.append(document.add_paragraph(chunk))
    return paragraphs


def _validate_authors(authors: Optional[Sequence[AuthorDict]]) -> None:
    """Raise ValueError when an author entry is malformed."""
    if not authors:
        return
    for index, author in enumerate(authors):
        if not isinstance(author, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "authors[%d] must be a mapping with a 'name' key" % index
            )
        name = author.get("name")
        if not name or not str(name).strip():
            raise ValueError(
                "authors[%d] is missing a non-empty 'name'" % index
            )


def _validate_sections(
    sections: Optional[Sequence[SectionDict]],
) -> None:
    """Raise ValueError when a section entry is malformed."""
    if not sections:
        return
    for index, section in enumerate(sections):
        if not isinstance(section, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "sections[%d] must be a mapping with a 'heading' key"
                % index
            )
        heading = section.get("heading")
        if not heading or not str(heading).strip():
            raise ValueError(
                "sections[%d] is missing a non-empty 'heading'" % index
            )


def _add_body_sections(
    document: DocumentCls,
    sections: Optional[Sequence[SectionDict]],
) -> None:
    """Render every section entry as a ``Heading 1`` plus body."""
    if not sections:
        return
    for section in sections:
        heading = str(section.get("heading", "")).strip()
        body = section.get("body", "")
        if body is None:
            body = ""
        # body is Union[str, Sequence[str]]; the helper handles either
        _add_section(document, heading, body)  # type: ignore[arg-type]


def _format_authors_list(authors: Sequence[AuthorDict]) -> str:
    """Return ``"Alice, Bob and Carol"`` from an authors list."""
    names = [str(a.get("name", "")).strip() for a in authors]
    names = [n for n in names if n]
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    if len(names) == 2:
        return f"{names[0]} and {names[1]}"
    return ", ".join(names[:-1]) + f", and {names[-1]}"


def _enable_two_columns(
    document: DocumentCls,
    space_pt: int = 12,
) -> None:
    """Push the body of ``document`` into a two-column section layout.

    The default first section becomes the (single-column) title-banner
    section; a fresh continuous section break is appended that switches
    every subsequent paragraph to two columns. Callers should call this
    after rendering the title / authors / abstract block — that block
    sits in the leading section so the banner spans the full text width.
    """
    body_section = document.add_section(WD_SECTION.CONTINUOUS)
    body_section.set_columns(count=2, space=Pt(space_pt))


# -- IEEE -----------------------------------------------------------------


def ieee_paper(
    title: str,
    authors: Optional[Sequence[AuthorDict]] = None,
    abstract: Optional[str] = None,
    keywords: Optional[Sequence[str]] = None,
    sections: Optional[Sequence[SectionDict]] = None,
    references: Optional[Sequence[ReferenceDict]] = None,
) -> DocumentCls:
    """Build an IEEE-style scientific paper and return the |Document|.

    Generates a paper in the conventional IEEE Conference / Transactions
    layout:

    * Title centred in the leading single-column banner.
    * Author block: per-author *name* on its own line, *affiliation*
      below in italics, *email* below in monospace — all centred.
    * ``Abstract—`` lead-in (an em dash, IEEE house style) before the
      abstract text in a single bold-leading paragraph.
    * ``Index Terms—`` lead-in (italic) before the keywords list,
      keywords joined by commas.
    * Body switched to two columns via a continuous section break so
      every section / reference renders in IEEE's compact two-column
      shape.
    * Body sections rendered as ``Heading 1`` + body paragraphs.
    * References rendered as a numbered list under the ``References``
      heading in IEEE citation format::

         [N] Authors, "Title," Venue, Year.

    Parameters
    ----------
    title
        Paper title. Required — rendered into the title banner.
    authors
        Sequence of author dicts. Each dict must have a ``name`` key;
        ``affiliation`` and ``email`` are optional. Authors are rendered
        in the order given.
    abstract
        Abstract body text. ``None`` skips the abstract block.
    keywords
        Sequence of keyword strings. Joined by ``, `` and rendered after
        an italic ``Index Terms—`` lead-in. ``None`` skips the block.
    sections
        Sequence of section dicts. Each entry needs a ``heading`` key;
        the optional ``body`` key can be a string or sequence of strings.
    references
        Sequence of reference dicts. Each may have ``authors`` /
        ``title`` / ``venue`` / ``year`` keys; missing keys collapse to
        empty.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with :meth:`Document.save`.

    Raises
    ------
    ValueError
        When ``title`` is empty or when any author / section entry is
        malformed.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")
    _validate_authors(authors)
    _validate_sections(sections)

    document = Document()

    # -- Title banner (single-column section) --
    _add_title(document, title)

    if authors:
        for author in authors:
            name = str(author.get("name", "")).strip()
            affiliation = str(author.get("affiliation", "")).strip()
            email = str(author.get("email", "")).strip()
            name_para = _add_centered(document, name)
            for run in name_para.runs:
                run.bold = True
            if affiliation:
                aff_para = _add_centered(document, affiliation)
                for run in aff_para.runs:
                    run.italic = True
            if email:
                _add_centered(document, email)

    if abstract:
        para = document.add_paragraph()
        label = para.add_run("Abstract—")
        label.bold = True
        para.add_run(abstract)

    if keywords:
        para = document.add_paragraph()
        label = para.add_run("Index Terms—")
        label.italic = True
        para.add_run(", ".join(keywords))

    # -- Switch to two-column layout for the body --
    _enable_two_columns(document)

    _add_body_sections(document, sections)

    if references:
        document.add_heading("References", level=1)
        for index, ref in enumerate(references, start=1):
            authors_str = str(ref.get("authors", "")).strip()
            ref_title = str(ref.get("title", "")).strip()
            venue = str(ref.get("venue", "")).strip()
            year = ref.get("year", "")
            year_str = str(year).strip() if year != "" else ""
            pieces = [f"[{index}]"]
            if authors_str:
                pieces.append(f"{authors_str},")
            if ref_title:
                pieces.append(f"“{ref_title},”")
            if venue:
                pieces.append(f"{venue},")
            if year_str:
                pieces.append(f"{year_str}.")
            elif pieces[-1].endswith(","):
                pieces[-1] = pieces[-1][:-1] + "."
            document.add_paragraph(" ".join(pieces))

    return document


# -- ACM ------------------------------------------------------------------


def acm_paper(
    title: str,
    authors: Optional[Sequence[AuthorDict]] = None,
    abstract: Optional[str] = None,
    keywords: Optional[Sequence[str]] = None,
    ccs_concepts: Optional[Sequence[str]] = None,
    sections: Optional[Sequence[SectionDict]] = None,
    references: Optional[Sequence[ReferenceDict]] = None,
) -> DocumentCls:
    """Build an ACM ``sigconf``-style scientific paper.

    Generates a paper in the conventional ACM ``acmart`` ``sigconf``
    layout. The output stays single-column at draft time (the ACM Word /
    LaTeX template handles the camera-ready two-column rendering); the
    structural skeleton — title, authors, abstract, CCS Concepts,
    Keywords, body sections, references — is laid out in the order ACM
    expects and labelled with the section names ``acmart`` enforces.

    Parameters
    ----------
    title
        Paper title. Required.
    authors
        Sequence of author dicts (``name`` / ``affiliation`` / ``email``).
    abstract
        Abstract body text. Preceded by an ``Abstract`` ``Heading 1``.
    keywords
        Sequence of keyword strings. Rendered under a ``Keywords``
        heading.
    ccs_concepts
        Sequence of CCS-Concepts strings (e.g.
        ``"Computing methodologies → Parallel computing"``). Rendered
        under a ``CCS Concepts`` heading. ACM requires at least one
        CCS Concept for camera-ready submissions; the template emits
        a placeholder when none is supplied.
    sections
        Sequence of section dicts.
    references
        Sequence of reference dicts. Rendered as a numbered list in
        ACM citation format::

            [N] Authors. Year. Title. Venue.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``title`` is empty or any author / section entry is bad.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")
    _validate_authors(authors)
    _validate_sections(sections)

    document = Document()

    _add_title(document, title)

    if authors:
        for author in authors:
            name = str(author.get("name", "")).strip()
            affiliation = str(author.get("affiliation", "")).strip()
            email = str(author.get("email", "")).strip()
            line = name
            if affiliation:
                line = f"{line}, {affiliation}" if line else affiliation
            name_para = _add_centered(document, line)
            for run in name_para.runs:
                run.bold = True
            if email:
                _add_centered(document, email)

    if abstract:
        document.add_heading("Abstract", level=1)
        document.add_paragraph(abstract)

    document.add_heading("CCS Concepts", level=1)
    if ccs_concepts:
        for concept in ccs_concepts:
            if concept:
                document.add_paragraph(f"• {concept}")
    else:
        document.add_paragraph(
            "[Insert at least one CCS Concept "
            "(see https://dl.acm.org/ccs).]"
        )

    if keywords:
        document.add_heading("Keywords", level=1)
        document.add_paragraph(", ".join(keywords))

    _add_body_sections(document, sections)

    if references:
        document.add_heading("References", level=1)
        for index, ref in enumerate(references, start=1):
            authors_str = str(ref.get("authors", "")).strip()
            ref_title = str(ref.get("title", "")).strip()
            venue = str(ref.get("venue", "")).strip()
            year = ref.get("year", "")
            year_str = str(year).strip() if year != "" else ""
            pieces: List[str] = [f"[{index}]"]
            if authors_str:
                pieces.append(f"{authors_str}.")
            if year_str:
                pieces.append(f"{year_str}.")
            if ref_title:
                pieces.append(f"{ref_title}.")
            if venue:
                pieces.append(f"{venue}.")
            document.add_paragraph(" ".join(pieces))

    return document


# -- APA ------------------------------------------------------------------


def apa_paper(
    title: str,
    authors: Optional[Sequence[AuthorDict]] = None,
    abstract: Optional[str] = None,
    keywords: Optional[Sequence[str]] = None,
    running_head: Optional[str] = None,
    sections: Optional[Sequence[SectionDict]] = None,
    references: Optional[Sequence[ReferenceDict]] = None,
) -> DocumentCls:
    """Build an APA 7th-edition manuscript-style paper.

    Generates a paper in APA 7th-edition manuscript layout: single
    column with double line-spacing applied to every body paragraph,
    title and author block centred at the top, ``Abstract`` heading
    centred per the manual, ``Keywords:`` italic lead-in. References
    are rendered in APA author-date format::

        Authors (Year). Title. Venue.

    Parameters
    ----------
    title
        Paper title. Required.
    authors
        Sequence of author dicts. ``affiliation`` is rendered under
        the name on its own line per the APA title-page convention.
    abstract
        Abstract body. Rendered under a centred ``Abstract`` heading.
    keywords
        Sequence of keyword strings. Rendered after an italic
        ``Keywords:`` lead-in.
    running_head
        Optional running-head string. Rendered as the first paragraph
        of the document, prefixed ``Running head: `` per APA 6th
        edition (APA 7th drops "Running head:" from the page header
        but the text still appears on the title page).
    sections
        Sequence of section dicts.
    references
        Sequence of reference dicts.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``title`` is empty or any entry is malformed.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")
    _validate_authors(authors)
    _validate_sections(sections)

    document = Document()

    if running_head:
        rh_para = document.add_paragraph()
        label = rh_para.add_run("Running head: ")
        label.bold = True
        rh_para.add_run(running_head.upper())

    title_para = _add_title(document, title)
    # Force double-spacing on the title too.
    title_para.paragraph_format.line_spacing = 2.0

    if authors:
        for author in authors:
            name = str(author.get("name", "")).strip()
            affiliation = str(author.get("affiliation", "")).strip()
            name_para = _add_centered(document, name)
            name_para.paragraph_format.line_spacing = 2.0
            if affiliation:
                aff_para = _add_centered(document, affiliation)
                aff_para.paragraph_format.line_spacing = 2.0

    if abstract:
        ab_heading = document.add_heading("Abstract", level=1)
        ab_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ab_para = document.add_paragraph(abstract)
        ab_para.paragraph_format.line_spacing = 2.0

    if keywords:
        kw_para = document.add_paragraph()
        label = kw_para.add_run("Keywords: ")
        label.italic = True
        kw_para.add_run(", ".join(keywords))
        kw_para.paragraph_format.line_spacing = 2.0

    if sections:
        for section in sections:
            heading = str(section.get("heading", "")).strip()
            body = section.get("body", "")
            if body is None:
                body = ""
            paragraphs = _add_section(document, heading, body)
            for para in paragraphs:
                para.paragraph_format.line_spacing = 2.0

    if references:
        ref_heading = document.add_heading("References", level=1)
        ref_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for ref in references:
            authors_str = str(ref.get("authors", "")).strip()
            ref_title = str(ref.get("title", "")).strip()
            venue = str(ref.get("venue", "")).strip()
            year = ref.get("year", "")
            year_str = str(year).strip() if year != "" else ""
            pieces: List[str] = []
            if authors_str:
                if year_str:
                    pieces.append(f"{authors_str} ({year_str}).")
                else:
                    pieces.append(f"{authors_str}.")
            elif year_str:
                pieces.append(f"({year_str}).")
            if ref_title:
                pieces.append(f"{ref_title}.")
            if venue:
                pieces.append(f"{venue}.")
            ref_para = document.add_paragraph(" ".join(pieces))
            ref_para.paragraph_format.line_spacing = 2.0

    return document


# -- Nature ---------------------------------------------------------------


def nature_paper(
    title: str,
    authors: Optional[Sequence[AuthorDict]] = None,
    abstract: Optional[str] = None,
    sections: Optional[Sequence[SectionDict]] = None,
    references: Optional[Sequence[ReferenceDict]] = None,
) -> DocumentCls:
    """Build a Nature-style scientific paper.

    Generates a paper in Nature's compact display style: title and
    author byline in a leading single-column banner, abstract paragraph
    in italics under the byline, body switched to two-column layout via
    a continuous section break for the article proper. Nature style
    omits a separate keywords block — keywords are inferred from the
    title + abstract by the Nature indexing pipeline. References are
    rendered as a numbered list under the ``References`` heading in
    Nature citation format::

        N. Authors. Title. Venue Year.

    Parameters
    ----------
    title
        Paper title. Required.
    authors
        Sequence of author dicts. Names are joined with commas into a
        single byline paragraph (the Nature article-page convention);
        affiliations are listed on a second line below the byline.
    abstract
        Abstract body text. Rendered in italics directly under the
        byline.
    sections
        Sequence of section dicts.
    references
        Sequence of reference dicts.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``title`` is empty or any entry is malformed.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")
    _validate_authors(authors)
    _validate_sections(sections)

    document = Document()

    _add_title(document, title)

    if authors:
        byline = _format_authors_list(authors)
        if byline:
            byline_para = _add_centered(document, byline)
            for run in byline_para.runs:
                run.bold = True
        affiliations: List[str] = []
        seen: set = set()
        for author in authors:
            affiliation = str(author.get("affiliation", "")).strip()
            if affiliation and affiliation not in seen:
                affiliations.append(affiliation)
                seen.add(affiliation)
        if affiliations:
            aff_para = _add_centered(document, "; ".join(affiliations))
            for run in aff_para.runs:
                run.italic = True

    if abstract:
        ab_para = document.add_paragraph(abstract)
        for run in ab_para.runs:
            run.italic = True

    # -- Switch to two-column layout for the article body --
    _enable_two_columns(document)

    _add_body_sections(document, sections)

    if references:
        document.add_heading("References", level=1)
        for index, ref in enumerate(references, start=1):
            authors_str = str(ref.get("authors", "")).strip()
            ref_title = str(ref.get("title", "")).strip()
            venue = str(ref.get("venue", "")).strip()
            year = ref.get("year", "")
            year_str = str(year).strip() if year != "" else ""
            pieces: List[str] = [f"{index}."]
            if authors_str:
                pieces.append(f"{authors_str}.")
            if ref_title:
                pieces.append(f"{ref_title}.")
            tail = " ".join(part for part in (venue, year_str) if part)
            if tail:
                pieces.append(f"{tail}.")
            document.add_paragraph(" ".join(pieces))

    return document


__all__ = [
    "ieee_paper",
    "acm_paper",
    "apa_paper",
    "nature_paper",
]
