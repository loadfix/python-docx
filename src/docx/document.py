# pyright: reportImportCycles=false
# pyright: reportPrivateUsage=false

"""|Document| and closely related objects."""

from __future__ import annotations

import datetime as dt
import os
import re
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator, Mapping, Sequence

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_SECTION
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.section import Section, Sections
from docx.shared import ElementProxy, Emu, Inches, Length, RGBColor
from docx.text.run import Run
from docx.watermark import Watermark

if TYPE_CHECKING:
    import docx.types as t

    from ooxml_comments import CommentIds, CommentsExtensible

    from docx.accessibility import HeadingIssue
    from docx.alt_chunk import AltChunk
    from docx.attachments import Attachment
    from docx.bibliography import Bibliography, Citation, Source
    from docx.bookmarks import Bookmark, Bookmarks
    from docx.chart import Chart, WD_CHART_TYPE
    from docx.comments import Comment, Comments
    from docx.content_controls import ContentControl, ContentControlType
    from docx.custom_properties import CustomProperties
    from docx.data_sources import DataSource
    from docx.extended_properties import ExtendedProperties
    from docx.custom_xml import CustomXmlPart
    from docx.drawing import Canvas
    from docx.embedded_objects import EmbeddedObject
    from docx.endnotes import Endnotes, EndnoteProperties
    from docx.equations import Equation
    from docx.fields import Field
    from docx.linked_content import LinkedTarget
    from docx.font_table import FontTable
    from docx.footnotes import FootnoteProperties, Footnotes
    from docx.form_fields import FormField
    from docx.glossary import Glossary
    from docx.ink import InkAnnotation
    from ooxml_math import MathExpr
    from docx.oxml.content_controls import CT_Sdt
    from docx.outline import Outline, OutlineNode
    from docx.oxml.document import CT_Body, CT_Document
    from docx.oxml.table import CT_Tbl
    from docx.parts.document import DocumentPart
    from docx.permissions import PermissionRange
    from docx.search import SearchMatch
    from docx.enum.text import WD_PROTECTION
    from docx.settings import DocumentProtection, Settings
    from docx.signatures import SignatureInfo
    from docx.smart_art import SmartArt
    from docx.readability import ReadabilityReport
    from docx.statistics import DocumentStatistics
    from docx.styles.style import ParagraphStyle, _TableStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.theme import Theme
    from docx.tracked_changes import _TrackedChangesCtx
    from docx.web_settings import WebSettings


class Document(ElementProxy):
    """WordprocessingML (WML) document.

    Not intended to be constructed directly. Use :func:`docx.Document` to open or create
    a document.
    """

    def __init__(self, element: CT_Document, part: DocumentPart):
        super().__init__(element)
        self._element = element
        self._part = part
        self.__body = None
        # -- Name of the style queued for the *next* ``add_paragraph`` call
        # -- that doesn't pass an explicit `style`. Set when the previous
        # -- paragraph was added with a style whose ``w:next`` points at
        # -- another style. See :meth:`add_paragraph`. upstream#888.
        self._pending_next_style: str | None = None
        # -- active `tracked_changes()` context, if any. Stored as a simple
        # -- LIFO stack so nested `with document.tracked_changes(...)` blocks
        # -- shadow outer ones.
        from docx.tracked_changes import _TrackedChangesCtx

        self._tracked_changes_stack: list[_TrackedChangesCtx] = []
        # -- expose this proxy on the part so deep children (Paragraph,
        # -- Run, etc.) can look up the active tracked-changes state via
        # -- `self.part._track_changes_doc_proxy`. Safe to set: the part
        # -- has exactly one Document proxy. The try/except tolerates
        # -- `spec_set` mocks used in unit tests. --
        try:
            setattr(part, "_track_changes_doc_proxy", self)
        except AttributeError:  # pragma: no cover -- test fixtures only
            pass

    def __enter__(self) -> Document:
        """Enter context-manager; returns `self`.

        Pairs with :meth:`__exit__` / :meth:`close` to provide lifecycle
        symmetry with other file-like resources. The document itself does
        not hold an OS file handle after opening — |docx| fully parses the
        underlying ``.docx`` zip on construction — so the context-manager
        protocol exists mainly to let callers use the familiar ``with``
        idiom (upstream#379)::

            with Document('example.docx') as document:
                document.add_paragraph('Added in context.')
                document.save('out.docx')

        .. versionadded:: 2026.05.0
        """
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit context-manager; delegates to :meth:`close`.

        .. versionadded:: 2026.05.0
        """
        self.close()

    def close(self) -> None:
        """Release resources associated with this document.

        python-docx reads the ``.docx`` package eagerly on construction and
        does not retain the source file handle, so :meth:`close` is a no-op
        today. It exists to give callers a symmetric lifecycle API — useful
        for code that treats a ``Document`` like any other closeable
        resource — and is safe to call multiple times.

        .. versionadded:: 2026.05.0
        """
        # -- drop any stale tracked-changes context state --
        self._tracked_changes_stack = []

    def add_comment(
        self,
        runs: Run | Sequence[Run],
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
        date: dt.datetime | None = None,
    ) -> Comment:
        """Add a comment to the document, anchored to the specified runs.

        `runs` can be a single `Run` object or a non-empty sequence of `Run` objects. Only the
        first and last run of a sequence are used, it's just more convenient to pass a whole
        sequence when that's what you have handy, like `paragraph.runs` for example. When `runs`
        contains a single `Run` object, that run serves as both the first and last run.

        A comment can be anchored only on an even run boundary, meaning the text the comment
        "references" must be a non-zero integer number of consecutive runs. The runs need not be
        _contiguous_ per se, like the first can be in one paragraph and the last in the next
        paragraph, but all runs between the first and the last will be included in the reference.

        The comment reference range is delimited by placing a `w:commentRangeStart` element before
        the first run and a `w:commentRangeEnd` element after the last run. This is why only the
        first and last run are required and why a single run can serve as both first and last.
        Word works out which text to highlight in the UI based on these range markers.

        `text` allows the contents of a simple comment to be provided in the call, providing for
        the common case where a comment is a single phrase or sentence without special formatting
        such as bold or italics. More complex comments can be added using the returned `Comment`
        object in much the same way as a `Document` or (table) `Cell` object, using methods like
        `.add_paragraph()`, .add_run()`, etc.

        The `author` and `initials` parameters allow that metadata to be set for the comment.
        `author` is a required attribute on a comment and is the empty string by default.
        `initials` is optional on a comment and may be omitted by passing |None|, but Word adds an
        `initials` attribute by default and we follow that convention by using the empty string
        when no `initials` argument is provided.

        `date` is the timestamp recorded on the comment's ``w:date`` attribute. When
        omitted or |None|, ``datetime.now(timezone.utc)`` is used. Pass an explicit
        datetime for reproducible-save scenarios where an implicit wall-clock
        timestamp would defeat byte-identical output.

        .. versionchanged:: 2026.05.5
           Added the ``date`` parameter — previously only the lower-level
           ``Document.comments.add_comment(date=...)`` accepted it, so
           reproducible fixtures had to poke ``comment._element.date`` by hand.
        """
        # -- normalize `runs` to a sequence of runs --
        runs = [runs] if isinstance(runs, Run) else runs
        first_run = runs[0]
        last_run = runs[-1]

        # -- Note that comments can only appear in the document part --
        comment = self.comments.add_comment(
            text=text, author=author, initials=initials, date=date
        )

        # -- let the first run orchestrate placement of the comment range start and end --
        first_run.mark_comment_range(last_run, comment.comment_id)

        return comment

    def add_citation(
        self,
        tag: str,
        title: "str | None" = None,
        author: "str | None" = None,
        year: "str | int | None" = None,
        source_type: str = "Book",
        **extra: str,
    ) -> "Source":
        """Append a bibliographic source to the document's bibliography.

        Creates the backing ``/customXml/item{N}.xml`` part (and its
        ``itemProps{N}.xml`` sibling) on first use. `tag` is the unique
        citation key that :meth:`Paragraph.add_citation_reference` looks up;
        re-using an existing tag raises :class:`ValueError`.

        `source_type` defaults to ``"Book"``; pass e.g. ``"JournalArticle"``
        or ``"InternetSite"`` for richer fixtures. Any ``**extra`` kwargs
        become text-only ``<b:{Capitalized}>`` children of the source
        element (e.g. ``city="London"`` → ``<b:City>London</b:City>``).

        Returns the newly-added :class:`Source`.

        .. versionadded:: 2026.05.7
        """
        return self.bibliography.add_source(
            tag,
            title=title,
            author=author,
            year=year,
            source_type=source_type,
            **extra,
        )

    def add_bookmark(self, runs: Run | Sequence[Run], name: str) -> Bookmark:
        """Add a bookmark spanning `runs`, and return the |Bookmark| proxy.

        `runs` may be a single |Run| or a non-empty sequence of |Run| objects.
        Only the first and last run of a sequence are used — just as with
        :meth:`add_comment` — so the caller can pass a whole
        ``paragraph.runs`` or a selection that spans paragraphs.

        A ``w:bookmarkStart`` element is inserted immediately before the
        first run and a matching ``w:bookmarkEnd`` is inserted immediately
        after the last run. A fresh ``@w:id`` is allocated from the body's
        existing bookmark ids; `name` must be unique within the document
        (caller enforced — Word accepts duplicates silently but treats
        them as ambiguous cross-reference targets).

        This is the symmetric counterpart to :meth:`Paragraph.add_bookmark`
        — use that helper for a single-paragraph bookmark and this one when
        the range must span multiple paragraphs without dropping into the
        oxml layer.

        .. versionadded:: 2026.05.0
        """
        from docx.bookmarks import Bookmark
        from docx.text.paragraph import Paragraph

        runs = [runs] if isinstance(runs, Run) else list(runs)
        if not runs:
            raise ValueError("runs must be a non-empty sequence of Run objects")
        first_run = runs[0]
        last_run = runs[-1]

        body = self._element.body
        bookmark_id = Paragraph._next_bookmark_id(body)

        first_run._r.insert_bookmark_start_before(bookmark_id, name)
        last_run._r.insert_bookmark_end_after(bookmark_id)

        bookmarkStart = body.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")[0]
        return Bookmark(bookmarkStart, body)

    def add_caption(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Return a new caption |Paragraph| appended to the end of the document body.

        A Word caption is a paragraph styled with the ``Caption`` style that
        auto-numbers itself via a ``SEQ`` field. The resulting paragraph has
        the shape ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field; Word updates the number when the
        document is opened or fields are refreshed.

        `label` groups captions into a named sequence (e.g. ``"Figure"`` or
        ``"Table"``). `style` selects the paragraph style applied to the
        caption paragraph and defaults to the built-in ``"Caption"`` style.

        .. versionadded:: 2026.05.0
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.add_paragraph()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def add_heading(self, text: str = "", level: int = 1):
        """Return a heading paragraph newly added to the end of the document.

        The heading paragraph will contain `text` and have its paragraph style
        determined by `level`. If `level` is 0, the style is set to `Title`. If `level`
        is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
        {level}`. Raises |ValueError| if `level` is outside the range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = "Title" if level == 0 else "Heading %d" % level
        return self.add_paragraph(text, style)

    # -- Fluent chainable shortcuts (issue #77) -----------------------------
    # Sugar over :meth:`add_heading` / :meth:`add_paragraph`. Each
    # returns the freshly-appended Paragraph so the caller can keep
    # chaining (e.g. ``doc.h1("Title").bold().align("center")``).

    def h1(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 1`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=1``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=1)

    def h2(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 2`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=2``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=2)

    def h3(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 3`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=3``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=3)

    def h4(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 4`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=4``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=4)

    def h5(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 5`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=5``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=5)

    def h6(self, text: str = "") -> "Paragraph":
        """Append an ``Heading 6`` paragraph and return it.

        Thin wrapper over :meth:`add_heading` with ``level=6``.

        .. versionadded:: 2026.05.12
        """
        return self.add_heading(text, level=6)

    def p(self, text: str = "") -> "Paragraph":
        """Append a body paragraph and return it.

        Thin wrapper over :meth:`add_paragraph` (no style applied) so
        the fluent chain reads naturally::

            doc.h1("Q1 Review").p("Revenue grew 8.7% YoY").bold().align("center")

        .. versionadded:: 2026.05.12
        """
        return self.add_paragraph(text)

    def add_page_break(self) -> Paragraph:
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_page_break()
        return paragraph

    def add_paragraph(
        self,
        text: str = "",
        style: str | ParagraphStyle | None = None,
        track_author: str | None = None,
        bind_to: object | None = None,
    ) -> Paragraph:
        """Return paragraph newly added to the end of the document.

        The paragraph is populated with `text` and having paragraph style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break.

        If `track_author` is supplied, the inserted run is wrapped in a
        `w:ins` element with ``@w:author`` set to that string (closes
        upstream#1025). When the document has an active
        :meth:`tracked_changes` context, `track_author` defaults to the
        author of the innermost context; pass ``track_author=""`` explicitly
        if you need to opt out of the active context for one call.

        If `bind_to` is supplied, ``text`` is treated as a smart-placeholder
        template. Tokens such as ``{customer.name}``, ``{date:short}`` or
        ``{property:Title}`` are resolved against the bound record on every
        :meth:`save` (closes #68). The original token-source string is
        preserved in a fork-scoped ``<lfxbind:src>`` child element so that
        ``load -> bind -> save`` cycles re-resolve cleanly instead of
        carrying the previously-resolved literal forward. ``bind_to`` also
        sets the document-level bound record (overriding any prior call to
        :meth:`bind`), so subsequent ``add_paragraph(text)`` calls without
        an explicit ``bind_to`` still resolve against the same record.

        If the previously-added paragraph had a style whose ``w:next`` pointed
        at another style, and `style` is |None| on this call, that "next"
        style is applied automatically — mirroring the behaviour Word exhibits
        when the user presses Enter. An explicit `style` argument (including
        an explicit ``style=None`` passed positionally) always takes
        precedence. Closes upstream#888.

        .. versionadded:: 2026.05.0
           Added ``track_author`` keyword argument.
        .. versionadded:: 2026.05.13
           Added ``bind_to`` keyword argument for smart-placeholder fields (#68).
        """
        effective_style = style
        if effective_style is None and self._pending_next_style is not None:
            effective_style = self._pending_next_style
        # -- reset pending before recomputing to avoid runaway chains --
        self._pending_next_style = None

        if track_author is None:
            paragraph = self._body.add_paragraph(text, effective_style)
        else:
            paragraph = self._body.add_paragraph(
                text, effective_style, track_author=track_author
            )

        # -- bind-token wiring (#68): record the supplied record, stamp
        # -- the source marker on the freshly-created run, and let the
        # -- save-time resolver render the live values. --
        if bind_to is not None:
            from docx.bind_tokens import has_token, reseat_token_source, set_bound_record

            set_bound_record(self, bind_to)
            if text and has_token(text):
                # -- locate the freshly-added run carrying ``text`` and
                # -- stamp it. ``add_paragraph`` produces at most one run
                # -- when text is non-empty.
                runs = paragraph._p.xpath("./w:r")
                for r in runs:
                    reseat_token_source(r, text)

        # -- queue the `w:next` style, if any, for the subsequent call --
        if effective_style is not None:
            next_style_id = self._resolve_next_style_id(effective_style)
            if next_style_id is not None:
                self._pending_next_style = next_style_id

        return paragraph

    def bind(
        self,
        record: object | None = None,
        iteration: int | None = None,
    ) -> "Document":
        """Bind ``record`` to this document for smart-placeholder resolution.

        Every ``{token}`` written into a paragraph via
        :meth:`add_paragraph(text, bind_to=...)` is preserved in
        the OOXML as a ``<lfxbind:src>`` source marker. On the next
        :meth:`save`, those markers are walked and the displayed text
        re-resolved against ``record`` — so a single saved document
        can be re-bound to a different record by:

        .. code-block:: python

            doc = Document("template.docx")
            doc.bind(record=customer)
            doc.save("out.docx")

        Returns ``self`` for method chaining.

        ``iteration`` (default |None|) is the value made available to
        the ``{i}`` token, intended for callers driving a mail-merge
        loop where ``i`` reports the current row index.

        .. versionadded:: 2026.05.13
        """
        from docx.bind_tokens import set_bound_record

        set_bound_record(self, record, iteration=iteration)
        return self

    def _resolve_next_style_id(self, style: str | ParagraphStyle) -> str | None:
        """Return the ``w:styleId`` of `style`'s ``w:next`` style, if any.

        `style` may be a style name, a style id, or a |ParagraphStyle| proxy.
        Returns |None| when the style is not found, has no ``w:next``
        element, or ``w:next`` points at an undefined style id.
        """
        from docx.styles.style import BaseStyle

        if isinstance(style, BaseStyle):
            style_elm = style._element
        else:
            # -- try name first, fall back to style_id lookup --
            styles = self._part.styles
            try:
                resolved = styles[style]
            except KeyError:
                resolved = None
            if resolved is None:
                return None
            style_elm = resolved._element
        return style_elm.next_val

    def tag_revisions(self, rsid: "str | None" = None) -> str:
        """Stamp every paragraph, run, and section in the body with ``rsid``.

        Mirrors what Microsoft Word does automatically when it saves a
        document: tags every element that changed in the current editing
        session with a session-wide revision-save ID so later diff/merge
        tools can correlate edits back to their originating session.

        When ``rsid`` is |None| a fresh random 8-character-hex token is
        minted via :meth:`RsidList.new_session` and used. Otherwise the
        caller-supplied token is used verbatim and added to the
        ``w:rsids`` table if not already present. The chosen token is
        returned so the caller can record it.

        Attributes written:

        - ``w:p`` — ``w:rsidR`` (run-properties insertion rsid) and
          ``w:rsidRDefault`` (default rsid for newly inserted runs).
        - ``w:pPr`` — ``w:rsidP`` (paragraph-mark insertion rsid) and,
          when a ``w:rPr`` child is present, ``w:rsidRPr`` on it.
        - ``w:r`` — ``w:rsidR`` and, when a ``w:rPr`` child is present,
          ``w:rsidRPr`` on it.
        - ``w:sectPr`` — ``w:rsidR``, ``w:rsidSect`` and, when a
          ``w:rPr`` child is present, ``w:rsidRPr`` on it.

        Existing rsid attributes are overwritten so every tagged element
        reflects the same editing session. Revision-save IDs are
        cosmetic metadata (Word does not let users disable them) and
        this method does not interact with the tracked-changes machinery
        (``w:ins`` / ``w:del``).

        Returns the rsid token that was stamped on the document.

        .. versionadded:: 2026.05.12
        """
        from docx.oxml.ns import qn

        if rsid is None:
            rsid = self.settings.rsids.new_session()
        else:
            self.settings.rsids.add(rsid)

        rsidR = qn("w:rsidR")
        rsidRDefault = qn("w:rsidRDefault")
        rsidP = qn("w:rsidP")
        rsidRPr = qn("w:rsidRPr")
        rsidSect = qn("w:rsidSect")
        w_pPr = qn("w:pPr")
        w_rPr = qn("w:rPr")
        w_sectPr = qn("w:sectPr")

        body = self._element.body
        for p in body.iter(qn("w:p")):
            p.set(rsidR, rsid)
            p.set(rsidRDefault, rsid)
            pPr = p.find(w_pPr)
            if pPr is not None:
                pPr.set(rsidP, rsid)
                pPr_rPr = pPr.find(w_rPr)
                if pPr_rPr is not None:
                    pPr_rPr.set(rsidRPr, rsid)
                sectPr = pPr.find(w_sectPr)
                if sectPr is not None:
                    sectPr.set(rsidR, rsid)
                    sectPr.set(rsidSect, rsid)
                    sectPr_rPr = sectPr.find(w_rPr)
                    if sectPr_rPr is not None:
                        sectPr_rPr.set(rsidRPr, rsid)
            for r in p.iter(qn("w:r")):
                r.set(rsidR, rsid)
                r_rPr = r.find(w_rPr)
                if r_rPr is not None:
                    r_rPr.set(rsidRPr, rsid)

        # Also the trailing body-level ``w:sectPr`` (sits after the last
        # ``w:p`` in the body, outside any paragraph) — every section in
        # the doc should carry the current session's rsid.
        for sectPr in body.findall(w_sectPr):
            sectPr.set(rsidR, rsid)
            sectPr.set(rsidSect, rsid)
            sectPr_rPr = sectPr.find(w_rPr)
            if sectPr_rPr is not None:
                sectPr_rPr.set(rsidRPr, rsid)

        return rsid

    def tracked_changes(
        self, author: str, date: "dt.datetime | None" = None
    ) -> "_TrackedChangesCtx":
        """Return a context manager that wraps new content in `w:ins` markers.

        Every call to :meth:`add_paragraph` / :meth:`BlockItemContainer.add_paragraph`
        or :meth:`Paragraph.add_run` made while the returned context is
        active wraps the new ``w:r`` element in a ``w:ins`` element whose
        ``@w:author`` is `author` and whose ``@w:date`` is `date` (defaulting
        to the current UTC time). Contexts can be nested; the innermost
        author/date wins.

        Closes upstream#1025::

            with document.tracked_changes(author="Reviewer"):
                document.add_paragraph("A new paragraph under review.")

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _TrackedChangesCtx

        return _TrackedChangesCtx(self, author, date)

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Add a block-level content control at the end of the document body.

        `type` selects the kind of content control (see :class:`ContentControlType`).
        `tag` and `title` map to `w:sdtPr/w:tag/@w:val` and `w:sdtPr/w:alias/@w:val`
        respectively.

        .. versionadded:: 2026.05.0
        """
        return self._body.add_content_control(type, tag=tag, title=title)

    def add_text_control(
        self,
        kind: "str | ContentControlType" = "rich-text",
        name: str | None = None,
        placeholder: str | None = None,
        value: str | None = None,
        locked: "bool | str | None" = None,
        bind_to: str | None = None,
        bind_source: str | None = None,
        items: "Sequence[str] | None" = None,
        title: str | None = None,
    ) -> ContentControl:
        """Append a block-level content control of `kind` to the document body.

        `kind` is one of ``"text"``, ``"rich-text"``, ``"dropdown"``,
        ``"combo"``, ``"date"``, ``"checkbox"``, ``"picture"``,
        ``"repeating-section"``, or a :class:`ContentControlType` member.
        See :func:`docx.content_controls.build_text_control` for the full
        argument contract.

        Block-level controls are commonly used for whole sections (executive
        summaries, signature blocks, etc.) where the user can edit a region
        of multi-paragraph rich content. Use :meth:`Paragraph.add_text_control`
        for inline controls.

        ``bind_source`` names a data source previously registered with
        :meth:`bind_data_source`; when supplied, the emitted ``<w:dataBinding>``
        is anchored to that source's store-item id and ``bind_to`` is treated
        as an XPath into the source's payload.

        Returns the typed |ContentControl| proxy (e.g. :class:`PlainTextControl`,
        :class:`DateControl`).

        .. versionadded:: 2026.05.13
        """
        from docx.content_controls import ContentControl, build_text_control

        sdt = build_text_control(
            kind,
            name=name,
            placeholder=placeholder,
            value=value,
            locked=locked,
            bind_to=bind_to,
            items=items,
            title=title,
            inline=False,
        )
        if bind_source is not None and bind_to is not None:
            self._anchor_sdt_binding_to_source(sdt, bind_source, bind_to)
        self._body._body._insert_sdt(sdt)  # pyright: ignore[reportPrivateUsage]
        return ContentControl.proxy_for(sdt)

    def bind_data_source(
        self,
        path: "str | bytes | os.PathLike[str] | IO[bytes]",
        name: str,
        schema: "str | bytes | os.PathLike[str] | IO[bytes] | None" = None,
    ) -> "DataSource":
        """Attach (or replace) a custom-XML data source under logical id ``name``.

        Loads ``path`` (a filesystem path, bytes, or open binary file-like)
        as the payload of a fresh ``/customXml/item{N}.xml`` part and pairs
        it with a sibling ``itemProps{N}.xml`` properties part carrying a
        ``{GUID}`` store-item id. Subsequent
        :meth:`Paragraph.add_text_control` /
        :meth:`Document.add_text_control` calls reference the source via
        ``bind_source=name``.

        Re-binding with the same ``name`` overwrites the existing payload
        in-place — the data part, props part, and store-item id are all
        preserved, so SDTs already wired to the source continue to resolve
        against the new payload on the next save. This is the *replace
        underlying part* contract from issue #80.

        When ``schema`` is supplied the payload is validated against the XSD
        via :func:`ooxml_customxml.validate` *before* the rewrite. A
        validation failure raises :class:`docx.data_sources.DataSourceValidationError`
        and leaves the prior payload intact.

        Returns the :class:`DataSource` describing the bound source.

        .. versionadded:: 2026.05.13
        """
        from docx.data_sources import bind_data_source as _bind

        return _bind(self._part, path, name, schema=schema)

    @property
    def data_sources(self) -> "list[DataSource]":
        """List of |DataSource| proxies for every bound custom-XML source.

        Empty when the document has not yet bound any sources via
        :meth:`bind_data_source`. Plain ``customXml`` parts loaded from a
        package (bibliography sources, Power BI datasets, etc.) without a
        logical-name marker are not surfaced here — use
        :attr:`custom_xml_parts` for the read-only inventory.

        .. versionadded:: 2026.05.13
        """
        from docx.data_sources import iter_bound_sources

        return iter_bound_sources(self._part)

    def _anchor_sdt_binding_to_source(
        self, sdt: "CT_Sdt", source_name: str, xpath: str
    ) -> None:
        """Wire a ``<w:dataBinding>`` on ``sdt`` to the named data source.

        Looks up the source's store-item id, rewrites the SDT's
        ``<w:dataBinding>`` to point at it (with the canonical
        ``xmlns:ns0`` prefix mapping for the payload's default namespace),
        and inlines the resolved value into the SDT's ``<w:sdtContent>``.

        Raises :class:`KeyError` when ``source_name`` has not been bound.
        """
        from docx.content_controls import ContentControl
        from docx.data_sources import _replace_sdt_content
        from docx.oxml.ns import nsmap

        sources = {src.name: src for src in self.data_sources}
        if source_name not in sources:
            raise KeyError(
                f"data source {source_name!r} has not been bound; "
                "call bind_data_source() first"
            )
        source = sources[source_name]

        proxy = ContentControl.proxy_for(sdt)
        ns_uri: str | None = None
        root = source.root_element
        if root is not None and isinstance(root.tag, str) and root.tag.startswith("{"):
            ns_uri = root.tag[1 : root.tag.index("}")]
        prefix_mappings = (
            f"xmlns:ns0='{ns_uri}'" if ns_uri else ""
        )
        proxy.set_data_binding(
            xpath,
            prefix_mappings=prefix_mappings,
            store_item_id=source.store_item_id,
        )
        # -- inline current resolved value (best-effort) --
        try:
            from ooxml_customxml import CustomXmlMapping, resolve_binding

            db = proxy.data_binding
            if db is not None and root is not None:
                mapping = CustomXmlMapping(db.element)
                value = resolve_binding(mapping, root)
                if value is not None:
                    _replace_sdt_content(sdt, value)
        except Exception:
            # -- best-effort; the save-time resolver will retry --
            _ = nsmap  # silence unused import in py3.9 pyflakes

    def add_repeating_section(
        self,
        name: str | None = None,
        section_title: str | None = None,
        schema: "dict[str, str] | Sequence[tuple[str, str]] | None" = None,
        locked: "bool | str | None" = None,
    ) -> "ContentControl":
        """Append a block-level repeating-section content control.

        `name` becomes the SDT's programmatic ``w:tag/@w:val`` (typically the
        repeating-section's identifier — e.g. ``"line_items"``).
        `section_title` populates the ``@w15:sectionTitle`` attribute Word
        shows in the *Insert New Item* drop-down.

        `schema` is a mapping of ``field_name -> kind`` describing the
        per-row fields the helper will stamp into each row when callers
        invoke :meth:`RepeatingSectionControl.add`. Kinds reuse the same
        strings :meth:`add_text_control` accepts (plus ``"number"``, which
        maps to plain text — Word has no dedicated number SDT).

        Returns a :class:`RepeatingSectionControl`.

        .. versionadded:: 2026.05.13
        """
        from docx.content_controls import (
            ContentControl,
            ContentControlType,
            RepeatingSectionControl,
            build_text_control,
        )

        sdt = build_text_control(
            ContentControlType.REPEATING_SECTION,
            name=name,
            locked=locked,
            inline=False,
        )
        self._body._body._insert_sdt(sdt)  # pyright: ignore[reportPrivateUsage]
        proxy = ContentControl.proxy_for(sdt)
        assert isinstance(proxy, RepeatingSectionControl)
        if section_title is not None:
            proxy.section_title = section_title
        if schema is not None:
            proxy.set_schema(schema)  # type: ignore[attr-defined]
        return proxy

    def add_chart(
        self,
        chart_type: WD_CHART_TYPE,
        categories: list[str],
        series_data: dict[str, list[float]],
        width: Length | None = None,
        height: Length | None = None,
    ) -> Chart:
        """Append a new chart to the end of the document and return it.

        `chart_type` selects the kind of chart (see :class:`docx.chart.WD_CHART_TYPE`).
        Only ``BAR``, ``BAR_STACKED``, ``COLUMN``, ``COLUMN_STACKED``, ``LINE``,
        and ``PIE`` are supported for creation.

        `categories` is a list of category labels (for example x-axis labels).
        `series_data` maps each series name to its list of numeric values; every
        value list must be the same length as `categories`.

        `width` and `height` are |Length| instances controlling the display size
        of the inline chart. When omitted a 6" x 3" default is used (similar to
        Word's default inline chart size).

        .. versionadded:: 2026.05.0
        """
        from docx.chart import Chart
        from docx.opc.constants import RELATIONSHIP_TYPE as _RT
        from docx.oxml.shape import CT_Inline
        from docx.parts.chart import ChartPart
        from docx.shared import Emu, Inches

        cx = width if width is not None else Inches(6)
        cy = height if height is not None else Inches(3)
        cx = Emu(int(cx))
        cy = Emu(int(cy))

        # -- create the chart part and relate it to the document part --
        package = self._part.package
        assert package is not None
        chart_part = ChartPart.new(package, chart_type, categories, series_data)
        rId = self._part.relate_to(chart_part, _RT.CHART)

        # -- build the wp:inline drawing pointing at the chart part --
        shape_id = self._part.next_id
        inline = CT_Inline.new_chart_inline(shape_id, rId, cx, cy)

        # -- append a new paragraph and drawing to the body --
        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        run._r.add_drawing(inline)

        return Chart(chart_part)

    def add_chart_inline(
        self,
        kind: str = "column",
        data: t.Any = None,
        *,
        x: str | None = None,
        y: str | Sequence[str] | None = None,
        title: str | None = None,
        subtitle: str | None = None,
        size: t.Any = None,
        show_values: bool = False,
        show_legend: bool | str = "auto",
        secondary_axis: Sequence[str] | None = None,
    ) -> Chart:
        """Append an inline chart with ergonomic data-shape input.

        Wraps :meth:`add_chart` with three input shapes (dict, list-of-dicts,
        ``pandas.DataFrame``) and the v1 chart-kind catalogue documented in
        :mod:`docx.chart_inline` — ``bar``, ``column``, ``line``, ``area``,
        ``pie``, ``donut``, ``scatter``, ``bubble``, ``combo``,
        ``stacked-bar``, ``stacked-column``, ``stacked-area``, ``sparkline``
        (plus ``grouped-bar`` / ``grouped-column`` aliases).

        ``data``
            One of: ``Mapping[str, float]`` (single-series, keys are
            categories), ``Sequence[Mapping[str, Any]]`` (list-of-dicts,
            ``x`` / ``y`` keys select columns), or a ``pandas.DataFrame``.
            ``pandas`` is **not** a hard dependency — DataFrame input is
            sniffed at runtime.

        ``x`` / ``y``
            Required for list-of-dicts and DataFrame input.  ``y`` may be a
            sequence of column names for a multi-series chart.

        ``title`` / ``subtitle``
            Rendered into ``c:title/c:tx/c:rich``; subtitle ships as a
            second paragraph at 12pt.

        ``size``
            Optional ``(width, height)`` pair (|Length| or float-inches).
            Defaults to 6" x 4".

        ``show_values``
            Reserved hook for per-point data labels.  Currently a no-op;
            wire to ``ChartFormat`` once the v2 helpers land.

        ``show_legend``
            ``"auto"`` (default) — legend on multi-series only.  Booleans
            force on / off.

        ``secondary_axis``
            Sequence of series names to plot against a secondary value-axis
            (right-hand side).  Combined with ``kind="combo"`` for the
            classic column + line dual-axis chart.

        Closes upstream#76.

        .. versionadded:: 2026.05.13
        """
        from docx.chart_inline import add_chart_inline as _impl

        if data is None:
            raise TypeError("add_chart_inline() missing required argument: 'data'")

        return _impl(
            self,
            kind=kind,
            data=data,
            x=x,
            y=y,
            title=title,
            subtitle=subtitle,
            size=size,
            show_values=show_values,
            show_legend=show_legend,
            secondary_axis=secondary_axis,
        )

    def add_dataframe(
        self,
        df: t.Any,
        *,
        style: str = "executive",
        alternating_rows: bool | None = None,
        header_color: "RGBColor | str | None" = None,
        header_text_color: "RGBColor | str | None" = None,
        autofit: bool = True,
        align: "Mapping[str, str] | None" = None,
        number_format: "Mapping[str, str] | None" = None,
        show_total_row: "bool | str | Mapping[str, str]" = False,
        table_style: str | None = None,
    ) -> "Table":
        """Append a ``pandas.DataFrame`` to this document as a styled Word table.

        Provides four built-in presets (``executive`` / ``minimal`` /
        ``boxed`` / ``striped``) plus per-column alignment, per-column
        number-format DSL, theme-aware header colours, and an optional
        total row.

        ``df``
            A :class:`pandas.DataFrame`. Pandas is **NOT** a hard
            dependency of python-docx — this method imports it lazily
            and raises :class:`ImportError` with an actionable message
            when it is missing. The DataFrame argument itself is
            sniffed via duck-typing (matching the pattern used by
            :meth:`add_chart_inline`).

        ``style``
            One of ``"executive"`` (bold header bar in theme primary,
            alternating row tint, total row at bottom), ``"minimal"``
            (header underline only, no fills, monospace numbers),
            ``"boxed"`` (full grid borders, light header tint), or
            ``"striped"`` (zebra rows, no borders).

        ``alternating_rows``
            Force-on / force-off alternating row tints. |None|
            (default) defers to the preset.

        ``header_color`` / ``header_text_color``
            Override the header row's fill and text colour. Accept an
            :class:`docx.shared.RGBColor` or a ``"RRGGBB"`` hex string.
            Pull straight from the active theme via
            ``document.theme.colors.accent_1`` /
            ``document.theme.colors.light_1``.

        ``autofit``
            Forwarded to :attr:`Table.autofit`.

        ``align``
            ``{column_name: "left"|"right"|"center"|"justify"}``.
            Numeric columns default to right-aligned, everything else
            left-aligned.

        ``number_format``
            ``{column_name: format_spec}``. Accepts the standard Python
            mini-language for numeric values (``"$,.1f"``, ``"0.0%"``,
            ``",d"`` …) plus a small DSL for date columns
            (``"MMM YYYY"``, ``"YYYY-MM-DD"`` …) modelled on the
            sibling ``python-xlsx`` number-format helpers.

        ``show_total_row``
            |False| (default) for no total row; |True| / ``"sum"`` to
            sum every numeric column; ``"mean"``, ``"count"``, or
            ``"none"`` for the corresponding aggregation; or a mapping
            ``{column_name: aggregator}`` for explicit per-column
            overrides.

        ``table_style``
            Optional Word table style name (e.g. ``"Light Grid"``)
            applied before the preset's inline formatting.

        Closes #40.

        .. versionadded:: 2026.05.13
        """
        from docx.dataframe import add_dataframe as _impl

        return _impl(
            self,
            df,
            style=style,
            alternating_rows=alternating_rows,
            header_color=header_color,
            header_text_color=header_text_color,
            autofit=autofit,
            align=align,
            number_format=number_format,
            show_total_row=show_total_row,
            table_style=table_style,
        )

    def add_shape(
        self,
        shape_type,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append an inline DrawingML preset shape in its own paragraph.

        `shape_type` is a :class:`docx.enum.shape.WD_SHAPE` member (e.g.
        ``WD_SHAPE.ROUNDED_RECTANGLE``). `width` and `height` are |Length|
        values; they default to 2" x 1" when omitted. When `text` is provided a
        minimal text-frame is attached so Word renders the string inside the
        shape. The shape is emitted as a ``wps:wsp`` with ``a:prstGeom`` inside
        a ``w:drawing/wp:inline``.

        Returns a :class:`docx.drawing.WordprocessingShape` proxy for the new
        shape. Closes upstream#1112 and upstream#517.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        return paragraph.add_shape(shape_type, width, height, text=text)

    def add_canvas(
        self,
        width: Length | None = None,
        height: Length | None = None,
    ) -> Canvas:
        """Append a DrawingML canvas (``wpc:wpc``) in its own paragraph.

        A canvas groups one or more shapes or pictures under a single
        ``w:drawing/wp:inline/a:graphic/a:graphicData`` whose URI is the
        WordprocessingCanvas namespace. `width` and `height` default to
        6" x 3".

        Returns a :class:`docx.drawing.Canvas` proxy. Callers can build up the
        canvas contents via :meth:`Canvas.add_shape`. Closes upstream#411.

        .. versionadded:: 2026.05.0
        """
        from docx.drawing import Canvas
        from docx.oxml.drawing import new_inline_canvas_drawing

        cx = int(width) if width is not None else int(Inches(6))
        cy = int(height) if height is not None else int(Inches(3))

        shape_id = self._part.next_id
        name = "Canvas %d" % shape_id
        drawing = new_inline_canvas_drawing(cx, cy, shape_id, name)

        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        run._r.append(drawing)

        wpc = drawing.xpath(".//wp:inline/a:graphic/a:graphicData/wpc:wpc")[0]
        return Canvas(wpc, paragraph)

    def add_text_box(
        self,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append an inline DrawingML text box in its own paragraph.

        A text box is a ``wps:wsp`` whose preset geometry is a simple rectangle
        carrying a ``wps:txbx/w:txbxContent`` text frame. `width` and `height`
        default to 3" x 1.5". When `text` is provided, the text box is
        initialised with a single paragraph containing that string; otherwise
        callers can append paragraphs via
        :meth:`~docx.drawing.WordprocessingShape.add_paragraph`.

        Returns a :class:`docx.drawing.WordprocessingShape` proxy. Closes
        upstream#524.

        .. versionadded:: 2026.05.0
        """
        from docx.enum.shape import WD_SHAPE

        cx = width if width is not None else Inches(3)
        cy = height if height is not None else Inches(1.5)

        paragraph = self.add_paragraph()
        shape = paragraph.add_shape(
            WD_SHAPE.RECTANGLE, cx, cy, text=text if text is not None else ""
        )
        return shape

    def add_picture(
        self,
        image_path_or_stream: "str | os.PathLike[str] | IO[bytes] | None" = None,
        width: int | Length | None = None,
        height: int | Length | None = None,
        link: bool = False,
        save_with_document: bool = True,
        url: str | None = None,
    ):
        """Return new picture shape added in its own paragraph at end of the document.

        The picture contains the image at `image_path_or_stream`, scaled based on
        `width` and `height`. If neither width nor height is specified, the picture
        appears at its native size. If only one is specified, it is used to compute a
        scaling factor that is then applied to the unspecified dimension, preserving the
        aspect ratio of the image. The native size of the picture is calculated using
        the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
        if no value is specified, as is often the case.

        `image_path_or_stream` may be a ``str`` path, an ``os.PathLike`` (e.g.
        :class:`pathlib.Path`), or a binary file-like object.

        When `link` is |True| and `save_with_document` is |False|, a linked
        (external) picture is inserted: the `a:blip` uses ``r:link`` and no
        image part is added to the package. `url` may be supplied to link
        a remote image rather than a local path. See
        :meth:`docx.text.run.Run.add_picture` for details.

        .. versionchanged:: 2026.05.0
           Accepts :class:`os.PathLike` path arguments.

        .. versionadded:: 2026.05.0
            ``link``, ``save_with_document``, and ``url`` parameters.
        """
        if isinstance(image_path_or_stream, os.PathLike):
            image_path_or_stream = os.fspath(image_path_or_stream)
        run = self.add_paragraph().add_run()
        return run.add_picture(
            image_path_or_stream,
            width,
            height,
            link=link,
            save_with_document=save_with_document,
            url=url,
        )

    def add_section(self, start_type: WD_SECTION = WD_SECTION.NEW_PAGE):
        """Return a |Section| object newly added at the end of the document.

        The optional `start_type` argument must be a member of the :ref:`WdSectionStart`
        enumeration, and defaults to ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

    def section(
        self,
        *,
        orientation=None,
        margins=None,
        page_size=None,
        page_numbering=None,
        header=None,
        footer=None,
        columns=None,
        line_numbering=None,
    ):
        """Return a context manager that scopes content to a fresh section.

        On enter, a continuous section break is appended and the
        keyword-supplied page setup applied to the new section; on exit,
        a second break is appended that reverts to the prior layout::

            with doc.section(orientation='landscape', margins='narrow'):
                doc.add_paragraph("Wide table follows")
                doc.add_table(rows=2, cols=12)

        ``orientation`` accepts ``"portrait"`` / ``"landscape"`` or a
        :class:`WD_ORIENTATION` member. ``margins`` accepts presets
        (``"narrow"`` / ``"normal"`` / ``"moderate"`` / ``"wide"``), a
        4-tuple ``(top, right, bottom, left)``, or a dict.
        ``page_size`` accepts presets (``"letter"`` / ``"legal"`` /
        ``"a4"`` / ``"a3"`` / ``"tabloid"``), a 2-tuple, or a dict.
        ``page_numbering`` is a dict with ``style`` / ``start`` /
        ``restart`` keys. ``columns`` is an int or dict forwarded to
        :meth:`Section.set_columns`. ``line_numbering`` is a bool or
        dict forwarded to :meth:`Section.set_line_numbering`.
        ``header`` / ``footer`` are plain text strings.

        Raises :class:`docx.exceptions.NestedSectionError` when entered
        inside another active section context — OOXML sections cannot
        nest.

        .. versionadded:: 2026.05.13
        """
        from docx.section_context import open_section

        return open_section(
            self,
            orientation=orientation,
            margins=margins,
            page_size=page_size,
            page_numbering=page_numbering,
            header=header,
            footer=footer,
            columns=columns,
            line_numbering=line_numbering,
        )

    def audit_styles(self):
        """Audit the document's styles and return a :class:`StyleAudit`.

        Issues surfaced: ``duplicate-styles``, ``direct-formatting``,
        ``mixed-fonts``, ``unstyled-paragraph``, ``heading-without-style``,
        ``orphan-style``. Use :meth:`StyleAudit.consolidate_styles` to
        rewrite paragraphs and drop redundant style definitions::

            audit = doc.audit_styles()
            audit.consolidate_styles("Heading 1", drop=["H1", "Heading1"])

        .. versionadded:: 2026.05.13
        """
        from docx.audit import audit_styles

        return audit_styles(self)

    def lint(self, rules=None):
        """Run lint rules and return a list of :class:`LintFinding`.

        Default rules: ``heading-skip``, ``heading-multiple-h1``,
        ``heading-no-h1``, ``heading-direct-formatting``,
        ``heading-empty``, ``heading-too-long``. ``rules`` may be |None|
        (defaults), a list of rule-id strings, or a list of callables.

        Accessibility rules (issue #15) are off by default but available
        by id: ``image-no-alt-text``, ``table-no-caption``,
        ``no-language-tag``, ``low-contrast``, ``no-document-title``.
        Pass :data:`docx.lint.ACCESSIBILITY_RULES` to enable the full
        set, or :data:`docx.lint.ALL_RULES` for every shipped rule.

        .. versionadded:: 2026.05.13
        .. versionchanged:: 2026.05.dev0
           Added accessibility rules.
        """
        from docx.lint import lint_document

        return lint_document(self, rules)

    def add_table_of_contents(
        self, levels: tuple[int, int] = (1, 3)
    ) -> Paragraph:
        """Append a table-of-contents paragraph at the end of the document body.

        Creates a new paragraph containing a ``TOC`` complex field whose
        cached *result text* is a preview built from the document's existing
        heading paragraphs. ``levels`` is a ``(min_level, max_level)`` tuple
        (default ``(1, 3)``) that selects which ``"Heading N"`` paragraphs
        contribute to the preview; headings outside the range are skipped.

        The preview lists one heading per line in the form
        ``"{text}\\t{index}"`` — the tab-separated trailing integer is a
        1-based position in the filtered heading list, not a page number
        (python-docx has no layout engine). Word rebuilds the TOC the next
        time the document is opened or when fields are refreshed, so the
        cached preview is purely a convenience for raw-XML consumers and
        non-Word viewers.

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        from docx.toc import populate_toc_paragraph

        # -- snapshot the current headings before appending the new paragraph
        #    so the empty TOC paragraph doesn't self-include if its style
        #    ever matched the heading regex. --
        source_paragraphs = list(self.paragraphs)
        paragraph = self.add_paragraph()
        return populate_toc_paragraph(paragraph, source_paragraphs, levels)

    def add_table_copy(self, other_table: Table) -> Table:
        """Append a deep copy of `other_table` (possibly from another document) to this body.

        The entire ``w:tbl`` element is deep-copied, then scanned for
        cross-document references that must be rewired into this document's
        package:

        - Embedded images: every ``a:blip/@r:embed`` (plus SVG sibling references
          via ``asvg:svgBlip/@r:embed``) is resolved against `other_table`'s
          document part. The image part is copied into this document's package
          (de-duplicated by SHA-1) and the ``r:embed`` attribute is rewritten to
          the freshly-minted ``rId`` in this part's relationships.
        - Table-style reference: when ``w:tblStyle/@w:val`` names a style that
          does not yet exist in this document, the source style's ``w:style``
          element is deep-copied into this document's styles part. *Advanced
          style cascades (``w:basedOn`` / ``w:link`` / ``w:next`` chains),
          numbering references, and conditional table-style formatting are
          not recursively imported; those remain TODO.*

        Returns the |Table| wrapping the inserted ``w:tbl``. When `other_table`
        belongs to this same document the copy is inserted without any rewiring
        (rIds are already valid in the current part).

        Closes upstream#612, #270.

        .. versionadded:: 2026.05.0
        """
        from copy import deepcopy

        from docx.table import Table

        src_tbl = other_table._tbl
        src_part = other_table.part
        dest_part = self._part

        new_tbl = deepcopy(src_tbl)

        # -- rewire image references (a:blip and asvg:svgBlip @r:embed) --
        if src_part is not dest_part:
            self._rewire_blip_refs(new_tbl, src_part)
            self._import_table_style(new_tbl, src_part)

        # -- insert the copied w:tbl at the end of the body (before sectPr) --
        body = self._element.body
        body._insert_tbl(new_tbl)  # pyright: ignore[reportPrivateUsage]

        return Table(new_tbl, self._body)

    def add_table_from(self, other_table: Table) -> Table:
        """Alias for :meth:`add_table_copy`.

        .. versionadded:: 2026.05.0
        """
        return self.add_table_copy(other_table)

    def _rewire_blip_refs(self, tbl: "CT_Tbl", src_part: DocumentPart) -> None:
        """Rewire ``a:blip/@r:embed`` refs on ``tbl`` to this document's rels.

        Copies the referenced image parts into this document's package
        (de-duplicating by SHA-1 via
        :meth:`docx.package.Package.get_or_add_image_part`) and rewrites every
        matching ``r:embed`` attribute on the copied tree. Unresolvable refs
        (the source rId is missing or points to a non-image part) are left
        untouched so Word still reads the file; the caller can flag them.

        Not a public API.
        """
        from docx.opc.constants import RELATIONSHIP_TYPE as _RT
        from docx.oxml.ns import qn
        from docx.parts.image import ImagePart

        dest_part = self._part
        package = dest_part.package

        # -- collect every blip-style element with @r:embed --
        embed_attr = qn("r:embed")
        blips = tbl.xpath(
            ".//a:blip[@r:embed] | .//asvg:svgBlip[@r:embed]"
        )
        rid_map: dict[str, str] = {}
        for blip in blips:
            old_rid = blip.get(embed_attr)
            if old_rid is None:
                continue
            new_rid = rid_map.get(old_rid)
            if new_rid is None:
                try:
                    src_img_part = src_part.related_parts[old_rid]
                except KeyError:
                    continue
                if not isinstance(src_img_part, ImagePart):
                    continue
                # -- copy blob into this package's image_parts (dedup by sha1) --
                import io as _io

                new_img_part = package.get_or_add_image_part(
                    _io.BytesIO(src_img_part.blob)
                )
                new_rid = dest_part.relate_to(new_img_part, _RT.IMAGE)
                rid_map[old_rid] = new_rid
            blip.set(embed_attr, new_rid)

    def _import_table_style(self, tbl: "CT_Tbl", src_part: DocumentPart) -> None:
        """Ensure the table's ``w:tblStyle/@w:val`` exists in this document.

        If the referenced styleId is not already defined in this document's
        styles part, deep-copy the source ``w:style`` element across. Styles
        that cascade through ``w:basedOn`` / ``w:link`` / ``w:next`` are *not*
        imported recursively — a V1 limitation noted in the method's caller.

        Not a public API.
        """
        from copy import deepcopy

        from docx.oxml.ns import qn

        styleId = tbl.xpath("string(./w:tblPr/w:tblStyle/@w:val)")
        if not styleId:
            return
        # -- resolve styles parts on both sides (create empty one on dest if needed) --
        dest_styles_part = self._part._styles_part  # pyright: ignore[reportPrivateUsage]
        dest_styles_elm = dest_styles_part.element
        if dest_styles_elm.get_by_id(styleId) is not None:
            return
        try:
            src_styles_part = src_part._styles_part  # pyright: ignore[reportPrivateUsage]
        except Exception:
            try:
                from docx.opc.constants import RELATIONSHIP_TYPE as _RT

                src_styles_part = src_part.part_related_by(_RT.STYLES)
            except Exception:
                return
        src_style = src_styles_part.element.get_by_id(styleId)
        if src_style is None:
            return
        dest_styles_elm.append(deepcopy(src_style))

    def add_alt_chunk(
        self,
        content: bytes | str,
        content_type: str = "text/html",
        match_src: bool | None = None,
    ) -> AltChunk:
        """Append an ``altChunk`` import reference and return an |AltChunk| proxy.

        Creates a new alternate-format import part carrying `content` with
        the given `content_type` (``text/html`` by default), wires it to
        the main document through an ``aFChunk`` relationship, and appends
        a ``w:altChunk`` element at the end of the body that references
        the relationship. Word substitutes the payload when the document
        is opened — python-docx does not evaluate the content itself.

        `content` may be :class:`bytes` or a UTF-8-decodable :class:`str`.
        Pass `match_src=True` to write a ``w:altChunkPr/w:matchSrc`` child
        asking Word to preserve the source's character formatting during
        the import. Closes upstream#1317, upstream#1103, and PR#649.

        .. versionadded:: 2026.05.0
        """
        from docx.alt_chunk import add_alt_chunk_to_document

        return add_alt_chunk_to_document(
            self._part, content, content_type, match_src=match_src
        )

    def add_html_chunk(
        self, html: str, match_src: bool | None = None
    ) -> AltChunk:
        """Append an ``altChunk`` carrying an XHTML payload (R14-5).

        Convenience wrapper over :meth:`add_alt_chunk` that fixes the
        content-type to ``application/xhtml+xml``. Word's HTML import
        filter renders the payload on open; python-docx does not parse
        the markup.

        `html` is encoded as UTF-8. Pass `match_src=True` to write
        ``w:altChunkPr/w:matchSrc`` asking Word to preserve the source's
        character formatting during the import.

        .. warning::
            altChunk payloads execute fully inside Word's rendering
            engine (scripts, external resource fetches, ActiveX
            depending on the user's macro-security settings). Never
            embed untrusted HTML without caller-side sanitisation.
            See the project ``SECURITY.md`` for the full threat model.

        .. versionadded:: 2026.05.10
        """
        return self.add_alt_chunk(
            html, content_type="application/xhtml+xml", match_src=match_src
        )

    def add_text_chunk(
        self,
        text: str,
        encoding: str = "utf-8",
        match_src: bool | None = None,
    ) -> AltChunk:
        """Append an ``altChunk`` carrying a plain-text payload (R14-5).

        Convenience wrapper over :meth:`add_alt_chunk` that fixes the
        content-type to ``text/plain``. `text` is encoded with
        `encoding` before being stored in the part (default ``utf-8``).

        .. versionadded:: 2026.05.10
        """
        payload = text.encode(encoding)
        return self.add_alt_chunk(
            payload, content_type="text/plain", match_src=match_src
        )

    def add_rtf_chunk(
        self, rtf: bytes, match_src: bool | None = None
    ) -> AltChunk:
        """Append an ``altChunk`` carrying an RTF payload (R14-5).

        Convenience wrapper over :meth:`add_alt_chunk` that fixes the
        content-type to ``application/rtf``. `rtf` must already be RTF
        bytes (the helper does not validate the ``{\\rtf1}`` header).

        .. warning::
            RTF payloads can carry embedded OLE objects, external data
            links, and control-word sequences that have been used as
            remote-code-execution vectors in Word (CVE-2017-0199,
            CVE-2023-21716, and similar). Never embed untrusted RTF —
            see the project ``SECURITY.md`` for the threat model.

        .. versionadded:: 2026.05.10
        """
        return self.add_alt_chunk(
            rtf, content_type="application/rtf", match_src=match_src
        )

    def add_mhtml_chunk(
        self, mhtml: bytes, match_src: bool | None = None
    ) -> AltChunk:
        """Append an ``altChunk`` carrying an MHTML payload (R14-5).

        Convenience wrapper over :meth:`add_alt_chunk` that fixes the
        content-type to ``message/rfc822`` (Word's dispatcher for
        multi-part MHTML archives).

        .. versionadded:: 2026.05.10
        """
        return self.add_alt_chunk(
            mhtml, content_type="message/rfc822", match_src=match_src
        )

    @property
    def alt_chunks(self) -> list[AltChunk]:
        """List of |AltChunk| proxies for every ``w:altChunk`` in this document.

        Returns an empty list when the document has no ``altChunk``
        imports. Order matches the document order of the ``w:altChunk``
        elements.

        .. versionadded:: 2026.05.0
        """
        from docx.alt_chunk import iter_alt_chunks

        return iter_alt_chunks(self._part)

    def add_list_of_figures(self, caption_label: str = "Figure") -> Paragraph:
        """Append a "List of Figures" field paragraph at the end of the document.

        Emits a ``TOC \\c "Figure"`` complex field which Word evaluates to
        build a table of items whose caption label matches `caption_label`.
        The cached result text is left empty; the field is marked *dirty*
        so Word rebuilds it on open. Closes upstream#723.

        `caption_label` defaults to ``"Figure"`` but may be set to any
        caption label in the document (e.g. ``"Illustration"``).

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        instr = f' TOC \\c "{caption_label}" '
        field = paragraph.add_complex_field(instr)
        field.mark_dirty()
        return paragraph

    def add_list_of_tables(self, caption_label: str = "Table") -> Paragraph:
        """Append a "List of Tables" field paragraph at the end of the document.

        Emits a ``TOC \\c "Table"`` complex field which Word evaluates to
        build a table of items whose caption label matches `caption_label`.
        The cached result text is left empty; the field is marked *dirty*
        so Word rebuilds it on open. Closes upstream#723.

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        instr = f' TOC \\c "{caption_label}" '
        field = paragraph.add_complex_field(instr)
        field.mark_dirty()
        return paragraph

    def add_table(self, rows: int, cols: int, style: str | _TableStyle | None = None):
        """Add a table having row and column counts of `rows` and `cols` respectively.

        `style` may be a table style object or a table style name. If `style` is |None|,
        the table inherits the default table style of the document.

        Raises |KeyError| (or whatever the style lookup surfaces) when `style`
        names a style that does not exist or is of the wrong type. In that case
        no ``w:tbl`` element is left behind in the body — the freshly-appended
        table is rolled back before the exception propagates. See upstream#563.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        try:
            table.style = style
        except Exception:
            # -- rollback: remove the freshly-added w:tbl so a bad style name
            # -- doesn't leave an orphan table in the body. --
            table.delete()
            raise
        return table

    @property
    def background_color(self) -> RGBColor | None:
        """Document-wide page background color, or |None| if not set.

        Maps to the ``w:color`` attribute on the ``w:background`` child of the
        ``w:document`` root element. Assigning an |RGBColor| writes (or updates)
        the ``w:background`` element. Assigning |None| removes the element.

        .. versionadded:: 2026.05.0
        """
        background = self._element.background
        if background is None:
            return None
        color = background.color
        if not isinstance(color, RGBColor):
            return None
        return color

    @background_color.setter
    def background_color(self, value: RGBColor | None) -> None:
        if value is None:
            self._element._remove_background()
            return
        background = self._element.get_or_add_background()
        background.color = value

    @property
    def page_background_color(self) -> str | None:
        """Document-wide page background color as a 6-char hex string, or |None|.

        Spec-friendly accessor that mirrors :attr:`background_color` but uses
        plain ``"RRGGBB"`` strings (no ``#`` prefix) on both ends. Writing
        |None| removes the ``w:background`` element.

        .. versionadded:: 2026.05.0
        """
        rgb = self.background_color
        if rgb is None:
            return None
        return str(rgb)

    @page_background_color.setter
    def page_background_color(self, value: str | None) -> None:
        if value is None:
            self.background_color = None
            return
        if not isinstance(value, str):
            raise TypeError(
                "page_background_color must be a 'RRGGBB' hex string or None, "
                "got %r" % type(value).__name__
            )
        hex_value = value.lstrip("#").strip()
        if len(hex_value) != 6:
            raise ValueError(
                "page_background_color must be a 6-char hex string, got %r"
                % value
            )
        self.background_color = RGBColor.from_string(hex_value)

    def add_text_watermark(
        self,
        text: str,
        font_name: str = "Calibri",
        font_size: int = 36,
        color_rgb: str = "808080",
        diagonal: bool = True,
    ) -> "Watermark":
        """Append a text watermark to every section's default page header.

        Word renders a watermark by embedding a VML shape in a header paragraph.
        This helper wires one up on each section, replacing any existing
        watermark. Returns the |Watermark| proxy for the first section's shape
        (all sections receive identical watermarks).

        `font_size` is in points, `color_rgb` is an ``"RRGGBB"`` hex string
        (with or without leading ``#``), and `diagonal=True` rotates the
        watermark by 45 degrees.

        .. versionadded:: 2026.05.0
        """
        from docx.shared import Pt

        color_hex = color_rgb.lstrip("#").strip()
        if len(color_hex) != 6:
            raise ValueError(
                "color_rgb must be a 6-char hex string, got %r" % color_rgb
            )
        rgb = RGBColor.from_string(color_hex)
        layout = "diagonal" if diagonal else "horizontal"
        first_watermark: Watermark | None = None
        for section in self.sections:
            wm = section.add_text_watermark(
                text,
                font=font_name,
                size=Pt(font_size),
                color=rgb,
                layout=layout,
            )
            if first_watermark is None:
                first_watermark = wm
        assert first_watermark is not None, "document has no sections"
        return first_watermark

    def add_picture_watermark(
        self,
        image_path: "str | IO[bytes]",
        scale: float = 1.0,
    ) -> "Watermark":
        """Append a picture watermark to every section's default page header.

        `image_path` may be a filesystem path or a file-like stream. `scale`
        is a positive multiplier applied to the image's native dimensions —
        ``scale=0.5`` produces a half-size watermark.

        Returns the |Watermark| proxy for the first section's shape.

        .. versionadded:: 2026.05.0
        """
        if scale <= 0:
            raise ValueError("scale must be > 0, got %r" % scale)

        # -- read native dimensions once so we can size each section identically --
        from docx.image.image import Image as _Image
        from docx.shared import Emu

        img = _Image.from_file(image_path)
        # -- rewind file-like streams so each section can re-read if needed --
        if hasattr(image_path, "seek"):
            try:
                image_path.seek(0)  # type: ignore[union-attr]
            except Exception:
                pass
        px_w, px_h = img.px_width, img.px_height
        h_dpi = img.horz_dpi or 72
        v_dpi = img.vert_dpi or 72
        # -- EMU = inches * 914400; inches = px / dpi --
        width_emu = Emu(int(px_w / h_dpi * 914400 * scale))
        height_emu = Emu(int(px_h / v_dpi * 914400 * scale))

        first_watermark: Watermark | None = None
        for section in self.sections:
            if hasattr(image_path, "seek"):
                try:
                    image_path.seek(0)  # type: ignore[union-attr]
                except Exception:
                    pass
            wm = section.add_image_watermark(
                image_path, width=width_emu, height=height_emu
            )
            if first_watermark is None:
                first_watermark = wm
        assert first_watermark is not None, "document has no sections"
        return first_watermark

    @property
    def watermarks(self) -> "list[Watermark]":
        """List of |Watermark| proxies currently present across all sections.

        Each section contributes at most one watermark (its default header's
        first VML shape). Sections without a watermark are skipped.

        .. versionadded:: 2026.05.0
        """
        result: list[Watermark] = []
        for section in self.sections:
            wm = section.watermark
            if wm is not None:
                result.append(wm)
        return result

    @property
    def bibliography(self) -> "Bibliography":
        """A |Bibliography| of citation sources backing this document.

        Lazily materializes a ``/customXml/item{N}.xml`` part carrying an empty
        ``<b:Sources>`` root if no bibliography is already related to the
        document. Subsequent calls return a proxy over the same underlying
        part.

        .. versionadded:: 2026.05.7
        """
        return self._part.bibliography

    @property
    def citations(self) -> "list[Citation]":
        """All ``CITATION`` fields found in the document body, in document order.

        Walks every top-level body paragraph looking for ``w:instrText``
        runs that begin with ``CITATION`` — whether the field is wrapped in
        a ``<w:sdt>`` citation control (see
        :meth:`Paragraph.add_citation_reference`) or emitted as a plain
        complex field (see :meth:`Paragraph.add_citation`). Each hit is
        surfaced as a :class:`docx.bibliography.Citation` proxy exposing
        ``source_tag`` plus the ``\\p``/``\\f``/``\\s`` switch values
        (``pages``, ``prefix``, ``suffix``).

        .. versionadded:: 2026.05.10
        """
        from docx.bibliography import Citation, is_citation_instruction
        from docx.fields import Field
        from docx.oxml.ns import qn

        result: "list[Citation]" = []
        # -- first: fields at the top level of each paragraph --
        for paragraph in self.paragraphs:
            for field in paragraph.fields:
                if is_citation_instruction(field.instruction):
                    result.append(Citation(field))
            # -- also walk fields wrapped in <w:sdt> citation controls
            # -- that Paragraph.fields / iter_field_elements intentionally
            # -- skip (their begin-run is a descendant, not a direct child). --
            p_element = paragraph._p  # type: ignore[attr-defined]
            for begin_run in p_element.xpath(
                ".//w:sdt//w:r[w:fldChar[@w:fldCharType='begin']]"
            ):
                _ = qn  # keep the import used for lxml's qn caching
                field = Field.for_complex(begin_run)
                if is_citation_instruction(field.instruction):
                    result.append(Citation(field))
        return result

    @property
    def bookmarks(self) -> Bookmarks:
        """A |Bookmarks| object providing access to the bookmarks in this document.

        .. versionadded:: 2026.05.0
        """
        from docx.bookmarks import Bookmarks

        return Bookmarks(self._element.body)

    @property
    def charts(self) -> list[Chart]:
        """List of |Chart| for each chart referenced from the document body.

        A chart is any ``c:chart`` reference element inside a drawing in the
        body (inline or floating). Each reference is resolved to its
        :class:`docx.parts.chart.ChartPart` via the document's relationship
        graph. References whose target part is missing or of the wrong type
        are skipped. Empty list when no charts are present.

        .. versionadded:: 2026.05.0
        """
        from docx.chart import Chart
        from docx.parts.chart import ChartPart

        result: list[Chart] = []
        rIds = self._element.body.xpath(
            ".//w:drawing/wp:inline/a:graphic/a:graphicData/c:chart/@r:id"
            " | .//w:drawing/wp:anchor/a:graphic/a:graphicData/c:chart/@r:id"
        )
        seen: set[str] = set()
        for rId in rIds:
            if rId in seen:
                continue
            seen.add(rId)
            try:
                chart_part = self._part.related_parts[rId]
            except KeyError:
                continue
            if isinstance(chart_part, ChartPart):
                result.append(Chart(chart_part))
        return result

    @property
    def comments(self) -> Comments:
        """A |Comments| object providing access to comments added to the document."""
        return self._part.comments

    @property
    def comments_ids(self) -> "CommentIds":
        """|CommentIds| proxy over ``word/commentsIds.xml``.

        Lazily creates the part on first access so callers writing new
        comments don't need to manage the relationship by hand. The
        returned proxy wraps the live ``<w16cid:commentsIds>`` element
        so mutations persist on save.

        .. versionadded:: 2026.05.10
        """
        return self._part.comments_ids

    @property
    def comments_extensible(self) -> "CommentsExtensible":
        """|CommentsExtensible| proxy over ``word/commentsExtensible.xml``.

        Lazily creates the part on first access. See :attr:`comments_ids`
        for semantics.

        .. versionadded:: 2026.05.10
        """
        return self._part.comments_extensible

    @property
    def permission_ranges(self) -> list[PermissionRange]:
        """All rich-text permission ranges (`w:permStart`) in this document body.

        Returned list is ordered by document-order of the `w:permStart` elements
        in the body.

        .. versionadded:: 2026.05.0
        """
        from docx.oxml.permissions import CT_PermStart
        from docx.permissions import PermissionRange

        body = self._element.body
        return [
            PermissionRange(cast("CT_PermStart", ps), body)
            for ps in body.xpath(".//w:permStart")
        ]

    @property
    def content_controls(self) -> list[ContentControl]:
        """All block-level |ContentControl| objects in this document body, in order.

        Only block-level content controls (direct children of `w:body`) are returned.
        Inline content controls are accessible via :attr:`Paragraph.content_controls`.

        .. versionadded:: 2026.05.0
        """
        return self._body.content_controls

    @property
    def endnotes(self) -> Endnotes:
        """A |Endnotes| object providing access to endnotes in the document.

        .. versionadded:: 2026.05.0
        """
        return self._part.endnotes

    @property
    def equations(self) -> list[Equation]:
        """List of |Equation| for each OMML expression in the document body.

        Walks the body for both inline ``m:oMath`` elements and display-mode
        ``m:oMathPara`` wrappers. Each equation is returned once — an
        ``m:oMath`` nested inside an ``m:oMathPara`` is represented by its
        enclosing wrapper, not separately. Equations inside headers, footers,
        footnotes, endnotes, or comments are not included here; those stories
        are accessible via the corresponding container objects.

        .. versionadded:: 2026.05.0
        """
        from docx.equations import Equation

        body = self._element.body
        result: list[Equation] = []
        for el in body.xpath(
            ".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]"
        ):
            result.append(Equation(el))
        return result

    def iter_math_expressions(self) -> "Iterator[MathExpr]":
        """Yield a :class:`~docx.math.MathExpr` proxy for every OMML expression in the body.

        Document-wide walk that dispatches each ``m:oMathPara`` /
        ``m:oMath`` (not already inside an ``m:oMathPara``) through
        :func:`ooxml_math.from_element`, so callers receive the richest
        typed proxy known to ``ooxml_math`` 0.4.0 — including the deferred
        :class:`~docx.math.Bar`, :class:`~docx.math.Box`,
        :class:`~docx.math.BorderBox`, :class:`~docx.math.Phantom`,
        :class:`~docx.math.GroupChar` and :class:`~docx.math.EqArray`
        wrappers — falling back to :class:`~docx.math.Raw` for any
        unrecognised tag.

        Equations inside headers, footers, footnotes, endnotes, or comments
        are not included; those stories are accessible via the corresponding
        container objects.

        .. versionadded:: 2026.05.10
        """
        from ooxml_math import from_element as _from_element

        body = self._element.body
        for el in body.xpath(
            ".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]"
        ):
            yield _from_element(el)

    @property
    def form_fields(self) -> list[FormField]:
        """All legacy form fields (``w:ffData``) found in the document body, in order.

        Walks top-level body paragraphs only. Form fields nested inside table
        cells, headers, footers, footnotes, or endnotes are not included in
        this collection — callers can access those via the ``form_fields``
        property on the enclosing paragraph.

        .. versionadded:: 2026.05.0
        """
        result: list[FormField] = []
        for paragraph in self.paragraphs:
            result.extend(paragraph.form_fields)
        return result

    @property
    def fields(self) -> "list[Field]":
        """All fields found in the document body, in document order.

        Includes both simple (``w:fldSimple``) and complex (``w:fldChar``)
        fields. Walks top-level body paragraphs only — fields nested inside
        table cells, headers, footers, footnotes, or endnotes are not
        included; callers can access those via the ``fields`` property on
        the enclosing paragraph.

        .. versionadded:: 2026.05.10
        """
        from docx.fields import Field  # noqa: F401 (re-exported in type hint)

        result: "list[Field]" = []
        for paragraph in self.paragraphs:
            result.extend(paragraph.fields)
        return result

    def rebuild_tocs(self, page_number_placeholder: str = "?") -> int:
        """Recompute every TOC field's cached result from current content.

        Calls :meth:`docx.fields.TocField.rebuild` on every TOC-family
        field in the document body — plain tables of contents, lists of
        figures / tables (:class:`TableOfFiguresField`), and tables of
        authorities (:class:`TableOfAuthoritiesField`). Each field's
        cached result (the runs between the ``separate`` and ``end``
        markers) is replaced with a tab-separated preview built from the
        current document content.

        `page_number_placeholder` fills in the column where Word would
        show a real page number — ``"?"`` by default, matching the
        placeholder Word itself displays for a dirty TOC before the first
        refresh. python-docx has no layout engine, so accurate page
        numbers cannot be produced; :meth:`Paragraph.add_toc` marks the
        field dirty so Word recomputes the real numbers on open.

        Returns the number of TOC-family fields rebuilt.

        .. versionadded:: 2026.05.10
        """
        count = 0
        for field in self.fields:
            toc = field.as_toc
            if toc is None:
                continue
            toc.rebuild(page_number_placeholder)
            count += 1
        return count

    @property
    def linked_targets(self) -> "list[LinkedTarget]":
        """All linked external targets (``INCLUDETEXT`` fields) in document order.

        Each entry is a :class:`docx.linked_content.LinkedTarget` proxy
        wrapping one ``INCLUDETEXT`` complex field. The list is built
        on demand from :attr:`fields`, so newly-added links via
        :meth:`docx.text.paragraph.Paragraph.link_to` show up on the
        next access without a manual reload.

        Walks top-level body paragraphs only — links nested inside
        table cells, headers, footers, footnotes, or endnotes are not
        currently included; callers can access those via the
        ``fields`` property on the enclosing paragraph and wrap each
        ``INCLUDETEXT`` field manually with
        :class:`~docx.linked_content.LinkedTarget`.

        .. versionadded:: 2026.05.13
        """
        from docx.linked_content import iter_linked_targets

        return list(iter_linked_targets(self))

    def update_links(self, base_dir: "str | None" = None) -> int:
        """Re-resolve every link target and rewrite the cached field result.

        Walks every ``INCLUDETEXT`` field in the document body
        (:attr:`linked_targets`), calls
        :meth:`docx.linked_content.LinkedTarget.refresh` on each, and
        returns the count of fields whose cached text was actually
        rewritten. Failed resolutions are silently skipped — the
        existing cached text is preserved so the document never loses
        information on a broken refresh.

        `base_dir` scopes relative paths in the link URLs. When
        |None|, paths are resolved against the current working
        directory.

        Resolution behaviour by link kind:

        * ``xlsx-cell`` / ``xlsx-table-column`` — opens the workbook
          via the sibling ``xlsx`` package and fetches the live value.
          Returns the empty string when the cell is blank.
        * ``pptx-slide`` — returns ``"[Slide N]"`` (or
          ``"[Slide N: <title>]"`` when the sibling ``pptx`` package is
          installed). Real slide rendering requires PowerPoint /
          LibreOffice and is intentionally out of scope.
        * ``unknown`` — left untouched.

        .. versionadded:: 2026.05.13
        """
        from docx.linked_content import update_document_links

        return update_document_links(self, base_dir=base_dir)

    @property
    def has_macros(self) -> bool:
        """True if this document contains a VBA project (macros).

        .. versionadded:: 2026.05.0
        """
        try:
            self._part.part_related_by(RT.VBA_PROJECT)
            return True
        except KeyError:
            return False

    @property
    def is_signed(self) -> bool:
        """True when this document's package contains digital-signature parts.

        python-docx does not verify signatures; this reports only whether a
        ``_xmlsignatures/origin.sigs`` or signature relationship is present at the
        package level.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).is_signed

    @property
    def signatures(self) -> list[SignatureInfo]:
        """List of |SignatureInfo| for each digital signature in the package.

        Empty list when the document is unsigned. See :class:`docx.signatures.SignatureInfo`
        for the available metadata.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).signatures

    def add_signature_line(
        self,
        signer_name: str,
        signer_title: str | None = None,
        email: str | None = None,
    ) -> "SignatureInfo":
        """Attach an unsigned signature-line placeholder to this document.

        Creates an unsigned ``/_xmlsignatures/sigN.xml`` placeholder part
        declaring *signer_name* (with optional *signer_title* / *email*
        encoded in ``mdssi:SignatureComments``). The placeholder is *not*
        a cryptographically valid signature — python-docx does not have
        access to a signing key. Round-trips through save + reload so
        downstream signing tools (or
        :class:`ooxml_signatures.Signer` / Word) can finalise it.

        .. versionadded:: 2026.05.10
        """
        from docx.package import Package

        return cast("Package", self._part.package).add_signature_line(
            signer_name=signer_name,
            signer_title=signer_title,
            email=email,
        )

    @property
    def font_table(self) -> FontTable | None:
        """A |FontTable| collection, or |None| if no font-table part is related.

        Returns |None| when the document has no ``fontTable`` relationship. For
        authoring workflows that need to embed a font use
        :attr:`font_table_or_new` instead, which materialises an empty
        ``fontTable.xml`` on demand.

        .. versionadded:: 2026.05.0
        """
        return self._part.font_table

    @property
    def font_table_or_new(self) -> FontTable:
        """A |FontTable| collection, creating an empty ``fontTable.xml`` if needed.

        Use this when the caller intends to *add* to the font table (e.g. via
        :meth:`FontTable.add_embedded_font`) — unlike :attr:`font_table` it
        never returns |None|.

        .. versionadded:: 2026.05.0
        """
        return self._part.font_table_or_new

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| object providing access to footnotes in the document.

        .. versionadded:: 2026.05.0
        """
        return self._part.footnotes

    def accept_all_changes(self) -> int:
        """Accept every tracked change in the document body.

        Insertions are flattened into live content, deletions are removed, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are discarded
        (the current, post-edit formatting is retained).

        Returns the number of change elements resolved.

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=True)

    def reject_all_changes(self) -> int:
        """Reject every tracked change in the document body.

        Insertions are removed, deletions are restored as live content, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are unwound so
        the prior formatting is restored.

        Returns the number of change elements resolved.

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=False)

    def accept_revisions(self) -> int:
        """Alias of :meth:`accept_all_changes`.

        Matches the ECMA-376 "revision" vocabulary used by
        :attr:`Document.revisions`, :attr:`Paragraph.revisions`, and
        :attr:`Run.revisions`. Returns the number of change elements resolved.

        .. versionadded:: 2026.05.11
        """
        return self.accept_all_changes()

    def reject_revisions(self) -> int:
        """Alias of :meth:`reject_all_changes`.

        Matches the ECMA-376 "revision" vocabulary used by
        :attr:`Document.revisions`, :attr:`Paragraph.revisions`, and
        :attr:`Run.revisions`. Returns the number of change elements resolved.

        .. versionadded:: 2026.05.11
        """
        return self.reject_all_changes()

    def accept_all_revisions(self) -> int:
        """Bulk-accept every tracked revision in the document body.

        Equivalent to :meth:`accept_all_changes`; provided under the
        ECMA-376 "revision" spelling that matches
        :attr:`Document.revisions`, :meth:`accept_revisions_by_author`, and
        related accessors. Returns the number of change elements resolved.

        .. versionadded:: 2026.05.13
        """
        return self.accept_all_changes()

    def reject_all_revisions(self) -> int:
        """Bulk-reject every tracked revision in the document body.

        Equivalent to :meth:`reject_all_changes`; provided under the
        ECMA-376 "revision" spelling that matches
        :attr:`Document.revisions`, :meth:`reject_revisions_by_author`, and
        related accessors. Returns the number of change elements resolved.

        .. versionadded:: 2026.05.13
        """
        return self.reject_all_changes()

    def accept_revisions_by_author(self, author: str) -> int:
        """Accept every tracked revision whose ``w:author`` is `author`.

        Resolves run-level (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`),
        cell-level (`w:cellIns`, `w:cellDel`), and formatting-level
        (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tcPrChange`,
        `w:trPrChange`, `w:tblPrChange`) revisions whose ``@w:author`` matches
        `author` exactly. Revisions with any other author survive untouched.

        Returns the number of change elements resolved.

        .. versionadded:: 2026.05.13
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(
            self._element.body, accept=True, author=author
        )

    def reject_revisions_by_author(self, author: str) -> int:
        """Reject every tracked revision whose ``w:author`` is `author`.

        Mirror of :meth:`accept_revisions_by_author`. Returns the number of
        change elements resolved.

        .. versionadded:: 2026.05.13
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(
            self._element.body, accept=False, author=author
        )

    @property
    def revisions(self) -> "list":
        """All run-level revisions in the document body, in document order.

        Returns a list of :class:`~docx.tracked_changes.Insertion`,
        :class:`~docx.tracked_changes.Deletion`, and
        :class:`~docx.tracked_changes.Move` proxies wrapping the body's
        `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo` descendants. Formatting
        revisions (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`,
        `w:tcPrChange`, `w:trPrChange`, `w:tblPrChange`) and cell markers
        (`w:cellIns`, `w:cellDel`) are excluded — those are exposed through
        per-type proxies (:attr:`Run.formatting_change`, etc.).

        Nested revisions (e.g. a `w:ins` inside an existing `w:del`) are
        included and appear after their enclosing ancestor in the list.

        .. versionadded:: 2026.05.11
        """
        from docx.tracked_changes import TrackedChange, _wrap_revision

        body = self._element.body
        result: list[TrackedChange] = []
        for elm in body.xpath(
            ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"
        ):
            result.append(_wrap_revision(elm))
        return result

    @property
    def footnote_properties(self) -> FootnoteProperties | None:
        """Document-level |FootnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:footnotePr`` element exists in the document settings.
        Use :meth:`add_footnote_properties` to add one and configure it.

        .. versionadded:: 2026.05.0
        """
        return self.settings.footnote_properties

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return document-level |FootnoteProperties|, adding a ``w:footnotePr`` if needed.

        .. versionadded:: 2026.05.0
        """
        return self.settings.add_footnote_properties()

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """Document-level |EndnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:endnotePr`` element exists in the document settings.
        Use :meth:`add_endnote_properties` to add one and configure it.

        .. versionadded:: 2026.05.0
        """
        return self.settings.endnote_properties

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return document-level |EndnoteProperties|, adding a ``w:endnotePr`` if needed.

        .. versionadded:: 2026.05.0
        """
        return self.settings.add_endnote_properties()

    @property
    def core_properties(self):
        """A |CoreProperties| object providing Dublin Core properties of document."""
        return self._part.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| collection providing access to custom document properties.

        Custom properties are user-defined, typed name/value pairs stored in the
        ``docProps/custom.xml`` part of the package. They are distinct from the fixed
        "core" Dublin-Core properties available via :attr:`core_properties`.

        .. versionadded:: 2026.05.0
        """
        return self._part.custom_properties

    def set_language(
        self,
        latin: str | None,
        east_asian: str | None = None,
        bidi: str | None = None,
    ) -> None:
        """Set the document-level theme-font language tags.

        Convenience wrapper for assigning
        :attr:`Settings.theme_font_language`. Pass ``latin`` as a BCP-47
        language tag (e.g. ``"en-US"``); the East-Asian and bidi tags are
        optional and default to |None| (meaning leave unset).

        .. versionadded:: 2026.05.0
        """
        self.settings.theme_font_language = (latin, east_asian, bidi)

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| proxy for this document's ``docProps/app.xml``.

        Exposes application-written metadata such as ``Company``, ``Manager``,
        ``Application``, ``AppVersion``, ``TotalTime``, and the cached
        ``Pages`` / ``Words`` / ``Characters`` statistics. A default (empty)
        extended-properties part is created on demand when none is present.

        .. versionadded:: 2026.05.0
        """
        return self._part.extended_properties

    @property
    def custom_xml_parts(self) -> list[CustomXmlPart]:
        """List of |CustomXmlPart| proxies for the custom XML data parts in the package.

        Empty when the document has no ``customXml`` relationships. Each entry
        exposes the data part's :attr:`~docx.custom_xml.CustomXmlPart.item_id`,
        :attr:`~docx.custom_xml.CustomXmlPart.schema_refs`,
        :attr:`~docx.custom_xml.CustomXmlPart.root_element`, and
        :attr:`~docx.custom_xml.CustomXmlPart.blob` read-only.

        .. versionadded:: 2026.05.0
        """
        return self._part.custom_xml_parts

    @property
    def numbering(self):
        """A |Numbering| object providing read/write access to the list-style
        numbering definitions for this document.

        Creates a default (empty) numbering part if one is not already related to the
        document.

        .. versionadded:: 2026.05.0
        """
        return self._part.numbering_part.numbering

    def list_labels(self) -> dict[int, str]:
        """Return ``{id(p_element): label}`` for every numbered paragraph in the body.

        Walks the document body top-to-bottom exactly once and, for each
        paragraph that resolves to a list (via a direct ``w:numPr`` or a
        paragraph style that declares one), computes the Word-rendered label
        (``"1."``, ``"a)"``, ``"I."``, ``"•"``, ``"1.1."``, ...) using the
        level's ``w:lvlText`` pattern and ``w:numFmt`` value.
        Paragraphs that are not part of any list are omitted.
        The mapping key is ``id(paragraph._p)`` — stable for the lifetime
        of the underlying element. To look up by |Paragraph| use
        ``labels[id(paragraph._p)]`` or use :attr:`Paragraph.list_label`.
        Returns an empty mapping when the document has no numbering part or
        no paragraph resolves to a list.

        Supported ``numFmt`` values: ``decimal``, ``decimalZero``,
        ``upperRoman``, ``lowerRoman``, ``upperLetter``, ``lowerLetter``,
        ``bullet``. Other formats (``cardinalText``, ``ordinalText``, ...)
        fall back to decimal.

        .. versionadded:: 2026.05.0
        """
        from docx.numbering import ListLabelRenderer

        numbering_part = getattr(self._part, "numbering_part", None)
        numbering_elm = (
            numbering_part.numbering_element if numbering_part is not None else None
        )
        styles_elm = None
        try:
            styles_part = self._part.part_related_by(RT.STYLES)
        except (KeyError, AttributeError):
            styles_part = None
        if styles_part is not None:
            styles_elm = getattr(styles_part, "element", None)

        renderer = ListLabelRenderer(numbering_elm, styles_elm)
        return renderer.label_map(self._element.body.xpath(".//w:p"))

    @property
    def ink_annotations(self) -> list[InkAnnotation]:
        """List of |InkAnnotation| objects for each ink annotation in the body.

        An ink annotation is any ``w:contentPart`` element that targets an ink part
        (content type ``application/inkml+xml``). The list is empty when no ink
        annotations are present. Read-only — python-docx does not support creating
        or modifying ink annotations.

        .. versionadded:: 2026.05.0
        """
        from docx.text.paragraph import Paragraph

        result: list[InkAnnotation] = []
        for p in self._element.body.xpath(".//w:p[.//w:contentPart]"):
            paragraph = Paragraph(p, self._body)
            result.extend(paragraph.ink_annotations)
        return result

    @property
    def attachments(self) -> list[Attachment]:
        """List of |Attachment| for each ``w:altChunk`` in the document body.

        An ``altChunk`` is an arbitrary foreign payload (HTML, RTF, another
        docx, etc.) that Word merges into the document on open. python-docx
        exposes them read-only: callers can iterate the altChunk elements,
        inspect their content-type (via the related part), and retrieve the
        raw payload bytes for further processing.

        Returns an empty list when the document has no altChunks.

        .. versionadded:: 2026.05.0
        """
        from docx.attachments import Attachment

        result: list[Attachment] = []
        doc_part = self._part
        for alt_elm in self._element.body.xpath(".//w:altChunk"):
            rId = alt_elm.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target_part = None
            if rId is not None:
                target_part = doc_part.related_parts.get(rId)
            result.append(Attachment(alt_elm, target_part))
        return result

    @property
    def embedded_objects(self) -> list[EmbeddedObject]:
        """List of |EmbeddedObject| for each embedded OLE object in the body.

        An embedded object is any ``w:object`` element containing an
        ``o:OLEObject`` descendant (content type
        ``application/vnd.openxmlformats-officedocument.oleObject`` for the
        related binary). The list is empty when no embedded objects are
        present. Read-only — python-docx does not support creating or
        modifying embedded objects, or extracting the ``w:pict`` image that
        Word displays in place of the OLE content.

        .. versionadded:: 2026.05.0
        """
        from docx.text.paragraph import Paragraph

        result: list[EmbeddedObject] = []
        for p in self._element.body.xpath(".//w:p[.//w:object/o:OLEObject]"):
            paragraph = Paragraph(p, self._body)
            result.extend(paragraph.embedded_objects)
        return result

    def add_smart_art(
        self,
        layout_name: str = "list",
        width: Length | None = None,
        height: Length | None = None,
    ) -> SmartArt:
        """Append a SmartArt diagram of `layout_name` and return a |SmartArt| proxy.

        `layout_name` is one of ``"list"``, ``"cycle"`` or ``"process"``
        (case-insensitive). Each selects a Word built-in layout family whose
        URN is baked into the freshly-minted ``data1.xml`` part. Word uses
        its internal layout engine keyed by that URN, so the embedded
        ``layout1.xml`` serves mainly to satisfy the OOXML package
        requirements rather than to drive rendering.

        `width` and `height` are |Length| values for the inline drawing's
        display size. When omitted a 5.5" x 3" default is used — a shape
        close to Word's own default SmartArt frame.

        The returned |SmartArt| is empty; populate it with
        :meth:`SmartArt.add_node` one string at a time. The diagram is
        appended in its own paragraph at the end of the document body,
        wrapped in a ``wp:inline`` so it flows with text like any other
        inline picture.

        Raises :class:`ValueError` when `layout_name` is not one of the
        supported families.

        .. versionadded:: 2026.05.7
        """
        from docx.smart_art import add_smart_art_to_document

        cx = int(width) if width is not None else int(Inches(5.5))
        cy = int(height) if height is not None else int(Inches(3))
        return add_smart_art_to_document(self, layout_name, cx, cy)

    @property
    def smart_art(self) -> list[SmartArt]:
        """List of |SmartArt| proxies for every SmartArt diagram in the body.

        Walks top-level body paragraphs (including paragraphs nested inside
        body-level tables) and returns one entry for each ``w:drawing`` that
        references a SmartArt diagram (i.e. contains a ``dgm:relIds``
        element). Empty list when the document has no SmartArt. Read-only —
        python-docx does not support creating or modifying SmartArt.

        .. versionadded:: 2026.05.0
        """
        return list(self.iter_smart_arts())

    @property
    def smart_arts(self) -> list[SmartArt]:
        """Plural alias for :attr:`smart_art`.

        Returns the same materialised list of |SmartArt| proxies. Provided
        so call sites that iterate naturally with plural naming read more
        cleanly — ``for sa in document.smart_arts: ...`` — while keeping
        the historical :attr:`smart_art` attribute stable.

        .. versionadded:: 2026.05.10
        """
        return self.smart_art

    def iter_smart_arts(self):
        """Yield each |SmartArt| proxy in the document body in document order.

        Streams the same sequence as :attr:`smart_arts` / :attr:`smart_art`
        without materialising a list. Useful for documents carrying many
        SmartArt diagrams (e.g. long research reports) where the caller
        only needs the first matching diagram.

        .. versionadded:: 2026.05.10
        """
        from docx.drawing import Drawing

        for d in self._element.body.xpath(".//w:drawing"):
            drawing = Drawing(d, self._body)
            sa = drawing.smart_art
            if sa is not None:
                yield sa

    def to_html(
        self,
        include_styles: bool = True,
        embed_images: bool = True,
    ) -> str:
        """Return an HTML5 rendering of this document as a string.

        A minimal, preview-grade exporter. Maps the main structural
        elements (paragraphs, runs, hyperlinks, tables, inline
        pictures, headings, lists) to equivalent HTML5 constructs with
        inline CSS for alignment, margins, borders, and colour.

        When ``include_styles`` is |True| (the default) a single
        ``<style>`` block carrying coarse CSS rules derived from this
        document's style definitions (``font-family``, ``font-size``,
        ``color``, margin) is emitted in ``<head>``. Pass |False| for
        a style-free document.

        When ``embed_images`` is |True| (the default) inline pictures
        are emitted with ``data:<content-type>;base64,…`` ``<img
        src>`` URLs, producing a self-contained HTML string. When
        |False|, pictures use ``cid:{rId}`` placeholders so a MIME
        assembler (email, MHTML, etc.) can attach the parts
        separately.

        Text content is HTML-escaped at every text node (including
        hyperlink URLs) to guard against XSS from document content.

        Not round-trippable: there is no HTML→docx import path. For
        fidelity beyond simple structure — fields, shapes, text
        boxes, anchored pictures, and equations — this exporter
        emits an ``<!-- unsupported: … -->`` comment and continues.

        .. versionadded:: 2026.05.10
        """
        from docx.html_export import document_to_html

        return document_to_html(
            self, include_styles=include_styles, embed_images=embed_images
        )

    def to_markdown(self) -> str:
        r"""Return a GitHub-Flavoured-Markdown rendering of this document as a string.

        A minimal, preview-grade exporter. Maps the main structural
        elements to GFM constructs:

        * ``Heading 1`` .. ``Heading 6`` -> ``#`` .. ``######``
        * Bold runs                       -> ``**text**``
        * Italic runs                     -> ``_text_``
        * Inline code (``Code`` /         -> backtick-wrapped text
          ``HTMLCode`` style, or
          monospace font)
        * Hyperlinks                      -> ``[text](url)``
        * Bullet list items               -> ``- ``
        * Numbered list items             -> ``1. ``
        * Tables                          -> GFM ``| col | col |``
        * Block quotes                    -> ``> ``
        * Inline pictures                 -> ``![alt](archive-path)``
          where the path is the .docx
          zip-relative location (e.g.
          ``word/media/image1.png``)
        * Page breaks                     -> ``---``
        * Footnotes / endnotes            -> ``[^N]`` references with
          ``[^N]: text`` blocks at the
          end

        Lossy conversions (Markdown is a strict subset of Word's
        expressiveness):

        * Run-level fonts, sizes, and colours collapse -- only bold,
          italic, and inline-code survive.
        * Paragraph alignment, indentation, and spacing collapse to
          plain paragraph breaks.
        * Drawing anchors, text boxes, OMML equations, fields, and
          SmartArt are skipped.
        * Tables flatten multi-paragraph cells to space-joined text --
          GFM cells cannot carry block content.
        * Image bytes are *not* embedded; the reference is the archive
          path. Consumers that need the raw bytes should extract them
          from the .docx zip alongside the Markdown output.

        Not round-trippable: there is no Markdown -> docx import path.

        .. versionadded:: 2026.05.29
        """
        from docx.markdown_export import document_to_markdown

        return document_to_markdown(self)

    def diff(self, other: "Document", level: str = "content"):
        """Return a :class:`~docx.semantic_diff.SemanticDiff` against `other`.

        Compares the *content* of two documents — paragraph adds /
        removes / modifications, table mutations, image counts, and
        (at ``level="formatting"``) style changes — rather than the
        raw XML, which would over-report whitespace and ordering noise
        with no visible impact.

        ``level`` selects the granularity:

        * ``"structural"`` — paragraph add / remove / move only.
        * ``"content"`` (default) — adds per-paragraph text edits.
        * ``"formatting"`` — adds style / font / colour changes.

        Example::

            old = Document("q1-review-v1.docx")
            new = Document("q1-review-v2.docx")
            diff = old.diff(new)
            print(diff.summary)
            # {'paragraphs_added': 3, 'paragraphs_removed': 1,
            #  'paragraphs_modified': 7, 'tables_modified': 1,
            #  'images_added': 0, 'styles_changed': 0,
            #  'total_changes': 12}
            for change in diff.changes:
                print(change.kind, change.target, change.before, change.after)

        The returned object exposes :meth:`~docx.semantic_diff.SemanticDiff.to_markdown`
        (for PR comments), :meth:`~docx.semantic_diff.SemanticDiff.to_html`
        (for web UIs), and
        :meth:`~docx.semantic_diff.SemanticDiff.to_word_track_changes`
        (best-effort visible-marker docx) for downstream rendering.

        .. versionadded:: 2026.05.13
        """
        from docx.semantic_diff import compute_diff

        return compute_diff(self, other, level=level)

    @property
    def inline_shapes(self):
        """The |InlineShapes| collection for this document.

        An inline shape is a graphical object, such as a picture, contained in a run of
        text and behaving like a character glyph, being flowed like other text in a
        paragraph.
        """
        return self._part.inline_shapes

    def iter_inner_content(
        self, include_sdt_flat: bool = False
    ) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this document in document order.

        When `include_sdt_flat` is |True|, any block-level ``w:sdt`` wrapper
        found in the body is flattened: the paragraphs and tables inside the
        sdt's ``w:sdtContent`` are yielded in place, as if the wrapper were
        transparent. This surfaces paragraphs that live inside content
        controls (including a TOC's own ``w:sdt`` wrapper), which the default
        iteration does not reach — closes upstream#1280.

        .. versionchanged:: 2026.05.0
            Added ``include_sdt_flat`` parameter.
        """
        if not include_sdt_flat:
            return self._body.iter_inner_content()
        return self._body._iter_inner_content_flat_sdt()

    @property
    def text(self) -> str:
        """Concatenated text of every top-level paragraph in the document body.

        Paragraphs are joined with a single ``"\\n"`` separator; tables in the
        body are skipped. This is the quick "give me the body text" helper
        requested in upstream#252 / upstream#72. For a breakdown that walks
        tables or non-body stories, iterate :attr:`paragraphs` /
        :attr:`tables` manually.

        .. versionadded:: 2026.05.0
        """
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def recovery_warnings(self) -> list[str]:
        """List of parse-warning strings collected while opening this document.

        Populated only when the document was opened via :func:`docx.Document`
        with ``recover=True``. Empty for normally-opened documents and for
        well-formed documents opened in recovery mode.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).recovery_warnings

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The |Paragraph| instances in the document, in document order.

        Note that paragraphs within revision marks such as ``<w:ins>`` or ``<w:del>`` do
        not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self) -> DocumentPart:
        """The |DocumentPart| object of this document."""
        return self._part

    def select(self, selector: str) -> list:
        """Return every proxy matching CSS-style ``selector`` in document order.

        Supports the eight element kinds ``p``, ``r``, ``tbl``, ``tr``,
        ``td``, ``hyperlink``, ``bookmark``, and ``comment``; the four
        attribute operators ``=`` / ``^=`` / ``$=`` / ``*=`` (plus bare
        ``[name]`` for "exists / is True"); the descendant (``" "``),
        child (``">"``), and adjacent-sibling (``"+"``) combinators; and
        the ``:first-child`` / ``:last-child`` / ``:nth-child(N)`` /
        ``:not(...)`` pseudo-classes. See :mod:`docx.selectors` for the
        full cheatsheet.

        Examples::

            doc.select('p[style="Heading 1"]')
            doc.select('p[style^="Heading "] r[bold]')
            doc.select('tbl tr td:nth-child(2)')

        Raises :class:`docx.selectors.SelectorSyntaxError` on malformed
        selectors. Closes #78.

        .. versionadded:: 2026.05.13
        """
        from docx.selectors import select as _select

        return _select(self, selector)

    def select_one(self, selector: str):
        """Return the first proxy matching ``selector`` or |None|.

        Convenience wrapper over :meth:`select` that stops after the
        first hit.

        .. versionadded:: 2026.05.13
        """
        from docx.selectors import select_one as _select_one

        return _select_one(self, selector)

    def replace(
        self,
        old_text: str,
        new_text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> int:
        """Replace occurrences of `old_text` with `new_text` in the document body paragraphs.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not affected.

        Preserves the run formatting of the first character's run for each replacement.
        Returns the number of replacements made.

        When `case_sensitive` is False, matching is case-insensitive. When `whole_word` is
        True, only whole-word matches are replaced.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_paragraphs

        return replace_in_paragraphs(
            self.paragraphs, old_text, new_text, case_sensitive, whole_word
        )

    def replace_all(
        self,
        old_text: str,
        new_text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> int:
        """Replace `old_text` with `new_text` in every story in this document.

        Unlike :meth:`replace`, which updates only top-level body paragraphs, this
        method walks every "story" in the package — the body (including top-level
        body tables), each section's non-inherited headers and footers, footnotes,
        endnotes, and comments — and applies the replacement to each.

        Paragraphs nested inside tables that live within a header, footer,
        footnote, endnote, or comment story are not descended into, and neither
        are tables nested inside body-level table cells; see
        :func:`docx.search._iter_all_paragraphs` for the full iteration contract.

        Returns the total number of replacements made across all stories.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_all_paragraphs

        return replace_in_all_paragraphs(
            self, old_text, new_text, case_sensitive, whole_word
        )

    def replace_regex(
        self,
        pattern: str | re.Pattern[str],
        replacement: str,
        flags: int = 0,
    ) -> int:
        """Replace all regex matches of `pattern` with `replacement` in body paragraphs.

        `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
        `flags` (e.g. `re.IGNORECASE`) is applied during compilation; when `pattern` is
        already compiled, `flags` is ignored. `replacement` follows `re.sub` semantics —
        backreferences such as ``\\1`` and ``\\g<name>`` are expanded per match.

        Note: Only top-level body paragraphs are processed. Text inside table cells,
        headers, footers, footnotes, and endnotes is not affected.

        Preserves the run formatting of the first character's run for each replacement.
        Returns the number of replacements made.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_paragraphs_regex

        return replace_in_paragraphs_regex(
            self.paragraphs, pattern, replacement, flags
        )

    def replace_regex_all(
        self,
        pattern: str | re.Pattern[str],
        replacement: str,
        flags: int = 0,
    ) -> int:
        """Replace regex `pattern` with `replacement` across every story in this document.

        Like :meth:`replace_all` but using regex semantics. ``replacement`` follows
        :func:`re.sub` semantics — backreferences such as ``\\1`` and ``\\g<name>``
        are expanded per match.

        Returns the total number of replacements made across all stories.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_all_paragraphs_regex

        return replace_in_all_paragraphs_regex(self, pattern, replacement, flags)

    def resolve_cross_references(self) -> int:
        """Resolve ``REF`` and ``PAGEREF`` fields in the document body.

        For each ``REF`` field whose bookmark exists, replaces the field's
        rendered :attr:`~docx.fields.Field.result_text` with the concatenated
        text of the referenced bookmark range. For each ``PAGEREF`` field
        whose cached result is empty, replaces it with ``"?"``; python-docx
        cannot compute real page numbers because it has no layout engine.

        Other field types (``PAGE``, ``DATE``, ``SEQ``, etc.) are ignored.
        Fields whose bookmark cannot be found are left unchanged. Returns the
        number of fields whose ``result_text`` was updated in place.

        Walks every field element in the body — including those inside
        tables, hyperlinks, and other containers — by descending all
        ``w:fldSimple`` elements and every run that opens a complex field.

        .. versionadded:: 2026.05.0
        """
        from docx.fields import Field

        body = self._element.body
        updated = 0
        for el in body.xpath(
            ".//w:fldSimple | .//w:r[w:fldChar[@w:fldCharType='begin']]"
        ):
            tag = el.tag.rsplit("}", 1)[-1]
            field = (
                Field.for_simple(el) if tag == "fldSimple" else Field.for_complex(el)
            )
            if field.type not in (
                "REF",
                "PAGEREF",
                "DOCPROPERTY",
                "AUTHOR",
                "TITLE",
                "SUBJECT",
                "KEYWORDS",
                "COMMENTS",
                "LASTSAVEDBY",
            ):
                continue
            resolved = field.resolve(self)
            if resolved == field.result_text:
                continue
            field.update_result_text(resolved)
            updated += 1
        return updated

    def evaluate_fields(self, context: "dict[str, object] | None" = None) -> int:
        """Evaluate complex-type fields in the document body against `context`.

        For every ``w:fldSimple`` or complex field in the body, calls
        :meth:`Field.evaluate` with ``context`` augmented with
        ``context["document"] = self`` (so property / cross-reference fields
        can still resolve). When the evaluated text differs from the cached
        :attr:`~docx.fields.Field.result_text`, it is written back in place
        via :meth:`Field.update_result_text`.

        Supported field types include ``IF`` (with nested ``MERGEFIELD``),
        ``MERGEFIELD``, ``HYPERLINK``, the ``=``-prefix formula field, and
        the runtime-dynamic ``PAGE`` / ``NUMPAGES`` / ``DATE`` / ``TIME``
        placeholders (which fall through to the cached result when present).
        ``REF`` / ``PAGEREF`` / ``DOCPROPERTY`` / core-property fields are
        delegated to :meth:`Document.resolve_cross_references` semantics.

        Returns the number of fields whose ``result_text`` was updated.

        .. versionadded:: 2026.05.7
        """
        from docx.fields import Field

        ctx: dict[str, object] = dict(context) if context else {}
        ctx.setdefault("document", self)

        body = self._element.body
        updated = 0
        for el in body.xpath(
            ".//w:fldSimple | .//w:r[w:fldChar[@w:fldCharType='begin']]"
        ):
            tag = el.tag.rsplit("}", 1)[-1]
            field = (
                Field.for_simple(el) if tag == "fldSimple" else Field.for_complex(el)
            )
            try:
                evaluated = field.evaluate(ctx)
            except Exception:
                continue
            if evaluated == field.result_text:
                continue
            field.update_result_text(evaluated)
            updated += 1
        return updated

    def revision_marks_text(
        self,
        open_ins: str = "[+",
        close_ins: str = "+]",
        open_del: str = "[-",
        close_del: str = "-]",
    ) -> str:
        """Return the document body's text with tracked-change markers applied.

        Each top-level paragraph is rendered via
        :meth:`Paragraph.revision_marks_text` and the results are joined with a
        blank line separator (``"\\n\\n"``). Tables in the body are skipped; this
        helper is intended as a quick CLI preview of inline insertions and
        deletions in the running prose.

        Pass ANSI escape sequences in place of the default square-bracket markers
        if you want styled terminal output.

        .. versionadded:: 2026.05.0
        """
        return "\n\n".join(
            p.revision_marks_text(open_ins, close_ins, open_del, close_del)
            for p in self.paragraphs
        )

    def append_document(self, other: Document) -> int:
        """Append the body of `other` to this document and return the number of
        block elements copied.

        Paragraphs, tables, and block-level SDT (structured document tag) elements
        from ``other`` are deep-copied into this document's body, inserted before
        the current section's sentinel ``w:sectPr`` so the destination's page
        setup is preserved. Relationships carried by the copied content —
        images, embedded objects, hyperlinks, charts, etc. — are imported into
        this document's package and rewritten to point at the new destination
        rIds. Paragraph / run styles referenced by the copied content are
        likewise copied (plus any ``basedOn`` / ``next`` / ``link`` dependencies).
        List numbering definitions referenced by the copied content are cloned
        under fresh numIds.

        Closes upstream#1457, upstream#558, upstream#543, upstream#437,
        upstream#460, upstream#44, upstream#709.

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_document

        return append_document(self, other)

    def append_body(self, other: Document) -> int:
        """Alias for :meth:`append_document`.

        Provided as a second entry-point for users who think of the operation
        as "append the body" rather than "append the document".

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_body

        return append_body(self, other)

    def append_paragraph(self, paragraph: Paragraph) -> Paragraph:
        """Copy `paragraph` from its owning document into this one and return the new paragraph.

        Any relationships referenced by the paragraph (images, hyperlinks,
        embedded objects) and any style / numbering references it carries are
        imported into this document the same way as for :meth:`append_document`.

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_paragraph

        return append_paragraph(self, paragraph)

    @property
    def is_strict(self) -> bool:
        """``True`` when this package was opened in ECMA-376 Strict mode.

        Delegates to :attr:`docx.opc.package.OpcPackage.is_strict`. A
        fresh authoring-path |Document| is Transitional by default.
        Assigning a value flips the flag so a subsequent :meth:`save`
        (with ``strict=None``) preserves the Strict class on emit.

        python-docx translates Strict → Transitional byte-level on
        open at the :class:`~docx.opc.pkgreader.PackageReader` layer,
        so the in-memory part tree carries Transitional URIs either
        way; the flag is surfaced for programmatic introspection and
        explicit round-trip preservation.

        .. versionadded:: 2026.05.11
        """
        return self._part.package.is_strict

    @is_strict.setter
    def is_strict(self, value: bool) -> None:
        self._part.package.is_strict = bool(value)

    def save(
        self,
        path_or_stream: str | IO[bytes],
        flat_opc: bool = False,
        reproducible: bool = False,
        password: str | None = None,
        strict: bool | None = None,
    ):
        """Save this document to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.

        When `path_or_stream` is a string, the filename component (the last path
        segment) is validated against the set of characters Windows disallows in
        file names (``< > : " | ? *``). If one of those characters is present,
        an :class:`OSError` is raised rather than writing a silently-empty or
        mis-named file (closes upstream#1111). The rest of the path — including
        drive-letter colons and forward/backward directory separators — is left
        to the underlying file system.

        When `flat_opc` is True, the document is serialised as Flat-OPC — the
        ``<pkg:package>`` single-XML-file representation defined in ECMA-376
        Part 2 — rather than a zip package. Closes upstream#892.

        When `reproducible` is True, the emitted zip archive uses a fixed
        timestamp for every member and writes members in sorted order, so
        repeated saves of the same content produce byte-identical output.
        This is the single bit of plumbing that closes upstream#1042 and
        upstream-PR#810.

        When `password` is provided, the saved ``.docx`` is password-protected
        using ECMA-376 Agile Encryption (the scheme Word uses when a user sets
        a password in the desktop app). Encryption requires the optional
        ``python-ooxml-crypto`` dependency. ``flat_opc`` and ``password`` are
        mutually exclusive: the Flat-OPC format is an XML document, not a zip
        archive, and is not password-protectable. ``reproducible`` and
        ``password`` compose normally — the fixed-timestamp zip members are
        produced first and the encryption wrapper is applied to that buffer.

        `strict` records the ECMA-376 conformance class on the package.
        ``None`` (default) preserves :attr:`is_strict`; ``True`` /
        ``False`` override it for this call. The bytes written to
        disk are always Transitional — python-docx does not currently
        perform Transitional → Strict byte-level translation on emit.

        .. versionadded:: 2026.05.0
           The `flat_opc` and `reproducible` parameters.
        .. versionadded:: 2026.05.10
           The `password` parameter.
        .. versionadded:: 2026.05.11
           The `strict` parameter.
        """
        # -- resolve smart-placeholder bind tokens (#68) immediately
        # -- before handing off to the part-level save so every text
        # -- run carrying ``{customer.name}`` / ``{date:short}`` /
        # -- ``{property:Title}`` etc. picks up the live bound record.
        # -- Best-effort by contract: a failure must not block save.
        try:
            from docx.bind_tokens import apply_bind_tokens as _apply_bind_tokens

            _apply_bind_tokens(self)
        except Exception:  # pragma: no cover - defensive guard
            pass

        if flat_opc:
            if password is not None:
                raise ValueError(
                    "flat_opc and password are mutually exclusive; the Flat-OPC "
                    "format is an XML document, not a zip archive, and is not "
                    "password-protectable."
                )
            import io as _io

            from docx.opc.flat_opc import write_flat_opc

            buf = _io.BytesIO()
            self._part.save(buf, reproducible=reproducible, strict=strict)
            write_flat_opc(path_or_stream, buf.getvalue())
            return
        self._part.save(
            path_or_stream, reproducible=reproducible, password=password,
            strict=strict,
        )

    def search(
        self,
        text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> list[SearchMatch]:
        """Find all occurrences of `text` in the document body paragraphs.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not included.

        Returns a list of |SearchMatch| objects, one for each occurrence found. Each match
        provides access to the paragraph, run indices, and character offsets.

        When `case_sensitive` is False, matching is case-insensitive. When `whole_word` is
        True, only whole-word matches are returned.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_paragraphs

        return search_paragraphs(self.paragraphs, text, case_sensitive, whole_word)

    def search_all(
        self,
        text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> list[SearchMatch]:
        """Find `text` in every story across this document.

        Unlike :meth:`search`, which only looks at top-level body paragraphs, this
        walks all document "stories" — the body and its top-level tables, each
        section's non-inherited headers and footers, footnotes, endnotes, and
        comments — and returns a |SearchMatch| for every hit. Each match's
        :attr:`SearchMatch.location` identifies which story produced it (e.g.
        ``"body"``, ``"table:0:row:1:col:2"``, ``"header:section0:primary"``,
        ``"footnote:2"``, ``"endnote:3"``, or ``"comment:5"``).

        Tables nested inside other stories (headers, footers, footnotes, etc.) and
        tables nested inside body-table cells are not descended into; see
        :func:`docx.search._iter_all_paragraphs` for details.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_all_paragraphs

        return search_all_paragraphs(self, text, case_sensitive, whole_word)

    def search_regex(
        self,
        pattern: str | re.Pattern[str],
        flags: int = 0,
    ) -> list[SearchMatch]:
        """Find all regex matches of `pattern` in the document body paragraphs.

        `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
        `flags` (e.g. `re.IGNORECASE`) is applied during compilation; when `pattern` is
        already compiled, `flags` is ignored.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not included.

        Returns a list of |SearchMatch| objects, one for each match found. Each match
        provides access to the paragraph, run indices, and character offsets.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_paragraphs_regex

        return search_paragraphs_regex(self.paragraphs, pattern, flags)

    def search_regex_all(
        self,
        pattern: str | re.Pattern[str],
        flags: int = 0,
    ) -> list[SearchMatch]:
        """Find regex matches of `pattern` in every story in this document.

        Like :meth:`search_all` but using regex semantics. ``pattern`` may be a
        string or a compiled :class:`re.Pattern`. Each returned |SearchMatch|
        carries a :attr:`SearchMatch.location` identifying the story that
        produced it.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_all_paragraphs_regex

        return search_all_paragraphs_regex(self, pattern, flags)

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    def outline(self) -> "Outline":
        """Return a hierarchical heading-tree snapshot of this document.

        Walks :attr:`paragraphs` once. Paragraphs styled ``Title`` map to
        outline level 0; ``Heading 1``..``Heading 9`` map to levels 1..9.
        Each heading becomes an :class:`~docx.outline.OutlineNode` carrying
        its ``text``, ``paragraph_index`` (position in
        :attr:`paragraphs`), a stable 8-char ``id``, the section's
        ``word_count`` (whitespace-token count of the heading and the
        body paragraphs that follow it up to the next same-or-shallower
        heading), and a list of nested ``children``.

        The wrapper :class:`~docx.outline.Outline` exposes
        ``walk()`` for depth-first traversal, ``to_dict()`` for
        JSON-serialisable output, and ``find(heading)`` for
        text-based lookup. ``Outline.title`` is sourced from the
        first ``Title``-styled paragraph, falling back to the
        document's core-properties title.

        ``Outline.total_pages_estimated`` reads the cached ``<Pages>``
        value from ``docProps/app.xml`` (Word's last-saved page count);
        python-docx has no layout engine so individual heading page
        numbers are intentionally **not** computed. Callers that need
        approximate page positions can fall back to the cached count
        for the whole document.

        Use this as a compact map of structure before mutating a long
        document — particularly useful for LLM agents that would
        otherwise need to ingest the full body.

        .. versionadded:: 2026.05.7
        """
        from docx.outline import build_outline

        return build_outline(self)

    def slice(
        self,
        start: "str | OutlineNode",
        end: "str | OutlineNode | None" = None,
    ) -> "Document":
        """Return a new |Document| containing one heading-bounded section.

        `start` is a heading's exact text (matched against
        :attr:`OutlineNode.text`) or an
        :class:`~docx.outline.OutlineNode` returned by
        :meth:`Outline.find` / :meth:`Outline.walk`. The slice runs
        from `start`'s paragraph (inclusive) to but not including
        `end`'s paragraph; when `end` is |None| the slice runs to the
        end of the document.

        The new document inherits python-docx's bundled default
        template; copied paragraphs are inserted via
        :meth:`append_paragraph`, which rewires image, hyperlink, and
        style references the same way :meth:`append_document` does.

        Raises :class:`ValueError` when `start` or `end` does not
        match any heading.

        .. versionadded:: 2026.05.7
        """
        from docx.outline import slice_document

        return slice_document(self, start, end)

    def validate_heading_structure(self) -> list[HeadingIssue]:
        """Return a list of |HeadingIssue| objects describing heading problems.

        Scans the body paragraphs for common accessibility issues in the heading
        outline: skipped levels (e.g. a Heading 3 directly following a Heading 1),
        multiple Heading 1 paragraphs, empty heading paragraphs, and documents that
        start below Heading 1. Returns an empty list when no issues are found.

        .. versionadded:: 2026.05.0
        """
        from docx.accessibility import validate_heading_structure

        return validate_heading_structure(self.paragraphs)

    def check_alt_text(self):
        """Return a list of ``(shape_or_table, issue_type)`` a11y issues.

        Flags two classes of problem:

        * ``"missing_alt_text"`` — an :class:`~docx.shape.InlineShape` whose
          ``wp:docPr/@descr`` is absent or empty, and which is *not* marked as
          ``decorative`` via :attr:`~docx.shape.InlineShape.a11y_role`. Purely
          decorative images are deliberately exempt: screen readers should
          skip them.
        * ``"missing_summary"`` — a :class:`~docx.table.Table` with no
          :attr:`~docx.table.Table.accessibility_summary`
          (``w:tblPr/w:tblDescription``).

        Results are returned in document order: all inline shapes first,
        then tables. Returns an empty list when no issues are found.

        .. versionadded:: 2026.05.0
        """
        issues: list[tuple[object, str]] = []
        for shape in self.inline_shapes:
            if shape.a11y_role == "decorative":
                continue
            alt = shape.alt_text
            if alt is None or not alt.strip():
                issues.append((shape, "missing_alt_text"))
        for table in self.tables:
            summary = table.accessibility_summary
            if summary is None or not summary.strip():
                issues.append((table, "missing_summary"))
        return issues

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the document-level settings."""
        return self._part.settings

    def protect(
        self,
        edit_mode: "WD_PROTECTION | None" = None,
        password: str | None = None,
        enforcement: bool = True,
    ) -> DocumentProtection:
        """Apply document protection in one call.

        Populates ``w:documentProtection`` with ``@w:edit=<edit_mode>`` and
        ``@w:enforcement=<enforcement>``. When `password` is given, the
        password is hashed with Word's legacy SHA-1 algorithm (using a fresh
        random salt and 100,000 iterations) and written to the ``@w:hash``,
        ``@w:salt``, and ``@w:crypt*`` attributes. `edit_mode` defaults to
        :attr:`WD_PROTECTION.READ_ONLY`.

        Returns the :class:`DocumentProtection` proxy for further tuning.

        .. versionadded:: 2026.05.10
        """
        from docx.enum.text import WD_PROTECTION

        mode = WD_PROTECTION.READ_ONLY if edit_mode is None else edit_mode
        return self.settings.enable_protection(
            mode=mode, enforce=enforcement, password=password
        )

    def unprotect(self) -> None:
        """Remove document protection, clearing ``w:documentProtection``.

        Equivalent to :meth:`Settings.disable_protection`. Leaves any
        ``w:writeProtection`` element untouched; call
        :meth:`Settings.disable_write_protection` separately to clear that.

        .. versionadded:: 2026.05.10
        """
        self.settings.disable_protection()

    def readability(self) -> "ReadabilityReport":
        """Return a |ReadabilityReport| of standard readability metrics.

        Computes Flesch Reading Ease, Flesch-Kincaid Grade, Gunning Fog,
        SMOG, Coleman-Liau, and Automated Readability Index for the
        whole body story, plus the underlying word, sentence, syllable,
        character, and complex-word counts. The report's ``sections``
        list partitions the body by ``Heading 1`` boundaries -- body
        text before the first ``Heading 1`` becomes a synthetic
        ``Introduction`` section. ``Title`` and ``Heading 2..9``
        paragraphs are folded into the surrounding section so the
        breakdown stays compact.

        Tokenisation uses a stdlib-only heuristic (no ``nltk`` /
        ``textstat`` dependency): vowel-group syllable counting,
        regex-based sentence splitting on ``[.!?]+``. Scores agree
        with published values to within a few percent for natural
        prose -- good enough for the Word "Readability Statistics"
        dialog use case the formulas were designed for.

        .. versionadded:: 2026.05.12
        """
        from docx.readability import build_report

        return build_report(self)

    @property
    def statistics(self) -> DocumentStatistics:
        """A |DocumentStatistics| summarizing the document body's text.

        Provides counts of non-empty paragraphs, words, characters (including
        spaces), and characters excluding spaces for the main document story
        (the ``w:body``). Text in headers, footers, footnotes, endnotes, and
        comments is not included, mirroring Word's default "Word Count"
        behavior.

        A "word" is a whitespace-delimited token, matching ``str.split()``
        semantics. Paragraphs nested inside tables or block-level content
        controls are included because they are part of the body story.

        :attr:`DocumentStatistics.pages` is populated from the cached
        ``<Pages>`` value in the extended-properties part (``docProps/app.xml``)
        when present, otherwise |None| -- python-docx does not lay the document
        out so it cannot compute a true page count.

        .. versionadded:: 2026.05.0
        """
        from docx.statistics import compute_statistics

        pages: int | None = None
        try:
            raw_pages = self.extended_properties.pages
        except Exception:
            # -- if the extended-properties part is missing or malformed,
            # -- fall back silently to a |None| page count rather than
            # -- propagating the error into the caller's statistics call.
            raw_pages = None
        # -- only accept genuine integer / None values; anything else (e.g.
        # -- a Mock surfaced by test doubles) is treated as "unknown" --
        if isinstance(raw_pages, int) and not isinstance(raw_pages, bool):
            pages = raw_pages
        else:
            pages = None

        return compute_statistics(self._element.body, pages=pages)

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in this document."""
        return self._part.styles

    @property
    def glossary(self) -> Glossary | None:
        """A |Glossary| proxy, or |None| when no ``glossaryDocument`` part is related.

        The glossary-document part carries the AutoText / Quick Parts /
        cover-page building blocks. Returns |None| when the document has no
        ``glossaryDocument`` relationship — which is the case for
        documents created via :func:`docx.Document` with the default
        template. To create a fresh glossary on demand and gain write
        access (`add_building_block` / `remove_building_block`), call
        :meth:`ensure_glossary` instead.

        .. versionadded:: 2026.05.0
        """
        return self._part.glossary

    def ensure_glossary(self) -> Glossary:
        """A writable |Glossary|, lazily creating the glossary part if needed.

        Returns the existing :class:`~docx.glossary.Glossary` when the
        document already has a ``glossaryDocument`` relationship. Otherwise
        a fresh, empty :class:`~docx.parts.glossary.GlossaryPart` is
        created, related to the document under the ``glossaryDocument``
        relationship type, and wrapped as a |Glossary|. Subsequent
        ``document.glossary`` accesses return the same proxy.

        .. versionadded:: 2026.05.10
        """
        return self._part.ensure_glossary()

    @property
    def glossary_document(self) -> Glossary | None:
        """The |GlossaryDocument| proxy, or |None| when no glossary part exists.

        Alias for :attr:`glossary` that matches the ECMA-376 vocabulary used
        by the R9-21 advanced API. The setter accepts a
        :class:`~docx.glossary.GlossaryDocument` value — assigning |None|
        drops the glossary-document relationship (removing the glossary
        part from the package); assigning an existing proxy is equivalent
        to :meth:`ensure_glossary` (the part is created when absent and
        left alone otherwise — the passed-in proxy's XML is **not** copied
        in at this pass, keeping the surface minimal).

        .. versionadded:: 2026.05.10
        """
        return self._part.glossary

    @glossary_document.setter
    def glossary_document(self, value: Glossary | None) -> None:
        if value is None:
            self._part.remove_glossary()
            return
        # -- lazily create when absent; ignore the passed proxy's XML --
        # -- (assignment semantics are "ensure a glossary exists") --
        self._part.ensure_glossary()

    @property
    def theme(self) -> Theme | None:
        """A |Theme| proxy, or |None| when no ``theme`` part is related.

        The theme part is owned by Word, so python-docx exposes it read-only.
        Returns |None| when the document has no ``theme`` relationship — which
        is uncommon for documents created by Word but possible for minimal
        documents synthesized by other tools.

        .. versionadded:: 2026.05.0
        """
        return self._part.theme

    @property
    def web_settings(self) -> WebSettings | None:
        """A |WebSettings| proxy, or |None| when no ``webSettings`` part is related.

        The web-settings part is owned by Word, so python-docx exposes it
        read-oriented. Returns |None| when the document has no ``webSettings``
        relationship — for example, documents created via :func:`docx.Document`
        with no template.

        .. versionadded:: 2026.05.0
        """
        return self._part.web_settings

    @property
    def tables(self) -> list[Table]:
        """All |Table| instances in the document, in document order.

        Note that only tables appearing at the top level of the document appear in this
        list; a table nested inside a table cell does not appear. A table within
        revision marks such as ``<w:ins>`` or ``<w:del>`` will also not appear in the
        list.
        """
        return self._body.tables

    @property
    def _block_width(self) -> Length:
        """A |Length| object specifying the space between margins in last section.

        Falls back to the US-Letter default (8.5" page width, 1" margins — a
        6.5" usable block) when the document body contains no ``w:sectPr``,
        which some third-party generators emit. See upstream#514.
        """
        sections = self.sections
        if len(sections) == 0:
            # -- no sectPr present: use standard US-Letter defaults --
            return Emu(Inches(8.5) - Inches(1) - Inches(1))
        section = sections[-1]
        page_width = section.page_width or Inches(8.5)
        left_margin = section.left_margin or Inches(1)
        right_margin = section.right_margin or Inches(1)
        return Emu(page_width - left_margin - right_margin)

    @property
    def _body(self) -> _Body:
        """The |_Body| instance containing the content for this document."""
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body


class _Body(BlockItemContainer):
    """Proxy for `<w:body>` element in this document.

    It's primary role is a container for document content.
    """

    def __init__(self, body_elm: CT_Body, parent: t.ProvidesStoryPart):
        super().__init__(body_elm, parent)
        self._body = body_elm

    def _iter_inner_content_flat_sdt(self) -> Iterator[Paragraph | Table]:
        """Yield `Paragraph`/`Table` walking into `w:sdt/w:sdtContent` wrappers.

        Block-level ``w:sdt`` elements have a ``w:sdtContent`` child holding
        the wrapped paragraphs and tables; descend into it so those items
        surface in the iteration. Non-block-level `w:sdt` elements (inline
        inside a paragraph) are not reached here because they are not direct
        children of the body.
        """
        from docx.oxml.ns import qn
        from docx.oxml.text.paragraph import CT_P as _CT_P
        from docx.oxml.table import CT_Tbl as _CT_Tbl
        from docx.table import Table as _Table
        from docx.text.paragraph import Paragraph as _Paragraph

        for child in self._body:
            tag = child.tag
            if isinstance(child, _CT_P):
                yield _Paragraph(child, self)
            elif isinstance(child, _CT_Tbl):
                yield _Table(child, self)
            elif tag == qn("w:sdt"):
                for sub in child:
                    if sub.tag != qn("w:sdtContent"):
                        continue
                    for inner in sub:
                        if isinstance(inner, _CT_P):
                            yield _Paragraph(inner, self)
                        elif isinstance(inner, _CT_Tbl):
                            yield _Table(inner, self)

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Add a block-level content control at the end of the body.

        The new `w:sdt` is inserted before any trailing `w:sectPr` element, mirroring
        how paragraphs and tables are appended.

        .. versionadded:: 2026.05.0
        """
        from docx.content_controls import ContentControl, new_sdt

        sdt = new_sdt(type, tag=tag, title=title, inline=False)
        # -- insert before trailing sectPr, if present --
        self._body._insert_sdt(sdt)  # pyright: ignore[reportPrivateUsage]
        return ContentControl.proxy_for(sdt)

    def clear_content(self) -> _Body:
        """Return this |_Body| instance after clearing it of all content.

        Section properties for the main document story, if present, are preserved.
        """
        self._body.clear_content()
        return self

    @property
    def content_controls(self) -> list[ContentControl]:
        """List of block-level |ContentControl| objects in this body, in document order.

        .. versionadded:: 2026.05.0
        """
        from docx.content_controls import ContentControl

        return [
            ContentControl.proxy_for(cast("CT_Sdt", sdt))
            for sdt in self._body.xpath("./w:sdt")
        ]
