# pyright: reportImportCycles=false
# pyright: reportPrivateUsage=false

"""|Document| and closely related objects."""

from __future__ import annotations

import re
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator, Sequence

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_SECTION
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.section import Section, Sections
from docx.shared import ElementProxy, Emu, Inches, Length, RGBColor
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.accessibility import HeadingIssue
    from docx.bookmarks import Bookmarks
    from docx.comments import Comment, Comments
    from docx.content_controls import ContentControl, ContentControlType
    from docx.custom_properties import CustomProperties
    from docx.endnotes import Endnotes, EndnoteProperties
    from docx.font_table import FontTable
    from docx.footnotes import FootnoteProperties, Footnotes
    from docx.ink import InkAnnotation
    from docx.oxml.content_controls import CT_Sdt
    from docx.oxml.document import CT_Body, CT_Document
    from docx.parts.document import DocumentPart
    from docx.search import SearchMatch
    from docx.settings import Settings
    from docx.signatures import SignatureInfo
    from docx.statistics import DocumentStatistics
    from docx.styles.style import ParagraphStyle, _TableStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph
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

    def add_comment(
        self,
        runs: Run | Sequence[Run],
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
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
        """
        # -- normalize `runs` to a sequence of runs --
        runs = [runs] if isinstance(runs, Run) else runs
        first_run = runs[0]
        last_run = runs[-1]

        # -- Note that comments can only appear in the document part --
        comment = self.comments.add_comment(text=text, author=author, initials=initials)

        # -- let the first run orchestrate placement of the comment range start and end --
        first_run.mark_comment_range(last_run, comment.comment_id)

        return comment

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

    def add_page_break(self) -> Paragraph:
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_page_break()
        return paragraph

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the document.

        The paragraph is populated with `text` and having paragraph style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break.
        """
        return self._body.add_paragraph(text, style)

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
        """
        return self._body.add_content_control(type, tag=tag, title=title)

    def add_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
    ):
        """Return new picture shape added in its own paragraph at end of the document.

        The picture contains the image at `image_path_or_stream`, scaled based on
        `width` and `height`. If neither width nor height is specified, the picture
        appears at its native size. If only one is specified, it is used to compute a
        scaling factor that is then applied to the unspecified dimension, preserving the
        aspect ratio of the image. The native size of the picture is calculated using
        the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
        if no value is specified, as is often the case.
        """
        run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    def add_section(self, start_type: WD_SECTION = WD_SECTION.NEW_PAGE):
        """Return a |Section| object newly added at the end of the document.

        The optional `start_type` argument must be a member of the :ref:`WdSectionStart`
        enumeration, and defaults to ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

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
        """
        from docx.toc import populate_toc_paragraph

        # -- snapshot the current headings before appending the new paragraph
        #    so the empty TOC paragraph doesn't self-include if its style
        #    ever matched the heading regex. --
        source_paragraphs = list(self.paragraphs)
        paragraph = self.add_paragraph()
        return populate_toc_paragraph(paragraph, source_paragraphs, levels)

    def add_table(self, rows: int, cols: int, style: str | _TableStyle | None = None):
        """Add a table having row and column counts of `rows` and `cols` respectively.

        `style` may be a table style object or a table style name. If `style` is |None|,
        the table inherits the default table style of the document.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        table.style = style
        return table

    @property
    def background_color(self) -> RGBColor | None:
        """Document-wide page background color, or |None| if not set.

        Maps to the ``w:color`` attribute on the ``w:background`` child of the
        ``w:document`` root element. Assigning an |RGBColor| writes (or updates)
        the ``w:background`` element. Assigning |None| removes the element.
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
    def bookmarks(self) -> Bookmarks:
        """A |Bookmarks| object providing access to the bookmarks in this document."""
        from docx.bookmarks import Bookmarks

        return Bookmarks(self._element.body)

    @property
    def comments(self) -> Comments:
        """A |Comments| object providing access to comments added to the document."""
        return self._part.comments

    @property
    def content_controls(self) -> list[ContentControl]:
        """All block-level |ContentControl| objects in this document body, in order.

        Only block-level content controls (direct children of `w:body`) are returned.
        Inline content controls are accessible via :attr:`Paragraph.content_controls`.
        """
        return self._body.content_controls

    @property
    def endnotes(self) -> Endnotes:
        """A |Endnotes| object providing access to endnotes in the document."""
        return self._part.endnotes

    @property
    def has_macros(self) -> bool:
        """True if this document contains a VBA project (macros)."""
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
        """
        from docx.package import Package

        return cast("Package", self._part.package).is_signed

    @property
    def signatures(self) -> list[SignatureInfo]:
        """List of |SignatureInfo| for each digital signature in the package.

        Empty list when the document is unsigned. See :class:`docx.signatures.SignatureInfo`
        for the available metadata.
        """
        from docx.package import Package

        return cast("Package", self._part.package).signatures

    @property
    def font_table(self) -> FontTable | None:
        """A |FontTable| collection, or |None| if no font-table part is related.

        The font-table part is owned by Word, so python-docx exposes it read-only.
        Returns |None| when the document has no ``fontTable`` relationship — for
        example, when the document was created via :func:`docx.Document` with no
        template.
        """
        return self._part.font_table

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| object providing access to footnotes in the document."""
        return self._part.footnotes

    def accept_all_changes(self) -> int:
        """Accept every tracked change in the document body.

        Insertions are flattened into live content, deletions are removed, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are discarded
        (the current, post-edit formatting is retained).

        Returns the number of change elements resolved.
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=True)

    def reject_all_changes(self) -> int:
        """Reject every tracked change in the document body.

        Insertions are removed, deletions are restored as live content, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are unwound so
        the prior formatting is restored.

        Returns the number of change elements resolved.
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=False)

    @property
    def footnote_properties(self) -> FootnoteProperties | None:
        """Document-level |FootnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:footnotePr`` element exists in the document settings.
        Use :meth:`add_footnote_properties` to add one and configure it.
        """
        return self.settings.footnote_properties

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return document-level |FootnoteProperties|, adding a ``w:footnotePr`` if needed."""
        return self.settings.add_footnote_properties()

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """Document-level |EndnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:endnotePr`` element exists in the document settings.
        Use :meth:`add_endnote_properties` to add one and configure it.
        """
        return self.settings.endnote_properties

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return document-level |EndnoteProperties|, adding a ``w:endnotePr`` if needed."""
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
        """
        return self._part.custom_properties

    @property
    def numbering(self):
        """A |Numbering| object providing read/write access to the list-style
        numbering definitions for this document.

        Creates a default (empty) numbering part if one is not already related to the
        document.
        """
        return self._part.numbering_part.numbering

    @property
    def ink_annotations(self) -> list[InkAnnotation]:
        """List of |InkAnnotation| objects for each ink annotation in the body.

        An ink annotation is any ``w:contentPart`` element that targets an ink part
        (content type ``application/inkml+xml``). The list is empty when no ink
        annotations are present. Read-only — python-docx does not support creating
        or modifying ink annotations.
        """
        from docx.text.paragraph import Paragraph

        result: list[InkAnnotation] = []
        for p in self._element.body.xpath(".//w:p[.//w:contentPart]"):
            paragraph = Paragraph(p, self._body)
            result.extend(paragraph.ink_annotations)
        return result

    @property
    def inline_shapes(self):
        """The |InlineShapes| collection for this document.

        An inline shape is a graphical object, such as a picture, contained in a run of
        text and behaving like a character glyph, being flowed like other text in a
        paragraph.
        """
        return self._part.inline_shapes

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this document in document order."""
        return self._body.iter_inner_content()

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
            if field.type not in ("REF", "PAGEREF"):
                continue
            resolved = field.resolve(self)
            if resolved == field.result_text:
                continue
            field.update_result_text(resolved)
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
        """
        return "\n\n".join(
            p.revision_marks_text(open_ins, close_ins, open_del, close_del)
            for p in self.paragraphs
        )

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this document to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self._part.save(path_or_stream)

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
        """
        from docx.search import search_all_paragraphs_regex

        return search_all_paragraphs_regex(self, pattern, flags)

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    def validate_heading_structure(self) -> list[HeadingIssue]:
        """Return a list of |HeadingIssue| objects describing heading problems.

        Scans the body paragraphs for common accessibility issues in the heading
        outline: skipped levels (e.g. a Heading 3 directly following a Heading 1),
        multiple Heading 1 paragraphs, empty heading paragraphs, and documents that
        start below Heading 1. Returns an empty list when no issues are found.
        """
        from docx.accessibility import validate_heading_structure

        return validate_heading_structure(self.paragraphs)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the document-level settings."""
        return self._part.settings

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
        """
        from docx.statistics import compute_statistics

        return compute_statistics(self._element.body)

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in this document."""
        return self._part.styles

    @property
    def web_settings(self) -> WebSettings | None:
        """A |WebSettings| proxy, or |None| when no ``webSettings`` part is related.

        The web-settings part is owned by Word, so python-docx exposes it
        read-oriented. Returns |None| when the document has no ``webSettings``
        relationship — for example, documents created via :func:`docx.Document`
        with no template.
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
        """A |Length| object specifying the space between margins in last section."""
        section = self.sections[-1]
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

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Add a block-level content control at the end of the body.

        The new `w:sdt` is inserted before any trailing `w:sectPr` element, mirroring
        how paragraphs and tables are appended.
        """
        from docx.content_controls import ContentControl, new_sdt

        sdt = new_sdt(type, tag=tag, title=title, inline=False)
        # -- insert before trailing sectPr, if present --
        self._body._insert_sdt(sdt)  # pyright: ignore[reportPrivateUsage]
        return ContentControl(sdt)

    def clear_content(self) -> _Body:
        """Return this |_Body| instance after clearing it of all content.

        Section properties for the main document story, if present, are preserved.
        """
        self._body.clear_content()
        return self

    @property
    def content_controls(self) -> list[ContentControl]:
        """List of block-level |ContentControl| objects in this body, in document order."""
        from docx.content_controls import ContentControl

        return [
            ContentControl(cast("CT_Sdt", sdt)) for sdt in self._body.xpath("./w:sdt")
        ]
