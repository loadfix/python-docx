"""Paragraph-related proxy types."""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator

from docx.drawing import Drawing
from docx.enum.section import WD_SECTION_START
from docx.enum.shape import WD_ANCHOR_H, WD_ANCHOR_V, WD_SHAPE, WD_WRAP_TYPE
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.fields import Field
from docx.form_fields import FormField
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.drawing import CT_Drawing
from docx.oxml.shape import CT_Anchor
from docx.oxml.table import CT_Tbl
from docx.oxml.text.run import CT_R
from docx.shape import FloatingImage
from docx.shared import Inches, StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.tracked_changes import MoveRevision, TrackedChange
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.bookmarks import Bookmark
    from docx.content_controls import ContentControl, ContentControlType
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.embedded_objects import EmbeddedObject
    from docx.equations import Equation
    from docx.ink import InkAnnotation
    from docx.oxml.content_controls import CT_Sdt
    from docx.oxml.document import CT_Body
    from docx.oxml.math import CT_OMath, CT_OMathPara
    from docx.oxml.text.paragraph import CT_P
    from docx.permissions import PermissionRange
    from docx.section import Section
    from docx.shared import Length
    from docx.styles.style import CharacterStyle
    from docx.table import Table as _Table
    from docx.styles.style import _TableStyle  # pyright: ignore[reportPrivateUsage]


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._p = self._element = p

    def add_bookmark(
        self,
        name: str,
        start_run: Run | None = None,
        end_run: Run | None = None,
    ) -> Bookmark:
        """Add a bookmark to this paragraph and return it.

        `name` is the bookmark name, which must be unique within the document.

        When `start_run` and `end_run` are both |None|, the bookmark wraps the entire
        paragraph content. When `start_run` is provided, the bookmark starts before that
        run. When `end_run` is provided, the bookmark ends after that run. When only
        `start_run` is provided, `end_run` defaults to `start_run`.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.bookmarks import Bookmark

        body = self._get_body()
        bookmark_id = self._next_bookmark_id(body)

        if start_run is None and end_run is None:
            self._p.add_bookmark(bookmark_id, name)
        else:
            if start_run is None:
                start_run = end_run
            if end_run is None:
                end_run = start_run
            assert start_run is not None
            assert end_run is not None
            start_run._r.insert_bookmark_start_before(bookmark_id, name)
            end_run._r.insert_bookmark_end_after(bookmark_id)

        bookmarkStart = self._p.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")
        return Bookmark(bookmarkStart[0], body)

    def add_permission_range(
        self,
        name: str | None = None,
        user: str | None = None,
        edit_group: str | None = None,
    ) -> PermissionRange:
        """Add a permission range wrapping this paragraph and return it.

        `user` is the single-user restriction (`w:ed`), and `edit_group`
        is a group restriction (`w:edGrp`, e.g. ``"everyone"`` or
        ``"current"``). At least one should typically be supplied.

        `name` is accepted for symmetry with ``add_bookmark()`` but is not
        persisted on the element — `w:permStart` has no `@w:name` attribute in
        OOXML; it is kept in the signature purely for call-site readability.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.permissions import PermissionRange
        from docx.oxml.permissions import CT_PermStart

        body = self._get_body()
        perm_id = self._next_permission_range_id(body)

        self._p.add_permission_range(perm_id, edit_group=edit_group, user=user)

        permStart = self._p.xpath(f".//w:permStart[@w:id='{perm_id}']")[0]
        return PermissionRange(cast(CT_PermStart, permStart), body)

    @property
    def permission_ranges(self) -> list[PermissionRange]:
        """List of |PermissionRange| objects rooted at `w:permStart` in this paragraph.


.. versionadded:: 1.3.0.dev0

"""
        from docx.permissions import PermissionRange
        from docx.oxml.permissions import CT_PermStart

        body = self._get_body()
        return [
            PermissionRange(cast(CT_PermStart, ps), body)
            for ps in self._p.xpath(".//w:permStart")
        ]

    @staticmethod
    def _next_permission_range_id(body) -> int:
        """Return the next available `w:permStart/@w:id` in the document body."""
        used_ids = [int(x) for x in body.xpath(".//w:permStart/@w:id")]
        return max(used_ids, default=-1) + 1

    def _get_body(self) -> CT_Body:
        """Return the w:body ancestor element."""
        from docx.oxml.document import CT_Body

        ancestor = self._p.getparent()
        while ancestor is not None and not isinstance(ancestor, CT_Body):
            ancestor = ancestor.getparent()
        if ancestor is None:
            raise ValueError("paragraph is not contained in a document body")
        return ancestor

    @staticmethod
    def _next_bookmark_id(body) -> int:
        """Return the next available bookmark ID in the document body."""
        used_ids = [int(x) for x in body.xpath(".//w:bookmarkStart/@w:id")]
        return max(used_ids, default=-1) + 1

    def add_hyperlink(
        self,
        url: str | None = None,
        text: str | None = None,
        style: str | CharacterStyle | None = "Hyperlink",
        anchor: str | None = None,
    ) -> Hyperlink:
        """Append a hyperlink to this paragraph and return a |Hyperlink| object.

        `url` is the target URL for an external hyperlink (e.g. "https://example.com").
        `text` is the visible link text; defaults to `url` or `anchor` when not provided.
        `style` is the character style for the hyperlink run, defaulting to "Hyperlink".
        `anchor` is a bookmark name for an internal document link.

        Either `url` or `anchor` must be provided, but not both.

        .. versionadded:: 1.3.0.dev0
        """
        if url is None and anchor is None:
            raise ValueError("Either url or anchor must be provided")
        if url is not None and anchor is not None:
            raise ValueError("Only one of url or anchor may be provided, not both")

        display_text = text if text is not None else (url or anchor or "")

        rId = None
        if url is not None:
            rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)

        rPr = None
        if style is not None:
            from docx.oxml.ns import qn
            from docx.oxml.parser import OxmlElement

            style_id = self.part.get_style_id(style, WD_STYLE_TYPE.CHARACTER)
            if style_id is not None:
                rPr = OxmlElement("w:rPr")
                rStyle = OxmlElement("w:rStyle")
                rStyle.set(qn("w:val"), style_id)
                rPr.append(rStyle)

        hyperlink_elm = self._p.add_hyperlink(rId, anchor, display_text, rPr)
        return Hyperlink(hyperlink_elm, self)

    def add_run(
        self,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        track_author: str | None = None,
    ) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.

        If `track_author` is supplied (or if an enclosing
        :meth:`Document.tracked_changes` context is active), the new run is
        wrapped in a `w:ins` tracked-insertion marker attributed to that
        author. Closes upstream#1025.

        .. versionadded:: 1.3.0.dev0
           Added ``track_author`` keyword argument.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        _maybe_wrap_tracked_run(r, track_author, self)
        return run

    def add_text(self, text: str) -> Run:
        """Append `text` to the last run of this paragraph, or create a new run.

        When the paragraph already contains at least one run, a ``w:t`` element
        containing `text` is appended to that last run. The run's existing
        character formatting (``w:rPr``) is preserved. When the paragraph has no
        runs, a new run is created and `text` assigned to it. Returns the run
        that now holds the appended text.

        Unlike :meth:`add_run`, this method does not split ``\\t``, ``\\n`` or
        ``\\r`` characters into separate elements — the entire `text` is placed
        in a single ``w:t`` element, with ``xml:space="preserve"`` applied if
        the text has leading or trailing whitespace.

        .. versionadded:: 1.3.0.dev0
        """
        runs = self._p.xpath("./w:r")
        if runs:
            r = runs[-1]
            r.add_t(text)
            return Run(cast(CT_R, r), self)
        # -- no existing run; create one and set its text --
        return self.add_run(text)

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Append an inline content control (structured document tag) to this paragraph.

        `type` is a :class:`ContentControlType` member. `tag` becomes the programmatic
        `w:sdtPr/w:tag/@w:val` value, and `title` becomes `w:sdtPr/w:alias/@w:val`.
        Returns the newly appended |ContentControl|.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.content_controls import ContentControl, new_sdt

        sdt = new_sdt(type, tag=tag, title=title, inline=True)
        self._p.append(sdt)
        return ContentControl(sdt)

    def add_page_break(self) -> Paragraph:
        """Append a page-break run to this paragraph and return self.

        .. versionadded:: 1.3.0.dev0
        """
        run = self.add_run()
        run.add_break(WD_BREAK.PAGE)
        return self

    def add_simple_field(self, instr: str, text: str | None = None) -> Field:
        """Append a ``<w:fldSimple>`` field to this paragraph and return a |Field|.

        `instr` is the field instruction (e.g. ``"PAGE"`` or ``"REF bookmark1 \\h"``).
        `text` is the optional current rendered result, added as a single run
        inside the fldSimple element.

        .. versionadded:: 1.3.0.dev0
        """
        fldSimple = self._p.add_fldSimple(instr, text)
        return Field.for_simple(fldSimple)

    def add_complex_field(self, instr: str, result_text: str | None = None) -> Field:
        """Append a complex field (begin/separate/end) to this paragraph.

        Returns a |Field| wrapping the run that contains the ``begin``
        ``<w:fldChar>`` marker. `instr` is the field instruction (e.g.
        ``"PAGE"``) and `result_text`, if provided, is added as a plain
        ``<w:r><w:t>`` run between the ``separate`` and ``end`` markers.

        .. versionadded:: 1.3.0.dev0
        """
        begin_run = self._p.add_complex_field(instr, result_text)
        return Field.for_complex(begin_run)

    def add_text_form_field(
        self,
        name: str,
        default: str = "",
        maxlength: int | None = None,
    ) -> FormField:
        """Append a legacy text form field (``FORMTEXT``) and return it.

        `name` becomes the form field's ``w:name/@w:val`` — the programmatic
        identifier used by Word macros and REF fields to retrieve the value.
        `default` is the initial value; it is written both to
        ``w:textInput/w:default`` and used as the rendered result text so
        Word displays it immediately without a field update. `maxlength` is
        the character limit (|None| means no limit).

        .. versionadded:: 1.3.0.dev0
        """
        from docx.form_fields import _append_form_field, new_text_form_field_ffData

        ffData = new_text_form_field_ffData(name, default=default, maxlength=maxlength)
        begin_run = _append_form_field(
            self._p, " FORMTEXT ", ffData, result_text=default
        )
        return FormField(begin_run)

    def add_checkbox_form_field(
        self,
        name: str,
        checked: bool = False,
    ) -> FormField:
        """Append a legacy checkbox form field (``FORMCHECKBOX``) and return it.

        `name` becomes the form field's ``w:name/@w:val``. `checked` sets both
        the default and current checked state. The rendered result region of
        the complex field is left empty — Word shows a checkbox glyph for
        ``FORMCHECKBOX`` regardless of the result runs.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.form_fields import _append_form_field, new_checkbox_form_field_ffData

        ffData = new_checkbox_form_field_ffData(name, checked=checked)
        begin_run = _append_form_field(self._p, " FORMCHECKBOX ", ffData, result_text="")
        return FormField(begin_run)

    def add_dropdown_form_field(
        self,
        name: str,
        options: list[str],
        default_index: int = 0,
    ) -> FormField:
        """Append a legacy dropdown form field (``FORMDROPDOWN``) and return it.

        `name` becomes the form field's ``w:name/@w:val``. `options` are the
        list entries the dropdown offers, in display order. `default_index` is
        the 0-based index of the option that is initially selected; it is
        written to both ``w:default`` and ``w:result``. The rendered result
        text is set to the option at `default_index` when that index is in
        range, so Word displays the initial value immediately.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.form_fields import _append_form_field, new_dropdown_form_field_ffData

        ffData = new_dropdown_form_field_ffData(
            name, options=options, default_index=default_index
        )
        initial_text = (
            options[default_index]
            if 0 <= default_index < len(options)
            else ""
        )
        begin_run = _append_form_field(
            self._p, " FORMDROPDOWN ", ffData, result_text=initial_text
        )
        return FormField(begin_run)

    def add_equation(
        self, omml_xml: str | bytes, display_mode: bool = False
    ) -> Equation:
        """Append an OMML equation to this paragraph and return the |Equation|.

        `omml_xml` is an OMML XML string (or bytes) whose root element is
        either ``m:oMath`` or ``m:oMathPara``. Namespace declarations for the
        ``m`` prefix must be present on the root. When `display_mode` is
        |True| and the root is a bare ``m:oMath``, it is wrapped in
        ``m:oMathPara`` to render in display mode.

        Raises :class:`ValueError` when the root element is neither
        ``m:oMath`` nor ``m:oMathPara``.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.equations import Equation, _make_equation_element

        element = _make_equation_element(omml_xml, display_mode=display_mode)
        self._p.append(element)
        return Equation(cast("CT_OMath | CT_OMathPara", element))

    @property
    def equations(self) -> list[Equation]:
        """List of |Equation| objects for each OMML expression in this paragraph.

        Includes both paragraph-level ``m:oMathPara`` wrappers and loose
        inline ``m:oMath`` descendants. Each ``m:oMath`` nested inside an
        ``m:oMathPara`` is represented once — by the enclosing
        ``m:oMathPara`` — so an equation is not counted twice.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.equations import Equation

        result: list[Equation] = []
        # -- top-level (or oMathPara-wrapped) matches first --
        for el in self._p.xpath(
            ".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]"
        ):
            result.append(Equation(cast("CT_OMath | CT_OMathPara", el)))
        return result

    def add_shape(
        self,
        shape_type,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append an inline `wps:wsp` DrawingML shape to this paragraph.

        `shape_type` is a :class:`docx.enum.shape.WD_SHAPE` member identifying
        the preset geometry. `width` and `height` are |Length| values; they
        default to 2 inches by 1 inch when omitted. When `text` is provided the
        shape gets a minimal text frame containing that string.

        Returns a :class:`docx.drawing.WordprocessingShape` proxy for the newly
        created shape.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.drawing import WordprocessingShape
        from docx.oxml.drawing import new_inline_shape_drawing

        if not isinstance(shape_type, WD_SHAPE):
            raise TypeError(
                "shape_type must be a WD_SHAPE member, got %r" % (shape_type,)
            )

        cx = int(width) if width is not None else int(Inches(2))
        cy = int(height) if height is not None else int(Inches(1))

        story_part = self.part
        shape_id = story_part.next_id
        name = "%s %d" % (_shape_name_for(shape_type), shape_id)

        drawing = new_inline_shape_drawing(
            shape_type.value, cx, cy, shape_id, name, text=text
        )

        run = self.add_run()
        run._r.append(drawing)

        wsp = drawing.xpath(
            ".//wp:inline/a:graphic/a:graphicData/wps:wsp"
        )[0]
        return WordprocessingShape(wsp, self)

    def add_floating_shape(
        self,
        image_path_or_stream: str | IO[bytes],
        x: int | Length = 0,
        y: int | Length = 0,
        width: int | Length | None = None,
        height: int | Length | None = None,
        h_anchor: WD_ANCHOR_H | str = WD_ANCHOR_H.COLUMN,
        v_anchor: WD_ANCHOR_V | str = WD_ANCHOR_V.PARAGRAPH,
        wrap: WD_WRAP_TYPE | str = WD_WRAP_TYPE.SQUARE,
    ) -> FloatingImage:
        """Add a floating image anchored at explicit coordinates and return it.

        `x` / `y` are horizontal / vertical offsets (EMU, or |Length|).
        `h_anchor` / `v_anchor` are the horizontal / vertical frame of
        reference; accepted as |WD_ANCHOR_H| / |WD_ANCHOR_V| members or the
        matching OOXML attribute strings (e.g. ``"page"``). `wrap` is the
        text-wrap style as a |WD_WRAP_TYPE| member or its string value.

        This is a coordinate-first counterpart to :meth:`add_floating_image`:
        use this method when you want to place a shape at a specific x/y
        offset (upstream #1414) rather than fall back to square-wrap with a
        zero offset.

        .. versionadded:: 1.3.0.dev0
        """
        return self.add_floating_image(
            image_path_or_stream,
            width=width,
            height=height,
            position={
                "h_anchor": h_anchor,
                "v_anchor": v_anchor,
                "horizontal": int(x),
                "vertical": int(y),
                "wrap": wrap,
            },
        )

    def add_floating_image(
        self,
        image_path_or_stream: "str | os.PathLike[str] | IO[bytes]",
        width: int | Length | None = None,
        height: int | Length | None = None,
        position: dict | None = None,
    ) -> FloatingImage:
        """Add a floating (anchored) image to this paragraph and return it.

        `image_path_or_stream` is a ``str`` path, an :class:`os.PathLike` (e.g.
        :class:`pathlib.Path`), or a binary file-like object for the image.
        `width` and `height` work the same way as for `add_picture`.

        `position` is an optional dict that may contain any of these keys:
        - `horizontal`: horizontal offset (int EMU or |Length|)
        - `vertical`: vertical offset (int EMU or |Length|)
        - `h_anchor`: |WD_ANCHOR_H| member (defaults to `COLUMN`)
        - `v_anchor`: |WD_ANCHOR_V| member (defaults to `PARAGRAPH`)
        - `wrap`: |WD_WRAP_TYPE| member (defaults to `SQUARE`)

        .. versionadded:: 1.3.0.dev0

        .. versionchanged:: 1.3.0.dev0
           Accepts :class:`os.PathLike` path arguments.
        """
        if isinstance(image_path_or_stream, os.PathLike):
            image_path_or_stream = os.fspath(image_path_or_stream)
        anchor = self.part.new_pic_anchor(image_path_or_stream, width, height)

        # -- apply optional positioning overrides --
        if position is not None:
            h_anchor = position.get("h_anchor", WD_ANCHOR_H.COLUMN)
            v_anchor = position.get("v_anchor", WD_ANCHOR_V.PARAGRAPH)
            horizontal = position.get("horizontal", 0)
            vertical = position.get("vertical", 0)
            wrap = position.get("wrap", WD_WRAP_TYPE.SQUARE)

            if isinstance(h_anchor, WD_ANCHOR_H):
                h_anchor_value = h_anchor.value
            else:
                h_anchor_value = str(h_anchor)
            if isinstance(v_anchor, WD_ANCHOR_V):
                v_anchor_value = v_anchor.value
            else:
                v_anchor_value = str(v_anchor)
            if isinstance(wrap, WD_WRAP_TYPE):
                wrap_value = wrap.value
            else:
                wrap_value = str(wrap)

            anchor.set_horizontal_position(h_anchor_value, int(horizontal))
            anchor.set_vertical_position(v_anchor_value, int(vertical))
            anchor.set_wrap(wrap_value)

        # -- append the anchor inside a new run's `w:drawing` --
        run = self.add_run()
        run._r.add_drawing(anchor)
        return FloatingImage(anchor)

    @property
    def fields(self) -> list[Field]:
        """List of |Field| objects for each field in this paragraph.

        Includes both simple (``w:fldSimple``) and complex (``w:fldChar``)
        fields, in document order.

        .. versionadded:: 1.3.0.dev0
        """
        result: list[Field] = []
        for kind, el in self._p.iter_field_elements():
            if kind == "simple":
                result.append(Field.for_simple(el))
            else:
                result.append(Field.for_complex(el))
        return result

    @property
    def form_fields(self) -> list[FormField]:
        """List of |FormField| objects for each legacy form field in this paragraph.

        A legacy form field is a complex field whose ``begin`` ``w:fldChar``
        carries a ``w:ffData`` child. Returned in document order. Complex
        fields without ``w:ffData`` (e.g. ``PAGE``, ``REF``) are ignored —
        those remain accessible via :attr:`fields`.

        .. versionadded:: 1.3.0.dev0
        """
        result: list[FormField] = []
        begin_runs = self._p.xpath(
            "./w:r[w:fldChar[@w:fldCharType='begin' and w:ffData]]"
        )
        for r in begin_runs:
            result.append(FormField(cast(CT_R, r)))
        return result

    @property
    def floating_images(self) -> list[FloatingImage]:
        """A |FloatingImage| instance for each `wp:anchor` in this paragraph.

    .. versionadded:: 1.3.0.dev0
    """
        return [
            FloatingImage(cast(CT_Anchor, a))
            for a in self._p.xpath(".//w:r/w:drawing/wp:anchor")
        ]

    @property
    def content_controls(self) -> list[ContentControl]:
        """List of inline |ContentControl| objects in this paragraph, in document order.

    .. versionadded:: 1.3.0.dev0
    """
        from docx.content_controls import ContentControl

        return [
            ContentControl(cast("CT_Sdt", sdt)) for sdt in self._p.xpath("./w:sdt")
        ]

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def delete(self) -> None:
        """Remove this paragraph from the document.

        The paragraph element is removed from its parent. After calling this method,
        this |Paragraph| object is "defunct" and should not be used further.

        .. versionadded:: 1.3.0.dev0
        """
        p = self._p
        parent = p.getparent()
        if parent is None:
            return
        parent.remove(p)

    def clear_page_breaks(self) -> None:
        """Remove all ``<w:br w:type="page"/>`` elements from this paragraph.

        If a run contains only a page break and no other content, the entire run is
        removed. If a run contains other content alongside the page break, only the
        ``<w:br>`` element is removed. Does nothing when no page breaks are present.

        .. versionadded:: 1.3.0.dev0
        """
        for br in self._p.xpath('.//w:br[@w:type="page"]'):
            r = br.getparent()
            r.remove(br)
            # --- remove the run if it's now empty (no child elements and no text) ---
            if len(r) == 0 and not r.text:
                r.getparent().remove(r)

    @property
    def has_section_break(self) -> bool:
        """``True`` if this paragraph contains a section break (``<w:sectPr>`` in its
        ``<w:pPr>``).

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._p.pPr
        if pPr is None:
            return False
        return pPr.sectPr is not None

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def has_page_break(self) -> bool:
        """`True` if this paragraph contains at least one ``<w:br w:type="page"/>``.

        .. versionadded:: 1.3.0.dev0
        """
        return bool(self._p.xpath('.//w:br[@w:type="page"]'))

    @property
    def drawings(self) -> list[Drawing]:
        """A |Drawing| instance for each `<w:drawing>` element in this paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        return [
            Drawing(cast(CT_Drawing, d), self)
            for d in self._p.xpath(".//w:drawing")
        ]

    @property
    def ink_annotations(self) -> list[InkAnnotation]:
        """List of |InkAnnotation| objects for each ``w:contentPart`` in this paragraph.

        Returns an empty list when the paragraph contains no ink annotations. Read-only
        — python-docx does not support creating or modifying ink annotations.

        A ``w:contentPart`` whose relationship cannot be resolved (for example because
        the referenced part is missing from the package) is silently skipped rather
        than raising.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ink import InkAnnotation
        from docx.oxml.ns import qn
        from docx.parts.ink import InkPart

        result: list[InkAnnotation] = []
        part = self.part
        for cp in self._p.xpath(".//w:contentPart"):
            rId = cp.get(qn("r:id"))
            if not rId:
                continue
            try:
                ink_part = part.related_parts[rId]
            except KeyError:
                continue
            if not isinstance(ink_part, InkPart):
                continue
            result.append(InkAnnotation(self, ink_part))
        return result

    @property
    def embedded_objects(self) -> list[EmbeddedObject]:
        """List of |EmbeddedObject| for each ``w:object/o:OLEObject`` in this paragraph.

        Returns an empty list when the paragraph contains no embedded OLE
        objects. Read-only — python-docx does not support creating or
        modifying embedded objects.

        An ``o:OLEObject`` whose ``r:id`` cannot be resolved (for example when
        the referenced part is missing from the package or is of an unexpected
        type) still produces an |EmbeddedObject|, but its
        :attr:`EmbeddedObject.blob` returns ``b""`` and
        :attr:`EmbeddedObject.embedded_partname` returns |None|.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.embedded_objects import EmbeddedObject
        from docx.oxml.ns import qn
        from docx.parts.embedded_object import EmbeddedObjectPart

        result: list[EmbeddedObject] = []
        part = self.part
        for ole_elm in self._p.xpath(".//w:object/o:OLEObject"):
            rId = ole_elm.get(qn("r:id"))
            embedded_part: EmbeddedObjectPart | None = None
            if rId:
                candidate = part.related_parts.get(rId)
                if isinstance(candidate, EmbeddedObjectPart):
                    embedded_part = candidate
            result.append(EmbeddedObject(self, ole_elm, embedded_part))
        return result

    @property
    def hyperlinks(self) -> list[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_section_break(
        self, start_type: WD_SECTION_START = WD_SECTION_START.NEW_PAGE
    ) -> Section:
        """Insert a section break in this paragraph and return the new |Section|.

        `start_type` is a member of :ref:`WdSectionStart` and defaults to
        ``WD_SECTION.NEW_PAGE``. If this paragraph already contains a section break,
        its type is replaced rather than a new one being added.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.section import Section as SectionCls

        pPr = self._p.get_or_add_pPr()
        sectPr = pPr.get_or_add_sectPr()
        sectPr.start_type = start_type
        return SectionCls(sectPr, self.part)

    def remove_section_break(self) -> None:
        """Remove the section break from this paragraph, if one is present.

        Calling this on a paragraph that has no section break is a no-op.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._p.pPr
        if pPr is None:
            return
        if pPr.sectPr is not None:
            pPr._remove_sectPr()

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def insert_paragraph_after(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly after this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph. The new
        paragraph is inserted into the same parent element as this paragraph (which
        may be a body, cell, header/footer, or other block-level container).

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.parser import OxmlElement

        new_p = cast("CT_P", OxmlElement("w:p"))
        self._p.addnext(new_p)
        paragraph = Paragraph(new_p, self._parent)
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def add_caption_before(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Insert a caption paragraph directly before this paragraph and return it.

        This is the common shape for a caption that sits *above* a figure
        or table. The inserted paragraph has the standard caption structure:
        ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field. See
        :meth:`docx.document.Document.add_caption` for details.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.insert_paragraph_before()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def add_caption_after(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Insert a caption paragraph directly after this paragraph and return it.

        This is the common shape for a caption that sits *below* a figure
        or table. The inserted paragraph has the standard caption structure:
        ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field. See
        :meth:`docx.document.Document.add_caption` for details.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.insert_paragraph_after()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def insert_table_of_contents_before(
        self, levels: tuple[int, int] = (1, 3)
    ) -> Paragraph:
        """Insert a TOC paragraph directly before this paragraph and return it.

        `levels` is a ``(min_level, max_level)`` tuple (default ``(1, 3)``)
        controlling which ``"Heading N"`` paragraphs contribute to the cached
        preview text. See
        :meth:`docx.document.Document.add_table_of_contents` for the full
        contract — this method uses the same helper and the same semantics,
        just placed before this paragraph rather than appended.

        The preview scans the document body for headings; the headings that
        appear *before* or *after* this paragraph are both included, since
        Word will rebuild the real TOC on open.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.toc import populate_toc_paragraph

        body = self._get_body()
        source_paragraphs = [Paragraph(p, self._parent) for p in body.xpath(".//w:p")]
        paragraph = self.insert_paragraph_before()
        return populate_toc_paragraph(paragraph, source_paragraphs, levels)

    def insert_table_of_contents_after(
        self, levels: tuple[int, int] = (1, 3)
    ) -> Paragraph:
        """Insert a TOC paragraph directly after this paragraph and return it.

        See :meth:`insert_table_of_contents_before` for the full contract;
        this variant places the TOC paragraph immediately after this one.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.toc import populate_toc_paragraph

        body = self._get_body()
        source_paragraphs = [Paragraph(p, self._parent) for p in body.xpath(".//w:p")]
        paragraph = self.insert_paragraph_after()
        return populate_toc_paragraph(paragraph, source_paragraphs, levels)

    def insert_table_before(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> _Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        before this paragraph.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this paragraph in its parent element.
        `width` is an optional total table width; if not provided it defaults to 6
        inches (a reasonable default for a US-Letter page with 1" margins).

        .. versionadded:: 1.3.0.dev0
        """
        from docx.table import Table

        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._p.addprevious(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    def insert_table_after(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> _Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        after this paragraph.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this paragraph in its parent element.
        `width` is an optional total table width; if not provided it defaults to 6
        inches.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.table import Table

        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._p.addnext(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        for r_or_hlink in self._p.inner_content_elements:
            yield (
                Run(r_or_hlink, self)
                if isinstance(r_or_hlink, CT_R)
                else Hyperlink(r_or_hlink, self)
            )

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def font(self):
        """A |Font| object providing access to the paragraph-mark character formatting.

        This exposes the ``w:pPr/w:rPr`` element — the "paragraph mark" character
        properties that control the font used to render the pilcrow (paragraph
        mark) itself and, in some contexts, the default run formatting for the
        paragraph. When no ``w:pPr`` or ``w:rPr`` element is present, reads
        return |None| for inheritable properties; writes create the chain of
        parent elements as needed.

        Note this is distinct from the per-run :attr:`Run.font`, which controls
        the appearance of the text runs inside the paragraph.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.text.font import Font

        pPr = self._p.get_or_add_pPr()
        return Font(pPr)  # type: ignore[arg-type]

    @property
    def list_level(self) -> int | None:
        """The integer list-level of this paragraph (``w:numPr/w:ilvl/@w:val``).

        Returns |None| when the paragraph has no ``w:numPr`` or ``w:ilvl``
        child. Valid values are ``0`` through ``8``.

        Assigning |None| removes the ``w:ilvl`` child. Assigning an integer
        outside the range 0..8 raises ``ValueError``.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None:
            return None
        return pPr.numPr.ilvl_val

    @list_level.setter
    def list_level(self, value: int | None) -> None:
        if value is not None:
            if not isinstance(value, int) or not 0 <= value <= 8:
                raise ValueError(
                    "list_level must be an int in 0..8 or None, got %r" % (value,)
                )
        pPr = self._p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.ilvl_val = value

    @property
    def list_format(self):
        """Named tuple ``(numbering_definition, level)`` describing this paragraph's
        list settings.

        Both fields are |None| when the paragraph is not part of a list. The
        ``numbering_definition`` is resolved by looking up the paragraph's
        ``numId`` in the document's numbering part.

        To set a paragraph's list format, use
        :meth:`NumberingDefinition.apply_to`.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.numbering import ListFormat, Numbering, NumberingDefinition

        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None:
            return ListFormat(None, None)
        numPr = pPr.numPr
        num_id = numPr.numId_val
        level = numPr.ilvl_val
        if num_id is None:
            return ListFormat(None, level)

        numbering_part = getattr(self.part, "numbering_part", None)
        if numbering_part is None:
            return ListFormat(None, level)

        numbering_elm = numbering_part.numbering_element
        try:
            num = numbering_elm.num_having_numId(num_id)
        except KeyError:
            return ListFormat(None, level)

        abstractNumId_elm = num.abstractNumId
        abstract_num_id = abstractNumId_elm.val
        try:
            abstractNum = numbering_elm.abstractNum_having_abstractNumId(
                abstract_num_id
            )
        except KeyError:
            return ListFormat(None, level)

        numbering_proxy = Numbering(numbering_elm, numbering_part)
        return ListFormat(
            NumberingDefinition(abstractNum, numbering_proxy), level
        )

    @property
    def list_label(self) -> str | None:
        """The rendered number/bullet string Word would display for this paragraph.

        Resolves this paragraph's ``numId`` and ``ilvl`` (directly or
        style-inherited), walks the document body from the start, and returns
        the formatted label — for example ``"1."``, ``"a)"``, ``"I."``,
        ``"1.1."``, or ``"•"`` — computed from the level's ``lvlText``
        pattern and ``numFmt`` (``decimal``, ``decimalZero``, ``upperRoman``,
        ``lowerRoman``, ``upperLetter``, ``lowerLetter``, ``bullet``).
        Returns |None| when this paragraph is not part of any numbered list
        or the referenced numbering cannot be resolved.

        Note that the returned label reflects the paragraph's *current*
        position in the document body. Counters propagate across siblings at
        the same level and reset when a deeper level is entered. This
        property walks the full body on each access — cache the result on
        the caller's side, or use :meth:`Document.list_labels` for a single
        bulk traversal, when label lookup is hot.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.numbering import ListLabelRenderer

        pPr = self._p.pPr
        has_direct_numPr = pPr is not None and pPr.numPr is not None
        # -- quickly bail out when neither direct numPr nor a pStyle pointing at
        # -- a numbered style is present: avoids walking the body unnecessarily --
        has_pStyle = pPr is not None and pPr.style is not None
        if not has_direct_numPr and not has_pStyle:
            return None

        # -- locate the body element that contains this paragraph --
        try:
            body = self._get_body()
        except ValueError:
            return None

        numbering_part = getattr(self.part, "numbering_part", None)
        numbering_elm = (
            numbering_part.numbering_element if numbering_part is not None else None
        )

        styles_elm = None
        try:
            styles_part = self.part.part_related_by(RT.STYLES)
        except (KeyError, AttributeError):
            styles_part = None
        if styles_part is not None:
            styles_elm = getattr(styles_part, "element", None)

        renderer = ListLabelRenderer(numbering_elm, styles_elm)

        # -- walk body paragraphs in order until we hit self --
        target_id = id(self._p)
        for p in body.xpath(".//w:p"):
            label = renderer.label_for(cast("CT_P", p))
            if id(p) == target_id:
                return label
        return None

    @property
    def numbering_format(self):
        """Read-only |Level| describing this paragraph's current level in its list.

        Returns |None| if the paragraph is not part of a numbered list, or if the
        list-level entry cannot be found in the document's numbering part.

        .. versionadded:: 1.3.0.dev0
        """
        list_format = self.list_format
        if list_format.numbering_definition is None:
            return None
        level = list_format.level if list_format.level is not None else 0
        return list_format.numbering_definition.level(level)

    def restart_numbering(self, start: int = 1) -> None:
        """Create a new numbering instance that restarts the current list at `start`.

        The new ``w:num`` reuses the existing abstract definition but adds a
        ``w:lvlOverride/w:startOverride`` for this paragraph's level. The
        paragraph's ``w:numPr/w:numId`` is rewritten to point at the new
        instance, so subsequent siblings at the same level continue the fresh
        count.

        Raises ``ValueError`` when the paragraph is not currently part of a
        numbered list.

        .. versionadded:: 1.3.0.dev0
        """
        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.numId_val is None:
            raise ValueError(
                "paragraph is not part of a numbered list; apply a numbering "
                "definition before calling restart_numbering()"
            )
        numPr = pPr.numPr
        num_id = numPr.numId_val
        ilvl = numPr.ilvl_val or 0

        try:
            numbering_part = self.part.numbering_part  # type: ignore[attr-defined]
        except AttributeError as err:
            raise ValueError(
                "cannot locate numbering part for this paragraph"
            ) from err

        numbering_elm = numbering_part.numbering_element
        try:
            existing_num = numbering_elm.num_having_numId(num_id)
        except KeyError as err:
            raise ValueError(
                "paragraph's numId %d does not match any w:num" % num_id
            ) from err

        abstract_num_id = existing_num.abstractNumId.val
        new_num = numbering_elm.add_num(abstract_num_id)
        override = new_num.add_lvlOverride(ilvl=ilvl)
        override.add_startOverride(val=start)

        numPr.numId_val = new_num.numId

    @property
    def rsid(self) -> str | None:
        """The paragraph's revision-save ID (``w:p/@w:rsidR``) or |None|.

        Read-only. Returns the 8-character hex string Word assigns to mark the
        editing session in which this paragraph was last modified, or |None|
        when the ``@w:rsidR`` attribute is not present.

        .. versionadded:: 1.3.0.dev0
        """
        return self._p.rsidR

    @property
    def stable_id(self) -> str:
        """A 16-character hex stable identifier for this paragraph.

        The ID is derived from the paragraph's ``w:rsidR`` (when present), its
        position within its parent, and its text content. It is stable across
        save/reload *when the paragraph keeps the same position with the same
        text*; it changes if the paragraph is reordered or edited. The value
        is recomputed on each access and never persisted on the element.

        This is intended for tools that need to correlate paragraphs across a
        save/reload cycle in a single editing session. For more robust cross-
        session tracking, compare :attr:`rsid` combined with :attr:`text`.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ids import compute_stable_id

        return compute_stable_id(self._p, self._p.text, self._p.rsidR)

    @property
    def rendered_page_breaks(self) -> list[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def page_breaks_inside(self) -> list[RenderedPageBreak]:
        """All ``w:lastRenderedPageBreak`` positions inside this paragraph.

        Same data as :attr:`rendered_page_breaks`, exposed under the name used
        by upstream issue #744 so the pagination-detection use case is
        discoverable. ``w:lastRenderedPageBreak`` is written by Word when it
        renders a document; a programmatically-created document typically has
        none until Word opens and re-saves it. The explicit page-break markers
        used by :meth:`add_page_break` (``<w:br w:type="page"/>``) are a
        different element and are reported via :attr:`has_page_break`.

        .. versionadded:: 1.3.0.dev0
        """
        return self.rendered_page_breaks

    @property
    def runs(self) -> list[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph.

        Descends transparently through ``w:smartTag`` and ``w:customXml``
        wrappers so runs nested inside those elements are reported alongside
        direct-child runs (upstream #932, #225). Runs nested inside
        ``w:hyperlink``, ``w:fldSimple``, or ``w:sdt`` are not included here
        (they surface via :attr:`hyperlinks`, :attr:`fields`, and
        :attr:`content_controls` respectively); use :attr:`all_runs` when a
        flat view over *every* visible run is needed.
        """
        return [Run(r, self) for r in self._p.iter_r_elements()]

    @property
    def all_runs(self) -> list[Run]:
        """Every visible |Run| in this paragraph, including those nested inside
        ``w:hyperlink``, ``w:fldSimple``, ``w:sdt/w:sdtContent``, complex-field
        ``separate``..``end`` regions, tracked insertions (``w:ins``), move-
        destinations (``w:moveTo``), and smartTag / customXml wrappers.

        Runs whose only content is ``w:instrText`` (the field *code*, not the
        rendered result) are excluded. This is the iterator routed through by
        the Find/Replace helpers in :mod:`docx.search` so that search and
        replace work on the text the user actually sees (upstream #1370,
        #1021).

        .. versionadded:: 1.3.0.dev0
        """
        return [Run(r, self) for r in self._p.iter_all_r_elements()]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def formatting_change(self):
        """A |FormattingChange| for this paragraph's `w:pPrChange`, or |None|.

        Present when the paragraph's formatting (its `w:pPr`) has been edited while
        track-changes is enabled. The returned object exposes the author, date, and
        the prior `w:pPr` via ``old_properties``.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import FormattingChange

        pPr = self._p.pPr
        if pPr is None:
            return None
        pPrChange = pPr.pPrChange  # pyright: ignore[reportAttributeAccessIssue]
        if pPrChange is None:
            return None
        return FormattingChange(pPrChange)

    @property
    def tracked_changes(self) -> list[TrackedChange]:
        """A list of |TrackedChange| objects for each run-level track change.

        Yields proxies for `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo` children
        of this paragraph in document order. Move-revision elements are wrapped
        in |MoveRevision|, exposing the `@w:name` pairing attribute and
        ``.peer`` lookup.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.tracked_changes import CT_MoveFrom, CT_MoveTo

        result: list[TrackedChange] = []
        for tc in self._p.tracked_change_elements:
            if isinstance(tc, (CT_MoveFrom, CT_MoveTo)):
                result.append(MoveRevision(tc))
            else:
                result.append(TrackedChange(tc))
        return result

    def revision_marks_text(
        self,
        open_ins: str = "[+",
        close_ins: str = "+]",
        open_del: str = "[-",
        close_del: str = "-]",
    ) -> str:
        """Return this paragraph's text with tracked-change markers applied.

        Inserted runs (inside ``<w:ins>``) are wrapped with `open_ins`/`close_ins`
        and deleted runs (inside ``<w:del>``) with `open_del`/`close_del`. Runs
        outside of any track-change wrapper are rendered as plain text.

        When the paragraph contains no tracked changes, the return value matches
        :attr:`text`. The defaults are CLI-friendly square-bracket markers; callers
        can pass ANSI escape sequences (e.g. ``"\\033[4m"`` / ``"\\033[0m"``) to
        style terminal output instead.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import _render_paragraph_marks

        return _render_paragraph_marks(
            self._p, open_ins, close_ins, open_del, close_del
        )

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)


# -- human-readable prefixes for generated shape names (used in @name attrs) --
_SHAPE_NAME_PREFIX: dict[WD_SHAPE, str] = {
    WD_SHAPE.RECTANGLE: "Rectangle",
    WD_SHAPE.ROUNDED_RECTANGLE: "Rounded Rectangle",
    WD_SHAPE.OVAL: "Oval",
    WD_SHAPE.ARROW_RIGHT: "Right Arrow",
    WD_SHAPE.CALLOUT_ROUNDED_RECTANGLE: "Callout",
}


def _shape_name_for(shape_type: WD_SHAPE) -> str:
    """Return a human-readable prefix used to build a `wps:cNvPr/@name`."""
    return _SHAPE_NAME_PREFIX.get(shape_type, "Shape")


def _maybe_wrap_tracked_run(
    r: CT_R,
    track_author: str | None,
    paragraph: "Paragraph",
) -> None:
    """Wrap `r` in a `w:ins` revision marker when tracked-change writing is active.

    Resolution order for the author/date:

    1. An explicit `track_author` keyword argument (from
       :meth:`Paragraph.add_run` or :meth:`BlockItemContainer.add_paragraph`)
       always wins. The paired date is taken from the active
       :meth:`Document.tracked_changes` context, or `None` (``now()``) when
       no context is active.
    2. When `track_author` is |None|, the active
       :meth:`Document.tracked_changes` context supplies both author and
       date.

    A paragraph is considered to have no active context when
    ``paragraph.part`` does not reference a :class:`Document` proxy. In that
    case, only the explicit `track_author` branch applies.
    """
    import datetime as _dt

    from docx.tracked_changes import _active_track_author, wrap_run_in_ins

    # -- dial in author/date --
    author: str | None = None
    date: _dt.datetime | None = None
    try:
        part = paragraph.part
    except Exception:  # pragma: no cover -- detached paragraphs in tests
        part = None

    active = _active_track_author(part) if part is not None else None
    if track_author is not None:
        author = track_author
        if active is not None:
            date = active[1]
    elif active is not None:
        author, date = active

    if not author:
        return
    wrap_run_in_ins(r, author, date)
