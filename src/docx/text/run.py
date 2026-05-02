"""Run-related proxy objects for python-docx, Run in particular."""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator

from docx.drawing import Drawing
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.oxml.drawing import CT_Drawing
from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
from docx.shape import InlineShape
from docx.shared import StoryChild
from docx.styles.style import CharacterStyle
from docx.text.font import Font
from docx.text.pagebreak import RenderedPageBreak
from docx.text.symbol import Symbol

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_UNDERLINE
    from docx.oxml.text.run import CT_R, CT_Text
    from docx.ruby import RubyAnnotation
    from docx.shared import Length


class Run(StoryChild):
    """Proxy object wrapping `<w:r>` element.

    Several of the properties on Run take a tri-state value, |True|, |False|, or |None|.
    |True| and |False| correspond to on and off respectively. |None| indicates the
    property is not specified directly on the run and its effective value is taken from
    the style hierarchy.
    """

    def __init__(self, r: CT_R, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._r = self._element = self.element = r

    def add_break(self, break_type: WD_BREAK = WD_BREAK.LINE):
        """Add a break element of `break_type` to this run.

        `break_type` can take the values `WD_BREAK.LINE`, `WD_BREAK.PAGE`, and
        `WD_BREAK.COLUMN` where `WD_BREAK` is imported from `docx.enum.text`.
        `break_type` defaults to `WD_BREAK.LINE`.
        """
        type_, clear = {
            WD_BREAK.LINE: (None, None),
            WD_BREAK.PAGE: ("page", None),
            WD_BREAK.COLUMN: ("column", None),
            WD_BREAK.LINE_CLEAR_LEFT: ("textWrapping", "left"),
            WD_BREAK.LINE_CLEAR_RIGHT: ("textWrapping", "right"),
            WD_BREAK.LINE_CLEAR_ALL: ("textWrapping", "all"),
        }[break_type]
        br = self._r.add_br()
        if type_ is not None:
            br.type = type_
        if clear is not None:
            br.clear = clear

    def add_picture(
        self,
        image_path_or_stream: "str | os.PathLike[str] | IO[bytes] | None" = None,
        width: int | Length | None = None,
        height: int | Length | None = None,
        link: bool = False,
        save_with_document: bool = True,
        url: str | None = None,
    ) -> InlineShape:
        """Return |InlineShape| containing image identified by `image_path_or_stream`.

        The picture is added to the end of this run.

        `image_path_or_stream` can be a ``str`` path, an :class:`os.PathLike`
        (e.g. :class:`pathlib.Path`), or a binary file-like object containing an image.

        If neither width nor height is specified, the picture appears at
        its native size. If only one is specified, it is used to compute a scaling
        factor that is then applied to the unspecified dimension, preserving the aspect
        ratio of the image. The native size of the picture is calculated using the dots-
        per-inch (dpi) value specified in the image file, defaulting to 72 dpi if no
        value is specified, as is often the case.

        When `link` is |True| and `save_with_document` is |False|, the
        picture is added as a linked (external) image: no image part is
        created in the package and the `a:blip` uses ``r:link`` referencing
        an external relationship. `url` may be supplied to link a remote
        image; when both `url` and `image_path_or_stream` are supplied,
        `url` becomes the link target while the local path is used only to
        probe the native dimensions.

        .. versionchanged:: 1.3.0.dev0
           Accepts :class:`os.PathLike` path arguments.

        .. versionadded:: 1.3.0.dev0
            ``link``, ``save_with_document``, and ``url`` parameters.
        """
        if isinstance(image_path_or_stream, os.PathLike):
            image_path_or_stream = os.fspath(image_path_or_stream)
        inline = self.part.new_pic_inline(
            image_path_or_stream,
            width,
            height,
            link=link,
            save_with_document=save_with_document,
            url=url,
        )
        self._r.add_drawing(inline)
        return InlineShape(inline, self.part)

    def add_text_box(
        self,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append a DrawingML text box (``wps:wsp`` + ``wps:txbx``) to this run.

        The text box is created with a rectangular preset geometry of `width`
        by `height` (defaults 3" x 1.5") and may be seeded with `text` in a
        single initial paragraph. Callers can add further paragraphs via
        :meth:`~docx.drawing.WordprocessingShape.add_paragraph`.

        Returns the :class:`~docx.drawing.WordprocessingShape` proxy.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.drawing import WordprocessingShape
        from docx.enum.shape import WD_SHAPE
        from docx.oxml.drawing import new_inline_shape_drawing
        from docx.shared import Inches

        cx = int(width) if width is not None else int(Inches(3))
        cy = int(height) if height is not None else int(Inches(1.5))

        story_part = self.part
        shape_id = story_part.next_id
        name = "Text Box %d" % shape_id

        drawing = new_inline_shape_drawing(
            WD_SHAPE.RECTANGLE.value,
            cx,
            cy,
            shape_id,
            name,
            text=text if text is not None else "",
        )
        self._r.append(drawing)

        wsp = drawing.xpath(
            ".//wp:inline/a:graphic/a:graphicData/wps:wsp"
        )[0]
        return WordprocessingShape(wsp, self)

    def add_tab(self) -> None:
        """Add a ``<w:tab/>`` element at the end of the run, which Word interprets as a
        tab character."""
        self._r.add_tab()

    def add_text(self, text: str):
        """Returns a newly appended |_Text| object (corresponding to a new ``<w:t>``
        child element) to the run, containing `text`.

        Compare with the possibly more friendly approach of assigning text to the
        :attr:`Run.text` property.
        """
        t = self._r.add_t(text)
        return _Text(t)

    def add_symbol(self, char_code: int | str, font: str) -> Symbol:
        """Append a ``<w:sym>`` element to this run and return a |Symbol| for it.

        `char_code` identifies the glyph's Unicode code point within `font`. It
        may be an ``int`` (e.g. ``0xF0E0``) or a hex ``str`` (e.g. ``"F0E0"``
        or ``"0xF0E0"``). Word always stores this value as a 4-character
        uppercase hex string in the XML; integer and lowercase-hex inputs are
        normalized on write. `font` is the name of the font supplying the
        glyph, for example ``"Wingdings"``.

        .. versionadded:: 1.3.0.dev0
        """
        if isinstance(char_code, str):
            code_int = int(char_code, 16)
        else:
            code_int = int(char_code)
        char_hex = format(code_int, "04X")
        sym = self._r.add_sym(char_hex, font)
        return Symbol(sym)

    @property
    def symbols(self) -> Iterator[Symbol]:
        """Generate a |Symbol| for each ``<w:sym>`` child of this run, in document
        order.

        .. versionadded:: 1.3.0.dev0
        """
        for sym in self._r.sym_lst:
            yield Symbol(sym)

    @property
    def text_with_symbols(self) -> str:
        """Run text including ``w:sym`` glyphs rendered as ``chr(@w:char)``.

        Alias for :attr:`text`; kept as a named property because the upstream
        issue request (upstream#1528) was specifically for a symbol-aware
        variant of ``run.text``. Provided so callers can opt into the intent
        explicitly even though ``.text`` now includes symbols too.

        .. versionadded:: 1.3.0.dev0
        """
        return self._r.text

    @property
    def equations(self):
        """List of |Equation| objects for OMML elements inside this run.

        OMML is almost always a paragraph-level sibling of ``w:r`` (not a run
        child), so this property is usually empty. It is provided for symmetry
        with :attr:`Paragraph.equations` so callers can query any run without
        a type check. Walks descendant ``m:oMath`` and ``m:oMathPara`` nodes.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.equations import Equation

        result: list[Equation] = []
        for el in self._r.xpath(
            ".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]"
        ):
            result.append(Equation(el))
        return result

    @property
    def ruby_annotations(self) -> list["RubyAnnotation"]:
        """A |RubyAnnotation| for each ``<w:ruby>`` child, in document order.

        Read-only. Ruby is used for phonetic annotation (Japanese furigana etc.)
        pairing base text with an above-the-line reading.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ruby import RubyAnnotation

        return [RubyAnnotation(r) for r in self._r.ruby_lst]

    @property
    def bold(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text of the run to appear in bold face. When |False|,
        the text unconditionally appears non-bold. When |None| the bold setting for this
        run is inherited from the style hierarchy.
        """
        return self.font.bold

    @bold.setter
    def bold(self, value: bool | None):
        self.font.bold = value

    def clear(self):
        """Return reference to this run after removing all its content.

        All run formatting is preserved.
        """
        self._r.clear_content()
        return self

    def copy_formatting_from(self, source: "Run") -> "Run":
        """Replace this run's character formatting with a deep copy of `source`'s.

        The source run's ``w:rPr`` is deep-copied onto this run, replacing any
        pre-existing character formatting on this run. The run's text content
        is untouched. Returns this run for chaining convenience.

        .. versionadded:: 1.3.0.dev0
        """
        source.font.copy_to(self.font)
        return self

    def delete(self) -> None:
        """Remove this run from its parent paragraph.

        The run element is removed from its parent. After calling this method,
        this |Run| object is "defunct" and should not be used further.

        .. versionadded:: 1.3.0.dev0
        """
        r = self._r
        parent = r.getparent()
        if parent is None:
            return
        parent.remove(r)

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this run.

        Note that "hard" page-breaks inserted by the author are not included. A hard
        page-break gives rise to a rendered page-break in the right position so if those
        were included that page-break would be "double-counted".

        It would be very rare for multiple rendered page-breaks to occur in a single
        run, but it is possible.
        """
        return bool(self._r.lastRenderedPageBreaks)

    @property
    def font(self) -> Font:
        """The |Font| object providing access to the character formatting properties for
        this run, such as font name and size."""
        return Font(self._element)

    @property
    def italic(self) -> bool | None:
        """Read/write tri-state value.

        When |True|, causes the text of the run to appear in italics. When |False|, the
        text unconditionally appears non-italic. When |None| the italic setting for this
        run is inherited from the style hierarchy.
        """
        return self.font.italic

    @italic.setter
    def italic(self, value: bool | None):
        self.font.italic = value

    def iter_inner_content(self) -> Iterator[str | Drawing | RenderedPageBreak]:
        """Generate the content-items in this run in the order they appear.

        NOTE: only content-types currently supported by `python-docx` are generated. In
        this version, that is text and rendered page-breaks. Drawing is included but
        currently only provides access to its XML element (CT_Drawing) on its
        `._drawing` attribute. `Drawing` attributes and methods may be expanded in
        future releases.

        There are a number of element-types that can appear inside a run, but most of
        those (w:br, w:cr, w:noBreakHyphen, w:t, w:tab) have a clear plain-text
        equivalent. Any contiguous range of such elements is generated as a single
        `str`. Rendered page-break and drawing elements are generated individually. Any
        other elements are ignored.
        """
        for item in self._r.inner_content_items:
            if isinstance(item, str):
                yield item
            elif isinstance(item, CT_LastRenderedPageBreak):
                yield RenderedPageBreak(item, self)
            elif isinstance(item, CT_Drawing):  # pyright: ignore[reportUnnecessaryIsInstance]
                yield Drawing(item, self)

    def mark_comment_range(self, last_run: Run, comment_id: int) -> None:
        """Mark the range of runs from this run to `last_run` (inclusive) as belonging to a comment.

        `comment_id` identfies the comment that references this range.
        """
        # -- insert `w:commentRangeStart` with `comment_id` before this (first) run --
        self._r.insert_comment_range_start_above(comment_id)

        # -- insert `w:commentRangeEnd` and `w:commentReference` run with `comment_id` after
        # -- `last_run`
        last_run._r.insert_comment_range_end_and_reference_below(comment_id)

    def split(self, offset: int) -> tuple[Run, Run]:
        """Return (left_run, right_run) after splitting this run at character `offset`.

        Text before `offset` stays in this run and text from `offset` onward moves
        to a new run inserted immediately after this one. Both runs share the same
        character formatting (`w:rPr`).

        .. versionadded:: 1.3.0.dev0
        """
        new_r = self._r.split_run(offset)
        right_run = Run(new_r, self._parent)
        return self, right_run

    @property
    def formatting_change(self):
        """A |FormattingChange| for this run's `w:rPrChange`, or |None|.

        Present when the run's formatting (its `w:rPr`) has been edited while
        track-changes is enabled. The returned object exposes the author, date,
        and the prior `w:rPr` via ``old_properties``.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.tracked_changes import FormattingChange

        rPr = self._r.rPr
        if rPr is None:
            return None
        rPrChange = rPr.rPrChange  # pyright: ignore[reportAttributeAccessIssue]
        if rPrChange is None:
            return None
        return FormattingChange(rPrChange)

    @property
    def rsid(self) -> str | None:
        """The run's revision-save ID (``w:r/@w:rsidR``) or |None|.

        Read-only. Returns the 8-character hex string Word assigns to mark the
        editing session in which this run was last modified, or |None| when
        the ``@w:rsidR`` attribute is not present.

        .. versionadded:: 1.3.0.dev0
        """
        return self._r.rsidR

    @property
    def stable_id(self) -> str:
        """A 16-character hex stable identifier for this run.

        The ID is derived from the run's ``w:rsidR`` (when present), its
        position within its parent element, and its text. It is stable across
        save/reload *when the run keeps the same position with the same text*;
        it changes if the run is reordered or edited. The value is recomputed
        on each access and never persisted on the element.

        For more robust cross-session tracking, compare :attr:`rsid` combined
        with :attr:`text`.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.ids import compute_stable_id

        return compute_stable_id(self._r, self._r.text, self._r.rsidR)

    @property
    def style(self) -> CharacterStyle:
        """Read/write.

        A |CharacterStyle| object representing the character style applied to this run.
        The default character style for the document (often `Default Character Font`) is
        returned if the run has no directly-applied character style. Setting this
        property to |None| removes any directly-applied character style.
        """
        style_id = self._r.style
        return cast(CharacterStyle, self.part.get_style(style_id, WD_STYLE_TYPE.CHARACTER))

    @style.setter
    def style(self, style_or_name: str | CharacterStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.CHARACTER)
        self._r.style = style_id

    @property
    def text(self) -> str:
        """String formed by concatenating the text equivalent of each run.

        Each `<w:t>` element adds the text characters it contains. A `<w:tab/>` element
        adds a `\\t` character. A `<w:cr/>` or `<w:br>` element each add a `\\n`
        character. Note that a `<w:br>` element can indicate a page break or column
        break as well as a line break. Only line-break `<w:br>` elements translate to
        a `\\n` character. Others are ignored. All other content child elements, such as
        `<w:drawing>`, are ignored.

        Assigning text to this property has the reverse effect, translating each `\\t`
        character to a `<w:tab/>` element and each `\\n` or `\\r` character to a
        `<w:cr/>` element. Any existing run content is replaced. Run formatting is
        preserved.
        """
        return self._r.text

    @text.setter
    def text(self, text: str):
        self._r.text = text

    @property
    def underline(self) -> bool | WD_UNDERLINE | None:
        """The underline style for this |Run|.

        Value is one of |None|, |True|, |False|, or a member of :ref:`WdUnderline`.

        A value of |None| indicates the run has no directly-applied underline value and
        so will inherit the underline value of its containing paragraph. Assigning
        |None| to this property removes any directly-applied underline value.

        A value of |False| indicates a directly-applied setting of no underline,
        overriding any inherited value.

        A value of |True| indicates single underline.

        The values from :ref:`WdUnderline` are used to specify other outline styles such
        as double, wavy, and dotted.
        """
        return self.font.underline

    @underline.setter
    def underline(self, value: bool | WD_UNDERLINE | None):
        self.font.underline = value


class _Text:
    """Proxy object wrapping `<w:t>` element."""

    def __init__(self, t_elm: CT_Text):
        super().__init__()
        self._t = t_elm
