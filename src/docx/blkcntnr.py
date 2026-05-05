# pyright: reportImportCycles=false

"""Block item container, used by body, cell, header, etc.

Block level items are things like paragraph and table, although there are a few other
specialized ones like structured document tags.
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from typing import TYPE_CHECKING, cast, overload

from typing_extensions import TypeAlias

from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import StoryChild
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.comments import CT_Comment
    from docx.oxml.document import CT_Body
    from docx.oxml.endnotes import CT_Endnote
    from docx.oxml.footnotes import CT_Footnote
    from docx.oxml.section import CT_HdrFtr
    from docx.oxml.table import CT_Tc
    from docx.shared import Length
    from docx.styles.style import ParagraphStyle
    from docx.table import Table

BlockItemElement: TypeAlias = "CT_Body | CT_Comment | CT_Endnote | CT_Footnote | CT_HdrFtr | CT_Tc"


class _ParagraphsView(Sequence[Paragraph]):
    """Lazy read-only view over the paragraphs in a block-item container.

    Each indexed access wraps *only* the requested ``CT_P`` in a
    ``Paragraph`` proxy, instead of materialising the whole list. The
    underlying ``p_lst`` (an lxml ``findall``) is computed once on first
    access and memoised on the view instance, so holding a view and
    indexing it many times is O(1) per access after the first.

    The class behaves like ``list[Paragraph]`` for all the common
    consumers inside this codebase — iteration, ``len()``, ``[i]``,
    ``[a:b]``, ``list(...)``, and ``==`` against a ``list``. It is
    deliberately *not* a ``list`` subclass: callers who actually mutate
    the returned list (``append``, ``sort``, ``[i] = …``) were already
    wrong because ``BlockItemContainer.paragraphs`` has always been a
    read-only snapshot rebuilt on every access.

    Note on idioms. ``container.paragraphs[i]`` re-evaluated inside a
    loop is still O(N^2) in the worst case because the underlying
    document can mutate between calls and we cannot safely cache the
    ``<w:p>`` child list at container scope. The standard idiom
    ``paras = container.paragraphs; paras[i]`` (or iteration) pays the
    O(N) ``findall`` cost once and is then O(1) per access.
    """

    __slots__ = ("_container", "_element", "_p_lst")

    def __init__(self, container: "BlockItemContainer") -> None:
        self._container = container
        self._element = container._element  # pyright: ignore[reportPrivateUsage]
        self._p_lst: list[CT_P] | None = None  # memoised on first access

    def _get_p_lst(self) -> list[CT_P]:
        p_lst = self._p_lst
        if p_lst is None:
            p_lst = self._element.p_lst
            self._p_lst = p_lst
        return p_lst

    def __len__(self) -> int:
        return len(self._get_p_lst())

    @overload
    def __getitem__(self, idx: int) -> Paragraph: ...
    @overload
    def __getitem__(self, idx: slice) -> list[Paragraph]: ...
    def __getitem__(self, idx: int | slice) -> Paragraph | list[Paragraph]:
        p_lst = self._get_p_lst()
        if isinstance(idx, slice):
            return [Paragraph(p, self._container) for p in p_lst[idx]]
        return Paragraph(p_lst[idx], self._container)

    def __iter__(self) -> Iterator[Paragraph]:
        container = self._container
        for p in self._get_p_lst():
            yield Paragraph(p, container)

    def __bool__(self) -> bool:
        # Avoids materialising all proxies just to answer truthiness.
        return bool(self._get_p_lst())

    def __contains__(self, item: object) -> bool:
        # Compare by the underlying <w:p> element so callers can ask
        # ``paragraph in container.paragraphs`` without caring that the
        # view re-wraps each access in a fresh Paragraph proxy.
        if isinstance(item, Paragraph):
            target = item._p  # pyright: ignore[reportPrivateUsage]
            return any(p is target for p in self._get_p_lst())
        return False

    def index(
        self, value: Paragraph, start: int = 0, stop: int | None = None
    ) -> int:
        """Return the first index of ``value`` in this view.

        Matches by underlying ``<w:p>`` element rather than Python
        identity, so ``paragraphs.index(some_paragraph)`` works even
        though the view regenerates ``Paragraph`` proxies on each access.
        The ``start``/``stop`` positional arguments accepted by
        :meth:`collections.abc.Sequence.index` are honoured.
        """
        target = value._p  # pyright: ignore[reportPrivateUsage]
        p_lst = self._get_p_lst()
        length = len(p_lst)
        if stop is None:
            stop = length
        if start < 0:
            start = max(0, length + start)
        if stop < 0:
            stop = max(0, length + stop)
        stop = min(stop, length)
        for idx in range(start, stop):
            if p_lst[idx] is target:
                return idx
        raise ValueError(f"{value!r} is not in paragraph view")

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _ParagraphsView):
            # Compare by underlying <w:p> identity — two distinct views
            # over the same container are semantically equal.
            return self._get_p_lst() == other._get_p_lst()
        if isinstance(other, list):
            p_lst = self._get_p_lst()
            other_lst = cast("list[object]", other)
            if len(other_lst) != len(p_lst):
                return False
            for a_p, b in zip(p_lst, other_lst):
                if (
                    not isinstance(b, Paragraph)
                    or b._p is not a_p  # pyright: ignore[reportPrivateUsage]
                ):
                    return False
            return True
        return NotImplemented

    def __ne__(self, other: object) -> bool:
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result

    def __repr__(self) -> str:
        return "_ParagraphsView(len=%d)" % len(self)


class BlockItemContainer(StoryChild):
    """Base class for proxy objects that can contain block items.

    These containers include _Body, _Cell, header, footer, footnote, endnote, comment,
    and text box objects. Provides the shared functionality to add a block item like a
    paragraph or table.
    """

    def __init__(self, element: BlockItemElement, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._element = element

    def add_paragraph(
        self,
        text: str = "",
        style: str | ParagraphStyle | None = None,
        track_author: str | None = None,
    ) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph
        style `style`.

        If `style` is |None|, no paragraph style is applied, which has the same effect
        as applying the 'Normal' style.

        If `track_author` is supplied (or if an enclosing
        :meth:`Document.tracked_changes` context is active), the freshly-inserted
        run containing `text` is wrapped in a `w:ins` revision marker
        attributed to that author. A paragraph added with empty `text` is not
        wrapped because it contains no run to mark. Closes upstream#1025.

        .. versionadded:: 2026.05.0
           Added ``track_author`` keyword argument.
        """
        paragraph = self._add_paragraph()
        if text:
            if track_author is None:
                paragraph.add_run(text)
            else:
                paragraph.add_run(text, track_author=track_author)
        if style is not None:
            paragraph.style = style
        return paragraph

    def add_table(self, rows: int, cols: int, width: Length) -> Table:
        """Return table of `width` having `rows` rows and `cols` columns.

        The table is appended appended at the end of the content in this container.

        `width` is evenly distributed between the table columns.
        """
        from docx.table import Table

        tbl = CT_Tbl.new_tbl(rows, cols, width)
        self._element._insert_tbl(tbl)  # pyright: ignore[reportPrivateUsage]
        return Table(tbl, self)

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this container in document order."""
        from docx.table import Table

        for element in self._element.inner_content_elements:
            yield (Paragraph(element, self) if isinstance(element, CT_P) else Table(element, self))

    @property
    def paragraphs(self):
        """A sequence of the paragraphs in this container, in document order.

        Read-only. The return value is a lightweight ``Sequence[Paragraph]``
        view that supports ``len()``, indexed access (``[i]`` and
        ``[a:b]``), iteration, ``list(...)`` coercion, and equality
        comparison with a ``list[Paragraph]``. Each indexed access wraps
        only the requested element, so a ``for i in range(N):
        container.paragraphs[i]`` loop runs in O(N) instead of O(N^2).
        """
        return _ParagraphsView(self)

    @property
    def tables(self):
        """A list containing the tables in this container, in document order.

        Read-only.
        """
        from docx.table import Table

        return [Table(tbl, self) for tbl in self._element.tbl_lst]

    def _add_paragraph(self):
        """Return paragraph newly added to the end of the content in this container."""
        return Paragraph(self._element.add_p(), self)
