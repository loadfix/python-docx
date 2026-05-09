"""Legacy form-field (``w:ffData``) proxy types.

A *legacy form field* is a pre-SDT Word form control embedded in a document
as a complex-field sequence. The ``begin`` ``w:fldChar`` of the sequence
carries a ``w:ffData`` child that holds the form-field metadata (name, help
text, enabled flag, calc-on-exit flag, and a type-specific options block).

This module exposes a high-level :class:`FormField` proxy that wraps the
``begin`` run of such a complex field. Three subclasses —
:class:`TextInputField`, :class:`CheckBoxField`, :class:`DropDownListField` —
are returned by :meth:`FormField.proxy_for` for pattern-matching convenience,
and three narrow views — :class:`TextInputFormField`,
:class:`CheckboxFormField`, :class:`DropdownFormField` — surface the
type-specific read-only options blocks.

The enum :class:`WD_FORM_FIELD_TYPE` discriminates the three supported form
types — ``TEXT``, ``CHECKBOX``, and ``DROPDOWN``.
"""

from __future__ import annotations

import enum
from typing import TYPE_CHECKING, Union, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.content_controls import ContentControl
    from docx.oxml.fields import CT_FldChar
    from docx.oxml.form_fields import (
        CT_FFCheckBox,
        CT_FFData,
        CT_FFDDList,
        CT_FFTextInput,
    )
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R
    from docx.oxml.xmlchemy import BaseOxmlElement


class WD_FORM_FIELD_TYPE(enum.Enum):
    """Enumerates the three legacy form-field types.

    .. versionadded:: 2026.05.0
    """

    TEXT = "text"
    """Free-form text input (``FORMTEXT``)."""

    CHECKBOX = "checkbox"
    """Boolean checkbox (``FORMCHECKBOX``)."""

    DROPDOWN = "dropdown"
    """Drop-down list of choices (``FORMDROPDOWN``)."""


# -- helpers ----------------------------------------------------------------


def _val_attr(el: "BaseOxmlElement | None", default: str = "") -> str:
    """Return the ``@w:val`` attribute of `el`, or `default` when absent.

    `el` may be ``None``, in which case `default` is returned.
    """
    if el is None:
        return default
    v = el.get(qn("w:val"))
    return v if v is not None else default


def _bool_val(el: "BaseOxmlElement | None", default: bool = True) -> bool:
    """Interpret an OnOff-shape element as a bool.

    An absent element returns `default`. An element with no ``@w:val`` is
    treated as ``True`` (the OOXML on-off convention). Recognised truthy
    values are ``"1"``, ``"true"``, and ``"on"``.
    """
    if el is None:
        return default
    v = el.get(qn("w:val"))
    if v is None:
        return True
    return v.lower() in ("1", "true", "on")


def _int_val(el: "BaseOxmlElement | None", default: int = 0) -> int:
    """Return the ``@w:val`` attribute of `el` parsed as an int, or `default`."""
    if el is None:
        return default
    v = el.get(qn("w:val"))
    if v is None:
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


# -- type-specific proxies --------------------------------------------------


#: Legal ``w:textInput/w:type/@w:val`` values per ``ST_FFTextType`` in the XSD.
#: The ECMA-376 enumeration is ``regular``, ``number``, ``date``,
#: ``currentTime``, ``currentDate``, ``calculated``. Note: the task spec
#: refers to the last as ``calculation``; the schema name is ``calculated``
#: and python-docx emits the schema spelling.
VALID_TEXT_INPUT_TYPES = (
    "regular",
    "number",
    "date",
    "currentTime",
    "currentDate",
    "calculated",
)


class TextInputFormField:
    """Read-only view onto a ``w:ffData/w:textInput`` block.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, textInput: "CT_FFTextInput"):
        self._textInput = textInput

    @property
    def type(self) -> str:
        """The text-input sub-type (``w:type/@w:val``), or ``"regular"``.

        Valid values per ECMA-376 ``ST_FFTextType``: ``"regular"``,
        ``"number"``, ``"date"``, ``"currentTime"``, ``"currentDate"``,
        ``"calculated"``. Absent element returns ``"regular"`` (the schema
        default).

        .. versionadded:: 2026.05.10
        """
        return _val_attr(self._textInput.type, "regular")

    @property
    def default(self) -> str:
        """The default text (``w:default/@w:val``), or the empty string.

        .. versionadded:: 2026.05.0
        """
        return _val_attr(self._textInput.default, "")

    @property
    def max_length(self) -> int | None:
        """The maximum length (``w:maxLength/@w:val``).

        A value of ``0`` in the XML conventionally means "no limit"; this
        property returns |None| in that case. Returns |None| when the element
        is absent.

        .. versionadded:: 2026.05.0
        """
        el = self._textInput.maxLength
        if el is None:
            return None
        n = _int_val(el, 0)
        return n if n > 0 else None

    @property
    def format(self) -> str:
        """The ``w:format/@w:val`` (e.g. ``"UPPERCASE"``), or empty string.

        .. versionadded:: 2026.05.0
        """
        return _val_attr(self._textInput.format, "")


class CheckboxFormField:
    """Read-only view onto a ``w:ffData/w:checkBox`` block.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, checkBox: "CT_FFCheckBox"):
        self._checkBox = checkBox

    @property
    def size_auto(self) -> bool:
        """|True| when the checkbox auto-sizes to the surrounding text.

        The XSD requires exactly one of ``w:size`` (explicit measure) or
        ``w:sizeAuto`` (auto-size flag). When neither child is present the
        checkbox conventionally auto-sizes, so the default is |True|.

        .. versionadded:: 2026.05.10
        """
        if self._checkBox.size is not None:
            return False
        return _bool_val(self._checkBox.sizeAuto, default=True)

    @property
    def size(self) -> int | None:
        """The explicit checkbox size in half-points (``w:size/@w:val``).

        Returns |None| when the checkbox is auto-sized (``w:sizeAuto`` is
        present or no size child is authored at all).

        .. versionadded:: 2026.05.10
        """
        el = self._checkBox.size
        if el is None:
            return None
        v = el.get(qn("w:val"))
        if v is None:
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    @property
    def default(self) -> bool:
        """The default checked state (``w:default``), defaulting to ``False``.

        .. versionadded:: 2026.05.0
        """
        return _bool_val(self._checkBox.default, default=False)

    @property
    def checked(self) -> bool:
        """The current checked state (``w:checked``), defaulting to ``False``.

        When ``w:checked`` is absent but ``w:default`` is present, the default
        is returned. This mirrors Word's behaviour: the initial value of a
        checkbox with no explicit ``w:checked`` is its default.

        .. versionadded:: 2026.05.0
        """
        if self._checkBox.checked is not None:
            return _bool_val(self._checkBox.checked, default=False)
        return self.default


class DropdownFormField:
    """Read-only view onto a ``w:ffData/w:ddList`` block.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, ddList: "CT_FFDDList"):
        self._ddList = ddList

    @property
    def options(self) -> list[str]:
        """The list-entry values (``w:listEntry/@w:val``), in document order.

        .. versionadded:: 2026.05.0
        """
        return [
            _val_attr(le, "") for le in self._ddList.xpath("./w:listEntry")
        ]

    @property
    def default_index(self) -> int:
        """The 0-based default-selection index (``w:default``), or ``0``.

        .. versionadded:: 2026.05.0
        """
        return _int_val(self._ddList.default, 0)

    @property
    def result_index(self) -> int:
        """The 0-based currently-selected index (``w:result``).

        Falls back to :attr:`default_index` when ``w:result`` is absent.

        .. versionadded:: 2026.05.0
        """
        if self._ddList.result is not None:
            return _int_val(self._ddList.result, 0)
        return self.default_index


# -- main form-field proxy --------------------------------------------------


class FormField:
    """High-level proxy for a legacy form field inside a document.

    Wraps the run that carries the ``begin`` ``w:fldChar`` of the form field's
    complex-field sequence. All metadata is read from the ``w:ffData`` child
    of that ``w:fldChar``.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, begin_run: "CT_R"):
        self._begin_run = begin_run

    # -- internal ----------------------------------------------------------

    @property
    def _fldChar(self) -> "CT_FldChar | None":
        """The ``begin`` ``w:fldChar`` inside the begin-run, or |None|."""
        for child in self._begin_run:
            if child.tag == qn("w:fldChar"):
                return cast("CT_FldChar", child)
        return None

    @property
    def _ffData(self) -> "CT_FFData | None":
        fldChar = self._fldChar
        if fldChar is None:
            return None
        return fldChar.ffData

    # -- type discrimination ----------------------------------------------

    @property
    def type(self) -> WD_FORM_FIELD_TYPE | None:
        """The :class:`WD_FORM_FIELD_TYPE` of this form field, or |None|.

        Determined by which of ``w:textInput``, ``w:checkBox``, or ``w:ddList``
        is present in the ``w:ffData`` block. Returns |None| when no type
        marker is present (a malformed but tolerated edge case).

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return None
        if ffData.textInput is not None:
            return WD_FORM_FIELD_TYPE.TEXT
        if ffData.checkBox is not None:
            return WD_FORM_FIELD_TYPE.CHECKBOX
        if ffData.ddList is not None:
            return WD_FORM_FIELD_TYPE.DROPDOWN
        return None

    # -- shared metadata ---------------------------------------------------

    @property
    def name(self) -> str:
        """The form field's name (``w:ffData/w:name/@w:val``), or empty string.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return ""
        return _val_attr(ffData.name, "")

    @property
    def help_text(self) -> str:
        """The help text (``w:ffData/w:helpText/@w:val``), or empty string.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return ""
        return _val_attr(ffData.helpText, "")

    @property
    def status_text(self) -> str:
        """The status text (``w:ffData/w:statusText/@w:val``), or empty string.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return ""
        return _val_attr(ffData.statusText, "")

    @property
    def enabled(self) -> bool:
        """The enabled flag (``w:ffData/w:enabled``).

        An absent element or the default ``True`` (no ``@w:val``) both mean
        the field is enabled. ``@w:val="0"`` means disabled.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return True
        return _bool_val(ffData.enabled, default=True)

    @property
    def calc_on_exit(self) -> bool:
        """The calc-on-exit flag (``w:ffData/w:calcOnExit``), defaulting False.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None:
            return False
        return _bool_val(ffData.calcOnExit, default=False)

    # -- type-specific views ----------------------------------------------

    @property
    def text_input(self) -> TextInputFormField | None:
        """A :class:`TextInputFormField` view, or |None| when not a text field.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None or ffData.textInput is None:
            return None
        return TextInputFormField(ffData.textInput)

    @property
    def checkbox(self) -> CheckboxFormField | None:
        """A :class:`CheckboxFormField` view, or |None| when not a checkbox.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None or ffData.checkBox is None:
            return None
        return CheckboxFormField(ffData.checkBox)

    @property
    def dropdown(self) -> DropdownFormField | None:
        """A :class:`DropdownFormField` view, or |None| when not a dropdown.

        .. versionadded:: 2026.05.0
        """
        ffData = self._ffData
        if ffData is None or ffData.ddList is None:
            return None
        return DropdownFormField(ffData.ddList)

    # -- current value -----------------------------------------------------

    @property
    def value(self) -> str | bool | None:
        """The form field's current value.

        * ``TEXT``: the rendered result text of the complex field (the run
          between the ``separate`` and ``end`` markers), as a ``str``.
        * ``CHECKBOX``: the checked state, as a ``bool``.
        * ``DROPDOWN``: the currently selected list entry, as a ``str``.
        * |None| when no form-field type marker is present.

        .. versionadded:: 2026.05.0
        """
        ff_type = self.type
        if ff_type == WD_FORM_FIELD_TYPE.CHECKBOX:
            cb = self.checkbox
            return cb.checked if cb is not None else False
        if ff_type == WD_FORM_FIELD_TYPE.DROPDOWN:
            dd = self.dropdown
            if dd is None:
                return ""
            options = dd.options
            idx = dd.result_index
            if 0 <= idx < len(options):
                return options[idx]
            return ""
        if ff_type == WD_FORM_FIELD_TYPE.TEXT:
            return self._result_text()
        return None

    @property
    def current_value(self) -> str | bool | None:
        """Alias for :attr:`value`.

        Provided for readability at call-sites that want to make it clear
        they are reading the *current* rendered value of the form field
        (as opposed to its default). See :attr:`value` for the full
        semantics.

        .. versionadded:: 2026.05.10
        """
        return self.value

    # -- SDT migration ----------------------------------------------------

    def to_sdt(self) -> "ContentControl":
        """Replace this legacy form field with an equivalent SDT; return it.

        The five-run complex-field sequence (``begin``/``instrText``/
        ``separate``/result/``end``) is removed from the owning paragraph
        and replaced *in-place* with a single ``<w:sdt>`` whose
        ``w:sdtContent`` contains a run holding the form field's current
        rendered value. Form-field metadata is mapped as follows:

        * ``w:name``  → ``w:sdtPr/w:tag/@w:val``
        * ``w:helpText`` → ``w:sdtPr/w:alias/@w:val``
        * ``TEXT`` fields → ``w:text`` marker (plain-text SDT)
        * ``CHECKBOX`` fields → ``w14:checkbox`` marker (Word 2010 extension)
        * ``DROPDOWN`` fields → ``w:dropDownList`` + ``w:listItem``
          children (one per legacy ``w:listEntry``)

        This is a *lossy* migration: the old ``w:ffData`` block is
        discarded. Callers who need to preserve the legacy element should
        copy it out of the begin-run before invoking this method.

        Returns the |ContentControl| (or a typed subclass) wrapping the
        new ``w:sdt``.

        .. versionadded:: 2026.05.10
        """
        # -- local imports to avoid a hard cycle with content_controls --
        from docx.content_controls import (
            ContentControl,
            ContentControlType,
            new_sdt,
        )

        ff_type = self.type
        cc_type: "ContentControlType"
        if ff_type == WD_FORM_FIELD_TYPE.TEXT:
            cc_type = ContentControlType.PLAIN_TEXT
        elif ff_type == WD_FORM_FIELD_TYPE.CHECKBOX:
            cc_type = ContentControlType.CHECKBOX
        elif ff_type == WD_FORM_FIELD_TYPE.DROPDOWN:
            cc_type = ContentControlType.DROPDOWN
        else:
            cc_type = ContentControlType.RICH_TEXT

        sdt = new_sdt(cc_type, tag=self.name or None,
                      title=self.help_text or None, inline=True)

        # -- populate the sdtContent's seed run with the current value --
        sdtContent = sdt.xpath("./w:sdtContent")[0]
        seed_r = sdtContent.xpath("./w:r")[0]
        if ff_type == WD_FORM_FIELD_TYPE.CHECKBOX:
            cb = self.checkbox
            glyph = "☒" if cb is not None and cb.checked else "☐"
            t = OxmlElement("w:t")
            t.text = glyph
            seed_r.append(t)
        else:
            v = self.value
            if isinstance(v, str) and v:
                t = OxmlElement("w:t")
                t.text = v
                if v != v.strip():
                    t.set(qn("xml:space"), "preserve")
                seed_r.append(t)

        # -- for DROPDOWN, populate the sdt's listItems from the legacy entries
        if ff_type == WD_FORM_FIELD_TYPE.DROPDOWN:
            dd = self.dropdown
            if dd is not None:
                dropDownList = sdt.xpath("./w:sdtPr/w:dropDownList")[0]
                for opt in dd.options:
                    listItem = OxmlElement("w:listItem")
                    listItem.set(qn("w:displayText"), opt)
                    listItem.set(qn("w:value"), opt)
                    dropDownList.append(listItem)

        # -- for CHECKBOX, set the w14:checked state from the legacy field --
        if ff_type == WD_FORM_FIELD_TYPE.CHECKBOX:
            checkbox_marker = sdt.xpath("./w:sdtPr/w14:checkbox")
            if checkbox_marker:
                cb_marker = checkbox_marker[0]
                checked_el = OxmlElement("w14:checked")
                cb = self.checkbox
                checked_el.set(
                    qn("w14:val"),
                    "1" if (cb is not None and cb.checked) else "0",
                )
                cb_marker.append(checked_el)

        # -- replace the legacy begin..end run sequence in the parent --
        parent = self._begin_run.getparent()
        if parent is None:
            raise ValueError("form field begin-run has no parent")
        # -- collect all runs from begin..end inclusive --
        victims: list = [self._begin_run]
        cursor = self._begin_run.getnext()
        while cursor is not None:
            if cursor.tag == qn("w:r"):
                victims.append(cursor)
                is_end = False
                for child in cursor:
                    if child.tag == qn("w:fldChar") and child.get(
                        qn("w:fldCharType")
                    ) == "end":
                        is_end = True
                        break
                if is_end:
                    break
            cursor = cursor.getnext()

        insert_at = parent.index(victims[0])
        parent.insert(insert_at, sdt)
        for v in victims:
            parent.remove(v)

        return ContentControl.proxy_for(sdt)

    def _result_text(self) -> str:
        """Concatenate the text of runs between ``separate`` and ``end`` markers."""
        seen_separate = False
        parts: list[str] = []
        sibling = self._begin_run.getnext()
        while sibling is not None:
            if sibling.tag == qn("w:r"):
                saw_end = False
                for child in sibling:
                    tag = child.tag
                    if tag == qn("w:fldChar"):
                        ft = child.get(qn("w:fldCharType"))
                        if ft == "separate":
                            seen_separate = True
                            break
                        if ft == "end":
                            saw_end = True
                            break
                else:
                    if seen_separate:
                        parts.append(sibling.text or "")
                if saw_end:
                    break
            sibling = sibling.getnext()
        return "".join(parts)


# -- builders ---------------------------------------------------------------


def _append_form_field(
    p: "CT_P",
    instr: str,
    ffData_elm: "CT_FFData",
    result_text: str = "",
) -> "CT_R":
    """Append a legacy-form-field complex-field sequence to `p`.

    The sequence is: ``begin`` run (with ``w:ffData`` attached to the
    ``w:fldChar``), ``instrText`` run, ``separate`` run, result run (always
    emitted, with ``result_text``), and ``end`` run.

    Returns the run carrying the ``begin`` ``w:fldChar``.
    """
    # -- begin run with fldChar + ffData --
    r_begin = p.add_r()
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")
    fldChar_begin.append(ffData_elm)
    r_begin.append(fldChar_begin)

    # -- instrText run --
    r_instr = p.add_r()
    instrText = OxmlElement("w:instrText")
    # -- instructions like " FORMTEXT " are customarily emitted with leading
    #    and trailing spaces; preserve whitespace so consumers don't trim. --
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = instr
    r_instr.append(instrText)

    # -- separate run --
    r_sep = p.add_r()
    fldChar_sep = OxmlElement("w:fldChar")
    fldChar_sep.set(qn("w:fldCharType"), "separate")
    r_sep.append(fldChar_sep)

    # -- result run (always present so the form field has a value slot) --
    r_result = p.add_r()
    t = OxmlElement("w:t")
    t.text = result_text
    if result_text != result_text.strip():
        t.set(qn("xml:space"), "preserve")
    r_result.append(t)

    # -- end run --
    r_end = p.add_r()
    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")
    r_end.append(fldChar_end)

    return r_begin


def _new_ffData(name: str) -> "CT_FFData":
    """Return a new ``w:ffData`` element with ``w:name`` set."""
    ffData = cast("CT_FFData", OxmlElement("w:ffData"))
    name_el = ffData.get_or_add_name()
    name_el.set(qn("w:val"), name)
    # -- always emit w:enabled so the form field is active by default --
    ffData.get_or_add_enabled()
    # -- always emit w:calcOnExit="0" per typical Word output --
    calc = ffData.get_or_add_calcOnExit()
    calc.set(qn("w:val"), "0")
    return ffData


def new_text_form_field_ffData(
    name: str,
    default: str = "",
    maxlength: int | None = None,
    type_: str = "regular",
    format: str = "",
) -> "CT_FFData":
    """Build a complete ``w:ffData`` element for a text form field.

    `type_` selects the ``ST_FFTextType`` sub-kind (``"regular"``,
    ``"number"``, ``"date"``, ``"currentTime"``, ``"currentDate"``,
    ``"calculated"``); invalid values raise |ValueError|. The schema
    default is ``"regular"`` so that value is not emitted (keeps the output
    byte-identical to the pre-2026.05.10 shape when `type_` is left as
    default). `format` populates ``w:format/@w:val`` (e.g. ``"UPPERCASE"``,
    ``"0.00"``, ``"MM/dd/yyyy"``).

    .. versionadded:: 2026.05.0
    .. versionchanged:: 2026.05.10
        Added ``type_`` and ``format`` parameters.
    """
    if type_ not in VALID_TEXT_INPUT_TYPES:
        raise ValueError(
            "type_ must be one of %r; got %r"
            % (list(VALID_TEXT_INPUT_TYPES), type_)
        )
    ffData = _new_ffData(name)
    textInput = ffData.get_or_add_textInput()
    # -- only emit w:type when non-default so existing round-trip shape holds --
    if type_ != "regular":
        type_el = textInput.get_or_add_type()
        type_el.set(qn("w:val"), type_)
    if default:
        default_el = textInput.get_or_add_default()
        default_el.set(qn("w:val"), default)
    # -- maxLength is a signed 16-bit value per the schema; 0 means no limit --
    max_el = textInput.get_or_add_maxLength()
    max_el.set(qn("w:val"), str(maxlength if maxlength is not None else 0))
    # -- always emit format; Word writes it too --
    fmt = textInput.get_or_add_format()
    fmt.set(qn("w:val"), format)
    return ffData


def new_checkbox_form_field_ffData(
    name: str,
    checked: bool = False,
    size_auto: bool = True,
    size: int | None = None,
) -> "CT_FFData":
    """Build a complete ``w:ffData`` element for a checkbox form field.

    When `size` is supplied (a half-points measure), ``w:size`` is emitted
    and ``w:sizeAuto`` is omitted (the XSD ``<xsd:choice>`` forbids both).
    When `size` is |None|, ``w:sizeAuto`` is emitted with ``@w:val="1"`` if
    `size_auto` is |True| (Word's default shape).

    .. versionadded:: 2026.05.0
    .. versionchanged:: 2026.05.10
        Added ``size_auto`` and ``size`` parameters.
    """
    ffData = _new_ffData(name)
    checkBox = ffData.get_or_add_checkBox()
    if size is not None:
        size_el = checkBox.get_or_add_size()
        size_el.set(qn("w:val"), str(int(size)))
    else:
        sizeAuto_el = checkBox.get_or_add_sizeAuto()
        sizeAuto_el.set(qn("w:val"), "1" if size_auto else "0")
    default_el = checkBox.get_or_add_default()
    default_el.set(qn("w:val"), "1" if checked else "0")
    checked_el = checkBox.get_or_add_checked()
    checked_el.set(qn("w:val"), "1" if checked else "0")
    return ffData


def new_dropdown_form_field_ffData(
    name: str, options: list[str], default_index: int = 0
) -> "CT_FFData":
    """Build a complete ``w:ffData`` element for a dropdown form field.

    .. versionadded:: 2026.05.0
    """
    ffData = _new_ffData(name)
    ddList = ffData.get_or_add_ddList()
    result_el = ddList.get_or_add_result()
    result_el.set(qn("w:val"), str(default_index))
    default_el = ddList.get_or_add_default()
    default_el.set(qn("w:val"), str(default_index))
    for opt in options:
        entry = ddList.add_listEntry()
        entry.set(qn("w:val"), opt)
    return ffData


# -- typed subclasses -------------------------------------------------------
#
# The base :class:`FormField` is a type-agnostic facade — it surfaces every
# accessor on a single class and returns |None| for the type-specific views
# that don't apply. The three subclasses below exist to let callers
# pattern-match on the Python type of a form field (``isinstance(ff,
# TextInputField)``) and to let static type-checkers narrow the typed-view
# return values (``TextInputField.text_input`` is always non-|None|).
#
# Subclasses are returned by :meth:`FormField.proxy_for`; construct them
# directly when the kind is known.


class TextInputField(FormField):
    """Legacy text form field (``FORMTEXT``) proxy.

    Subclass of :class:`FormField` returned by :meth:`FormField.proxy_for`
    when the wrapped form field's type is :attr:`WD_FORM_FIELD_TYPE.TEXT`.
    All :class:`FormField` accessors apply; only :attr:`text_input` is
    guaranteed non-|None|.

    .. versionadded:: 2026.05.10
    """


class CheckBoxField(FormField):
    """Legacy checkbox form field (``FORMCHECKBOX``) proxy.

    .. versionadded:: 2026.05.10
    """


class DropDownListField(FormField):
    """Legacy dropdown form field (``FORMDROPDOWN``) proxy.

    .. versionadded:: 2026.05.10
    """


def _proxy_for(begin_run: "CT_R") -> FormField:
    """Return the appropriate :class:`FormField` subclass for `begin_run`.

    Dispatches on the form-field type marker in the ``w:ffData`` block.
    Falls back to the base :class:`FormField` when no type marker is
    present (malformed / stub ``w:ffData``).
    """
    base = FormField(begin_run)
    ff_type = base.type
    if ff_type == WD_FORM_FIELD_TYPE.TEXT:
        return TextInputField(begin_run)
    if ff_type == WD_FORM_FIELD_TYPE.CHECKBOX:
        return CheckBoxField(begin_run)
    if ff_type == WD_FORM_FIELD_TYPE.DROPDOWN:
        return DropDownListField(begin_run)
    return base


# -- attach proxy_for as a classmethod; defined after subclasses so all three
#    are visible in the dispatch table --
FormField.proxy_for = classmethod(  # type: ignore[attr-defined]
    lambda cls, begin_run: _proxy_for(begin_run)
)


# -- unified authoring dispatcher ------------------------------------------


def _kind_key(kind: "str | WD_FORM_FIELD_TYPE") -> WD_FORM_FIELD_TYPE:
    """Coerce a `kind` argument into a :class:`WD_FORM_FIELD_TYPE`."""
    if isinstance(kind, WD_FORM_FIELD_TYPE):
        return kind
    # -- accept raw strings for convenience: "text", "checkbox", "dropdown",
    #    plus the instruction-text aliases "FORMTEXT", "FORMCHECKBOX",
    #    "FORMDROPDOWN" (case-insensitive). --
    s = str(kind).strip().lower()
    mapping = {
        "text": WD_FORM_FIELD_TYPE.TEXT,
        "formtext": WD_FORM_FIELD_TYPE.TEXT,
        "checkbox": WD_FORM_FIELD_TYPE.CHECKBOX,
        "formcheckbox": WD_FORM_FIELD_TYPE.CHECKBOX,
        "dropdown": WD_FORM_FIELD_TYPE.DROPDOWN,
        "formdropdown": WD_FORM_FIELD_TYPE.DROPDOWN,
    }
    if s not in mapping:
        raise ValueError(
            "kind must be one of WD_FORM_FIELD_TYPE or 'text'/"
            "'checkbox'/'dropdown'; got %r" % (kind,)
        )
    return mapping[s]


def _append_form_field_of_kind(
    p: "CT_P",
    kind: Union[str, "WD_FORM_FIELD_TYPE"],
    name: str,
    **kwargs,
) -> "CT_R":
    """Build and append the appropriate form field to `p`; return begin-run.

    Dispatches on `kind`. This is the worker the :meth:`Paragraph.add_form_field`
    facade calls.
    """
    ff_type = _kind_key(kind)
    if ff_type is WD_FORM_FIELD_TYPE.TEXT:
        default = kwargs.pop("default", "")
        maxlength = kwargs.pop("maxlength", None)
        type_ = kwargs.pop("type_", "regular")
        format = kwargs.pop("format", "")
        if kwargs:
            raise TypeError(
                "unexpected keyword arguments for text form field: %r"
                % (sorted(kwargs),)
            )
        ffData = new_text_form_field_ffData(
            name, default=default, maxlength=maxlength,
            type_=type_, format=format,
        )
        return _append_form_field(p, " FORMTEXT ", ffData, result_text=default)
    if ff_type is WD_FORM_FIELD_TYPE.CHECKBOX:
        checked = kwargs.pop("checked", False)
        size_auto = kwargs.pop("size_auto", True)
        size = kwargs.pop("size", None)
        if kwargs:
            raise TypeError(
                "unexpected keyword arguments for checkbox form field: %r"
                % (sorted(kwargs),)
            )
        ffData = new_checkbox_form_field_ffData(
            name, checked=checked, size_auto=size_auto, size=size,
        )
        return _append_form_field(p, " FORMCHECKBOX ", ffData, result_text="")
    # -- dropdown --
    options = kwargs.pop("options", None)
    if options is None:
        raise TypeError("'options' is required for a dropdown form field")
    default_index = kwargs.pop("default_index", 0)
    if kwargs:
        raise TypeError(
            "unexpected keyword arguments for dropdown form field: %r"
            % (sorted(kwargs),)
        )
    ffData = new_dropdown_form_field_ffData(
        name, options=list(options), default_index=default_index,
    )
    initial_text = (
        options[default_index]
        if 0 <= default_index < len(options)
        else ""
    )
    return _append_form_field(
        p, " FORMDROPDOWN ", ffData, result_text=initial_text,
    )
