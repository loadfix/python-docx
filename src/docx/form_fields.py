"""Legacy form-field (``w:ffData``) proxy types.

A *legacy form field* is a pre-SDT Word form control embedded in a document
as a complex-field sequence. The ``begin`` ``w:fldChar`` of the sequence
carries a ``w:ffData`` child that holds the form-field metadata (name, help
text, enabled flag, calc-on-exit flag, and a type-specific options block).

This module exposes a high-level :class:`FormField` proxy that wraps the
``begin`` run of such a complex field, plus three narrow proxies
(:class:`TextInputFormField`, :class:`CheckboxFormField`,
:class:`DropdownFormField`) that surface the type-specific read-only views.

The enum :class:`WD_FORM_FIELD_TYPE` discriminates the three supported form
types — ``TEXT``, ``CHECKBOX``, and ``DROPDOWN``.
"""

from __future__ import annotations

import enum
from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
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


class TextInputFormField:
    """Read-only view onto a ``w:ffData/w:textInput`` block.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, textInput: "CT_FFTextInput"):
        self._textInput = textInput

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
    name: str, default: str = "", maxlength: int | None = None
) -> "CT_FFData":
    """Build a complete ``w:ffData`` element for a text form field.

    .. versionadded:: 2026.05.0
    """
    ffData = _new_ffData(name)
    textInput = ffData.get_or_add_textInput()
    if default:
        default_el = textInput.get_or_add_default()
        default_el.set(qn("w:val"), default)
    # -- maxLength is a signed 16-bit value per the schema; 0 means no limit --
    max_el = textInput.get_or_add_maxLength()
    max_el.set(qn("w:val"), str(maxlength if maxlength is not None else 0))
    # -- emit an empty format element; Word writes this too --
    fmt = textInput.get_or_add_format()
    fmt.set(qn("w:val"), "")
    return ffData


def new_checkbox_form_field_ffData(
    name: str, checked: bool = False
) -> "CT_FFData":
    """Build a complete ``w:ffData`` element for a checkbox form field.

    .. versionadded:: 2026.05.0
    """
    ffData = _new_ffData(name)
    checkBox = ffData.get_or_add_checkBox()
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
