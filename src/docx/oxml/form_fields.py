"""Custom element classes for legacy form-field metadata (``w:ffData`` family).

A *legacy form field* is a pre-SDT form control embedded in a document as a
complex-field sequence. The ``begin`` ``w:fldChar`` carries a ``w:ffData``
child that holds the field's metadata (name, help text, enabled flag, and a
type-specific options block). The field type is determined by which of the
three mutually-exclusive marker children is present:

* ``w:textInput`` — free-form text
* ``w:checkBox``  — boolean checkbox
* ``w:ddList``    — drop-down list

The instruction text of the enclosing complex field is ``FORMTEXT``,
``FORMCHECKBOX``, or ``FORMDROPDOWN`` respectively.

Several of the leaf children here — ``w:default``, ``w:checked``, ``w:result``,
``w:maxLength``, ``w:helpText``, ``w:statusText``, ``w:format``,
``w:listEntry`` — are context-dependent: the *type* of the ``@w:val`` attribute
depends on which parent element they live inside. Rather than try to register
a single class against all of them, we leave them unregistered (they are
handled as generic ``BaseOxmlElement`` instances) and the proxy layer in
:mod:`docx.form_fields` reads ``@w:val`` directly and interprets it.
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Callable

from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_FFTextInput(BaseOxmlElement):
    """``<w:textInput>`` inside ``w:ffData``, holding text-input metadata."""

    default: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:default", successors=("w:maxLength", "w:format")
    )
    maxLength: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:maxLength", successors=("w:format",)
    )
    format: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:format", successors=()
    )

    get_or_add_default: Callable[[], BaseOxmlElement]
    get_or_add_maxLength: Callable[[], BaseOxmlElement]
    get_or_add_format: Callable[[], BaseOxmlElement]


class CT_FFCheckBox(BaseOxmlElement):
    """``<w:checkBox>`` inside ``w:ffData``, holding checkbox metadata.

    Note: the schema also allows a ``w:sizeAuto`` / ``w:size`` pair preceding
    ``w:default``/``w:checked``. python-docx does not currently model those;
    they pass through untouched when reading an existing document.
    """

    default: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:default", successors=("w:checked",)
    )
    checked: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:checked", successors=()
    )

    get_or_add_default: Callable[[], BaseOxmlElement]
    get_or_add_checked: Callable[[], BaseOxmlElement]


class CT_FFDDList(BaseOxmlElement):
    """``<w:ddList>`` inside ``w:ffData``, holding dropdown metadata."""

    result: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:result", successors=("w:default", "w:listEntry")
    )
    default: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:default", successors=("w:listEntry",)
    )
    listEntry = ZeroOrMore("w:listEntry", successors=())

    get_or_add_result: Callable[[], BaseOxmlElement]
    get_or_add_default: Callable[[], BaseOxmlElement]
    add_listEntry: Callable[[], BaseOxmlElement]
    listEntry_lst: "list[BaseOxmlElement]"


class CT_FFData(BaseOxmlElement):
    """``<w:ffData>`` element, the metadata block for a legacy form field.

    Lives as a child of the ``begin`` ``w:fldChar`` of a complex form-field
    sequence. The order of its children is fixed by the XSD schema.
    """

    # -- w:name is registered globally as CT_String (val: str RequiredAttribute),
    #    which matches this usage too.
    name: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name",
        successors=(
            "w:enabled",
            "w:calcOnExit",
            "w:helpText",
            "w:statusText",
            "w:textInput",
            "w:checkBox",
            "w:ddList",
        ),
    )
    enabled: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:enabled",
        successors=(
            "w:calcOnExit",
            "w:helpText",
            "w:statusText",
            "w:textInput",
            "w:checkBox",
            "w:ddList",
        ),
    )
    calcOnExit: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:calcOnExit",
        successors=(
            "w:helpText",
            "w:statusText",
            "w:textInput",
            "w:checkBox",
            "w:ddList",
        ),
    )
    helpText: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:helpText",
        successors=(
            "w:statusText",
            "w:textInput",
            "w:checkBox",
            "w:ddList",
        ),
    )
    statusText: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:statusText",
        successors=("w:textInput", "w:checkBox", "w:ddList"),
    )
    textInput: CT_FFTextInput | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:textInput", successors=("w:checkBox", "w:ddList")
    )
    checkBox: CT_FFCheckBox | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:checkBox", successors=("w:ddList",)
    )
    ddList: CT_FFDDList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ddList", successors=()
    )

    get_or_add_name: Callable[[], BaseOxmlElement]
    get_or_add_enabled: Callable[[], BaseOxmlElement]
    get_or_add_calcOnExit: Callable[[], BaseOxmlElement]
    get_or_add_helpText: Callable[[], BaseOxmlElement]
    get_or_add_statusText: Callable[[], BaseOxmlElement]
    get_or_add_textInput: Callable[[], CT_FFTextInput]
    get_or_add_checkBox: Callable[[], CT_FFCheckBox]
    get_or_add_ddList: Callable[[], CT_FFDDList]
