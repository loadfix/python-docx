"""Custom element classes related to custom document properties.

Custom document properties are stored in `docProps/custom.xml` (distinct from core
properties in `docProps/core.xml`). Each property is a typed name/value pair.
"""

from __future__ import annotations

import datetime as dt
from typing import cast

from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore


# -- Office format ID (fmtid) required on every `<property>` in custom.xml. --
# -- This constant value identifies "custom document properties" to Office.   --
CUSTOM_PROPERTIES_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# -- `pid` values 0 and 1 are reserved; user properties start at 2. --
_PID_MIN = 2


class CT_CustomProperties(BaseOxmlElement):
    """`<Properties>` element, the root element of the custom properties part.

    Contains a collection of `<property>` elements, each representing a single custom
    property. The order of the `<property>` elements is insignificant.
    """

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    property_lst: list["CT_CustomProperty"]

    property = ZeroOrMore("custprops:property")

    def add_property(
        self, name: str, fmtid: str = CUSTOM_PROPERTIES_FMTID
    ) -> "CT_CustomProperty":
        """Add a new `<property>` child with `name` and a unique `pid`.

        The returned element has no value child yet; the caller is expected to add an
        appropriately-typed child element (e.g. `vt:lpwstr`, `vt:i4`, `vt:bool`,
        `vt:r8`, or `vt:filetime`).
        """
        pid = self._next_available_pid()
        # -- build the `<property>` as a standalone fragment so the `vt` prefix is
        # -- available for value children added later.
        fragment = (
            f'<Properties xmlns="{nsmap["custprops"]}" '
            f'xmlns:vt="{nsmap["vt"]}">'
            f'<property fmtid="{fmtid}" pid="{pid}" name="{name}"/>'
            f'</Properties>'
        )
        tmp = cast(CT_CustomProperties, parse_xml(fragment))
        new_prop = tmp[0]
        self.append(new_prop)
        return cast("CT_CustomProperty", new_prop)

    def get_property_by_name(self, name: str) -> "CT_CustomProperty | None":
        """Return the first `<property>` whose `name` attribute matches, or |None|."""
        # -- iterate with getchildren-equivalent to avoid xpath namespace gymnastics --
        for prop in self.property_lst:
            if prop.name == name:
                return prop
        return None

    def _next_available_pid(self) -> int:
        """Return the next unused `pid` value, starting at 2."""
        used = {int(p.pid) for p in self.property_lst}
        candidate = _PID_MIN
        while candidate in used:
            candidate += 1
        return candidate


class CT_CustomProperty(BaseOxmlElement):
    """`<property>` element, representing a single custom document property.

    The element has required `fmtid`, `pid`, and `name` attributes and exactly one
    typed value child element in the `vt:` namespace.
    """

    @property
    def name(self) -> str:
        """The value of the `name` attribute."""
        return self.get("name") or ""

    @name.setter
    def name(self, value: str) -> None:
        self.set("name", value)

    @property
    def pid(self) -> int:
        """The value of the `pid` attribute."""
        pid_str = self.get("pid")
        return int(pid_str) if pid_str is not None else 0

    @pid.setter
    def pid(self, value: int) -> None:
        self.set("pid", str(value))

    @property
    def fmtid(self) -> str:
        """The value of the `fmtid` attribute."""
        return self.get("fmtid") or ""

    @fmtid.setter
    def fmtid(self, value: str) -> None:
        self.set("fmtid", value)

    @property
    def value(self) -> object:
        """The Python value represented by this property's typed child element.

        Returns |None| if the property has no recognized value child. Supported types:
        str (`vt:lpwstr`, `vt:lpstr`), int (`vt:i4`, `vt:int`), float (`vt:r8`),
        bool (`vt:bool`), datetime (`vt:filetime`), date (`vt:date`).
        """
        vt_child = self._vt_child
        if vt_child is None:
            return None
        localname = vt_child.tag.split("}", 1)[-1]
        text = vt_child.text or ""
        if localname in ("lpwstr", "lpstr", "bstr"):
            return text
        if localname in ("i4", "int", "i1", "i2", "i8", "ui1", "ui2", "ui4", "ui8"):
            return int(text)
        if localname in ("r4", "r8", "decimal"):
            return float(text)
        if localname == "bool":
            return text.strip().lower() in ("true", "1")
        if localname == "filetime":
            return _parse_filetime(text)
        if localname == "date":
            return _parse_date(text)
        # -- unrecognized type: return raw text --
        return text

    @value.setter
    def value(self, value: object) -> None:
        """Replace this property's value child with an element appropriate for `value`.

        Dispatches by Python type:
        * bool → `vt:bool`
        * int → `vt:i4`
        * float → `vt:r8`
        * datetime → `vt:filetime` (naive datetimes treated as UTC)
        * date (but not datetime) → `vt:date` (format `YYYY-MM-DD`)
        * str → `vt:lpwstr`
        """
        # -- NOTE: `bool` must be checked before `int` because `isinstance(True, int)`
        # -- is True in Python. Similarly `datetime` must be checked before `date`
        # -- because `datetime` is a subclass of `date`.
        if isinstance(value, bool):
            localname, text = "bool", "true" if value else "false"
        elif isinstance(value, int):
            localname, text = "i4", str(value)
        elif isinstance(value, float):
            localname, text = "r8", repr(value)
        elif isinstance(value, dt.datetime):
            localname, text = "filetime", _format_filetime(value)
        elif isinstance(value, dt.date):
            localname, text = "date", _format_date(value)
        elif isinstance(value, str):
            localname, text = "lpwstr", value
        else:
            raise TypeError(
                f"Unsupported custom-property value type: {type(value).__name__}"
            )

        # -- remove any existing value child --
        existing = self._vt_child
        if existing is not None:
            self.remove(existing)

        # -- build the new value child in the vt namespace --
        tag = qn(f"vt:{localname}")
        # -- construct via parse_xml to ensure the vt namespace declaration is picked
        # -- up from the root element (which should already declare it). We use the
        # -- lxml SubElement style via parse_xml of a tiny wrapper so that `tag` is
        # -- a Clark-notation name.
        from lxml import etree

        child = etree.SubElement(self, tag)
        child.text = text

    @property
    def _vt_child(self):
        """Return the first child element in the `vt:` namespace, or |None|."""
        vt_ns = nsmap["vt"]
        for child in self.iterchildren():
            tag = child.tag
            if isinstance(tag, str) and tag.startswith("{" + vt_ns + "}"):
                return child
        return None


def _parse_filetime(text: str) -> dt.datetime:
    """Parse an ISO-8601 UTC datetime as used in `vt:filetime`.

    Accepts either a trailing `Z` or `+00:00` to indicate UTC; returns a naive
    datetime for compatibility with `datetime` values passed in without tzinfo.
    """
    s = text.strip()
    if s.endswith("Z"):
        s = s[:-1]
    # -- strip a `+00:00` / `-00:00` suffix if present; we only support UTC on read --
    if len(s) >= 6 and s[-6] in ("+", "-") and s[-3] == ":":
        s = s[:-6]
    # -- try with and without fractional seconds --
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"could not parse filetime value '{text}'")


def _format_filetime(value: dt.datetime) -> str:
    """Format a datetime as an ISO-8601 UTC string with a trailing `Z`.

    Naive datetimes are treated as already being in UTC. Aware datetimes are
    converted to UTC before formatting.
    """
    if value.tzinfo is not None:
        value = value.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_date(text: str) -> dt.date:
    """Parse a `vt:date` value (ISO-8601 date, `YYYY-MM-DD`) into a `datetime.date`.

    Per ECMA-376 Part 1 §22.4.2.7 a `vt:date` is a `xsd:date` — just the date
    portion without a time component. Some producers tack on a trailing `Z` or
    time-zone suffix even though the spec forbids it; we tolerate both on read
    and discard the extraneous portion.
    """
    s = text.strip()
    # -- strip a trailing `Z` or timezone designator that some producers emit --
    if s.endswith("Z"):
        s = s[:-1]
    if len(s) >= 6 and s[-6] in ("+", "-") and s[-3] == ":":
        s = s[:-6]
    # -- if a time component sneaks in, keep only the date portion --
    if "T" in s:
        s = s.split("T", 1)[0]
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


def _format_date(value: dt.date) -> str:
    """Format a `datetime.date` as an ISO-8601 `YYYY-MM-DD` string for `vt:date`.

    No time or zone suffix is appended — `vt:date` carries only a calendar
    date per ECMA-376 Part 1 §22.4.2.7.
    """
    return value.strftime("%Y-%m-%d")
