.. _custom_properties:

Custom document properties
==========================

Beyond the fixed Dublin-Core "core" properties exposed by
:attr:`.Document.core_properties` (title, author, subject, ...), Word also
supports user-defined, typed **custom properties** stored in the
``docProps/custom.xml`` part. Examples of custom properties a document might
carry: a project code, a document revision number, a workflow status, a
client identifier, a budget figure.

|docx| surfaces these through :attr:`.Document.custom_properties`, which
returns a |CustomProperties| collection behaving like a Python ``dict``:
membership testing, indexed access, iteration, deletion, and a handful of
convenience methods.

A lazily-created ``custom.xml`` part is added to the document the first time
:attr:`custom_properties` is accessed, so callers never need to check whether
one already exists.


Supported value types
---------------------

Each custom property has a single, statically-typed value. Five Python types
are supported; each maps to a VT (Variant Type) element defined by the Office
``customProperties`` schema:

============================  ===================  ==============================================
Python type                   OOXML serialisation  Notes
============================  ===================  ==============================================
``str``                       ``vt:lpwstr``        Any length; Unicode.
``int``                       ``vt:i4``            32-bit signed integer.
``float``                     ``vt:r8``            IEEE-754 double.
``bool``                      ``vt:bool``          Stored as ``"true"``/``"false"``.
``datetime.datetime``         ``vt:filetime``      Serialised as ISO-8601 with a ``Z`` suffix.
============================  ===================  ==============================================

Assigning any other type (``list``, ``dict``, ``bytes``, custom objects, ...)
raises :class:`TypeError`.


Reading custom properties
-------------------------

::

    >>> from docx import Document
    >>> document = Document("contract.docx")
    >>> document.custom_properties["Project"]
    'Apollo'
    >>> document.custom_properties["Budget"]
    99.95
    >>> len(document.custom_properties)
    5

The collection also supports membership testing, iteration, and dict-style
``get()``::

    >>> "Project" in document.custom_properties
    True
    >>> "Unknown" in document.custom_properties
    False
    >>> document.custom_properties.get("Unknown")      # returns None
    >>> document.custom_properties.get("Unknown", "-") # returns '-'
    '-'
    >>> list(document.custom_properties)  # iteration yields names
    ['Project', 'Priority', 'Budget', 'Approved', 'Reviewed']

Use :meth:`.CustomProperties.names` to obtain a concrete list of property
names, or :meth:`.CustomProperties.items` to get ``(name, value)`` pairs —
both preserve document order::

    >>> document.custom_properties.names()
    ['Project', 'Priority', 'Budget', 'Approved', 'Reviewed']
    >>> document.custom_properties.items()
    [('Project', 'Apollo'), ('Priority', 5), ('Budget', 99.95), ...]


Adding and updating properties
------------------------------

Subscript assignment is the primary authoring API. If the named property does
not yet exist, it is created; if it does, its value (and serialised type) is
replaced::

    >>> document.custom_properties["Project"] = "Gemini"
    >>> document.custom_properties["Priority"] = 9

|CustomProperties| also offers :meth:`.CustomProperties.add`, which raises
:class:`ValueError` if the name is already in use — useful when the caller
wants to refuse accidental overwrites::

    >>> document.custom_properties.add("Owner", "alice@example.com")
    >>> document.custom_properties.add("Owner", "bob@example.com")
    Traceback (most recent call last):
      ...
    ValueError: a custom property named 'Owner' already exists


Deleting properties
-------------------

Use the ``del`` statement with the property name::

    >>> del document.custom_properties["Priority"]
    >>> "Priority" in document.custom_properties
    False

Deleting a property that doesn't exist raises :class:`KeyError`, mirroring
standard Python dict semantics.


Preserving order
----------------

Custom properties are stored in document order, not by any particular sort
key. Iteration, :meth:`.CustomProperties.names`, and
:meth:`.CustomProperties.items` all walk the underlying
``custom.xml/property`` children in their XML order. Adding a new property
appends it at the end; overwriting an existing property keeps its current
position.


Interoperability notes
----------------------

Word's UI exposes custom properties via **File > Info > Properties > Advanced
Properties > Custom**. Properties authored through |docx| show up in that
dialog with their declared types intact, and can be referenced from Word
fields (for example ``DOCPROPERTY "Project"``) or from other Office
applications reading the same document.

Some third-party readers ignore custom-property types and treat every value
as text. If your downstream tooling depends on the serialised type being
preserved, round-trip your document through the target reader once as part
of your test suite to confirm behaviour.
