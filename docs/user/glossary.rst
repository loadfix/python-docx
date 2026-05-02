.. _glossary:

Working with the Glossary Document
==================================

Word stores its *AutoText*, *Quick Parts*, *cover-page*, *header/footer*, and
similar reusable snippets in a dedicated part of the package called the
*glossary document*. At load time these show up in Word's *Insert > Quick
Parts* and *Insert > Cover Page* galleries. On disk they live under
``word/glossary/document.xml`` alongside the main document part, and each
individual snippet is a ``w:docPart`` element — a *building block*.

python-docx exposes this part **read-only**. The glossary document is almost
always authored by Word itself — for example, it's the part that carries the
built-in cover pages and headers — so this release surfaces it for inspection
without providing any creation or mutation API. Documents created via
``Document()`` with the default template do **not** ship with a glossary
part, so ``document.glossary`` returns |None| for them.

**What you get:**

- Discover whether a document has a glossary part at all.
- Iterate the building blocks it contains, in document order.
- Read each block's metadata (name, description, GUID, category, gallery).
- Walk the paragraphs and tables that make up each block's body.
- Filter and aggregate blocks by gallery and/or category name.


Accessing the glossary
----------------------

The glossary is reached via the :attr:`.Document.glossary` property::

    >>> from docx import Document
    >>> document = Document("briefing-with-cover-pages.docx")
    >>> glossary = document.glossary
    >>> glossary
    <docx.glossary.Glossary object at 0x1deadbeef>

For a document without a glossary part the same property returns |None|::

    >>> Document().glossary is None
    True

It is therefore worth a ``None`` check before working with the proxy::

    >>> glossary = document.glossary
    >>> if glossary is None:
    ...     print("no glossary part")
    ... else:
    ...     print(f"{len(glossary)} building blocks")
    ...
    7 building blocks


Iterating building blocks
-------------------------

The |Glossary| proxy behaves like a read-only collection. It supports
``len()``, iteration, and indexed lookup by building-block name::

    >>> len(glossary)
    7
    >>> for block in glossary:
    ...     print(block.name)
    ...
    Austere Cover Page
    Banded Cover Page
    Default Quick Part
    ...

The :attr:`.Glossary.building_blocks` property returns the same sequence as
a list, in document order. Indexed lookup is by **name** (exact, case
sensitive); a |KeyError| is raised when no building block with that name
exists::

    >>> block = glossary["Austere Cover Page"]
    >>> block.name
    'Austere Cover Page'
    >>> glossary["Does Not Exist"]
    Traceback (most recent call last):
      ...
    KeyError: 'Does Not Exist'


Building-block metadata
-----------------------

Each |BuildingBlock| exposes the metadata Word writes into
``w:docPart/w:docPartPr``::

    >>> block = glossary["Austere Cover Page"]
    >>> block.name
    'Austere Cover Page'
    >>> block.description
    'Cover page with bold title and heading frame.'
    >>> block.guid
    '{12345678-90AB-CDEF-1234-567890ABCDEF}'

Any of these may be |None| when the underlying metadata slot is absent, so
guard accordingly when working with arbitrary documents.

The *category* of a building block is available as a |BuildingBlockCategory|
proxy. The proxy is always returned — even when Word has not written a
``w:category`` element — but its slots will both be |None| in that case::

    >>> block = glossary["Austere Cover Page"]
    >>> block.category
    BuildingBlockCategory(gallery='coverPg', category_name='Built-In')
    >>> block.category.gallery
    'coverPg'
    >>> block.category.category_name
    'Built-In'

The ``gallery`` slot is the raw XML string that Word writes. For the
well-known galleries it can be mapped to a :ref:`WdBuildingBlockGallery`
enum member via :attr:`.BuildingBlockCategory.gallery_value`::

    >>> from docx.enum.text import WD_BUILDING_BLOCK_GALLERY
    >>> block.category.gallery_value is WD_BUILDING_BLOCK_GALLERY.COVER_PAGES
    True

Unknown or vendor-specific gallery values return |None| from
``gallery_value``; the raw string is still available via
:attr:`.BuildingBlockCategory.gallery` for manual inspection.


Reading the body of a building block
------------------------------------

A building block's content — the paragraphs and tables Word inserts when the
user picks the snippet from a gallery — is modelled as a block-item
container::

    >>> block = glossary["Banded Cover Page"]
    >>> for paragraph in block.paragraphs:
    ...     print(paragraph.text)
    ...
    Document Title
    2026-05-02
    >>> len(block.tables)
    1

When the block has no ``w:docPartBody`` element — a legitimate state for
placeholder entries in the glossary — both properties return empty lists.


Filtering and aggregating
-------------------------

The |Glossary| proxy includes a few convenience accessors for bulk
inspection. :meth:`.Glossary.by_category` filters building blocks by
gallery, category name, or both — passing neither is equivalent to
:attr:`.Glossary.building_blocks`::

    >>> from docx.enum.text import WD_BUILDING_BLOCK_GALLERY
    >>> [b.name for b in glossary.by_category(
    ...     gallery=WD_BUILDING_BLOCK_GALLERY.COVER_PAGES
    ... )]
    ['Austere Cover Page', 'Banded Cover Page']
    >>> [b.name for b in glossary.by_category(category_name="Built-In")]
    ['Austere Cover Page', 'Banded Cover Page']

The ``gallery`` argument also accepts a raw XML string, which is useful
when a document uses a gallery value that is not modelled by the enum::

    >>> glossary.by_category(gallery="custom1")
    [<docx.glossary.BuildingBlock object at 0x1deadbeef>]

Two more properties return deduplicated views of the set as a whole.
:attr:`.Glossary.galleries` returns each raw gallery value once, in
first-seen order::

    >>> glossary.galleries
    ['coverPg', 'quickParts']

:attr:`.Glossary.categories` returns one |BuildingBlockCategory| per unique
``(gallery, category_name)`` pair, again in first-seen order. Entries where
both slots are |None| are dropped::

    >>> for cat in glossary.categories:
    ...     print(cat)
    ...
    BuildingBlockCategory(gallery='coverPg', category_name='Built-In')
    BuildingBlockCategory(gallery='quickParts', category_name='General')
