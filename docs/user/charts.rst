.. _charts:

Working with Charts
===================

Word documents can embed *charts* authored in the drawing-markup chart format
(DrawingML ``c:chartSpace``). A chart might be a column chart summarizing
quarterly sales, a pie chart displaying market share, or any of the other
chart kinds Word supports. This page describes how to access charts already
present in a document and read their properties; building a new chart from
scratch is covered separately (see the forward reference at the end).

**Chart anatomy.** Each chart in a document lives in its own *chart part*
(for example ``/word/charts/chart1.xml``) and is referenced from the document
body by a ``<c:chart r:id="..."/>`` element nested inside a ``<w:drawing>``.
The drawing may be *inline* (flows with surrounding text) or *floating*
(anchored). Both are surfaced uniformly by |docx|.

A single chart-part contains one ``c:chartSpace`` XML tree. Inside it are a
*title* (optional), a *plot area* containing one chart-kind element
(``c:barChart``, ``c:lineChart``, ``c:pieChart``, etc.) and one or more
*series* (``c:ser``), and an optional *legend*. Each series carries a name
(its label in the legend), a list of numeric *values*, and a list of
*categories* (x-axis labels for bar / line / column charts, slice labels for
pie charts).

**Chart kinds.** |docx| exposes a subset of Word's chart-type enumeration in
:class:`docx.chart.WD_CHART_TYPE`. The read side recognizes ``BAR``,
``BAR_STACKED``, ``COLUMN``, ``COLUMN_STACKED``, ``LINE``, ``PIE``,
``DOUGHNUT``, ``SCATTER``, and ``AREA``. Charts authored with a chart-kind
outside this list still appear in :attr:`.Document.charts` but report a
|None| ``chart_type``.

**Scope.** The current API is deliberately narrow: it lets you *enumerate*
charts, ask each chart what kind it is and what its title reads, and iterate
its series to pull out the raw numbers. It does not expose axis formatting,
data-label styling, per-point color, or other presentation details; those
live in the underlying XML and can be reached via ``chart.part.element`` if
you need them.


Detecting charts in a document
------------------------------

Every chart referenced from the document body is surfaced through the
:attr:`.Document.charts` property. The list is empty when the document
contains no charts, so you can test for the presence of charts with a simple
truthiness check::

    >>> from docx import Document
    >>> document = Document("quarterly-report.docx")
    >>> if document.charts:
    ...     print("document contains %d chart(s)" % len(document.charts))
    document contains 3 chart(s)

Both *inline* and *floating* chart references are discovered, and duplicate
references to the same chart part (rare, but legal) are de-duplicated so each
chart appears exactly once. Broken references — those whose target part is
missing or of the wrong content-type — are silently skipped rather than
raising.


The Document.charts collection
------------------------------

:attr:`.Document.charts` returns a plain Python ``list`` of |Chart| objects
in document order::

    >>> charts = document.charts
    >>> charts
    [<docx.chart.Chart object at 0x02468ACE>,
     <docx.chart.Chart object at 0x02468BDF>,
     <docx.chart.Chart object at 0x02468CE0>]
    >>> len(charts)
    3
    >>> charts[0]
    <docx.chart.Chart object at 0x02468ACE>

Because it is a list, all the usual sequence operations are available —
indexing, slicing, iteration, and passing through ``len()``. Note that the
value is *recomputed* each time the property is accessed, so if you plan to
hit it repeatedly in a loop you should bind it to a local name.


Reading the chart type
----------------------

Every |Chart| exposes its chart kind via :attr:`.Chart.chart_type`, which
returns a :class:`docx.chart.WD_CHART_TYPE` member (or |None| when the
underlying chart kind is outside the enumerated subset)::

    >>> from docx.chart import WD_CHART_TYPE
    >>> chart = document.charts[0]
    >>> chart.chart_type
    <WD_CHART_TYPE.COLUMN: 'column'>
    >>> chart.chart_type is WD_CHART_TYPE.COLUMN
    True

The distinction between ``BAR`` and ``COLUMN`` (both authored with
``c:barChart`` in the XML) is decided by the ``c:barDir`` child: a *bar*
direction yields :class:`docx.chart.WD_CHART_TYPE`\ ``.BAR`` while *col*
yields :class:`docx.chart.WD_CHART_TYPE`\ ``.COLUMN``. The stacked variants
are likewise distinguished by the ``c:grouping`` value.


Reading the chart title
-----------------------

:attr:`.Chart.title` returns the concatenated text of the chart's title,
pulled from the ``c:title/c:tx/c:rich`` subtree. It returns |None| when no
title element is present::

    >>> chart.title
    'Quarterly Sales'
    >>> untitled_chart = document.charts[2]
    >>> untitled_chart.title is None
    True

Because the title text is built by concatenating every ``a:t`` descendant in
document order, rich-text titles with emphasized runs (for example a bold
word followed by a regular word) still round-trip as their plain-text
equivalent.


Iterating chart series
----------------------

Each chart owns an ordered list of series. :attr:`.Chart.series` returns a
list of :class:`docx.chart.ChartSeries`, one per ``c:ser`` element in the
plot area::

    >>> for series in chart.series:
    ...     print(series.name)
    East
    West

A :class:`docx.chart.ChartSeries` exposes three read-only properties:

- ``name`` — the series label (an empty string when no ``c:tx`` is set).
- ``values`` — a list of ``float`` read from the series' value cache.
- ``categories`` — the list of category labels (as strings) associated with
  this series. For charts with a shared category axis every series reports
  the same list; for scatter / pie charts the list may differ.

Pulling the numeric data out of a chart is therefore a one-liner per series::

    >>> east = chart.series[0]
    >>> east.name
    'East'
    >>> east.values
    [10.0, 20.0, 15.0, 25.0]
    >>> east.categories
    ['Q1', 'Q2', 'Q3', 'Q4']

When the chart XML does not carry a numeric cache (for example because the
authoring tool wrote only ``c:ref`` formulas and no ``c:numCache``), the
``values`` and ``categories`` lists will be empty. This is expected for
charts whose data source is an external embedded spreadsheet that was
subsequently removed.

For convenience, :attr:`.Chart.categories` is a shortcut equivalent to
``chart.series[0].categories`` for the common case of a shared category
axis, returning an empty list when the chart has no series.


Presence of a legend
--------------------

:attr:`.Chart.has_legend` is a ``bool`` that reports whether the chart has a
``c:legend`` element. This is useful when mirroring an existing chart's
formatting into a new document::

    >>> chart.has_legend
    True


Creating a new chart (forward reference)
----------------------------------------

|Document| exposes a narrow creation API as :meth:`.Document.add_chart`,
which takes a :class:`docx.chart.WD_CHART_TYPE`, a list of category labels,
and a mapping of series names to value lists::

    >>> from docx.chart import WD_CHART_TYPE
    >>> chart = document.add_chart(
    ...     WD_CHART_TYPE.COLUMN,
    ...     categories=["Q1", "Q2", "Q3", "Q4"],
    ...     series_data={"East": [10, 20, 15, 25], "West": [12, 18, 14, 22]},
    ... )

Only a subset of chart kinds are supported on the create side —
``BAR``, ``BAR_STACKED``, ``COLUMN``, ``COLUMN_STACKED``, ``LINE``, and
``PIE`` — and the chart is always appended to the document body as an
inline drawing. A more complete chart-authoring API, including titles,
legends, axis configuration, and placement control, is being developed in
a subsequent phase and will have its own user-guide page.
