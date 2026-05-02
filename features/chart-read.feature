Feature: Read charts embedded in a document
  In order to inspect charts authored by another tool or earlier process
  As a developer using python-docx
  I need read-side access to each chart's type, title, and series


  Scenario: Document.charts yields a Chart for each embedded chart
    Given a document having three embedded charts
     Then document.charts is a list of three Chart objects


  Scenario: Iterate Document.charts in document order
    Given a document having three embedded charts
     Then iterating document.charts yields Chart objects in document order


  Scenario Outline: Chart.chart_type reports the chart kind
    Given a document having three embedded charts
     Then charts[<idx>].chart_type == WD_CHART_TYPE.<member>

    Examples: chart type per chart
      | idx | member |
      | 0   | COLUMN |
      | 1   | BAR    |
      | 2   | LINE   |


  Scenario Outline: Chart.title exposes the chart title text
    Given a document having three embedded charts
     Then charts[<idx>].title == "<title>"

    Examples: chart title per chart
      | idx | title               |
      | 0   | Quarterly Sales     |
      | 1   | Headcount by Region |
      | 2   | Monthly Revenue     |


  Scenario: Chart.title is None when no title element is present
    Given a document having a chart with no title
     Then chart.title is None


  Scenario Outline: Chart.series exposes series names in document order
    Given a document having three embedded charts
     Then [s.name for s in charts[<idx>].series] == <names>

    Examples: series names per chart
      | idx | names                  |
      | 0   | ['East', 'West']       |
      | 1   | ['Employees']          |
      | 2   | ['2024', '2025']       |


  Scenario: ChartSeries exposes numeric values and categories
    Given a document having three embedded charts
     Then charts[0].series[0].values == [10.0, 20.0, 15.0, 25.0]
      And charts[0].series[0].categories == ['Q1', 'Q2', 'Q3', 'Q4']


  Scenario: Document with no charts returns an empty list
    Given a document having no charts
     Then document.charts is an empty list
