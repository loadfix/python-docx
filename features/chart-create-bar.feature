Feature: Create a bar chart via Document.add_chart
  In order to embed a clustered bar chart in a new document
  As a developer using python-docx
  I need Document.add_chart to append a c:barChart-backed chart part
  whose chart_type, series names, and series values round-trip cleanly


  Scenario: add_chart appends a clustered bar chart at end of document
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
     Then document.charts has one entry
      And the chart reference sits in the last body paragraph
      And the chart_type of the first chart is WD_CHART_TYPE.BAR


  Scenario: The created bar chart reports the expected series
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
     Then charts[0].series has 2 entries
      And charts[0].series[0].name == "North"
      And charts[0].series[1].name == "South"
      And charts[0].series[0].values == [10.0, 20.0, 15.0]
      And charts[0].series[1].values == [7.0, 14.0, 21.0]
      And charts[0].series[0].categories == ['Q1', 'Q2', 'Q3']


  Scenario: The generated chart part contains a c:barChart with c:barDir=bar
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
     Then the chart part XML contains a c:barChart element
      And the c:barChart has c:barDir with val "bar"
      And the c:barChart has c:grouping with val "clustered"
      And the chart part XML contains 2 c:ser entries


  Scenario: add_chart appends the chart after the existing paragraphs
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
     Then the chart paragraph is positioned after the base paragraphs


  Scenario: A created bar chart round-trips through save and reopen
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
      And I save and reopen the document
     Then document.charts has one entry
      And the chart_type of the first chart is WD_CHART_TYPE.BAR
      And charts[0].series has 2 entries
      And charts[0].series[0].values == [10.0, 20.0, 15.0]


  Scenario: Two bar charts can coexist in the same document
    Given the chart-create-bar base document
     When I add a BAR chart with 3 categories and 2 series
      And I add a BAR chart with 3 categories and 1 series
     Then document.charts has two entries
      And every chart has chart_type WD_CHART_TYPE.BAR
