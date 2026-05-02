Feature: Create a pie chart with Document.add_chart
  In order to author new pie charts from python-docx
  As a developer using the create-side chart API
  I need Document.add_chart(WD_CHART_TYPE.PIE, ...) to embed a pie chart
  with the supplied category labels and single series of values


  Scenario: Document.add_chart(PIE, ...) appends a chart at end of document
    Given a base document for pie-chart creation
     When I add a pie chart with slices "Red","Green","Blue" valued 10,20,30
     Then document.charts has length 1
      And the last paragraph in the document contains the new chart


  Scenario: The new chart reports chart_type as PIE
    Given a base document for pie-chart creation
     When I add a pie chart with slices "Red","Green","Blue" valued 10,20,30
     Then charts[0].chart_type == WD_CHART_TYPE.PIE


  Scenario: The pie chart exposes a single ChartSeries with the expected slices
    Given a base document for pie-chart creation
     When I add a pie chart with slices "Red","Green","Blue" valued 10,20,30
     Then document.charts[0] has 1 series
      And document.charts[0].series[0].name == "Slices"
      And charts[0].series[0].categories == ['Red', 'Green', 'Blue']
      And charts[0].series[0].values == [10.0, 20.0, 30.0]
      And the sum of charts[0].series[0].values is 60.0


  Scenario: The generated word/charts/chart1.xml contains c:pieChart and one c:ser
    Given a base document for pie-chart creation
     When I add a pie chart with slices "Red","Green","Blue" valued 10,20,30
      And I save the pie-chart document to scratch
     Then word/charts/chart1.xml contains a c:pieChart element
      And word/charts/chart1.xml contains exactly 1 c:ser element


  Scenario: Pie chart survives a save + reopen round-trip without loss
    Given a base document for pie-chart creation
     When I add a pie chart with slices "Red","Green","Blue" valued 10,20,30
      And I save and reopen the pie-chart document
     Then document.charts has length 1
      And charts[0].chart_type == WD_CHART_TYPE.PIE
      And charts[0].series[0].categories == ['Red', 'Green', 'Blue']
      And charts[0].series[0].values == [10.0, 20.0, 30.0]
