Feature: Create a line chart via Document.add_chart
  In order to author a document containing a line chart
  As a developer using python-docx
  I need Document.add_chart(WD_CHART_TYPE.LINE, ...) to embed a line chart


  Scenario: add_chart appends a line chart at the end of the document
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['Jan', 'Feb', 'Mar'] and series {'Revenue': [100.0, 120.0, 140.0]}
     Then document.charts has 1 chart
      And the added chart is the last embedded chart in the document
      And the added chart.chart_type == WD_CHART_TYPE.LINE


  Scenario: ChartSeries reports the authored names and values
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['Jan', 'Feb', 'Mar'] and series {'2024': [100.0, 110.0, 120.0], '2025': [130.0, 140.0, 150.0]}
     Then the added chart has 2 series
      And [s.name for s in added_chart.series] == ['2024', '2025']
      And added_chart.series[0].values == [100.0, 110.0, 120.0]
      And added_chart.series[1].values == [130.0, 140.0, 150.0]
      And added_chart.series[0].categories == ['Jan', 'Feb', 'Mar']


  Scenario: generated chart part XML contains c:lineChart and one c:ser per series
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['Jan', 'Feb', 'Mar'] and series {'2024': [100.0, 110.0, 120.0], '2025': [130.0, 140.0, 150.0]}
     Then the chart part XML contains a c:lineChart element
      And the chart part XML contains 2 c:ser elements


  Scenario: line chart accepts numeric-looking category labels
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['2021', '2022', '2023', '2024'] and series {'Users': [1000.0, 1500.0, 2100.0, 2800.0]}
     Then added_chart.series[0].categories == ['2021', '2022', '2023', '2024']
      And added_chart.series[0].values == [1000.0, 1500.0, 2100.0, 2800.0]


  Scenario: line chart survives a save-and-reopen round trip
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['Jan', 'Feb', 'Mar'] and series {'Revenue': [100.0, 120.0, 140.0]}
      And I save and reopen the chart-create-line document
     Then document.charts has 1 chart
      And document.charts[0].chart_type == WD_CHART_TYPE.LINE
      And document.charts[0].series[0].values == [100.0, 120.0, 140.0]
      And document.charts[0].series[0].categories == ['Jan', 'Feb', 'Mar']


  Scenario: line chart with three series emits three distinct c:ser entries
    Given a blank chart-create-line base document
     When I add a LINE chart with categories ['Q1', 'Q2', 'Q3', 'Q4'] and series {'North': [10.0, 20.0, 30.0, 40.0], 'South': [15.0, 25.0, 35.0, 45.0], 'East': [12.0, 22.0, 32.0, 42.0]}
     Then the added chart has 3 series
      And [s.name for s in added_chart.series] == ['North', 'South', 'East']
      And the chart part XML contains 3 c:ser elements
