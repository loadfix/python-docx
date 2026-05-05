Feature: Create a SmartArt diagram with Document.add_smart_art
  In order to author SmartArt diagrams without hand-crafting five XML parts
  As a developer using python-docx
  I need Document.add_smart_art(layout_name) to materialise the data, layout,
  colors, and quickStyle companion parts and attach them via an inline drawing


  Scenario Outline: Document.add_smart_art appends a SmartArt of the requested family
    Given a blank document for SmartArt authoring
     When I add a SmartArt diagram of family "<family>"
     Then document.smart_art has length 1
      And the last SmartArt's data partname ends with "data1.xml"

    Examples: SmartArt families
      | family  |
      | list    |
      | cycle   |
      | process |


  Scenario: SmartArt.add_node appends content nodes to the diagram
    Given a blank document for SmartArt authoring
     When I add a SmartArt diagram of family "list"
      And I add the SmartArt nodes "Alpha", "Beta", "Gamma"
     Then the SmartArt's node texts are ['Alpha', 'Beta', 'Gamma']


  Scenario: Authored SmartArt survives a save + reopen round-trip
    Given a blank document for SmartArt authoring
     When I add a SmartArt diagram of family "process"
      And I add the SmartArt nodes "Plan", "Build", "Ship"
      And I save and reopen the SmartArt document
     Then document.smart_art has length 1
      And the SmartArt's node texts are ['Plan', 'Build', 'Ship']


  Scenario: The generated package contains all four SmartArt companion parts
    Given a blank document for SmartArt authoring
     When I add a SmartArt diagram of family "cycle"
      And I add the SmartArt nodes "One", "Two"
      And I save the SmartArt document to scratch
     Then the package contains word/diagrams/data1.xml
      And the package contains word/diagrams/layout1.xml
      And the package contains word/diagrams/colors1.xml
      And the package contains word/diagrams/quickStyle1.xml
      And word/_rels/document.xml.rels references a diagramData relationship


  Scenario: An unknown layout name is rejected up front
    Given a blank document for SmartArt authoring
     Then add_smart_art("matrix") raises ValueError
