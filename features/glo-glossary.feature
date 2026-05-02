Feature: Read a document's glossary
  In order to inspect the AutoText / Quick Parts / cover-page building blocks
  As a developer using python-docx
  I need read-only access to the glossary-document part and its building blocks


  Scenario Outline: Access document.glossary
    Given a document having <a-or-no> glossary part
     Then document.glossary is <expected>

    Examples: glossary proxy availability
      | a-or-no | expected         |
      | a       | a Glossary object |
      | no      | None              |


  Scenario: Glossary.__len__()
    Given a Glossary object with 5 building blocks
     Then len(glossary) == 5


  Scenario: Glossary.__iter__()
    Given a Glossary object with 5 building blocks
     Then iterating glossary yields 5 BuildingBlock objects


  Scenario: Glossary.building_blocks preserves document order
    Given a Glossary object with 5 building blocks
     Then glossary.building_blocks names are Alpha, Beta, Gamma, Delta, Epsilon


  Scenario: Glossary.__getitem__() looks up a block by name
    Given a Glossary object with 5 building blocks
     When I call glossary["Gamma"]
     Then the result is a BuildingBlock named "Gamma"


  Scenario: Glossary.__getitem__() raises KeyError for an unknown name
    Given a Glossary object with 5 building blocks
     Then glossary["NoSuchBlock"] raises KeyError
