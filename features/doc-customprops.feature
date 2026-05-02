Feature: Read and write custom document properties
  In order to store and retrieve typed metadata outside the core Dublin Core fields
  As a developer using python-docx
  I need a mapping-like Document.custom_properties collection


  Scenario: Document.custom_properties enumerates property names
    Given a document having known custom properties
     Then document.custom_properties has length 5
      And document.custom_properties names are ["Project", "Priority", "Budget", "Approved", "Reviewed"]


  Scenario Outline: Document.custom_properties[name] returns the typed value
    Given a document having known custom properties
     Then document.custom_properties["<name>"] is <value>

    Examples: supported value types
      | name     | value    |
      | Project  | Apollo   |
      | Priority | 5        |
      | Budget   | 99.95    |
      | Approved | True     |


  Scenario: Document.custom_properties reports membership
    Given a document having known custom properties
     Then "Project" is in document.custom_properties
      And "NoSuchProperty" is not in document.custom_properties


  Scenario: Document.custom_properties.get() returns None for missing keys
    Given a document having known custom properties
     Then document.custom_properties.get("Missing") is None


  Scenario: Overwriting an existing custom property
    Given a document having known custom properties
     When I assign document.custom_properties["Project"] = "Gemini"
     Then document.custom_properties["Project"] is Gemini


  Scenario: Adding a new custom property via add()
    Given a fresh default document
     When I call document.custom_properties.add("Owner", "Alice")
     Then document.custom_properties["Owner"] is Alice
      And document.custom_properties has length 1


  Scenario: Deleting a custom property
    Given a document having known custom properties
     When I delete document.custom_properties["Priority"]
     Then "Priority" is not in document.custom_properties
      And document.custom_properties has length 4


  Scenario: Assigning an unsupported type raises TypeError
    Given a fresh default document
     Then assigning a list to document.custom_properties raises TypeError
