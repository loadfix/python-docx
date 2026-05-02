Feature: Read and write accessibility alt text on shapes
  In order to publish accessible documents
  As a developer using python-docx
  I need a way to get and set ``alt_text`` and ``title`` on pictures


  Scenario: Inline picture exposes its alt text and title
    Given a document known to contain an inline picture with alt text
     Then the first inline shape alt_text is "A pencil-drawing of a mountain peak"
      And the first inline shape title is "Mountain peak"


  Scenario: Inline picture without alt text returns None
    Given a document known to contain an inline picture with alt text
     Then the second inline shape alt_text is None
      And the second inline shape title is None


  Scenario: Floating picture exposes its alt text and title
    Given a document known to contain an inline picture with alt text
     Then the first floating picture alt_text is "Decorative floating mountain"
      And the first floating picture title is "Floating mountain"


  Scenario: Floating picture without alt text returns None
    Given a document known to contain an inline picture with alt text
     Then the second floating picture alt_text is None
      And the second floating picture title is None


  Scenario: Assigning alt text updates the attribute on inline shapes
    Given an inline shape of known dimensions
     When I set the inline shape's alt_text to "A new description"
     Then the inline shape alt_text is "A new description"
      When I set the inline shape's alt_text to None
      Then the inline shape alt_text is None


  Scenario: Assigning title updates the attribute on inline shapes
    Given an inline shape of known dimensions
     When I set the inline shape's title to "A caption"
     Then the inline shape title is "A caption"
      When I set the inline shape's title to None
      Then the inline shape title is None
