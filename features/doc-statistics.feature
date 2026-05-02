Feature: Access document statistics
  In order to report word, character, and paragraph counts the way Word does
  As a developer using python-docx
  I need Document.statistics to return a DocumentStatistics summarizing the body


  Scenario Outline: Document.statistics returns counts matching Word's Word Count dialog
    Given a document with known body text
     When I access document.statistics
     Then statistics is a DocumentStatistics object
      And statistics.<field> == <expected>

    Examples: DocumentStatistics field values
      | field      | expected |
      | paragraphs | 2        |
      | words      | 4        |
      | characters | 33       |
