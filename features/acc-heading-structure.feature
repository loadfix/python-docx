Feature: Validate document heading structure
  In order to author accessible documents
  As a developer using python-docx
  I need to detect common heading-outline issues such as missing or skipped headings


  Scenario Outline: Document.validate_heading_structure()
    Given a document with <outline> heading outline
     When I call document.validate_heading_structure()
     Then the result is a list of <expected-issues> HeadingIssue objects
      And the first reported issue has kind "<first-kind>"

    Examples: heading-outline cases
      | outline         | expected-issues | first-kind     |
      | a valid         | 0               |                |
      | a missing-H2    | 1               | no_h1          |
      | a skipped-level | 1               | skipped_level  |
