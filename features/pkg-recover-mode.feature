Feature: Open a malformed .docx using recover mode
  In order to salvage readable content from a corrupted package
  As a developer using python-docx
  I need Document(path, recover=True) and Document.recovery_warnings


  Scenario: Default open raises for a malformed document.xml
    Given a malformed .docx package
     Then Document(path) raises XMLSyntaxError


  Scenario: Recover mode loads the readable prefix
    Given a malformed .docx package
     When I call Document(path, recover=True)
     Then document.recovery_warnings is non-empty
      And at least one paragraph text contains "Readable prefix paragraph."


  Scenario: Normally opened documents have empty recovery_warnings
    Given a fresh default document
     Then document.recovery_warnings is empty
