Feature: Open macro-enabled .docm documents and detect their VBA project
  In order to let callers decide whether to trust macro-enabled documents before processing them
  As a developer using python-docx
  I need Document() to load the macroEnabled content type and Document.has_macros to detect VBA


  Scenario: Loading a .docm document succeeds
    Given a macro-enabled .docm document
     Then the document loads without error


  Scenario: Document.has_macros reports presence of a vbaProject relationship
    Given a macro-enabled .docm document
     Then document.has_macros is True


  Scenario: Document.has_macros is False for plain .docx
    Given a fresh default document
     Then document.has_macros is False
