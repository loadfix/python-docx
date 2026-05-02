Feature: Transparently open Strict-OOXML packages
  In order to support documents produced by conformance-strict OOXML tools
  As a developer using python-docx
  I need Document() to rewrite Strict namespaces to Transitional on open


  Scenario: Loading a Strict-OOXML .docx succeeds
    Given a Strict-OOXML .docx document
     Then the document loads without error


  Scenario: Paragraph text is readable through the Strict translation layer
    Given a Strict-OOXML .docx document
     Then document.paragraphs is iterable


  Scenario: Saving a Strict-OOXML document emits Transitional output
    Given a Strict-OOXML .docx document
     When I save the document to the scratch path
     Then the saved package contains no Strict namespace URIs
