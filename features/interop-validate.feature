Feature: Round-trip fidelity via ooxml-validate
  In order to catch regressions in library output that would confuse Office
  As a developer on python-docx
  I need automated validation that the library's output is as clean as the input

  Scenario: Round-tripping a real Word document introduces no new validator issues
    Given a Word-authored docx fixture
    When I round-trip it through python-docx
    Then ooxml-validate reports zero new issues

  Scenario: The round-trip is visually identical under LibreOffice
    Given a Word-authored docx fixture
    When I round-trip it through python-docx
    Then the LibreOffice PDF render is visually identical
