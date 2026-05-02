Feature: Stable element identifiers
  In order to correlate elements across a save-reload cycle
  As a developer using python-docx
  I need a stable id on paragraphs, runs, tables, and cells


  Scenario: Paragraph.stable_id returns a 16-char hex string
    Given a document containing five paragraphs
     Then each paragraph stable_id is a 16-char hex string


  Scenario: Run.stable_id returns a 16-char hex string
    Given a paragraph with three runs from par-multi.docx
     Then each run stable_id is a 16-char hex string


  Scenario: Table.stable_id returns a 16-char hex string
    Given a document containing three tables
     Then each table stable_id is a 16-char hex string


  Scenario: _Cell.stable_id returns a 16-char hex string
    Given a document containing three tables
     Then each cell stable_id is a 16-char hex string


  Scenario: Distinct paragraphs have distinct stable ids
    Given a document containing five paragraphs
     Then all paragraph stable_ids are unique


  Scenario: Stable id is deterministic across repeated access
    Given a document containing five paragraphs
     Then every paragraph stable_id is stable under repeated access
