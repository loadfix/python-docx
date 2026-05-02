Feature: FormattingChange for rPrChange / pPrChange / sectPrChange
  In order to inspect formatting revisions recorded by Word's track-changes
  As a developer using python-docx
  I need FormattingChange proxies on runs, paragraphs, and sections


  Scenario: Run.formatting_change exposes a rPrChange
    Given the trk-format document
     When I select the formatting_change of run 0 on paragraph 1
     Then the formatting change is not None
      And formatting_change.author == "Alice"
      And formatting_change.old_properties is not None


  Scenario: Paragraph.formatting_change exposes a pPrChange
    Given the trk-format document
     When I select the formatting_change of paragraph 1
     Then the formatting change is not None
      And formatting_change.author == "Bob"
      And formatting_change.old_properties is not None


  Scenario: Section.formatting_change exposes a sectPrChange
    Given the trk-format document
     When I select the formatting_change of section 0
     Then the formatting change is not None
      And formatting_change.author == "Carol"
      And formatting_change.old_properties is not None


  Scenario: Paragraph with no formatting change yields None
    Given the trk-format document
     When I select the formatting_change of paragraph 2
     Then the formatting change is None
