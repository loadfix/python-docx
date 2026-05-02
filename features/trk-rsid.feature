Feature: Revision-save IDs (rsidRoot + per-run/paragraph rsid)
  In order to correlate edits back to the Word editing sessions that produced them
  As a developer using python-docx
  I need read access to document-level w:rsids and per-run/paragraph w:rsidR attributes


  Scenario: Document settings expose rsid_root and rsids
    Given the trk-rsid document
     Then document.settings.rsid_root == "00CAFE00"
      And document.settings.rsids == ['00A1B2C3', '00DEAD00', '00BEEF00']


  Scenario: Paragraph.rsid reflects the paragraph's w:rsidR
    Given the trk-rsid document
     Then paragraph 1 rsid == "00A1B2C3"
      And paragraph 3 rsid == "00BEEF00"


  Scenario: Paragraph.rsid is None when no w:rsidR is set
    Given the trk-rsid document
     Then paragraph 2 rsid is None


  Scenario: Run.rsid reflects the run's w:rsidR
    Given the trk-rsid document
     Then paragraph 1 run 0 rsid == "00DEAD00"


  Scenario: Run.rsid is None when no w:rsidR is set
    Given the trk-rsid document
     Then paragraph 3 run 0 rsid is None
