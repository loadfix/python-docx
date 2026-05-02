Feature: revision_marks_text() preview output
  In order to eyeball tracked changes from the command line
  As a developer using python-docx
  I need Paragraph.revision_marks_text() and Document.revision_marks_text() to produce a readable preview


  Scenario: A plain paragraph renders identically to .text
    Given the trk-marks document
     Then paragraph 0 revision_marks_text() matches paragraph.text


  Scenario: An insertion is wrapped with the default "[+ +]" markers
    Given the trk-marks document
     Then paragraph 1 revision_marks_text() == "Please [+kindly +]consider."


  Scenario: A deletion is wrapped with the default "[- -]" markers
    Given the trk-marks document
     Then paragraph 2 revision_marks_text() == "Delete [-this part -]of the text."


  Scenario: An interleaved insertion and deletion render in document order
    Given the trk-marks document
     Then paragraph 3 revision_marks_text() == "The [-old-][+new+] value."


  Scenario: Custom markers override the defaults
    Given the trk-marks document
     When I call paragraph 1 revision_marks_text with custom <INS>/<DEL> markers
     Then the custom-marker preview == "Please <INS>kindly </INS>consider."


  Scenario: document.revision_marks_text() joins paragraphs with blank lines
    Given the trk-marks document
     Then document.revision_marks_text() ends with the final paragraph's preview
      And document.revision_marks_text() contains the insertion-only preview
