Feature: Structured document tag (content control) support
  In order to author and inspect Word content controls from Python
  As a developer using python-docx
  I need to create, read, and iterate block-level and inline content controls


  Scenario Outline: Document.add_content_control() for each SDT type
    Given a default document
     When I assign cc = document.add_content_control(<type>, tag="t", title="T")
     Then cc is a ContentControl object
      And cc.type is ContentControlType.<type>
      And cc.tag == "t"
      And cc.title == "T"
      And cc.sdt_id is a positive integer

    Examples: block-level types
      | type       |
      | RICH_TEXT  |
      | PLAIN_TEXT |
      | DATE       |
      | CHECKBOX   |
      | COMBO_BOX  |
      | DROPDOWN   |
      | PICTURE    |


  Scenario: Paragraph.add_content_control() adds an inline SDT
    Given a default document with a paragraph "Hello "
     When I assign cc = paragraph.add_content_control(RICH_TEXT, tag="inline")
      And I assign cc.text = "world"
     Then cc is a ContentControl object
      And cc.type is ContentControlType.RICH_TEXT
      And cc.tag == "inline"
      And cc.text == "world"
      And paragraph.content_controls[0] == cc


  Scenario: Iterate Document.content_controls in document order
    Given a document having 7 block-level content controls
     Then document.content_controls yields 7 ContentControl objects
      And the control tags are "rich-1, plain-1, date-1, cbx-1, cmb-1, dd-1, pic-1"
      And the control types are "RICH_TEXT, PLAIN_TEXT, DATE, CHECKBOX, COMBO_BOX, DROPDOWN, PICTURE"


  Scenario: Read a checkbox content control
    Given a document having 7 block-level content controls
     When I read the CHECKBOX control
     Then cc.checked is True


  Scenario: Write and re-read a checkbox content control
    Given a default document
     When I assign cc = document.add_content_control(CHECKBOX, tag="ok")
      And I assign cc.checked = True
     Then cc.checked is True


  Scenario: Inline content controls are exposed on Paragraph
    Given a document having 7 block-level content controls
     Then the final paragraph has 1 inline content control
      And the inline control's tag == "inline-rt"
      And the inline control's text == "inline value"


  Scenario: Read a rich-text control's text
    Given a document having 7 block-level content controls
     When I read the RICH_TEXT control
     Then cc.text == "Rich text body"


  Scenario: Setting cc.text replaces the content
    Given a default document
     When I assign cc = document.add_content_control(PLAIN_TEXT, tag="p")
      And I assign cc.text = "hello"
     Then cc.text == "hello"
      And cc.text is a single line


  Scenario: Default rich-text SDT has no explicit marker
    Given a default document
     When I assign cc = document.add_content_control(RICH_TEXT)
     Then cc.type is ContentControlType.RICH_TEXT
