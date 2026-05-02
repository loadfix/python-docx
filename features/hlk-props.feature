Feature: Access hyperlink properties
  In order to access the URL and other details for a hyperlink
  As a developer using python-docx
  I need properties on Hyperlink


  Scenario: Hyperlink.address has the URL of the hyperlink
    Given a hyperlink
     Then hyperlink.address is the URL of the hyperlink


  Scenario Outline: Hyperlink.contains_page_break reports presence of page-break
    Given a hyperlink having <zero-or-more> rendered page breaks
     Then hyperlink.contains_page_break is <value>

    Examples: Hyperlink.contains_page_break cases
      | zero-or-more | value |
      | no           | False |
      | one          | True  |


  Scenario: Hyperlink.fragment has the URI fragment of the hyperlink
    Given a hyperlink having a URI fragment
     Then hyperlink.fragment is the URI fragment of the hyperlink


  Scenario Outline: Hyperlink.runs contains Run for each run in hyperlink
    Given a hyperlink having <zero-or-more> runs
     Then hyperlink.runs has length <value>
      And hyperlink.runs contains only Run instances

    Examples: Hyperlink.runs cases
      | zero-or-more | value |
      | one          |   1   |
      | two          |   2   |


  Scenario: Hyperlink.text has the visible text of the hyperlink
    Given a hyperlink
     Then hyperlink.text is the visible text of the hyperlink


  Scenario Outline: Hyperlink.url is the full URL of an internet hyperlink
    Given a hyperlink having address <address> and fragment <fragment>
     Then hyperlink.url is <url>

    Examples: Hyperlink.url cases
      | address                   | fragment       | url                       |
      | ''                        | linkedBookmark | ''                        |
      | https://foo.com           | ''             | https://foo.com           |
      | https://foo.com?q=bar     | ''             | https://foo.com?q=bar     |
      | http://foo.com/           | intro          | http://foo.com/#intro     |
      | https://foo.com?q=bar#baz | ''             | https://foo.com?q=bar#baz |
      | court-exif.jpg            | ''             | court-exif.jpg            |


  Scenario: paragraph.add_hyperlink(url=...) creates an external hyperlink
    Given a fresh paragraph in a default document
     When I call paragraph.add_hyperlink(url="https://example.com", text="link")
     Then the returned hyperlink.address is "https://example.com"
      And the returned hyperlink.text is "link"


  Scenario: paragraph.add_hyperlink(url=..., text=None) uses url as display text
    Given a fresh paragraph in a default document
     When I call paragraph.add_hyperlink(url="https://example.com")
     Then the returned hyperlink.text is "https://example.com"


  Scenario: paragraph.add_hyperlink(anchor=...) creates an internal link
    Given a fresh paragraph in a default document
     When I call paragraph.add_hyperlink(anchor="intro", text="Intro")
     Then the returned hyperlink.fragment is "intro"
      And the returned hyperlink.text is "Intro"


  Scenario: paragraph.add_hyperlink without url or anchor raises ValueError
    Given a fresh paragraph in a default document
     Then calling paragraph.add_hyperlink() raises ValueError


  Scenario: paragraph.add_hyperlink with both url and anchor raises ValueError
    Given a fresh paragraph in a default document
     Then calling paragraph.add_hyperlink(url="x", anchor="y") raises ValueError
