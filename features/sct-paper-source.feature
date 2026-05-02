Feature: Section paper source (printer tray)
  In order to hint printer-tray selection per section
  As a developer using python-docx
  I need read/write access to first_page_paper_source and other_pages_paper_source


  Scenario: No paperSrc element present
    Given a Section with no paperSrc as section
     Then section.first_page_paper_source is None
      And section.other_pages_paper_source is None


  Scenario: Both first and other attributes set
    Given a Section with first=7 and other=15 paperSrc as section
     Then section.first_page_paper_source is 7
      And section.other_pages_paper_source is 15


  Scenario: Only first attribute set
    Given a Section with first=1 only paperSrc as section
     Then section.first_page_paper_source is 1
      And section.other_pages_paper_source is None


  Scenario: Only other attribute set
    Given a Section with other=2 only paperSrc as section
     Then section.first_page_paper_source is None
      And section.other_pages_paper_source is 2


  Scenario: Assigning creates paperSrc element
    Given a Section with no paperSrc as section
     When I assign 3 to section.first_page_paper_source
     Then section.first_page_paper_source is 3
      And section.other_pages_paper_source is None


  Scenario: Setting other preserves existing first
    Given a Section with first=1 only paperSrc as section
     When I assign 9 to section.other_pages_paper_source
     Then section.first_page_paper_source is 1
      And section.other_pages_paper_source is 9


  Scenario: Clearing last attribute removes paperSrc element
    Given a Section with first=1 only paperSrc as section
     When I assign None to section.first_page_paper_source
     Then section.first_page_paper_source is None
      And section has no paperSrc element


  Scenario: Clearing first while other remains keeps paperSrc
    Given a Section with first=7 and other=15 paperSrc as section
     When I assign None to section.first_page_paper_source
     Then section.first_page_paper_source is None
      And section.other_pages_paper_source is 15


  Scenario: Clearing when not present is a no-op
    Given a Section with no paperSrc as section
     When I assign None to section.first_page_paper_source
     Then section has no paperSrc element
