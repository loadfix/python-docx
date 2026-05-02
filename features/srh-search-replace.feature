Feature: Search and replace text
  In order to locate and transform text in Word documents
  As a developer using python-docx
  I need search and replace methods on Document


  Scenario: exact case-sensitive search in body paragraphs
    Given a Document loaded from srh-target-text.docx
     When I call document.search("SEARCH_IN_BODY")
     Then the result is a list of 3 SearchMatch objects
      And every match.location is None
      And every match.paragraph is a Paragraph
      And every match.text equals "SEARCH_IN_BODY"


  Scenario: case-insensitive search returns lowercased variants
    Given a Document loaded from srh-target-text.docx
     When I call document.search("search_in_body", case_sensitive=False)
     Then the result is a list of 4 SearchMatch objects


  Scenario: regex search matches invoice patterns
    Given a Document loaded from srh-target-text.docx
     When I call document.search_regex(r"INV-\d+")
     Then the result is a list of 2 SearchMatch objects
      And match_texts == ["INV-12345", "INV-99"]


  Scenario: plain replace preserves formatting of the first run
    Given a Document loaded from srh-target-text.docx
     When I call document.replace("SEARCH_IN_BODY", "REPLACED_TOKEN")
     Then the returned count is 3
      And document.search("SEARCH_IN_BODY") returns 0 matches
      And document.search("REPLACED_TOKEN") returns 3 matches


  Scenario: regex replace expands backreferences
    Given a Document loaded from srh-target-text.docx
     When I call document.replace_regex(r"INV-(\d+)", r"INVOICE[\1]")
     Then the returned count is 2
      And document.search("INVOICE[12345]") returns 1 matches
      And document.search("INVOICE[99]") returns 1 matches


  Scenario: multi-run match reports every spanned run index
    Given a Document loaded from srh-target-text.docx
     When I call document.search("SEARCH_IN_BODY")
     Then at least one match spans multiple runs


  Scenario: search_all finds text inside a body table cell
    Given a Document loaded from srh-target-text.docx
     When I call document.search_all("SEARCH_IN_TABLE")
     Then the result is a list of 1 SearchMatch objects
      And match.location starts with "table:"


  Scenario: search_all finds text in a section header
    Given a Document loaded from srh-target-text.docx
     When I call document.search_all("SEARCH_IN_HEADER")
     Then the result is a list of 1 SearchMatch objects
      And match.location starts with "header:"


  Scenario: search_all finds text in a section footer
    Given a Document loaded from srh-target-text.docx
     When I call document.search_all("SEARCH_IN_FOOTER")
     Then the result is a list of 1 SearchMatch objects
      And match.location starts with "footer:"


  Scenario: search_all finds text inside a footnote
    Given a Document loaded from srh-target-text.docx
     When I call document.search_all("SEARCH_IN_FOOTNOTE")
     Then the result is a list of 1 SearchMatch objects
      And match.location starts with "footnote:"


  Scenario: replace_all updates every story
    Given a Document loaded from srh-target-text.docx
     When I call document.replace_all("SEARCH_IN_FOOTER", "REPLACED_FOOTER")
     Then the returned count is 1
      And document.search_all("SEARCH_IN_FOOTER") returns 0 matches
      And document.search_all("REPLACED_FOOTER") returns 1 matches
