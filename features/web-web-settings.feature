Feature: Read and update the web-settings part
  In order to control how Word saves the document as a web page
  As a developer using python-docx
  I need Document.web_settings returning a WebSettings proxy


  Scenario: Document.web_settings is not None for the default template
    Given a fresh default document
     Then document.web_settings is not None


  Scenario: Reading writable flags from a fixture
    Given a document having non-default web settings
     Then web_settings.optimize_for_browser is True
      And web_settings.allow_png is True
      And web_settings.do_not_save_as_single_file is True


  Scenario Outline: Toggling a web-settings flag persists through a round-trip
    Given a fresh default document
     When I assign web_settings.<flag> = True
     Then web_settings.<flag> is True

    Examples: writable web-settings flags
      | flag                       |
      | optimize_for_browser       |
      | allow_png                  |
      | do_not_save_as_single_file |


  Scenario: Clearing a web-settings flag with None removes the XML element
    Given a document having non-default web settings
     When I assign web_settings.optimize_for_browser = None
     Then web_settings.optimize_for_browser is False
