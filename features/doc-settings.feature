Feature: Document.settings
  In order to operate on document-level settings
  As a developer using python-docx
  I need access to settings to the Settings object for the document
  And I need properties and methods on Settings


  Scenario Outline: Access document settings
    Given a document having <a-or-no> settings part
     Then document.settings is a Settings object

    Examples: having a settings part or not
      | a-or-no   |
      | a         |
      | no        |


  Scenario Outline: Settings.odd_and_even_pages_header_footer getter
    Given a Settings object <with-or-without> odd and even page headers as settings
     Then settings.odd_and_even_pages_header_footer is <value>

    Examples: Settings.odd_and_even_pages_header_footer states
      | with-or-without | value |
      | with            | True  |
      | without         | False |


  Scenario Outline: Settings.odd_and_even_pages_header_footer setter
    Given a Settings object <with-or-without> odd and even page headers as settings
     When I assign <value> to settings.odd_and_even_pages_header_footer
     Then settings.odd_and_even_pages_header_footer is <value>

    Examples: Settings.odd_and_even_pages_header_footer assignment cases
      | with-or-without | value |
      | with            | True  |
      | with            | False |
      | without         | True  |
      | without         | False |


  Scenario Outline: Settings.compat_flags getter for well-known flags
    Given a Settings object loaded from doc-compat
     Then settings.compat_flags[<flag>] is <value>

    Examples: well-known compat flag presence
      | flag                       | value |
      | "growAutofit"              | True  |
      | "doNotBreakWrappedTables"  | True  |
      | "useFELayout"              | True  |
      | "noTabHangInd"             | False |


  Scenario: Settings.compat_flags iteration and membership
    Given a Settings object loaded from doc-compat
     Then "growAutofit" is in settings.compat_flags
      And "noTabHangInd" is not in settings.compat_flags
      And len(settings.compat_flags) is 3


  Scenario: Settings.compat_flags.names() lists well-known flags
    Given a Settings object loaded from doc-compat
     Then CompatFlags.names() contains "growAutofit"
      And CompatFlags.names() contains "doNotBreakWrappedTables"


  Scenario: Settings.compat_flags round-trip set and clear
    Given a Settings object loaded from doc-word-default-blank
     When I assign True to settings.compat_flags["growAutofit"]
      And I assign False to settings.compat_flags["useFELayout"]
     Then settings.compat_flags["growAutofit"] is True
      And settings.compat_flags["useFELayout"] is False


  Scenario Outline: Settings.compat_settings getter for well-known keys
    Given a Settings object loaded from doc-compat
     Then settings.compat_settings[<name>] is <value>

    Examples: compatSetting entries
      | name                                    | value |
      | "compatibilityMode"                     | "15"  |
      | "differentiateMultirowTableHeaders"     | "1"   |
      | "useWord2013TrackBottomHyphenation"     | "1"   |


  Scenario: Settings.compat_settings round-trip assignment
    Given a Settings object loaded from doc-word-default-blank
     When I assign "1" to settings.compat_settings["noTabHangInd"]
     Then settings.compat_settings["noTabHangInd"] is "1"


  Scenario Outline: Settings.view getter
    Given a Settings object loaded from <testfile>
     Then settings.view is <value>

    Examples: Settings.view states
      | testfile               | value           |
      | doc-view               | WD_VIEW.OUTLINE |
      | doc-word-default-blank | None            |


  Scenario Outline: Settings.view setter
    Given a Settings object loaded from doc-word-default-blank
     When I assign <value> to settings.view
     Then settings.view is <value>

    Examples: Settings.view assignment cases
      | value           |
      | WD_VIEW.PRINT   |
      | WD_VIEW.WEB     |
      | WD_VIEW.NORMAL  |
      | WD_VIEW.OUTLINE |
      | WD_VIEW.READING |
      | None            |


  Scenario Outline: Settings.zoom_percent getter
    Given a Settings object loaded from <testfile>
     Then settings.zoom_percent is <value>

    Examples: Settings.zoom_percent states
      | testfile               | value |
      | doc-view               | 175   |
      | doc-word-default-blank | 130   |


  Scenario Outline: Settings.zoom_percent setter
    Given a Settings object loaded from doc-word-default-blank
     When I assign <value> to settings.zoom_percent
     Then settings.zoom_percent is <value>

    Examples: Settings.zoom_percent assignment cases
      | value |
      | 80    |
      | 200   |
      | None  |


  Scenario Outline: Settings.track_revisions getter
    Given a Settings object loaded from <testfile>
     Then settings.track_revisions is <value>

    Examples: Settings.track_revisions states
      | testfile               | value |
      | doc-view               | True  |
      | doc-word-default-blank | False |


  Scenario: Settings.track_revisions setter
    Given a Settings object loaded from doc-word-default-blank
     When I assign True to settings.track_revisions
     Then settings.track_revisions is True


  Scenario Outline: Settings.compatibility_mode getter
    Given a Settings object loaded from <testfile>
     Then settings.compatibility_mode is <value>

    Examples: Settings.compatibility_mode states
      | testfile               | value |
      | doc-odd-even-hdrs      | 15    |


  Scenario: Settings.default_tab_stop round-trip with a Length value
    Given a Settings object loaded from doc-word-default-blank
     When I assign Twips(720) to settings.default_tab_stop
     Then settings.default_tab_stop is 457200


  Scenario: Settings.default_tab_stop clear
    Given a Settings object loaded from doc-word-default-blank
     When I assign None to settings.default_tab_stop
     Then settings.default_tab_stop is None
