Feature: Settings.mail_merge
  In order to inspect and configure the mail-merge metadata stored in a document
  As a developer using python-docx
  I need access to the MailMerge proxy on Settings
  And I need methods to enable and disable mail merge on a document

  Scenario: Settings.mail_merge returns None when no mailMerge element is present
    Given a document having no mail-merge configuration
     Then settings.mail_merge is None

  Scenario: Settings.enable_mail_merge() with defaults
    Given a document having no mail-merge configuration
     When I call settings.enable_mail_merge()
     Then settings.mail_merge is a MailMerge object
      And mail_merge.main_document_type == WD_MAIL_MERGE_TYPE.FORM_LETTERS
      And mail_merge.destination is None
      And mail_merge.data_type is None
      And mail_merge.connect_string is None
      And mail_merge.query is None

  Scenario: Settings.enable_mail_merge() with all arguments
    Given a document having no mail-merge configuration
     When I call settings.enable_mail_merge() with realistic arguments
     Then mail_merge.main_document_type == WD_MAIL_MERGE_TYPE.EMAIL
      And mail_merge.destination == WD_MAIL_MERGE_DESTINATION.EMAIL
      And mail_merge.data_type == WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET
      And mail_merge.connect_string is the supplied connect_string
      And mail_merge.query is the supplied query
      And mail_merge.mail_subject == "Quarterly update"
      And mail_merge.address_field_name == "Email"

  Scenario Outline: Read MailMerge property from an enabled document
    Given a document with mail-merge enabled
     Then mail_merge.<property> == <expected>

    Examples: MailMerge property values
      | property                     | expected                                                       |
      | main_document_type           | WD_MAIL_MERGE_TYPE.EMAIL                                       |
      | destination                  | WD_MAIL_MERGE_DESTINATION.EMAIL                                |
      | data_type                    | WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET                            |
      | connect_string               | Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx    |
      | query                        | SELECT FirstName, Email FROM [Sheet1$]                         |
      | mail_subject                 | Quarterly update                                               |
      | address_field_name           | Email                                                          |
      | active_record                | 3                                                              |
      | check_errors                 | None                                                           |
      | link_to_query                | True                                                           |
      | do_not_suppress_blank_lines  | False                                                          |
      | mail_as_attachment           | False                                                          |
      | view_merged_data             | True                                                           |

  Scenario: Settings.disable_mail_merge() removes the mail-merge element
    Given a document with mail-merge enabled
     When I call settings.disable_mail_merge()
     Then settings.mail_merge is None
