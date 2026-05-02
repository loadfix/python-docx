Feature: Access and mutate legacy form fields
  In order to work with FORMTEXT, FORMCHECKBOX, and FORMDROPDOWN fields in a document
  As a developer using python-docx
  I need to read, iterate, mutate, and round-trip legacy form fields


  Scenario: Document.form_fields iterates all form fields in body order
    Given a document having 3 legacy form fields
     Then document.form_fields is a list of 3 FormField objects
      And the form fields are returned in document order


  Scenario: Read properties of a text-input form field
    Given a document having 3 legacy form fields
     When I select the text-input form field
     Then form_field.type is WD_FORM_FIELD_TYPE.TEXT
      And form_field.name == "FullName"
      And form_field.enabled is True
      And form_field.text_input.default == "Jane Doe"
      And form_field.text_input.max_length == 40
      And form_field.value == "Jane Doe"
      And form_field.checkbox is None
      And form_field.dropdown is None


  Scenario: Read properties of a checkbox form field
    Given a document having 3 legacy form fields
     When I select the checkbox form field
     Then form_field.type is WD_FORM_FIELD_TYPE.CHECKBOX
      And form_field.name == "Subscribe"
      And form_field.checkbox.checked is True
      And form_field.value is True
      And form_field.text_input is None


  Scenario: Read properties of a dropdown form field
    Given a document having 3 legacy form fields
     When I select the dropdown form field
     Then form_field.type is WD_FORM_FIELD_TYPE.DROPDOWN
      And form_field.name == "Country"
      And form_field.dropdown.options == ["US", "UK", "AU"]
      And form_field.dropdown.default_index == 1
      And form_field.value == "UK"


  Scenario: Add form fields and round-trip through save/load
    Given a freshly created document
     When I append a text form field named "Color" with default "blue"
      And I append a checkbox form field named "Agree" checked False
      And I append a dropdown form field named "Size" with options ["S", "M", "L"] default 2
      And I save and re-open the document
     Then document.form_fields is a list of 3 FormField objects
      And the text form field's default is "blue"
      And the checkbox form field is unchecked
      And the dropdown form field's selected value is "L"
