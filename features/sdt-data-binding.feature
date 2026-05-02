Feature: Data binding on content controls
  In order to tie a content control to an XML data payload
  As a developer using python-docx
  I need read/write access to the w:dataBinding element on an SDT


  Scenario: ContentControl.data_binding is None when unbound
    Given a default document
     When I assign cc = document.add_content_control(PLAIN_TEXT)
     Then cc.data_binding is None


  Scenario: set_data_binding() returns a DataBinding with the requested attributes
    Given a default document
     When I assign cc = document.add_content_control(PLAIN_TEXT)
      And I call cc.set_data_binding(xpath, prefix_mappings, store_item_id)
     Then cc.data_binding is a DataBinding object
      And cc.data_binding.xpath is the xpath I supplied
      And cc.data_binding.prefix_mappings is the prefix_mappings I supplied
      And cc.data_binding.store_item_id is the store_item_id I supplied


  Scenario: remove_data_binding() unbinds the control
    Given a default document
     When I assign cc = document.add_content_control(PLAIN_TEXT)
      And I call cc.set_data_binding(xpath, prefix_mappings, store_item_id)
      And I call cc.remove_data_binding()
     Then cc.data_binding is None


  Scenario: Read data binding metadata from a saved document
    Given a document containing a data-bound content control
     Then the bound control's data_binding.xpath == "/ns0:order[1]/ns0:customer[1]"
      And the bound control's data_binding.prefix_mappings == "xmlns:ns0='http://example.com/orders'"
      And the bound control's data_binding.store_item_id == "{11111111-2222-3333-4444-555555555555}"


  Scenario: The matching custom XML data part is reachable via Document.custom_xml_parts
    Given a document containing a data-bound content control
     Then document has a custom XML part whose item_id matches the binding's store_item_id
      And that custom XML part's schema_refs include "http://example.com/orders"
      And that custom XML part's root_element is not None


  Scenario: Data binding survives a round-trip through save/load
    Given a document containing a data-bound content control
     When I save and reload the document
     Then the bound control's data_binding.xpath == "/ns0:order[1]/ns0:customer[1]"
      And the bound control's data_binding.store_item_id == "{11111111-2222-3333-4444-555555555555}"


  Scenario: Updating an existing binding overwrites its attributes
    Given a document containing a data-bound content control
     When I call cc.set_data_binding with new values
     Then cc.data_binding.xpath == "/a"
      And cc.data_binding.prefix_mappings == ""
      And cc.data_binding.store_item_id == "{22222222-3333-4444-5555-666666666666}"
