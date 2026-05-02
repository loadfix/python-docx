Feature: Embed an OLE object into a document
  In order to attach spreadsheets, PDFs and other binary payloads to a document
  As a developer using python-docx
  I need to programmatically add ``<w:object>/<o:OLEObject>`` references plus
  their embedded parts


  Scenario: Embed a fake xlsx and round-trip the binary
    Given a fresh empty document
     When I add an OLE object for an xlsx payload with prog_id "Excel.Sheet.12"
      And I save and reload the document to a BytesIO buffer
     Then the document exposes 1 embedded objects
      And the resolved embedded object has prog_id "Excel.Sheet.12"
      And the resolved embedded object has type "Embed"
      And the resolved embedded object blob round-trips the original bytes
