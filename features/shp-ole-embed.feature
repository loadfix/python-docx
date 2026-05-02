Feature: Read embedded OLE objects from a document
  In order to inspect embedded workbooks, PDFs, equations, etc.
  As a developer using python-docx
  I need read access to ``<w:object>/<o:OLEObject>`` references and their parts


  Scenario: Enumerate embedded objects at the document level
    Given a document known to contain an embedded OLE object
     Then the document exposes 2 embedded objects
      And the resolved embedded object has prog_id "Excel.Sheet.12"
      And the resolved embedded object has type "Embed"
      And the resolved embedded object blob is non-empty
      And the resolved embedded object has embedded_partname "/word/embeddings/oleObject1.bin"


  Scenario: Unresolved OLE reference reports empty blob
    Given a document known to contain an embedded OLE object
     Then the unresolved embedded object has prog_id "AcroExch.Document"
      And the unresolved embedded object has type "Link"
      And the unresolved embedded object blob is empty
      And the unresolved embedded object has embedded_partname None


  Scenario: Embedded objects accessible via containing paragraph
    Given a document known to contain an embedded OLE object
     Then the first paragraph carrying an OLE reference has 1 embedded object
      And the embedded object paragraph attribute is the paragraph that contains it


  Scenario: Document without OLE has an empty embedded_objects list
    Given a pristine empty document
     Then the document exposes 0 embedded objects
