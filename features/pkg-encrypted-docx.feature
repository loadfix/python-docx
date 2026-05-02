Feature: Detect password-encrypted .docx files
  In order to fail loudly with an actionable message instead of a confusing zipfile error
  As a developer using python-docx
  I need Document() to raise EncryptedDocumentError for OLE-compound-file packages


  Scenario: Opening an OLE-compound-file package raises EncryptedDocumentError
    Given an OLE compound file masquerading as a .docx
     Then Document(path) raises EncryptedDocumentError


  Scenario: Recover mode still raises EncryptedDocumentError (not XMLSyntaxError)
    Given an OLE compound file masquerading as a .docx
     Then Document(path, recover=True) raises EncryptedDocumentError
