Feature: Detect digital signatures attached to a package
  In order to report whether a document is signed and surface signer metadata
  As a developer using python-docx
  I need Document.is_signed and Document.signatures returning SignatureInfo objects


  Scenario: Document.is_signed is False for unsigned packages
    Given a fresh default document
     Then document.is_signed is False
      And document.signatures is empty


  Scenario: Document.is_signed is True for a signed fixture
    Given a signed document
     Then document.is_signed is True
      And document.signatures has length 1


  Scenario: SignatureInfo exposes the signer subject name
    Given a signed document
     Then signatures[0].signer == "CN=Alice Example"


  Scenario: SignatureInfo exposes the signing time from the XAdES block
    Given a signed document
     Then signatures[0].signed_at == 2024-04-01T12:34:56Z


  Scenario: SignatureInfo.partname targets the sigN.xml part
    Given a signed document
     Then signatures[0].partname == "/_xmlsignatures/sig1.xml"
