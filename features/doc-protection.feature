Feature: Configure document protection and password hashing
  In order to prevent unauthorized modifications to a document when opened in Word
  As a developer using python-docx
  I need Settings.document_protection, Settings.enable_protection, Settings.disable_protection


  Scenario: Reading document_protection from a protected fixture
    Given a document having comments-only protection
     Then document_protection.mode is WD_PROTECTION.COMMENTS
      And document_protection.enforce is True
      And document_protection.password_hash is not None
      And document_protection.password_salt is not None
      And document_protection.spin_count == 100000


  Scenario: disable_protection clears enforce and removes the password
    Given a document having comments-only protection
     When I call settings.disable_protection()
     Then document_protection.enforce is False
      And document_protection.mode is None


  Scenario: enable_protection without a password omits the hash
    Given a fresh default document
     When I call settings.enable_protection(WD_PROTECTION.READ_ONLY, enforce=True)
     Then document_protection.mode is WD_PROTECTION.READ_ONLY
      And document_protection.password_hash is None


  Scenario: enable_protection with a password populates hash and salt
    Given a fresh default document
     When I call settings.enable_protection(WD_PROTECTION.COMMENTS, password="s3cret", enforce=True)
     Then document_protection.password_hash is not None
      And document_protection.password_salt is not None
      And document_protection.crypto_algorithm_sid == 4
      And document_protection.crypto_provider_type == "rsaAES"
