# Reference .docx Files

This directory contains reference `.docx` files used for Layer 4 (Reference File
Comparison) testing. These files are created in Microsoft Word and committed to the
repository as test fixtures.

## Files

### `comments-simple.docx`
A Word document containing basic comment scenarios:
- **Comment 0**: Simple comment on a single word ("Hello") by author "Author A" with
  initials "AA"
- **Comment 1**: Comment on a full paragraph by author "Author B" with initials "AB"
- The document body contains two paragraphs with runs that have comment range markers

Created in Microsoft Word to serve as the canonical reference for how Word structures
comment XML. Used to verify that python-docx produces compatible output and can read
Word-produced comments correctly.

## How to Add Reference Files

1. Create the document in Microsoft Word with the specific feature scenarios
2. Save as `.docx` format
3. Add the file to this directory
4. Document what the file contains in this README
5. Write tests in `tests/test_strategy_integration.py` that read and verify the file

## Notes

- Do NOT modify these files programmatically — they must remain as Word produced them
- If a reference file needs updating, re-create it in Word and replace the file
- Each file should test a focused set of scenarios for one feature area
