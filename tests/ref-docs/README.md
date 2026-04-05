# Reference Documents

This directory contains reference `.docx` files created in Microsoft Word for use in
testing. These files serve as ground truth for validating that python-docx can correctly
read documents produced by Word.

## How to Use

Reference files are used in Layer 4 (Reference File Comparison) tests. Test code reads
these files with python-docx and asserts the parsed content matches expectations.

```python
from docx import Document
from tests.helpers.refcmp import ref_docx_path

def it_reads_a_word_comments_file():
    doc = Document(ref_docx_path("comments-simple"))
    comments = doc.comments
    assert len(comments) == 1
    assert comments.get(0).author == "John Doe"
```

## Reference Files

### comments-simple.docx (planned)
- One comment on a single word
- Author: "John Doe", Initials: "JD"
- Comment text: "This is a simple comment."

### comments-threaded.docx (planned)
- Parent comment with 2 reply comments
- Multiple authors
- Demonstrates reply threading via `w16cid:paraIdParent`

### comments-multi-author.docx (planned)
- Comments by 3 different authors
- Each with distinct initials

### comments-formatted.docx (planned)
- Comment containing bold and italic text
- Comment containing multiple paragraphs

## Creating Reference Files

1. Open Microsoft Word (any recent version)
2. Create the document content described above
3. Save as `.docx` format
4. Place the file in this directory
5. Update this README with the actual content description

These files are committed to the repository and should only be recreated when
the expected content changes.
