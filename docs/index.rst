
python-docx
===========

Release v\ |version| (:ref:`Installation <install>`)

*python-docx* is a Python library for creating and updating Microsoft Word
(.docx) files.


What it can do
--------------

.. |img| image:: /_static/img/example-docx-01.png

Here's an example of what |docx| can do:

============================================  ===============================================================
|img|                                         ::

                                                from docx import Document
                                                from docx.shared import Inches

                                                document = Document()

                                                document.add_heading('Document Title', 0)

                                                p = document.add_paragraph('A plain paragraph having some ')
                                                p.add_run('bold').bold = True
                                                p.add_run(' and some ')
                                                p.add_run('italic.').italic = True

                                                document.add_heading('Heading, level 1', level=1)
                                                document.add_paragraph('Intense quote', style='Intense Quote')

                                                document.add_paragraph(
                                                    'first item in unordered list', style='List Bullet'
                                                )
                                                document.add_paragraph(
                                                    'first item in ordered list', style='List Number'
                                                )

                                                document.add_picture('monty-truth.png', width=Inches(1.25))

                                                records = (
                                                    (3, '101', 'Spam'),
                                                    (7, '422', 'Eggs'),
                                                    (4, '631', 'Spam, spam, eggs, and spam')
                                                )

                                                table = document.add_table(rows=1, cols=3)
                                                hdr_cells = table.rows[0].cells
                                                hdr_cells[0].text = 'Qty'
                                                hdr_cells[1].text = 'Id'
                                                hdr_cells[2].text = 'Desc'
                                                for qty, id, desc in records:
                                                    row_cells = table.add_row().cells
                                                    row_cells[0].text = str(qty)
                                                    row_cells[1].text = id
                                                    row_cells[2].text = desc

                                                document.add_page_break()

                                                document.save('demo.docx')
============================================  ===============================================================


User Guide
----------

.. toctree::
   :maxdepth: 1

   user/install
   user/quickstart
   user/documents
   user/tables
   user/text
   user/sections
   user/hdrftr
   user/api-concepts
   user/styles-understanding
   user/styles-using
   user/comments
   user/bookmarks
   user/charts
   user/content-controls
   user/endnotes
   user/fields
   user/footnotes
   user/form-fields
   user/glossary
   user/shapes
   user/accessibility
   user/statistics


API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   api/document
   api/settings
   api/style
   api/text
   api/table
   api/section
   api/comments
   api/shape
   api/dml
   api/shared
   api/accessibility
   api/bookmarks
   api/captions
   api/chart
   api/content-controls
   api/custom-properties
   api/custom-xml
   api/embedded-objects
   api/endnotes
   api/equations
   api/fields
   api/font-table
   api/footnotes
   api/form-fields
   api/glossary
   api/ink
   api/numbering
   api/permissions
   api/ruby
   api/search
   api/signatures
   api/smart-art
   api/stable-ids
   api/statistics
   api/theme
   api/toc
   api/tracked-changes
   api/watermark
   api/web-settings
   api/enum/index


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/analysis/index
