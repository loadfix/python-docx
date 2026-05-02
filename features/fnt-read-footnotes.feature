Feature: Read footnotes from a document
  In order to inspect the footnotes already present in a document
  As a developer using python-docx
  I need to iterate Document.footnotes and read each footnote's properties


  Scenario Outline: Access document footnotes
    Given a document having <a-or-no> footnotes part
     Then document.footnotes is a Footnotes object

    Examples: having a footnotes part or not
      | a-or-no |
      | a       |
      | no      |


  Scenario Outline: Footnotes.__len__()
    Given a document having <count> footnotes
     Then len(document.footnotes) == <count>

    Examples: len(document.footnotes) values
      | count |
      | 0     |
      | 3     |


  Scenario: Footnotes.__iter__() yields user footnotes only
    Given a document having 3 footnotes
     Then iterating document.footnotes yields 3 Footnote objects
      And the separator and continuation-separator footnotes are not yielded


  Scenario: Footnote.footnote_id
    Given a document having 3 footnotes
     Then the yielded footnote ids are [2, 3, 4]


  Scenario: Footnote.text
    Given a document having 3 footnotes
     Then footnote with id 2 has text "A common saying about Iberian weather."
      And footnote with id 3 has text "As of the loadfix fork."


  Scenario: Footnote.paragraphs
    Given a document having 3 footnotes
     Then each footnote has one paragraph
      And each footnote paragraph has the FootnoteText style
