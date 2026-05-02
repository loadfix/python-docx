Feature: Read ruby (phonetic) annotations
  In order to inspect Japanese furigana and similar above-the-line annotations
  As a python-docx developer
  I need to enumerate RubyAnnotation objects on a run


  Scenario Outline: Count ruby annotations on a run
    Given a run from txt-ruby run <run_idx>
     Then len(run.ruby_annotations) is <count>

    Examples: ruby annotation counts
      | run_idx | count |
      | 0       | 2     |
      | 1       | 1     |


  Scenario Outline: Read ruby base and annotation text
    Given the ruby annotation at run <run_idx> position <pos> in txt-ruby
     Then ruby.base_text is <base>
      And ruby.ruby_text is <ruby>

    Examples: base/ruby text pairs
      | run_idx | pos | base    | ruby        |
      | 0       | 0   | 日本    | にほん      |
      | 0       | 1   | 東京    | とうきょう  |
      | 1       | 0   | <empty> | <empty>     |


  Scenario Outline: Read ruby alignment and language
    Given the ruby annotation at run <run_idx> position <pos> in txt-ruby
     Then ruby.alignment is <align>
      And ruby.language is <lang>

    Examples: alignment and language values
      | run_idx | pos | align            | lang  |
      | 0       | 0   | distributeSpace  | ja-JP |
      | 0       | 1   | None             | None  |
      | 1       | 0   | None             | None  |


  Scenario: Ruby base text contributes to run.text
    Given a run from txt-ruby run 0
     Then run.text contains the ruby base strings
