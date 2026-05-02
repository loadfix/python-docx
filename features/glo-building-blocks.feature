Feature: Iterate and filter building blocks
  In order to discover the AutoText / Quick Parts / cover-page content stored
  in the glossary-document part
  As a developer using python-docx
  I need access to each building block's metadata and content, and a way to
  filter the collection by gallery and/or category name


  Scenario: BuildingBlock.name
    Given a BuildingBlock object named "Alpha"
     Then building_block.name == "Alpha"


  Scenario: BuildingBlock.description
    Given a BuildingBlock object named "Alpha"
     Then building_block.description == "a quick-parts alpha block"


  Scenario: BuildingBlock.guid
    Given a BuildingBlock object named "Alpha"
     Then building_block.guid is a non-empty string


  Scenario: BuildingBlock.category exposes name and gallery
    Given a BuildingBlock object named "Alpha"
     Then building_block.category.category_name == "General"
      And building_block.category.gallery == "quickParts"


  Scenario: BuildingBlockCategory.gallery_value maps to the enum
    Given a BuildingBlock object named "Alpha"
     Then building_block.category.gallery_value is WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS


  Scenario: BuildingBlock exposes its body paragraphs
    Given a BuildingBlock object named "Alpha"
     Then building_block.paragraphs[0].text == "Alpha body"


  Scenario: BuildingBlock exposes its body tables
    Given a BuildingBlock object named "Gamma"
     Then len(building_block.tables) == 1


  Scenario: A building block without a body returns empty paragraph and table lists
    Given a BuildingBlock object named "Epsilon"
     Then building_block.paragraphs == []
      And building_block.tables == []


  Scenario: A building block without a category returns a proxy with None slots
    Given a BuildingBlock object named "Epsilon"
     Then building_block.category.category_name is None
      And building_block.category.gallery is None


  Scenario Outline: Filter building blocks by gallery enum member
    Given a Glossary object with 5 building blocks
     When I call glossary.by_category(gallery=<gallery>)
     Then the result names are <names>

    Examples: filter by gallery
      | gallery                                 | names          |
      | WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS   | Alpha, Beta    |
      | WD_BUILDING_BLOCK_GALLERY.COVER_PAGES   | Gamma          |
      | WD_BUILDING_BLOCK_GALLERY.HEADERS       | Delta          |


  Scenario: Filter building blocks by raw gallery XML string
    Given a Glossary object with 5 building blocks
     When I call glossary.by_category(gallery="quickParts")
     Then the result names are Alpha, Beta


  Scenario: Filter building blocks by category name
    Given a Glossary object with 5 building blocks
     When I call glossary.by_category(category_name="Built-In")
     Then the result names are Gamma, Delta


  Scenario: Filter by gallery and category name intersects
    Given a Glossary object with 5 building blocks
     When I call glossary.by_category(gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS, category_name="General")
     Then the result names are Alpha, Beta


  Scenario: Filter by mismatched gallery and category name returns empty
    Given a Glossary object with 5 building blocks
     When I call glossary.by_category(gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS, category_name="Built-In")
     Then the result is an empty list


  Scenario: Glossary.categories deduplicates by (gallery, category_name)
    Given a Glossary object with 5 building blocks
     Then glossary.categories has 3 entries with keys (quickParts, General), (coverPg, Built-In), (hdrs, Built-In)


  Scenario: Glossary.galleries returns unique gallery strings in first-seen order
    Given a Glossary object with 5 building blocks
     Then glossary.galleries == ["quickParts", "coverPg", "hdrs"]
