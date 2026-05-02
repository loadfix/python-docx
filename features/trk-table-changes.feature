Feature: Cell and row tracked changes
  In order to reconcile inserted, deleted, and reformatted table content
  As a developer using python-docx
  I need flags on _Cell and _Row surfacing w:cellIns, w:cellDel, w:tcPrChange, w:trPrChange, and w:tblPrChange


  Scenario: _Cell.is_tracked_insertion is True when w:cellIns is present
    Given the trk-table document
     Then cell (0, 1).is_tracked_insertion is True
      And cell (0, 1).is_tracked_deletion is False


  Scenario: _Cell.is_tracked_deletion is True when w:cellDel is present
    Given the trk-table document
     Then cell (1, 0).is_tracked_deletion is True
      And cell (1, 0).is_tracked_insertion is False


  Scenario: Plain cell flags both False
    Given the trk-table document
     Then cell (1, 1).is_tracked_insertion is False
      And cell (1, 1).is_tracked_deletion is False


  Scenario: _Cell.formatting_change exposes a w:tcPrChange
    Given the trk-table document
     Then cell (0, 0).formatting_change.author == "Alice"


  Scenario: _Row.formatting_change exposes a w:trPrChange
    Given the trk-table document
     Then row 1 formatting_change.author == "Dave"


  Scenario: Table.formatting_change exposes a w:tblPrChange
    Given the trk-table document
     Then table.formatting_change.author == "Eve"


  Scenario: Row without trPrChange yields None
    Given the trk-table document
     Then row 0 has no formatting_change
