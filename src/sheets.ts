export const SheetUtils = {
  getOrCreateSheet(name: string) {
    const spreadsheet = SpreadsheetApp.getActive();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
    }
    return sheet;
  },

  /**
   * Applies a "drop-down" validation to the given range.
   * @param range the range to turn into a drop-down
   * @param values the drop-values
   */
  setRangeDropDown(
    range: GoogleAppsScript.Spreadsheet.Range,
    values: string[]
  ) {
    range.clearDataValidations();
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .build();
    range.setDataValidation(validation);
  }
};
