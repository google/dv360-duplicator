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
  },

  /**
   * Show the progress/status (e.g. 'Loading...') message to the user
   * 
   * @param message Text inside the message box
   * @param title Title of the message box
   */
  showProgressMessage(
    message = 'Loading...',
    title = 'Please wait',
  ) {
    SpreadsheetApp.getUi()
      .showModelessDialog(
        HtmlService.createHtmlOutput(message)
          .setWidth(250)
          .setHeight(90),
        title
      );
  },

  /**
   * Hide previosly created progress message
   * 
   * @param title Title of the message box. Will be shown for very short period 
   *  of time
   */
  hideProgressMessage(title = 'Done') {
    const hl = '<!DOCTYPE html><html><head><base target="_top"></head><script>'
      + 'window.onload=()=>{google.script.host.close();}'
      + '</script><body></body></html>';
    SpreadsheetApp.getUi().showModelessDialog(
      HtmlService.createHtmlOutput(hl), 
      title
    )
  }
};
