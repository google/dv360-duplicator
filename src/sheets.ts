/**
    Copyright 2023 Google LLC
    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at
        https://www.apache.org/licenses/LICENSE-2.0
    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.
*/

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
