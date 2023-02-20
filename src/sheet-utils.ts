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

import { CsvFile } from './csv-file';

export type StringKeyObject = {[key: string]: string};

export const SheetUtils = {
  /**
   * Returns the sheet object. If the sheet does not exist yet, creates it
   * 
   * @param name Sheet's name
   */
  getOrCreateSheet(name: string) {
    const spreadsheet = SpreadsheetApp.getActive();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
    }
    return sheet;
  },

  /**
   * Returns all sheet content as 2D array
   * 
   * @param name Sheet name
   */
  readSheet(name: string):string[][] {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (! sheet) {
      throw Error(`Sheet with name '${name}' not found`);
    }

    return sheet
      .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
      .getValues();
  },

  /**
   * Read all rows from the spreadsheet and return an array of JSON Objects
   *  with sheet headers as keys.
   * 
   * @param name Sheet name 
   */
  readSheetAsJSON(name: string) {
    const returnValues: StringKeyObject[] = [];
    const sheetData = SheetUtils.readSheet(name);
    if (sheetData && sheetData.length < 2) { 
      const headers = sheetData[0];
      sheetData
        .slice(1) // Skip header row
        .forEach(row => {
          const rowToJSON: StringKeyObject = {};
          for (let i=0; i<headers.length; i++) {
            rowToJSON[ headers[i] ] = row[i];
          }
          returnValues.push(rowToJSON);
        });
    }

    return returnValues;
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
   * Put CSV values to different sheets (one sheet per csv file)
   * 
   * @param csvValues Values to save to sheets
   * @param sheetPrefix Prefix that will be added to each created sheet 
   *  (e.g. useful if we want to store SDFs for several advertisers)
   * @param hideSheets Should we hide those created CVS containing sheets
   */
  putCsvValuesIntoSheets(
    csvValues: Array<CsvFile>, 
    sheetPrefix = '', 
    hideSheets = false
  ) {
    csvValues.forEach((file) => {
      const targetSheet = SheetUtils.getOrCreateSheet(sheetPrefix + file.fileName);
      targetSheet
        .clear()
        .getRange(1, 1, file.values.length, file.values[0].length)
        .setValues(file.values);

      if (hideSheets) {
        targetSheet.hideSheet();
      }
    });
  },
};
