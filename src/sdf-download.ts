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

export class SdfDownload {
  protected currentSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  protected csvFiles:GoogleAppsScript.Base.Blob [] = [];

  constructor(url: string) {
    this.currentSpreadsheet = SpreadsheetApp.openByUrl(url);
    if (!this.currentSpreadsheet) {
      throw Error(`Cannot open spreadsheet: ${url}`);
    }
  }

  getAllCSVsFromSpreadsheet() {
    this.currentSpreadsheet
      .getSheets()
      .forEach((sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
        if (sheet.getLastRow()) {
          this.csvFiles.push(this.sheetToCSV(sheet));
        }
      });
    
    return this.csvFiles;
  }

  sheetToCSV(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const data = sheet
      .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
      .getValues();
    const csv = this.arrayToCSV(data);
    
    const blob = Utilities.newBlob(csv).setName(`${sheet.getName()}.csv`);
    return blob;
  }

  arrayToCSV(arr: string[][]): string {
    const rows: string[] = [];
    arr.forEach(a => {
      rows.push(
        a.map(
          b => `"${b.toString().replace('"', '""')}"`
        )
        .join(",")
      );
    });
    
    return rows.join("\n");
  }
}