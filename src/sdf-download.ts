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
        this.csvFiles.push(this.sheetToCSV(sheet));
      });
    
    return this.csvFiles;
  }

  sheetToCSV(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const url = "https://docs.google.com/spreadsheets/d/"
      + this.currentSpreadsheet.getId() 
      + `/export?gid=${sheet.getSheetId().toString()}&format=csv`;
    
    const blob = this.fetchCSV(url);
    blob.setName(`${sheet.getName()}.csv`);

    return blob;
  }

  fetchCSV(url: string) {
    try {
      const res = UrlFetchApp.fetch(
        url,
        {
          method: 'get',
          headers: {
            Authorization: "Bearer " + ScriptApp.getOAuthToken(),
          }
        }
      );

      if (200 != res.getResponseCode() && 204 != res.getResponseCode()) {
        console.log('HTTP code: ' + res.getResponseCode());
        console.log('API error: ' + res.getContentText());
        console.log('URL: ' + url);
        throw Error(res.getContentText());
      }

      return res.getBlob();
    } catch (e: any) {
      console.log(e);
      console.log(e.stack);
      throw e;
    }
  }
}
