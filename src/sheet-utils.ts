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
import { dv360 } from './dv360-utils';
import { CacheValue, SheetCache } from './sheet-cache';
import { Config } from './config';
import { StringKeyObject } from './shared';

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
    if (sheetData && sheetData.length > 1) { 
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
  putCsvFilesIntoSheets(
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

  /**
   * Inform the user that archived campaign should be unarchived before it can 
   *  be copied
   * 
   * @param campaign Campaign selected from the drop down 
   * @param cache Where to search for info about that campaign
   */
  processArchivedCampaign(range: GoogleAppsScript.Spreadsheet.Range, cache: SheetCache) {
    const campaign = "" + range.getValues()[0];
    const extendedCampaignInfo = cache.find(campaign, 0);
    if (! extendedCampaignInfo.length) {
      throw Error('Campaign not found');
    }

    const advertiserId = extendedCampaignInfo[1];
    const campaignId = extendedCampaignInfo[2];
    const campaignStatus = extendedCampaignInfo[4];
    if (! advertiserId || !campaignId || !campaignStatus) {
      throw Error('Campaign info is not full');
    }

    switch (campaignStatus) {
      case Config.CampaignStatus.Archived:
        return SheetUtils.showUnarchiveDialog(
          advertiserId, campaignId, range, cache,
        );

      case Config.CampaignStatus.Active:
      case Config.CampaignStatus.Paused:
        return;
      
      default:
        throw Error('Unknown campaign status');
    }
  },

  /**
   * Shows the "Unarchive" dialog and if user selected "Yes", unarchives the 
   *  campaign
   * 
   * @param advertiserId Campaign's Advertiser ID
   * @param campaignId Campaign ID to be unarchived
   * @param range The cell where the campaign was selected
   * @param cache Campaigns cache, to update the status after the campaign was 
   *  unarchived
   */
  showUnarchiveDialog(
    advertiserId: CacheValue,
    campaignId: CacheValue,
    range: GoogleAppsScript.Spreadsheet.Range,
    cache: SheetCache,
  ) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'You selected an archived campaign', 
      'You can anarchive it (or duplicate) yourself in UI or we can do this for you. Do you want we unarchive it for you?',
      ui.ButtonSet.YES_NO
    );

    // TODO: "Loading" indicator
    if (ui.Button.YES == response) {
      try {
        dv360.unarchiveCampaign(
          advertiserId as string, 
          campaignId as string
        );
        cache.findAndUpdateOneElement(
          campaignId, 2, 
          Config.CampaignStatus.Active, 4,
        );

        ui.alert('Done! Now you can copy this campaign');
      } catch (e: any) {
        console.log(e);
        console.log(e.stack);
        ui.alert('Oops, something went wrong, please try again later');
      }
    }
  },
};
