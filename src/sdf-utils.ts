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

import { dv360 } from "./dv360-utils";
import { SheetUtils } from "./sheet-utils";
import { Config } from "./config";
import { SheetCache } from "./sheet-cache";
import { StringKeyObject } from './shared';

export type SdfEntityName = 'Campaigns'
    | 'InsertionOrders' 
    | 'LineItems' 
    | 'AdGroups'
    | 'AdGroupAds';

export class NotFoundInSDF extends Error {}

export const SdfUtils = {
    /**
     * Returns the SDF as an array of Objects. If the sheet does not exist
     * then first download all SDF files from DV360 API and save to the sheet.
     * 
     * @param entity SDF for which DV360 entity should be returned.
     * @param advertiserId Advertiser ID, used as prefix for the "cache" sheet 
     *  name
     * @param reloadCache If TRUE, then the SDF will be downloaded and the cache
     *  overwritten
     */
    sdfGetFromSheetOrDownload(
        entity: SdfEntityName,
        advertiserId: string,
        reloadCache?: boolean
    ) {
        const prefix = advertiserId + '-';
        const sheetName = prefix + 'SDF-' + entity + '.csv';
        if (reloadCache || ! SpreadsheetApp.getActive().getSheetByName(sheetName)) {
            const csvFiles = dv360.downloadSdfs(advertiserId);
            SheetUtils.putCsvFilesIntoSheets(csvFiles, prefix, true);
        }

        return SheetUtils.readSheetAsJSON(sheetName);
    },

    /**
     * Generates SDF files (including IOs, LIs, etc.) for the copied campaign
     *  and create sheets with those generated data.
     * 
     * @param sheetCache Cache with the information about existing entities
     * @param reloadCache If TRUE then download SDF files for to be copied 
     *  campaigns from DV360. Usually download process takes several minutes.
     */
    generateNewSDFForSelectedCampaigns(
        sheetCache: {[key: string]: SheetCache},
        reloadCache?: boolean
    ) {
        const sheetData = SheetUtils.readSheetAsJSON(Config.WorkingSheet.Campaigns);
        if (!sheetData || sheetData.length < 2) {
          // Nothing to generate
          return;
        }
      
        // We don't want to download SDFs several times for the same run
        const advertisersForWhichWeAlreadyGotSDF: string[] = [];
        const generatedSdfEntries: StringKeyObject[] = [];
        sheetData.forEach((row) => {
          // TODO: More Validation! Especially check that all mandatory fields are set
      
          if (! ('Advertiser' in row)) {
            throw Error('Advertiser is not defined');
          }
      
          const advertiserInfo = sheetCache.Advertisers.find(row['Advertiser'], 0);
          if (! advertiserInfo || advertiserInfo.length < 4 || ! advertiserInfo[2]) {
            throw Error(`Advertiser '${row['Advertiser']}' not found.`);
          }

          const advertiserId = advertiserInfo[2] as string;
          const campaignsSDF = SdfUtils.sdfGetFromSheetOrDownload(
            'Campaigns', 
            advertiserId, 
            reloadCache && !advertisersForWhichWeAlreadyGotSDF.includes(advertiserId)
          );
          // TODO: Validate the SDF structure for different SDF versions
          advertisersForWhichWeAlreadyGotSDF.push(advertiserId);
      
          const campaignInfo = sheetCache.Campaigns.find(row['Campaign'], 0);
          if (! campaignInfo || campaignInfo.length < 4 || ! campaignInfo[2]) {
            throw Error(`Campaign '${row['Campaign']}' not found.`);
          }
          const campaignId = campaignInfo[2] as string;
          
          const selectedCampaign = campaignsSDF.find(
            campaign => campaignId == campaign['Campaign Id']
          );
          if (!selectedCampaign) {
            throw new NotFoundInSDF(
              `Campaign with id '${campaignId}' not found in SDF. Do you want to try reloading the cache (this might take some time)?`
            );
          }
          
          console.log(
            `Generating SDFs for advertiser id:${advertiserId} and campaign id:${campaignId}`,
            `- old Campaign: ${row['Campaign']}, new campaign name ${row['New: Name']}`
          );
          
          const generatedSDFRow = SdfUtils.generateSDFRow(
            selectedCampaign, row
          );
          generatedSdfEntries.push(generatedSDFRow);
        });

        console.log(
            'SDF is generated, saving to the new spreadsheet',
            generatedSdfEntries
        );
        
        const newSheet = SdfUtils.saveToSpreadsheet(
            {'Campaigns': generatedSdfEntries}
        );

        const content = `
            Please check 
            <a href="${newSheet}" target="_blank">the SDF by clicking here</a>!
        `;
        const htmlOutput = HtmlService.createHtmlOutput(content);
        SpreadsheetApp.getUi().showModalDialog(
            htmlOutput,
            'SDF was generated'
        );
    },

    /**
     * Generates and returns one row (as JS Object) of the copied SDF for the 
     *  selected entity
     * 
     * @param templateEntity Entity that should be copied
     * @param entityChanges Changes to be made as JS Object
     */
    generateSDFRow(
        templateEntity: StringKeyObject,
        entityChanges: StringKeyObject
    ): StringKeyObject {
        const result: StringKeyObject = {};
        for (let key in templateEntity) {
            if (Config.SDFGeneration.ClearEntriesForNewSDF.includes(key)) {
                result[ key ] = '';
                continue;
            }

            const changedFieldName = 
                Config.SDFGeneration.ChangedEntryPrefix + key;
            if (changedFieldName in entityChanges) {
                result[ key ] = entityChanges[changedFieldName];
            } else {
                result[ key ] = templateEntity[key];
            }
        }

        return result;
    },

    saveToSpreadsheet(entities: {[key: string]: StringKeyObject[]}) {
        try {
            const newSpreadsheet = SpreadsheetApp.create(Config.SDFGeneration.NewSpreadsheetTitle);
            for (let entityType in entities) {
                const sheet = newSpreadsheet.insertSheet();
                const data = SdfUtils.JSONtoArray(entities[entityType])

                sheet.setName(entityType)
                    .getRange(1, 1, data.length, data[0].length)
                    .setValues(data);
            }
        
            const url = newSpreadsheet.getUrl();
            console.log(url);
            return url;
        } catch (e: any) {
            console.log(e);
            console.log(e.stack);
            throw e;
        }
    },

    JSONtoArray(json: StringKeyObject[]) {
        const result: string[][] = [];
        if (! json.length) {
            return result;
        }

        const headers = Object.keys(json[0]);
        result.push(headers);

        json.forEach(entity => {
            const row: string[] = [];
            headers.forEach(header => {
                row.push(entity[header]);
            });
            result.push(row);
        });

        return result;
    },
}