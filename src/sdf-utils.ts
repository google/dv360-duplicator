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

import { SheetUtils } from "./sheet-utils";
import { Config } from "./config";
import { SheetCache } from "./sheet-cache";
import { StringKeyObject, StringKeyArrayOfObjects, StringKeyObjectOfObjects } from './shared';
import { sdfGenerator } from "./sdf-generator";

export class NotFoundInSDF extends Error {}
export class EmptyDataSet extends Error {}

export const SdfUtils = {
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
        console.log('generateNewSDFForSelectedCampaigns sheetData', sheetData);
        if (!sheetData || sheetData.length < 1) {
          // Nothing to generate
          throw new EmptyDataSet('Nothing to generate');
        }
      
        // We don't want to download SDFs several times for the same run
        const advertisersForWhichWeAlreadyGotSDF: string[] = [];
        const campaigns: StringKeyObject[] = [];
        sheetData.forEach((row) => {
          console.log('generateNewSDFForSelectedCampaigns: Processing row:', row);
          // TODO: More Validation! Especially check that all mandatory fields are set
      
          if (! ('Advertiser' in row)) {
            throw Error('Advertiser is not defined');
          }
      
          const advertiserInfo = sheetCache.Advertisers.find(row['Advertiser'], 0);
          if (! advertiserInfo || advertiserInfo.length < 4 || ! advertiserInfo[2]) {
            throw Error(`Advertiser '${row['Advertiser']}' not found.`);
          }

          const advertiserId = advertiserInfo[2] as string;
          const campaignsSDF = sdfGenerator.sdfGetFromSheetOrDownload(
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
              `Campaign with id '${campaignId}' not found in SDF. 
              Do you want to try reloading the cache (this might take some time)?`
            );
          }
          
          campaigns.push(selectedCampaign);
        });

        const sdf = sdfGenerator.generateAll(sheetData, campaigns);
        
        console.log('SDF is generated, saving to the new spreadsheet', sdf);
        const newSheet = SdfUtils.saveToSpreadsheet(sdf).replace('"', '\"');

        const htmlContent = `
            <script>
                function initDownload(url) {
                    var d = document.createElement('a');
                    d.href = url;
                    d.click();
                }
            </script>
            Done! Here is what you can do next:<br /><br />
            <b>Step 1</b>: Please check 
            <a href="${newSheet}" target="_blank">the SDF (by clicking here)</a> 
            and adjust it if needed.<br /><br />

            <b>Step 2</b>:
            <a href="" target="_blank"
                onclick="google.script.run.withSuccessHandler(initDownload).downloadGeneratedSDFAsZIP('${newSheet}'); return false;">
            download generated files as one "zip"</a>.<br /><br />

            <b>Step 3</b>: Upload to DV360.
        `;
        const htmlOutput = HtmlService.createHtmlOutput(htmlContent);
        SpreadsheetApp.getUi().showModalDialog(
            htmlOutput, 
            'SDF was generated 🤗'
        );
    },

    /**
     * Creates a new spreadsheet and saves each element of "entities" to a separate
     *  sheet. The names of the sheets are the keys of the "entities".
     * 
     * @param entities Data to be saved
     */
    saveToSpreadsheet(entities: {[key: string]: StringKeyObject[]}) {
        if (! entities || ! Object.keys(entities).length) {
            throw Error('Empty entities list, nothing to save...');
        }

        try {
            const newSpreadsheet = SpreadsheetApp.create(Config.SDFGeneration.NewSpreadsheetTitle);
            for (let entityType in entities) {
                const sheet = newSpreadsheet.insertSheet();
                const data = SdfUtils.JSONtoArray(entities[entityType])
                console.log('saveToSpreadsheet data:', data);

                sheet.setName(entityType);
                if (data && data[0]) {
                    sheet.getRange(1, 1, data.length, data[0].length)
                        .setValues(data);
                }
            }

            // Removing the default sheet, since it's not needed
            const defaultSheet = newSpreadsheet.getSheetByName('Sheet1');
            if (defaultSheet) {
                newSpreadsheet.deleteSheet(defaultSheet);
            }
        
            const url = newSpreadsheet.getUrl();
            console.log('Generated SDF is saved to spreadsheet:', url);
            return url;
        } catch (e: any) {
            console.log(e);
            console.log(e.stack);
            throw e;
        }
    },

    /**
     * Converts an array of JSONs to the table array (and returns it) with the
     *  JSON keys as the headers
     * 
     * @param json JSONs to be converted
     */
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