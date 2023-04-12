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
import { StringKeyObject, StringKeyArrayOfObjects, StringKeyObjectOfObjects } from './shared';

export type SdfEntityName = 'Campaigns'
    | 'InsertionOrders' 
    | 'LineItems' 
    | 'AdGroups'
    | 'AdGroupAds';

export class NotFoundInSDF extends Error {}
export class NotFoundInCache extends Error {}

class ExtNameGenerator {
    static count = 1;
    static cache:StringKeyObjectOfObjects = {};

    static generateAndCache(entityName: SdfEntityName, id: string) {
        // First cache it
        if (!ExtNameGenerator.cache[entityName]) {
            ExtNameGenerator.cache[entityName] = {};
        }
        const extId = Config.SDFGeneration.NewIdPrefix + ExtNameGenerator.count++;
        ExtNameGenerator.cache[entityName][extId] = id;

        return extId;
    }

    static getCached(entityName: SdfEntityName, id: string) {
        if (
            !(entityName in ExtNameGenerator.cache) 
            || !(id in ExtNameGenerator.cache[entityName])
        ) {
            console.log('Current ExtNameGenerator.cache', ExtNameGenerator.cache);
            throw new NotFoundInCache(
                `ExtNameGenerator didn't find [${entityName}][${id}] in cache`
            );
        }

        return ExtNameGenerator.cache[entityName][id];
    }
}

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
        reloadCache?: boolean,
        filters?: StringKeyObject,
    ) {
        const prefix = advertiserId + '-';
        const sheetName = prefix + 'SDF-' + entity + '.csv';
        if (reloadCache || ! SpreadsheetApp.getActive().getSheetByName(sheetName)) {
            const csvFiles = dv360.downloadSdfs(advertiserId);
            SheetUtils.putCsvFilesIntoSheets(csvFiles, prefix, true);
        }

        const sdfArray = SheetUtils.readSheetAsJSON(sheetName);
        return filters 
            ? sdfArray.filter(
                x => x[ Object.keys(filters)[0] ] === Object.values(filters)[0]
            )
            : sdfArray;
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
        const generatedSdfEntries: StringKeyArrayOfObjects = {};
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
            `- Template Campaign: ${row['Campaign']}`
          );

          if (! generatedSdfEntries.Campaigns) {
            generatedSdfEntries.Campaigns = [];
          }

          const generatedCampaign = SdfUtils.generateSDFRow(
            selectedCampaign, row, "Campaigns"
          );
          generatedSdfEntries.Campaigns.push(generatedCampaign);

          if (! generatedSdfEntries.InsertionOrders) {
            generatedSdfEntries.InsertionOrders = [];
          }
          row['InsertionOrders:' + Config.SDFGeneration.SetNewExtId.Campaigns] 
            = generatedCampaign[Config.SDFGeneration.SetNewExtId.Campaigns];
        
          const generatedIOs = SdfUtils.generateSDFForInsertionOrders(
            selectedCampaign, row
          );
          generatedSdfEntries.InsertionOrders.push(...generatedIOs);

          if (! generatedSdfEntries.LineItems) {
            generatedSdfEntries.LineItems = [];
          }
          const generatedLIs = SdfUtils.generateSDFForLineItems(
            generatedIOs, row, advertiserId
          );
          generatedSdfEntries.LineItems.push(...generatedLIs);

        });

        console.log(
            'SDF is generated, saving to the new spreadsheet',
            generatedSdfEntries
        );
        
        const newSheet = SdfUtils.saveToSpreadsheet(generatedSdfEntries);

        const htmlContent = `Done! Here is what you can do next:<br /><br />
            <b>Step 1</b>: Please check 
            <a href="${newSheet}" target="_blank">the SDF (by clicking here)</a> 
            and adjust it if needed.<br /><br />

            <b>Step 2</b>: <a href="${newSheet}" target="_blank">download generated 
            files as one file.</a>!<br /><br />

            <b>Step 3</b>: Upload to DV360.
        `;
        const htmlOutput = HtmlService.createHtmlOutput(htmlContent);
        SpreadsheetApp.getUi().showModalDialog(
            htmlOutput, 
            'SDF was generated 🤗'
        );
    },

    generateSDFForInsertionOrders(
        campaign: StringKeyObject,
        changes: StringKeyObject
    ) {
        const entityName = 'InsertionOrders';
        const result:StringKeyObject[] = [];

        const ioSDFs = SdfUtils.sdfGetFromSheetOrDownload(
            entityName, 
            campaign['Advertiser Id'],
            false,
            {'Campaign Id': campaign['Campaign Id']}
        );

        ioSDFs.forEach(io => {
            result.push(SdfUtils.generateSDFRow(io, changes, entityName));
        });
        
        return result;
    },

    generateSDFForLineItems(
        ios: StringKeyObject[],
        changes: StringKeyObject,
        advertiserId: string
    ) {
        const entityName = 'LineItems';
        const result:StringKeyObject[] = [];

        ios.forEach(io => {
            console.log('+ Getting lineitems for IO ID:', io);
            const liSDFs = SdfUtils.sdfGetFromSheetOrDownload(
                entityName,
                advertiserId,
                false,
                {'Io Id': ExtNameGenerator.getCached("InsertionOrders", io['Io Id'])}
            );
    
            changes[entityName + ':' + Config.SDFGeneration.SetNewExtId.InsertionOrders]
                = io[Config.SDFGeneration.SetNewExtId.InsertionOrders];
            liSDFs.forEach(li => {
                console.log('+ Generating SDF for LI ID:', li['Line Item Id']);
                result.push(SdfUtils.generateSDFRow(li, changes, entityName));
            });
        });
        
        return result;
    },

    /**
     * Generates and returns one row (as JS Object) of the copied SDF for the 
     *  selected entity with the substituted values
     * 
     * @param templateEntity Entity that should be copied
     * @param entityChanges Changes to be made as JS Object
     * @param entityName DV360 Entity name for which the SDF row is generated 
     */
    generateSDFRow(
        templateEntity: StringKeyObject,
        entityChanges: StringKeyObject,
        entityName: SdfEntityName,
    ): StringKeyObject {
        const result: StringKeyObject = {};
        for (let key in templateEntity) {
            result[ key ] = SdfUtils.generateSDFCell(
                key, templateEntity, entityChanges, entityName
            );
        }

        return result;
    },

    generateSDFCell(
        key: string,
        templateEntity: StringKeyObject,
        entityChanges: StringKeyObject,
        entityName: SdfEntityName,
    ): string {
        const entityFieldName = 
            entityName 
            + Config.SDFGeneration.EntityNameDelimiter 
            + key;
        if (entityName && entityFieldName in entityChanges) {
            return entityChanges[entityFieldName];
        }

        if (Config.SDFGeneration.ClearEntriesForNewSDF.includes(key)) { 
            return '';
        }

        if (
            entityName in Config.SDFGeneration.SetNewExtId
            && key == Config.SDFGeneration.SetNewExtId[entityName]
        ) {    
            return ExtNameGenerator.generateAndCache(
                entityName, templateEntity[key]
            );
        }
        
        return templateEntity[key];
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