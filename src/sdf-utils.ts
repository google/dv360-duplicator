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

export type SdfEntityName = 'Campaigns'
    | 'InsertionOrders' 
    | 'LineItems' 
    | 'AdGroups'
    | 'AdGroupAds';

export const SdfUtils = {
    /**
     * Returns the SDF as an array of Objects. If the sheet does not exist
     * then first download all SDF files from DV360 API and save to the sheet.
     * 
     * @param entity SDF for which DV360 entity should be returned.
     * @param advertiserId Advertiser ID, used as prefix for the "cache" sheet 
     *  name
     */
    sdfGetFromSheetOrDownload(entity: SdfEntityName, advertiserId: string) {
        const prefix = advertiserId + '-';
        const sheetName = prefix + 'SDF-' + entity + '.csv';
        if (! SpreadsheetApp.getActive().getSheetByName(sheetName)) {
            const csvFiles = dv360.downloadSdfs(advertiserId);
            SheetUtils.putCsvFilesIntoSheets(csvFiles, prefix, true);
        }

        return SheetUtils.readSheetAsJSON(sheetName);
    }
}