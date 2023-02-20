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

export type SdfEntityName = 
    'Campaigns'
    | 'InsertionOrders' 
    | 'LineItems' 
    | 'AdGroups'
    | 'AdGroupAds';

export const SdfUtils = {
    /**
     * Get the SDF from the sheet or if the sheet does not exists then download
     *  it from DV360 API.
     * 
     * @param entity SDF for which DV360 entity should be returned.
     * @param advertiserId Advertiser ID, used as prefix for the "cache" sheet 
     *  name
     */
    downloadSDFOrGetFromSheet(entity: SdfEntityName, advertiserId: string) {

    }
}