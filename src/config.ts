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

const appName = 'DV360 Duplicator';

export const Config = {
    Menu: {
        Name: appName,
        GenerateSDFForActiveSheet: '🚀  Generate SDF(s)',
        ClearCache: '🧹  Clear cache',
        Install: '🏗️  Install',
        Help: '🆘  Help',
    },
    CacheSheetName: {
        Partners: '[DO NOT EDIT] Partners',
        Advertisers: '[DO NOT EDIT] Advertisers',
        Campaigns: '[DO NOT EDIT] Campaigns',
    },
    WorkingSheet: {
        Campaigns: 'Campaigns',
        CampaignsHeaders: [
            'Partner',
            'Advertiser',
            'Campaign',
            'Campaigns:Name'
        ],
    },
    DV360DefaultFilters: {
        // We list here all the statuses, since we want ot be able to list
        // and duplicate all (including archived) campaigns.
        CampaignStatus: [
            'ENTITY_STATUS_ACTIVE',
            'ENTITY_STATUS_PAUSED',
            'ENTITY_STATUS_ARCHIVED',
        ],
    },
    CampaignStatus: {
        Active: 'ENTITY_STATUS_ACTIVE',
        Archived: 'ENTITY_STATUS_ARCHIVED',
        Paused: 'ENTITY_STATUS_PAUSED',
    },
    SDFGeneration: {
        // In the stpreadsheet changes will have same column name as in SDF
        // but with this prefix
        EntityNameDelimiter: ':',
        NewIdPrefix: 'ext',

        // Some fields should be always empty for the items we create
        ClearEntriesForNewSDF: [
            'Timestamp',
        ],

        SetNewExtId: {
            Campaigns: "Campaign Id",
            InsertionOrders: "Io Id",
            LineItems: "Line Item Id",
            AdGroups: "Ad Group Id",
            AdGroupAds: "Ad Id",
        },

        NewSpreadsheetTitle: `[${appName}] Generated SDF (${Date()})`,
    }
};