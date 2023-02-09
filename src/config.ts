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

export const Config = {
    CacheSheetName: {
        Partners: '[DO NOT EDIT] Partners',
        Advertisers: '[DO NOT EDIT] Advertisers',
        Campaigns: '[DO NOT EDIT] Campaigns',
    },
    WorkingSheet: {
        Campaigns: 'Campaigns',
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
};