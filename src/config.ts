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