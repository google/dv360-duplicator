import { Config } from './config';
import { DV360 } from './dv360';
import { CacheUtils } from './cache-utils';
import { onEditEvent } from './trigger';
import { TriggerUtils } from './trigger-utils';

/**
 * Global cache container
 */
const SHEET_CACHE = CacheUtils.initCache(Config.CacheSheetName);
const dv360 = new DV360(ScriptApp.getOAuthToken());

const loadPartners = CacheUtils.generateDV360LoaderFunction(
  SHEET_CACHE.Partners,
  dv360.listPartners,
  {
    name: 'displayName',
    id: 'partnerId',
  }
);

const loadAdvertisers = CacheUtils.generateDV360LoaderFunction(
  SHEET_CACHE.Advertisers,
  dv360.listAdvertisers,
  {
    name: 'displayName',
    id: 'advertiserId',
    parent: 'partnerId',
  }
);

const loadCampaigns = CacheUtils.generateDV360LoaderFunction(
  SHEET_CACHE.Campaigns,
  dv360.listCampaigns,
  {
    name: 'displayName',
    id: 'campaignId',
    parent: 'advertiserId',
  }
);

/*
function loadAdvertisers(partnerId: string) {
  let advertisers = SHEET_CACHE.Advertisers.lookup(partnerId, 1);
  if (! advertisers.length) {
    advertisers = dv360.listAdvertisers(partnerId, { limit: 10 })
      .map((advertiser) => [
        `${advertiser.displayName} (${advertiser.advertiserId})`,
        advertiser.partnerId,
        advertiser.advertiserId,
        advertiser.displayName
      ]);
    
    // Do not overwrite other cached advertisers, use ".append" 
    SHEET_CACHE.Advertisers.append(advertisers);
  }

  return advertisers;
}

function loadCampaigns(advertiserId: string) {
  let campaigns = SHEET_CACHE.Campaigns.lookup(advertiserId, 1);
  if (!campaigns.length) {
    const dv360campaigns = dv360.listCampaigns(advertiserId, { limit: 10 });
    if (campaigns && campaigns.length && campaigns[0]) {
      campaigns = dv360campaigns.map((campaign) => {
        console.log('campaign', campaign); 
        return [
          `${campaign.displayName} (${campaign.campaignId})`,
          campaign.advertiserId,
          campaign.campaignId,
          campaign.displayName
        ]
      });
    }
    
    // Do not overwrite other cached campaigns, use ".append" 
    SHEET_CACHE.Campaigns.append(campaigns);
  }

  return campaigns;
}
*/

function run() {
  loadPartners();
}

function setup() {
  onEditEvent.install();
}

function teardown() {
  onEditEvent.uninstall();
}

const partnerChangedHandler = TriggerUtils.generateOnEditHandler(
  'Campaigns', 1, 1, SHEET_CACHE.Partners, loadAdvertisers
);

const advertiserChangedHandler = TriggerUtils.generateOnEditHandler(
  'Campaigns', 2, 2, SHEET_CACHE.Advertisers, loadCampaigns
);

onEditEvent.addHandler(partnerChangedHandler);
onEditEvent.addHandler(advertiserChangedHandler);