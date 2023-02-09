import { Config } from './config';
import { DV360 } from './dv360';
import { CacheUtils } from './cache-utils';
import { onEditEvent } from './trigger';
import { TriggerUtils } from './trigger-utils';
import { DV360Utils } from './dv360-utils';

/**
 * Global cache container
 */
const SHEET_CACHE = CacheUtils.initCache(Config.CacheSheetName);
const dv360 = new DV360(ScriptApp.getOAuthToken());

function loadPartners() {
  if (SHEET_CACHE.Partners.isEmpty()) {
    const partners = dv360.listPartners({ limit: 100 })
      .map((partner) => [
        `${partner.displayName} (${partner.partnerId})`,
        partner.partnerId,
        partner.displayName
      ]);
    
      SHEET_CACHE.Partners.set(partners);
  }
}

function loadAdvertisers(partnerId: string) {
  let advertisers = SHEET_CACHE.Advertisers.lookup(partnerId, 1);
  if (! advertisers.length) {
    const dv360advertisers = dv360.listAdvertisers(partnerId, { limit: 10 });
    if (dv360advertisers && dv360advertisers.length && dv360advertisers[0]) {
      advertisers = dv360advertisers.map((advertiser) => [
        `${advertiser.displayName} (${advertiser.advertiserId})`,
        advertiser.partnerId,
        advertiser.advertiserId,
        advertiser.displayName
      ]);
    }
    
    // Do not overwrite other cached advertisers, use ".append" 
    SHEET_CACHE.Advertisers.append(advertisers);
  }

  return advertisers;
}

function loadCampaigns(advertiserId: string) {
  let campaigns = SHEET_CACHE.Campaigns.lookup(advertiserId, 1);
  if (!campaigns.length) {
    const dv360campaigns = dv360.listCampaigns(
      advertiserId,
      {
        limit: 10,
        // List all campaigns
        filter: DV360Utils.generateFilterString(
          'entityStatus',
          Config.DV360DefaultFilters.CampaignStatus
        ),
      }
    );

    if (dv360campaigns && dv360campaigns.length && dv360campaigns[0]) {
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
  Config.WorkingSheet.Campaigns, 1, 1, SHEET_CACHE.Partners, loadAdvertisers
);

const advertiserChangedHandler = TriggerUtils.generateOnEditHandler(
  Config.WorkingSheet.Campaigns, 2, 2, SHEET_CACHE.Advertisers, loadCampaigns
);

onEditEvent.addHandler(partnerChangedHandler);
onEditEvent.addHandler(advertiserChangedHandler);