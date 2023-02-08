import { Config } from './config';
import { DV360 } from './dv360';
import { SheetUtils } from './sheet-utils';
import { SheetCache } from './sheet-cache';
import { CacheUtils } from './cache-utils';
import { onEditEvent, OnEditHandler } from './trigger';
import { TriggerUtils } from './trigger-utils';

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
    const dv360campaigns = dv360.listCampaigns(advertiserId, { limit: 10 });
    if (campaigns && campaigns.length && campaigns[0]) {
      campaigns = dv360campaigns.map((campaign) => [
          `${campaign.displayName} (${campaign.campaignId})`,
          campaign.advertiserId,
          campaign.campaignId,
          campaign.displayName
        ]);
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
  Config.WorkingSheets.campaigns,
  1,
  1,
  SHEET_CACHE.Partners,
  loadAdvertisers,
  'Loading advertisers... this might take some time'
);

const advertiserChangedHandler = TriggerUtils.generateOnEditHandler(
  Config.WorkingSheets.campaigns,
  2,
  2,
  SHEET_CACHE.Advertisers,
  loadCampaigns,
  'Loading campaigns... this might take some time'
);

onEditEvent.addHandler(partnerChangedHandler);
onEditEvent.addHandler(advertiserChangedHandler);