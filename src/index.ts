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
import { DV360 } from './dv360';
import { SheetUtils } from './sheet-utils';
import { Config } from './config';
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

function generateNewSDFForSelectedCampaigns() {
  const sheetData = SheetUtils.readSheetAsJSON(Config.WorkingSheet.Campaigns);
  if (! sheetData || !sheetData.length) {
    // Nothing to generate
    return;
  }

  sheetData.forEach((row) => {
    // TODO: More Validation!!!

    if (! ('Advertiser' in row)) {
      throw Error('Advertiser is not defined');
    }

    const advertiserInfo = SHEET_CACHE.Advertiser.find(row['Advertiser'], 1);
    if (! advertiserInfo || 4 != advertiserInfo.length || ! advertiserInfo[2]) {
      throw Error(`Advertiser '${row['Advertiser']}' not defined or not found.`);
    }

    const advertiserId = advertiserInfo[2];
  });
}