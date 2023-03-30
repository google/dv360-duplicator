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
import { SheetUtils } from './sheet-utils';
import { Config } from './config';
import { CacheUtils } from './cache-utils';
import { onEditEvent } from './trigger';
import { TriggerUtils } from './trigger-utils';
import { DV360Utils, dv360 } from './dv360-utils';
import { Setup } from './setup';
import { NotFoundInSDF, SdfUtils } from './sdf-utils';

/**
 * Global cache container
 */
const SHEET_CACHE = CacheUtils.initCache(Config.CacheSheetName);

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
          campaign.displayName,
          campaign.entityStatus,
        ]
      });
    }
    
    // Do not overwrite other cached campaigns, use ".append" 
    SHEET_CACHE.Campaigns.append(campaigns);
  }

  return campaigns;
}

const partnerChangedHandler = TriggerUtils.generateOnEditHandler(
  Config.WorkingSheet.Campaigns, 1, 1, SHEET_CACHE.Partners, loadAdvertisers
);

const advertiserChangedHandler = TriggerUtils.generateOnEditHandler(
  Config.WorkingSheet.Campaigns, 2, 2, SHEET_CACHE.Advertisers, loadCampaigns
);

const checkArchivedCampaignHandler = TriggerUtils.generateOnEditHandler(
  Config.WorkingSheet.Campaigns, 
  3, 
  2, // Not important since we use a custom run function
  SHEET_CACHE.Campaigns, 
  undefined,
  SheetUtils.processArchivedCampaign
);

onEditEvent.addHandler(partnerChangedHandler);
onEditEvent.addHandler(advertiserChangedHandler);
onEditEvent.addHandler(checkArchivedCampaignHandler);

function run() {
  loadPartners();
}

function setup() {
  onEditEvent.install();
}

function teardown() {
  onEditEvent.uninstall();
}

function onOpen(e: Event) {
  Setup.createMenu();
}

function generateSDFForActiveSheet(reloadCache?: boolean): void {
  const activeSheetName = SpreadsheetApp.getActiveSheet().getName();
  console.log(`generateSDFForActiveSheet for sheet '${activeSheetName}'`);
  switch (activeSheetName) {
    case Config.WorkingSheet.Campaigns:
      try {
        SdfUtils.generateNewSDFForSelectedCampaigns(SHEET_CACHE, reloadCache);
      } catch (e: any) {
        if (e instanceof NotFoundInSDF && !reloadCache) {
          const ui = SpreadsheetApp.getUi();
          const response = ui.alert(
            'Action needed', e.message, ui.ButtonSet.YES_NO
          );

          // TODO: "Loading" indicator
          if (ui.Button.YES == response) {
            return generateSDFForActiveSheet(true);
          }
        } else if (e instanceof NotFoundInSDF && reloadCache) {
          SpreadsheetApp.getUi().alert(
            'After reloading the cache the campaign is still not found. Please try selecting a different campaign.'
          );
          return;
        } else {
          SpreadsheetApp.getUi().alert(e.message);
          throw e;
        }
      }
      break;

    default:
      const message = `Unsupported sheet ('${activeSheetName}'), cannot generate SDF. Please select different sheet.`;
      SpreadsheetApp.getUi().alert(message);
      throw Error(message);
  }
}

function reloadCacheForActiveSheet() {
  const activeSheetName = SpreadsheetApp.getActiveSheet().getName();
  console.log(`reloadSDFForActiveSheet reloading cached SDFs for sheet '${activeSheetName}'`);


}