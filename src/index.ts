import { Config } from './config';
import { DV360 } from './dv360';
import { SheetUtils } from './sheets';
import { SheetCache } from "./sheet-cache";
import { onEditEvent, OnEditHandler } from './trigger';

const dv360 = new DV360(ScriptApp.getOAuthToken());

function loadPartners() {
  const partnersCache = new SheetCache(
    SheetUtils.getOrCreateSheet(Config.CacheSheetName.Partners)
  );

  if (partnersCache.isEmpty()) {
    const partners = dv360.listPartners({ limit: 100 })
      .map((partner) => [
        `${partner.displayName} (${partner.partnerId})`,
        partner.partnerId,
        partner.displayName
      ]);
    
    partnersCache.set(partners);
  }
}

function loadAdvertisers(partnerId: string) {
  const advertisersCache = new SheetCache(
    SheetUtils.getOrCreateSheet(Config.CacheSheetName.Advertisers)
  );

  let advertisers = advertisersCache.lookup(partnerId, 1);
  if (! advertisers.length) {
    advertisers = dv360.listAdvertisers(partnerId, { limit: 10 })
      .map((advertiser) => [
        `${advertiser.displayName} (${advertiser.advertiserId})`,
        advertiser.partnerId,
        advertiser.advertiserId,
        advertiser.displayName
      ]);
    
    // Do not overwrite other cached advertisers, use ".append" 
    advertisersCache.append(advertisers);
  }

  return advertisers;
}

function loadCampaigns(advertiserId: string) {
  const campaignsCache = new SheetCache(
    SheetUtils.getOrCreateSheet(Config.CacheSheetName.Campaigns)
  );

  let campaigns = campaignsCache.lookup(advertiserId, 1);
  if (!campaigns.length) {
    campaigns = dv360.listCampaigns(advertiserId, { limit: 10 })
      .map((campaign) => [
        `${campaign.displayName} (${campaign.campaignId})`,
        campaign.advertiserId,
        campaign.campaignId,
        campaign.displayName
      ]);
    
    // Do not overwrite other cached campaigns, use ".append" 
    campaignsCache.append(campaigns);
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

const partnerChangedHandler: OnEditHandler = {
  shouldRun({ range }) {
    return (
      range.getSheet().getName() === 'Campaigns' &&
      range.getColumn() === 1 &&
      !!range.getValue()
    );
  },
  run({ range }) {
    console.log('Running partner changed handler');
    console.log(`Source range: [${range.getColumn()}, ${range.getRow()}]`);
    const campaignSheet = range.getSheet();
    const targetRange = campaignSheet
      .getRange(range.getRow(), range.getColumn() + 1)
      .clearDataValidations()
      .clear();
    
    // Inform User about WIP
    targetRange.setValue('Loading ...');
    
    console.log(
      `Target range: [${targetRange.getColumn()}, ${targetRange.getRow()}]`
    );

    const partnerName = '' + range.getValues()[0];
    console.log(`Partner name: ${partnerName}`);
    const partnerCache = new SheetCache(
      SheetUtils.getOrCreateSheet(Config.CacheSheetName.Partners)
    );
    
    const partnerValues = partnerCache.find(partnerName, 0);
    console.log('Partner: ', partnerValues);
    if (partnerValues) {
      const partnerId = "" + partnerValues[1];
      console.log(partnerId);
      try {
        const advertisers = loadAdvertisers(partnerId);
        console.log('Advertisers: ', advertisers);
        const names = advertisers.map(
          (advertiser) => advertiser && Array.isArray(advertiser) && advertiser.length
            ? `${advertiser[1]} (${advertiser[0]})` 
            : ''
        );
        console.log('Advertiser names', names);
        SheetUtils.setRangeDropDown(targetRange, names);
      } catch (e: any) {
        SpreadsheetApp.getUi().alert('Error accured, try again...');
        range.setValue('');
        Logger.log(e?.message);
      }

      targetRange.setValue('');
    }
  }
};

/*
const advertiserChangedHandler: OnEditHandler = {
  shouldRun({ range }) {
    return (
      range.getSheet().getName() === 'Campaigns' &&
      range.getColumn() === 2 &&
      !!range.getValue()
    );
  },
  run({ range }) {
    console.log('Running advertiser changed handler');
    console.log(`Source range: [${range.getColumn()}, ${range.getRow()}]`);
    const campaignSheet = range.getSheet();
    const targetRange = campaignSheet
      .getRange(range.getRow(), range.getColumn() + 1)
      .clearDataValidations()
      .clear();
    console.log(
      `Target range: [${targetRange.getColumn()}, ${targetRange.getRow()}]`
    );

    const advertiserName = '' + range.getValues()[0];
    console.log(`Advertiser name: ${advertiserName}`);
    const advertiserSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        Config.CacheSheetName.Advertisers
      );
    const values = advertiserSheet?.getSheetValues(1, 1, -1, 3);
    const advertiserValues = values?.find(([name]) => name === advertiserName);
    console.log('Advertiser: ', values);
    if (advertiserValues) {
      const advertiserId = advertiserValues[2];
      console.log(`Selected AdvertiserId ${advertiserId}`);
      const campaigns = loadCampaigns(advertiserId);
      const names = campaigns.map(
        (campaign) => `${campaign.displayName} (${campaign.campaignId})`
      );
      console.log(`Campaigns fetched ${names}`);
      SheetUtils.setRangeDropDown(targetRange, names);
    }
  }
};
*/
onEditEvent.addHandler(partnerChangedHandler);
//onEditEvent.addHandler(advertiserChangedHandler);

/*
function test_SheetCache() {
    const cache = new SheetCache(SheetUtils.getOrCreateSheet('Cache sheet'));
    cache.set([
        [1, 2, 'A123'],
        [11, 22, 'T123'],
    ]); 
    cache.append([[333, 444, 'T123']]);
    cache.append([[11, 444, 'T321-1']]);
    cache.append([[99, 22, 'T321-2']]);
      
    const cachedValue1 = cache.lookup(11, 0);
    console.log('cachedValue1', cachedValue1);
    
    const cachedValue2 = cache.lookup(22, 1);
    console.log('cachedValue1', cachedValue2);

    const cachedValue3 = cache.lookup('22', 1);
    console.log('cachedValue1', cachedValue3);
  
    const cachedValue4 = cache.lookup('T123', 2);
    console.log('cachedValue1', cachedValue4);
}

function test_loadCampaigns() {
  loadCampaigns('2012934');
  loadCampaigns('2012934');

  loadCampaigns('676945619');
}

function test_loadAdvertisers() {
  loadAdvertisers('2015636');
  loadAdvertisers('2015636');

  loadAdvertisers('1839756');
}

*/