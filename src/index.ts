import { CsvFile, DV360 } from './dv360';
import { SheetUtils } from './sheets';
import { onEditEvent, OnEditHandler } from './trigger';

const dv360 = new DV360(ScriptApp.getOAuthToken());

function loadPartners() {
  const partners = dv360.listPartners({ limit: 100 });
  const sheet = SheetUtils.getOrCreateSheet('[DO NOT EDIT] Partners');
  const values = partners.map((partner) => [
    `${partner.displayName} (${partner.partnerId})`,
    partner.partnerId,
    partner.displayName
  ]);
  sheet.clear();
  sheet.getRange(1, 1, values.length, 3).setValues(values);
}

function loadAdvertisers(partnerId: string) {
  const advertisers = dv360.listAdvertisers(partnerId, { limit: 10 });
  const sheet = SheetUtils.getOrCreateSheet('[DO NOT EDIT] Advertisers');
  advertisers
    .map((advertiser) => [
      `${advertiser.displayName} (${advertiser.advertiserId})`,
      advertiser.partnerId,
      advertiser.advertiserId,
      advertiser.displayName
    ])
    .forEach((row) => sheet.appendRow(row));
  return advertisers;
}
function loadCampaigns(advertiserId: string) {
  const campaigns = dv360.listCampaigns(advertiserId, { limit: 10 });
  const sheet = SheetUtils.getOrCreateSheet('[DO NOT EDIT] Campaigns');
  campaigns
    .map((campaign) => [
      `${campaign.displayName} (${campaign.campaignId})`,
      campaign.advertiserId,
      campaign.campaignId,
      campaign.displayName
    ])
    .forEach((row) => sheet.appendRow(row));
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
    console.log(
      `Target range: [${targetRange.getColumn()}, ${targetRange.getRow()}]`
    );

    const partnerName = '' + range.getValues()[0];
    console.log(`Partner name: ${partnerName}`);
    const partnerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      '[DO NOT EDIT] Partners'
    );
    const values = partnerSheet?.getSheetValues(1, 1, -1, 2);
    console.log('Partners: ', values);
    const partnerValues = values?.find(([name]) => name === partnerName);
    console.log('Partner: ', partnerValues);
    if (partnerValues) {
      const partnerId = partnerValues[1];
      console.log(partnerId);
      const advertisers = loadAdvertisers(partnerId);
      console.log('Advertisers: ', advertisers);
      const names = advertisers.map(
        ({ displayName, advertiserId }) => `${displayName} (${advertiserId})`
      );
      console.log('Advertiser names', names);
      SheetUtils.setRangeDropDown(targetRange, names);
    }
  }
};
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
        '[DO NOT EDIT] Advertisers'
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
onEditEvent.addHandler(partnerChangedHandler);
onEditEvent.addHandler(advertiserChangedHandler);

function putCsvValuesIntoSheets(csvValues: Array<CsvFile>) {
  csvValues.forEach((file) => {
    const targetSheet = SheetUtils.getOrCreateSheet(file.fileName);
    targetSheet.clear();
    targetSheet
      .getRange(1, 1, file.values.length, file.values[0].length)
      .setValues(file.values);
  });
}

import { SheetCache } from './sheet-cache';
function test_SheetCache() {
  const cache = new SheetCache(SheetUtils.getOrCreateSheet('Cache sheet'));
  cache.set([
    [1, 2, 'A123'],
    [11, 22, 'T123']
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

function test() {
  const sdfValues = dv360.downloadSdfs('676945619');
  putCsvValuesIntoSheets(sdfValues);
}
