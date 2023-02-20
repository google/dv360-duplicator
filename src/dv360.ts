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

import { ApiClient } from './api-client';
import { CsvFile } from './csv-file';

interface ResourcePage {
  nextPageToken?: string;
}

export interface Dv360Partner {
  displayName: string;
  partnerId: string;
}
interface Dv360PartnersPage extends ResourcePage {
  partners: Dv360Partner[];
}

export interface Dv360Advertiser {
  displayName: string;
  partnerId: string;
  advertiserId: string;
}

export interface Dv360Campaign {
  displayName: string;
  partnerId: string;
  advertiserId: string;
  campaignId: string;
}

interface Dv360AdvertisersPage extends ResourcePage {
  advertisers: Dv360Advertiser[];
}

interface Dv360CampaignsPage extends ResourcePage {
  campaigns: Dv360Campaign[];
}

interface PagingOptions {
  pageSize?: number;
}

export interface ListPartnersOptions extends PagingOptions {
  orderBy?: 'displayName';
  filter?: string;
  limit?: number;
}

export interface ListAdvertiserOptions extends PagingOptions {
  orderBy?: 'displayName' | 'entityStatus' | 'updateTime';
  filter?: string;
  limit?: number;
}

export interface ListCampaignOptions extends PagingOptions {
  orderBy?: 'displayName' | 'entityStatus' | 'updateTime';
  filter?: string;
  limit?: number;
}

/**
 * DV360 API Wrapper class. Implements DV360 API calls.
 */
export class DV360 extends ApiClient {
  /**
   * Set the DV360 wrapper configuration
   *
   */
  constructor(
    authToken: string,
    dv360BaseUrl = 'https://displayvideo.googleapis.com/v2'
  ) {
    super(dv360BaseUrl, authToken);
  }

  listPartnersPage(
    pageToken?: string,
    options?: ListPartnersOptions
  ): Dv360PartnersPage {
    const { pageSize, filter, orderBy } = options ?? {};
    return this.fetchEntity(
      this.getUrl('partners', { pageToken, pageSize, filter, orderBy })
    );
  }

  listPartners(options?: ListPartnersOptions) {
    let partners: Dv360Partner[] = [];
    let nextPageToken: string | undefined;
    
    do {
      const response = this.listPartnersPage(nextPageToken, options);
      partners = partners.concat(response.partners);
      nextPageToken = response.nextPageToken;
    } while (
      nextPageToken &&
      partners.length < (options?.limit ?? Number.POSITIVE_INFINITY)
    );

    if (options?.limit) {
      partners = partners.slice(0, options.limit);
    }
    return partners;
  }

  listAdvertisersPage(
    partnerId: string,
    pageToken?: string,
    options?: ListAdvertiserOptions
  ): Dv360AdvertisersPage {
    const { pageSize, filter, orderBy } = options ?? {};
    return this.fetchEntity(
      this.getUrl('advertisers', {
        partnerId,
        pageToken,
        pageSize,
        filter,
        orderBy
      })
    );
  }

  listAdvertisers(partnerId: string, options?: ListPartnersOptions) {
    let advertisers: Dv360Advertiser[] = [];
    let nextPageToken: string | undefined;

    do {
      const response = this.listAdvertisersPage(
        partnerId,
        nextPageToken,
        options
      );
      advertisers = advertisers.concat(response.advertisers);
      nextPageToken = response.nextPageToken;
    } while (
      nextPageToken &&
      advertisers.length < (options?.limit ?? Number.POSITIVE_INFINITY)
    );

    if (options?.limit) {
      advertisers = advertisers.slice(0, options.limit);
    }
    return advertisers;
  }

  listCampaignsPage(
    advertiserId: string,
    pageToken?: string,
    options?: ListCampaignOptions
  ): Dv360CampaignsPage {
    const { pageSize, filter, orderBy } = options ?? {};
    return this.fetchEntity(
      this.getUrl(`advertisers/${advertiserId}/campaigns`, {
        pageToken,
        pageSize,
        filter,
        orderBy
      })
    );
  }

  listCampaigns(advertiserId: string, options?: ListAdvertiserOptions) {
    let campaigns: Dv360Campaign[] = [];
    let nextPageToken: string | undefined;

    do {
      const response = this.listCampaignsPage(
        advertiserId,
        nextPageToken,
        options
      );

      campaigns = campaigns.concat(response.campaigns);
      nextPageToken = response.nextPageToken;
    } while (
      nextPageToken &&
      campaigns.length < (options?.limit ?? Number.POSITIVE_INFINITY)
    );

    if (options?.limit) {
      campaigns = campaigns.slice(0, options.limit);
    }
    return campaigns;
  }

  protected createSdfDownloadOperation(advertiserId: string) {
    const downloadOperation = this.fetchEntity(
      this.getUrl('sdfdownloadtasks'),
      {
        method: 'post'
      },
      {
        version: 'SDF_VERSION_5_5',
        advertiserId: advertiserId,
        parentEntityFilter: {
          fileType: [
            'FILE_TYPE_CAMPAIGN',
            'FILE_TYPE_INSERTION_ORDER',
            'FILE_TYPE_LINE_ITEM',
            'FILE_TYPE_AD_GROUP',
            'FILE_TYPE_AD'
          ],
          filterType: 'FILTER_TYPE_NONE',
          filterIds: []
        }
      }
    );
    return downloadOperation;
  }

  protected waitForDownloadOperationResource(downloadOperationName: string) {
    const operationResourceUrl = this.getUrl(downloadOperationName);
    const initialDelay = 2000;
    const maxRetries = 10;
    const delayMultiplier = 2;

    let downloadOperation;
    let delay = initialDelay;
    let tryCount = 0;
    while (tryCount <= maxRetries) {
      downloadOperation = this.fetchEntity(operationResourceUrl);
      Logger.log(downloadOperation);
      if (downloadOperation.done) {
        return downloadOperation.response.resourceName;
      }
      Logger.log(`Backing off for ${delay}ms`);
      Utilities.sleep(delay);
      tryCount++;
      delay *= delayMultiplier;
    }
    Logger.log(`Try limit exceeded after ${tryCount} tries`);
    return undefined;
  }

  downloadMedia(resourceName: string): GoogleAppsScript.Base.Blob {
    //url doesn't contain version, just /media/
    const downloadMediaUrl = this.getUrl(
      `download/${resourceName}?alt=media`
    ).replace('/v2', '');
    const downloadMedia = this.fetchBlob(downloadMediaUrl);
    //TODO move somewhere else afterwards
    downloadMedia.setContentType('application/zip');
    return downloadMedia;
  }

  unzipSdfs(downloadMedia: GoogleAppsScript.Base.Blob): Array<CsvFile> {
    var unZippedfiles = Utilities.unzip(downloadMedia);
    const csvValues = Array<CsvFile>();
    unZippedfiles.forEach((blob) => {
      const fileName = blob.getName();
      const values = Utilities.parseCsv(blob.getDataAsString());
      Logger.log(`Found ${values.length} entities in file ${fileName}`);
      csvValues.push(new CsvFile(fileName, values));
    });
    return csvValues;
  }

  downloadSdfs(advertiserId: string): Array<CsvFile> {
    const downloadTask = this.createSdfDownloadOperation(advertiserId);
    const resourceName = this.waitForDownloadOperationResource(
      downloadTask.name
    );
    Logger.log(`Downloading media ${resourceName} and putting it into sheets`);
    const media = this.downloadMedia(resourceName);
    return this.unzipSdfs(media);
  }
}
