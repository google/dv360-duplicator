export abstract class ApiClient {
  /**
   *
   * @param {string} baseUrl
   * @param {string} authToken A token needed to connect to DV360 API
   */
  constructor(
    protected readonly baseUrl: string,
    protected readonly authToken: string
  ) {
    if (!authToken) {
      throw new Error('AuthToken must not be empty.');
    }
  }

  private encodeParam(key: string, value: string | number | boolean) {
    const encodedKey = encodeURIComponent(key);
    const encodedValue = encodeURIComponent('' + value);
    return `${encodedKey}=${encodedValue}`;
  }

  protected toUrlParamString(
    params?: Record<string, string | number | boolean | undefined>
  ) {
    const entries = params ? Object.entries(params) : [];
    const defined = entries.filter(([key, value]) => value !== undefined) as [
      string,
      string | number | boolean
    ][];
    if (defined.length > 0) {
      return defined
        .map(([key, value]) => this.encodeParam(key, value))
        .join('&');
    }
    return '';
  }

  protected getUrl(
    path: string,
    params?: Record<string, string | number | boolean | undefined>
  ) {
    const pathUrl = `${this.baseUrl}/${path}`;
    const paramString = this.toUrlParamString(params);
    return paramString ? `${pathUrl}?${paramString}` : pathUrl;
  }

  fetchEntity(
    url: string,
    options?: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions,
    payload?: Record<string, any>
  ) {
    const res = this.fetchResponse(url, options, payload);
    return res.getContentText() ? JSON.parse(res.getContentText()) : {};
  }

  fetchBlob(
    url: string,
    options?: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions,
    payload?: Record<string, any>
  ) {
    const res = this.fetchResponse(url, options, payload);
    return res.getBlob();
  }

  /**
   * Make an HTTPS API request using specified auth method (see 'Auth' class)
   */
  fetchResponse(
    url: string,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions | undefined,
    payload: Record<string, any> | undefined
  ) {
    const defaultOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      muteHttpExceptions: true,
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + this.authToken,
        Accept: '*/*'
      }
    };
    const requestOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      ...defaultOptions,
      ...options
    };
    if (payload) {
      requestOptions.headers!['Content-Type'] = 'application/json';
      requestOptions.payload = JSON.stringify(payload);
    }

    console.log(`Before fetch URL: ${url}`);
    console.log(`Before fetch payload: ${requestOptions.payload}`);
    const res = UrlFetchApp.fetch(url, requestOptions);

    if (200 != res.getResponseCode() && 204 != res.getResponseCode()) {
      Logger.log('HTTP code: ' + res.getResponseCode());
      Logger.log('API error: ' + res.getContentText());
      Logger.log('URL: ' + url);
      throw new Error(res.getContentText());
    }
    return res;
  }
}
