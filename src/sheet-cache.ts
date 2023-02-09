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

/**
 * What can be stored in the cache cells
 */
export type CacheValue = string|number|null;

/**
 * Implementation of the caching.
 * Data is cached in the google spreadsheet.
 */
export class SheetCache {
    /**
     * Create an instance
     * 
     * @param source Spreadsheet object
     */
    constructor(
        private cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
    ) {}

    /**
     * Returns TRUE if the cache is empty. 
     */
    isEmpty() {
        return ! this.cacheSheet.getLastRow();
    }

    /**
     * Save the array of data (the rows) to the sheet.
     * By default it will override the existing cache sheet.
     * 
     * @param data Any data with strings, numbers or empty elements
     * @param append If set to TRUE, then will append data to an existing cache
     */
    set(data: Array<Array<CacheValue>>, append: boolean = false) {
        if (!append) {
            this.clear();
        }

        if (! data.length) {
            // Nothing to save
            return;
        }

        const lastRow = this.cacheSheet.getLastRow();
        this.cacheSheet
            .hideSheet()
            .getRange(lastRow + 1, 1, data.length, data[0].length)
            .setValues(data);
    }

    /**
     * Save the array of data to the sheet
     * 
     * @param data Any data with strings, numbers or empty elements
     */
    append(data: Array<Array<CacheValue>>) {
        return this.set(data, true);
    }

    /**
     * Finds and returns all the rows in the cache where the 'index' column equal to 'key'
     * 
     * @param key The value we are looking for 
     * @param index The index (first element is "0") of the column where we search for it
     * @param returnOnlyOne If set to TRUE, then will return only first found element
     */
    lookup(
        key: CacheValue, 
        index: number, 
        returnOnlyOne: boolean = false
    ): Array<Array<CacheValue>> {
        if (this.isEmpty()) {
            return [];
        }

        const data = this.cacheSheet
            .getRange(
                1, 1,
                this.cacheSheet.getLastRow(),
                this.cacheSheet.getLastColumn()
            )
            .getValues();

        if (!data || !data.length) {
            Logger.log(`Empty cache in sheet "${this.cacheSheet.getName()}"`);
            return [];
        }

        if (index < 0 || index > data[0].length - 1) {
            throw Error(`Invalid index "${index}" (allowed index range 0..${data[0].length - 1})`);
        }

        const searchFunc = (row: Array<CacheValue>) => row[index] == key;
        const result = returnOnlyOne ? [ data.find(searchFunc) as Array<CacheValue> ] : data.filter(searchFunc);

        return result;
    }

    /**
     * Finds and returns only first row from the cache where the 'index' column equal to 'key'
     * 
     * @param key The value we are looking for
     * @param index The index of the column where we search for it
     */
    find(key: CacheValue, index: number) {
        const result = this.lookup(key, index, true);
        return result.length ? result[0] : [];
    }

    /**
     * Erase the cache
     */
    clear() {
        this.cacheSheet.clear();
    }
}