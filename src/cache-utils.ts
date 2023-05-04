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

import { SheetCache } from './sheet-cache';
import { SheetUtils } from './sheet-utils';

export const CacheUtils = {
    /**
     * Create all needed cache instances
     * 
     * @param entities Key-Value pairs for all cache entities to be created
     */
    initCache(entities: { [key: string]: string }) {
        const result: {[key: string]: SheetCache} = {};
        Object.keys(entities).forEach((key: string) => {
            result[key] = new SheetCache(
                SheetUtils.getOrCreateSheet(entities[key], true)
            );
        });
    
        return result;
    },

    /**
     * Clear all cache instances
     * 
     * @param entities Key-Value pairs for all cache entities to be created
     */
    clearCache(entities: {[key: string]: SheetCache}) {
        Object.keys(entities).forEach(key => {
            entities[key].clear();
        });
    }
}