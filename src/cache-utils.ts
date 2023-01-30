import { SheetCache } from './sheet-cache';
import { SheetUtils } from './sheets';

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
                SheetUtils.getOrCreateSheet(entities[key])
            );
        });
    
        return result;
    }
}