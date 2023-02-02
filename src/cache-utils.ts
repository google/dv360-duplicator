import { CacheValue, SheetCache } from './sheet-cache';
import { SheetUtils } from './sheets';
import { ListAdvertiserOptions, Dv360Campaign, Dv360Advertiser, Dv360Partner } from './dv360';

// E.g. limit, to output only 10 enities
const globalOptions: ListAdvertiserOptions = {limit: 10};

export type FieldsMapping = {
    name: string;
    id: string;
    parent?: string;
}

export type LoaderFunctionResult = Dv360Campaign | Dv360Advertiser | Dv360Partner;

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
    },

    generateDV360LoaderFunction(
        cache: SheetCache,
        dv360Loader: {object: {[k: string]: Function|string}, methodName: string},
        fieldsMapping: FieldsMapping,
    ) {
        return (id?: string) => {
            let entities: CacheValue[][] = [];
            if (id) {
                if (! ('parent' in fieldsMapping) || !fieldsMapping.parent) {
                    throw Error('fieldsMapping.parent is mandatory when id is set');
                }

                entities = cache.lookup(id, 1);
                if (
                    ! entities.length 
                    && dv360Loader.object[dv360Loader.methodName]
                    && 'function' == typeof dv360Loader.object[dv360Loader.methodName]
                ) {
                    const dv360entities = dv360Loader.object[dv360Loader.methodName](id, globalOptions);
                    if (
                        dv360entities 
                        && dv360entities.length 
                        && dv360entities[0]
                    ) {
                        entities = dv360entities.map((entity: LoaderFunctionResult) => [
                            `${entity[fieldsMapping.name]} (${entity[fieldsMapping.id]})`,
                            fieldsMapping.parent ? entity[fieldsMapping.parent] : '',
                            entity[fieldsMapping.id],
                            entity[fieldsMapping.name]
                        ]);
                    }

                    // Do not overwrite other cached entities, use ".append" 
                    cache.append(entities);
                }
            } else {
                entities = cache.getAll();
                if (!entities || !entities.length) {
                    const dv360entities = dv360Loader.object[dv360Loader.methodName](globalOptions);
                }
            }

            return entities;
        }
    }
}