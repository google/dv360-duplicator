import { Config } from "./config";
import { StringKeyObjectOfObjects } from './shared-types';
import { SdfEntityName } from "./sdf-generator";

export class NotFoundInCache extends Error {}

export class ExtNameGenerator {
    static count = 1;
    static cache:StringKeyObjectOfObjects = {};

    static generateAndCache(entityName: SdfEntityName, id: string) {
        // First cache it
        if (!ExtNameGenerator.cache[entityName]) {
            ExtNameGenerator.cache[entityName] = {};
        }
        const extId = Config.SDFGeneration.NewIdPrefix + ExtNameGenerator.count++;
        ExtNameGenerator.cache[entityName][extId] = id;

        return extId;
    }

    static getCached(entityName: SdfEntityName, id: string) {
        if (
            !(entityName in ExtNameGenerator.cache) 
            || !(id in ExtNameGenerator.cache[entityName])
        ) {
            const message = `ExtNameGenerator didn't find [${entityName}][${id}] in cache`;
            console.log(message, 'Current ExtNameGenerator.cache', ExtNameGenerator.cache);
            throw new NotFoundInCache(message);
        }

        return ExtNameGenerator.cache[entityName][id];
    }
}