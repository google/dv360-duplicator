
import { StringKeyObject, StringKeyObjectOfObjects } from './shared';
import { Config } from './config';
import { ExtNameGenerator } from './sdf-extname';
import { SheetUtils } from './sheet-utils';
import { dv360 } from './dv360-utils';

export type SdfEntityName = 
    'Campaigns'
    | 'InsertionOrders' 
    | 'LineItems' 
    | 'AdGroups'
    | 'AdGroupAds';

export type Sdf = {
    Campaigns?: StringKeyObject[];
    InsertionOrders?: StringKeyObject[];
    LineItems?: StringKeyObject[];
    AdGroups?: StringKeyObject[];
    AdGroupAds?: StringKeyObject[];
};

const entityHierarchy: StringKeyObjectOfObjects  = {
    InsertionOrders: {
        id: "Io Id",
        parent: "Campaigns",
        parentId: "Campaign Id",
    },
    LineItems: {
        id: "Line Item Id",
        parent: "InsertionOrders",
        parentId: "Io Id",
    },
    AdGroups: {
        id: "Ad Group Id",
        parent: "LineItems",
        parentId: "Line Item Id",
    },
    AdGroupAds: {
        id: "Ad Id",
        parent: "AdGroups",
        parentId: "Ad Group Id",
    },
};

export class SdfGenerator {
    protected sdf: Sdf = {};

    /**
     * Returns the generated SDFs
     */
    getSdf() {
        return this.sdf;
    }

    /**
     * Generate SDF for all entities and return the SDF.
     */
    generateAll(
        changes: StringKeyObject[],
        templateEntities: StringKeyObject[],
    ) {
        return this.generateSDFForParentEntities(
                "Campaigns", changes, templateEntities
            )
            .getSdf();
    }

    /**
     * Generates and returns one row (as JS Object) of the copied SDF for the 
     *  selected entity with the substituted values
     * 
     * @param templateEntity Entity that should be copied
     * @param entityChanges Changes to be made as JS Object
     * @param entityName DV360 Entity name for which the SDF row is generated 
     */
    generateSDFRow(
        templateEntity: StringKeyObject,
        entityChanges: StringKeyObject,
        entityName: SdfEntityName,
    ): StringKeyObject {
        const result: StringKeyObject = {};
        for (let key in templateEntity) {
            result[ key ] = this.generateSDFCell(
                key, templateEntity, entityChanges, entityName
            );
        }

        return result;
    }

    /**
     * Generates one entry in the SDF by substituting the values in the template
     *  entity with the proper end SDF values according to the column names  
     *  specified in the spreadsheet. Especially it substitutes the following
     *  types of entries: 1) prefixed with the entity type in the spreadsheet;
     *  2) those that should not have any value; 3) generates correct "ext*" 
     *  references.
     * 
     * @param key Column name in the spreadsheet
     * @param templateEntity The SDF row that should be used as a template
     * @param entityChanges What should be changed in the entity
     * @param entityName One of the supported entitie names (see type SdfEntityName)
     * @returns Substituted value (as a string) for this particular cell
     */
    generateSDFCell(
        key: string,
        templateEntity: StringKeyObject,
        entityChanges: StringKeyObject,
        entityName: SdfEntityName,
    ): string {
        const entityFieldName = 
            entityName 
            + Config.SDFGeneration.EntityNameDelimiter 
            + key;
        
        if (entityName && entityFieldName in entityChanges) {
            return entityChanges[entityFieldName];
        }

        if (Config.SDFGeneration.ClearEntriesForNewSDF.includes(key)) { 
            return '';
        }

        if (
            entityName in Config.SDFGeneration.SetNewExtId
            && key == Config.SDFGeneration.SetNewExtId[entityName]
        ) {    
            return ExtNameGenerator.generateAndCache(
                entityName, templateEntity[key]
            );
        }
        
        return templateEntity[key];
    }

    /**
     * Generates the SDF for a parent (usually a campaign)
     * 
     * @param entityName 
     * @param changes 
     * @param templateEntities 
     * @returns 
     */
    generateSDFForParentEntities(
        entityName: SdfEntityName,
        changes: StringKeyObject[],
        templateEntities: StringKeyObject[],
    ) {
        if (changes.length !== templateEntities.length) {
            throw Error('Arrays length cannot be different');
        }

        if (! this.sdf.Campaigns) {
            this.sdf.Campaigns = [];
        }

        for (let i=0; i<templateEntities.length; i++) {
            const generatedCampaign = this.generateSDFRow(
                templateEntities[i], changes[i], entityName
            );
            this.sdf.Campaigns.push(generatedCampaign);
            
            const advertiserId = templateEntities[i]['Advertiser Id'];
            this.generateSDFForAllChildEntities(
                advertiserId, changes[i], generatedCampaign
            );
        }

        return this;
    }

    /**
     * Generate the SDF for the whole entity hierarchy
     * 
     * @param advertiserId 
     * @param changes 
     * @param campaign 
     * @returns 
     */
    generateSDFForAllChildEntities(
        advertiserId: string,
        changes: StringKeyObject,
        campaign: StringKeyObject
    ) {
        // insertionOrders
        const insertionOrders = this.generateSDFForDirectChildren(
            "InsertionOrders", advertiserId, changes, [ campaign ]
        );
        
        if (!this.sdf.InsertionOrders) {
            this.sdf.InsertionOrders = [];
        }
        this.sdf.InsertionOrders.push(... insertionOrders);
        
        //  lineItems
        const lineItems = this.generateSDFForDirectChildren(
            "LineItems", advertiserId, changes, insertionOrders
        );
        
        if (!this.sdf.LineItems) {
            this.sdf.LineItems = [];
        }
        this.sdf.LineItems.push(... lineItems);

        //  AdGroups
        const adGroups = this.generateSDFForDirectChildren(
            "AdGroups", advertiserId, changes, lineItems
        );
        
        if (!this.sdf.AdGroups) {
            this.sdf.AdGroups = [];
        }
        this.sdf.AdGroups.push(... adGroups);

        //  Ads
        const adGroupAds = this.generateSDFForDirectChildren(
            "AdGroupAds", advertiserId, changes, adGroups
        );
        
        if (!this.sdf.AdGroupAds) {
            this.sdf.AdGroupAds = [];
        }
        this.sdf.AdGroupAds.push(... adGroupAds);

        return this;
    }

    /**
     * Generates and returns the SDFs for the direct children of the entities
     * 
     * @param entityName 
     * @param advertiserId 
     * @param changes 
     * @param parentEntities 
     * @returns 
     */
    generateSDFForDirectChildren(
        entityName: SdfEntityName,
        advertiserId: string,
        changes: StringKeyObject,
        parentEntities: StringKeyObject[],
    ) {
        console.log(
            'generateSDFForDirectChildren: Generating SDF for', entityName
        );

        const result:StringKeyObject[] = [];

        parentEntities.forEach(parent => {
            const parentIdKey = entityHierarchy[entityName]['parentId'];
            const parentIdValue = parent[ parentIdKey ];
            const parentIdCachedValue = ExtNameGenerator.getCached(
                entityHierarchy[entityName]['parent'] as SdfEntityName, 
                parentIdValue
            );

            // Fetching all children
            const currentSDF = this.sdfGetFromSheetOrDownload(
                entityName,
                advertiserId,
                false,
                { [ parentIdKey ]: parentIdCachedValue }
            );
    
            // Adding correct extId from the parent
            const extKey = entityName + Config.SDFGeneration.EntityNameDelimiter
                + entityHierarchy[entityName]['parentId'];
            
            currentSDF.forEach((entity:StringKeyObject) => {
                changes[ extKey ] = ExtNameGenerator.generateAndCache(
                    entityName,
                    entity[ entityHierarchy[entityName]['id'] ]
                );

                result.push( this.generateSDFRow(entity, changes, entityName) );
            });
        });
        
        return result;
    }

    /**
     * Returns the SDF as an array of Objects. If the sheet does not exist
     * then first download all SDF files from DV360 API and save to the sheet.
     * 
     * @param entity SDF for which DV360 entity should be returned.
     * @param advertiserId Advertiser ID, used as prefix for the "cache" sheet 
     *  name
     * @param reloadCache If TRUE, then the SDF will be downloaded and the cache
     *  overwritten
     */
    sdfGetFromSheetOrDownload(
        entity: SdfEntityName,
        advertiserId: string,
        reloadCache?: boolean,
        filters?: StringKeyObject,
    ) {
        const prefix = advertiserId + '-';
        const sheetName = prefix + 'SDF-' + entity + '.csv';
        if (reloadCache || ! SpreadsheetApp.getActive().getSheetByName(sheetName)) {
            const csvFiles = dv360.downloadSdfs(advertiserId);
            SheetUtils.putCsvFilesIntoSheets(csvFiles, prefix, true);
        }

        const sdfArray = SheetUtils.readSheetAsJSON(sheetName);
        return filters 
            ? sdfArray.filter(
                x => x[ Object.keys(filters)[0] ] === Object.values(filters)[0]
            )
            : sdfArray;
    }
}

export const sdfGenerator = new SdfGenerator();