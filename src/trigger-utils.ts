import { OnEditHandler } from "./trigger";
import { SheetCache, CacheValue } from "./sheet-cache";
import { SheetUtils } from "./sheets";

export const TriggerUtils = {
    generateOnEditHandler(
        sheetName: string, 
        columnNumber: number,
        parentIdIndex: number,
        cache: SheetCache,
        loadFunction: ((x: string) => CacheValue[] | CacheValue[][])
    ): OnEditHandler {
        return {
            shouldRun({ range }) {
                return (
                  range.getSheet().getName() === sheetName &&
                  range.getColumn() === columnNumber &&
                  !!range.getValue()
                );
            },
            run({ range }) {
                console.log(`Running changed handler`);
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
            
                const parentName = '' + range.getValues()[0];
                console.log(`Parent name: ${parentName}`);
            
                let defaultDropdownValue = '';
                const parentValues = cache.find(parentName, 0);
                if (parentValues) {
                  const parentId = "" + parentValues[parentIdIndex];
                  console.log('parentId', parentId);
                  try {
                    const children = loadFunction(parentId);
                    console.log('Children: ', children);
                    const names = children.map(
                      (child) => child && Array.isArray(child) && child.length
                        ? '' + child[0]
                        : ''
                    );
                    console.log('Children names', names);
                    SheetUtils.setRangeDropDown(targetRange, names);
                    
                    // Usability feature: if only one element in drop down, 
                    // then set it as a selected value
                    if (1 === names.length) {
                        defaultDropdownValue = names[0];
                    }
                  } catch (e: any) {
                    SpreadsheetApp.getUi().alert('Error accured, try again...');
                    range.setValue('');
                    console.log(e);
                  }
                  
                  targetRange.setValue(defaultDropdownValue);
                }
            }
        };
    }
};