import { onEditEvent, OnEditHandler } from "./trigger";
import { SheetCache, CacheValue } from "./sheet-cache";
import { SheetUtils } from "./sheets";

export const TriggerUtils = {
    /**
     * Generate onEdit trigger for the parent entity drop downs with 
     *  sub-entities (e.g. if the partner - parent is changed, then automatically 
     *  load the advertisers)
     *
     * @param sheetName The sheet name where we wait for the edit events
     * @param columnNumber Index of the column where we wait for the event
     *  (indexes start from 1)
     * @param parentIdIndex Column index of the parnet entity ID in the "cache"
     * @param cache Cache object (see SheetCache)
     * @param loadFunction Function that returns an array of values for the 
     *  sub-entity drop down
     * @returns Correct handler to be added in the installable trigger 
     *  (e.g. `onEditEvent.addHandler(handler)`)
     */
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
                  let triggerOnInit = false;

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
                      triggerOnInit = true;
                    }
                  } catch (e: any) {
                    SpreadsheetApp.getUi().alert('Error accured, try again...');
                    // On Error make all selects blank
                    range.setValue('');
                    targetRange.setValue('');
                    // Log error and stacktrace
                    console.log(e);
                    console.log(e.stack);
                  }
                  
                  targetRange.setValue(defaultDropdownValue);
                  // Usability feature: if only one element in drop down, 
                  // then (set it as a selected value and) trigger onEdit to load
                  // sub-entities for that one element.
                  if (triggerOnInit) {
                    onEditEvent({range: targetRange});
                  }
                }
            }
        };
    }
};