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

import { onEditEvent, OnEditHandler } from "./trigger";
import { SheetCache, CacheValue } from "./sheet-cache";
import { SheetUtils } from "./sheet-utils";

export const TriggerUtils = {
    /**
     * Generate onEdit trigger for the parent entity drop downs with 
     *  sub-entities (e.g. if the partner - parent is changed, then automatically 
     *  load the advertisers)
     *
     * @param sheetName The sheet name where we wait for the edit events
     * @param columnNumber Index of the column where we wait for the event
     *  (indexes start from 1)
     * @param parentIdIndex Column index of the parent entity ID in the "cache"
     * @param cache Cache object (see SheetCache)
     * @param loadFunction Function that returns an array of values for the 
     *  sub-entity drop down
     * @param runFunction If specified, this function will be triggered instead 
     *  of default "run"(loading sub-entities) function
     * @returns Correct handler to be added in the installable trigger 
     *  (e.g. `onEditEvent.addHandler(handler)`)
     */
    generateOnEditHandler(
        sheetName: string, 
        columnNumber: number,
        parentIdIndex: number,
        cache: SheetCache,
        loadFunction?: ((x: string) => CacheValue[] | CacheValue[][]),
        runFunction?: (r: GoogleAppsScript.Spreadsheet.Range, c: SheetCache) => any
    ): OnEditHandler {
        return {
            shouldRun({ range }) {
                const isOnlyOneColumnUpdated = !(range.getLastColumn() - range.getColumn());
                const isThisTheColumnWeWaitFor = 
                  range.getSheet().getName() === sheetName
                  && range.getColumn() === columnNumber;
                const isNewValueNonEmpty = !!range.getValue();

                return isOnlyOneColumnUpdated
                  && isThisTheColumnWeWaitFor
                  && isNewValueNonEmpty;
            },
            run({ range }) {
                console.log(
                  `Running changed handler. Source range: [${range.getColumn()}, ${range.getRow()}]`
                );

                if (runFunction) {
                  console.log(`Trigerring a custom runFunction ${runFunction}`);
                  return runFunction(range, cache)
                }

                if (! loadFunction) {
                  throw Error(
                    'Either "loadFunction" or "runFunction" should be specified'
                  );
                }

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
            
                const parentName = '' + range.getDisplayValues()[0];
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