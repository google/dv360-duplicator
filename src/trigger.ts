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

export type SheetsOnEditEventSimplified = {
  range: GoogleAppsScript.Spreadsheet.Range,
};

/**
 * Interface for a handler. This supports filtering all events for those that
 * should be handled.
 */
export interface OnEditHandler {
  /**
   * Determines if a handler should run given the onEdit event.
   * @param event the event to inspect
   * @returns true if the handler should run, false otherwise
   */
  shouldRun: (event: SheetsOnEditEventSimplified) => boolean;
  /**
   * Runs the actual handler.
   * @param event the event to handle
   */
  run: (event: SheetsOnEditEventSimplified) => void;
}

const onEditHandlers: OnEditHandler[] = [];

/**
 * The wrapper function which is called on every edit and determines which
 * handlers should be executed.
 * @param event an edit event
 */
export function onEditEvent(event: SheetsOnEditEventSimplified) {
  onEditHandlers.forEach(
    (handler) => handler.shouldRun(event) && handler.run(event)
  );
}
/**
 * Retrieve the trigger that targets the onEditEvent function.
 * @returns the trigger or undefined if the filter cannot be found
 */
onEditEvent.getTriggers = () => {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.filter(
    (trigger) => trigger.getHandlerFunction() === onEditEvent.name
  );
};

/**
 * Installs a trigger for the onEdit function to be called on edit.
 */
onEditEvent.install = () => {
  const triggers = onEditEvent.getTriggers();
  if (!triggers.length) {
    ScriptApp.newTrigger(onEditEvent.name)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
};
/**
 * Uninstalls the trigger for the onEdit function.
 */
onEditEvent.uninstall = () => {
  const triggers = onEditEvent.getTriggers();
  if (triggers.length) {
    triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  }
};

/**
 * Adds a handler to the onEdit trigger.
 * @param handler the handler
 */
onEditEvent.addHandler = (handler: OnEditHandler) => {
  onEditHandlers.push(handler);
};