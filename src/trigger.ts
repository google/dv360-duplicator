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
  shouldRun: (event: GoogleAppsScript.Events.SheetsOnEdit) => boolean;
  /**
   * Runs the actual handler.
   * @param event the event to handle
   */
  run: (event: GoogleAppsScript.Events.SheetsOnEdit) => void;
}

const onEditHandlers: OnEditHandler[] = [];

/**
 * The wrapper function which is called on every edit and determines which
 * handlers should be executed.
 * @param event an edit event
 */
export function onEditEvent(event: GoogleAppsScript.Events.SheetsOnEdit) {
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
