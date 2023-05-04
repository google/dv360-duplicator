import { Config } from "./config";

export const Setup = {
    createMenu() {
        SpreadsheetApp.getUi()
            .createMenu(Config.Menu.Name)
            .addItem(Config.Menu.GenerateSDFForActiveSheet, 'generateSDFForActiveSheet')
            .addSeparator()
            .addItem(Config.Menu.ClearCache, 'clearCache')
            .addToUi();
    }
}