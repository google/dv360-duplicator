import { Config } from "./config";

export const Setup = {
    createMenu() {
        SpreadsheetApp.getUi()
            .createMenu(Config.Menu.Name)
            //.addItem(Config.Menu.ReloadCacheForActiveSheet, 'reloadCacheForActiveSheet')
            .addItem(Config.Menu.GenerateSDFForActiveSheet, 'generateSDFForActiveSheet')
            .addToUi();
    }
}