import { Config } from "./config";
import { SheetUtils } from "./sheet-utils";

export const Setup = {
    createMenu() {
        SpreadsheetApp.getUi()
            .createMenu(Config.Menu.Name)
            .addItem(Config.Menu.GenerateSDFForActiveSheet, 'generateSDFForActiveSheet')
            .addSeparator()
            .addItem(Config.Menu.Setup, 'setup')
            .addItem(Config.Menu.ClearCache, 'clearCache')
            .addToUi();
    },

    setUpCampaignsSheet(partners: string[]) {
        const sheet = SheetUtils
            .getOrCreateSheet(Config.WorkingSheet.Campaigns)
            .activate();
        if (!sheet.getLastRow()) {
            sheet.getRange('A1:D1')
                .setValues([Config.WorkingSheet.CampaignsHeaders]);

            sheet.getRange('A1:J1')
                .setBackground('#a4c2f4')
                .setFontWeight('bold');
        }

        // Add partners drop down
        SheetUtils.setRangeDropDown(
            sheet.getRange('A2:A'),
            partners
        );
    }
}