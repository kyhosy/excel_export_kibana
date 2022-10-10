import { UiActionsSetup, UiActionsStart } from 'src/plugins/ui_actions/public';
import { CoreSetup, CoreStart, Plugin } from 'kibana/public';
import { GetExcelPanelAction } from './panel_action/get_excel_panel_action';
import { CONTEXT_MENU_TRIGGER } from '../../../src/plugins/embeddable/public';
import { DataPublicPluginStart } from '../../../src/plugins/data/public';


export interface SavedSearchExcelExportPluginSetupDependencies {
  uiActions: UiActionsSetup;
  data: DataPublicPluginStart;
}

export interface SavedSearchExcelExportPluginStartDependencies {
  uiActions: UiActionsStart;
}

export class ExcelExportPlugin
  implements
    Plugin<
      void,
      void,
      SavedSearchExcelExportPluginSetupDependencies,
      SavedSearchExcelExportPluginStartDependencies
    > {
  public setup(
    core: CoreSetup<SavedSearchExcelExportPluginSetupDependencies>,
    { uiActions, data }: SavedSearchExcelExportPluginSetupDependencies
  ) {
    const action = new GetExcelPanelAction(core, data);
    uiActions.registerAction(action);
    uiActions.attachAction(CONTEXT_MENU_TRIGGER, action.id);
    uiActions.addTriggerAction(CONTEXT_MENU_TRIGGER, action);
  }

  public start(core: CoreStart, plugins: SavedSearchExcelExportPluginSetupDependencies) {}

  public stop() {}
}
