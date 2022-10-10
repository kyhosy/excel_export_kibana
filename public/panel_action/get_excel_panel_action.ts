import moment from 'moment-timezone';
import { CoreSetup } from 'src/core/public';
import { writeFile, read } from 'xlsx';
import ExcelJS from "exceljs"
import Papa from "papaparse";

import { IncompatibleActionError } from '../../../../src/plugins/ui_actions/public';
import type { UiActionsActionDefinition as ActionDefinition } from '../../../../src/plugins/ui_actions/public';
import type { ISearchEmbeddable, SavedSearch } from '../../../../src/plugins/discover/public';
import {
  loadSharingDataHelpers,
  SEARCH_EMBEDDABLE_TYPE,
} from '../../../../src/plugins/discover/public';
import { IEmbeddable, ViewMode } from '../../../../src/plugins/embeddable/public';
import { API_GENERATE_IMMEDIATE } from '../../common/constants';
import type { JobParamsDownloadCSV } from '../../../../x-pack/plugins/reporting/server/export_types/csv_searchsource_immediate/types';
import { DataPublicPluginStart } from '../../../../src/plugins/data/public';
import ExcelHelpers from '../../common/ExcelHelpers';

interface ActionContext {
  embeddable: ISearchEmbeddable;
}

function isSavedSearchEmbeddable(
  embeddable: IEmbeddable | ISearchEmbeddable
): embeddable is ISearchEmbeddable {
  return embeddable.type === SEARCH_EMBEDDABLE_TYPE;
}

export class GetExcelPanelAction implements ActionDefinition<ActionContext> {
  readonly id = 'downloadExcelReport';

  private isDownloading: boolean;
  private core: any;
  constructor(core: CoreSetup, data: DataPublicPluginStart ) {
    this.isDownloading = false;
    this.core = core;
    this.data = data;
  }
  public isCompatible = async (context: ActionContext) => {
    const { embeddable } = context;
    return embeddable.getInput().viewMode !== ViewMode.EDIT && isSavedSearchEmbeddable(embeddable);
  };

  public getIconType(): string {
    return 'document';
  }

  public getDisplayName(): string {
    return 'Export as Excel';
  }

  // @ts-ignore
  private getHeaderFile(header: Array<string>) : Array<string> {

  }

  public async getSearchSource(savedSearch: SavedSearch, embeddable: ISearchEmbeddable) {
    const { getSharingData } = await loadSharingDataHelpers();
    const map1 = new Map(Object.entries(this.core.uiSettings.defaults));
     const ps={
      uiSettings: map1,
      data: this.data
    }
    const searchSource = savedSearch.searchSource;
    const index = searchSource.getField('index');
    const existingFilter = searchSource.getField('filter');
    const cloneSaveSearch = savedSearch;

    console.log('savedSearch=>>',savedSearch);
    const sorted = cloneSaveSearch.sort
    if(Array.isArray(sorted) && sorted.length == 0){
      cloneSaveSearch.sort = [
        "_score",
        "desc"
      ]
    }
    console.log('this.core.uiSettings=>>',this.core.uiSettings);
    console.log('cloneSaveSearch=>>',cloneSaveSearch);

    return await getSharingData(
        savedSearch.searchSource,
        cloneSaveSearch, // TODO: get unsaved state (using embeddale.searchScope): https://github.com/elastic/kibana/issues/43977
        ps
     );
  }

  public execute = async (context: ActionContext) => {
    const { embeddable } = context;

    if (!isSavedSearchEmbeddable(embeddable)) {
      throw new IncompatibleActionError();
    }

    if (this.isDownloading) {
      return;
    }

    const savedSearch = embeddable.getSavedSearch();
    const { columns,  getSearchSource} = await this.getSearchSource(savedSearch, embeddable);
    const searchSource = getSearchSource();
    const kibanaTimezone = this.core.uiSettings.get('dateFormat:tz');
    console.log(' kibanaTimezone==>>', kibanaTimezone);
    const browserTimezone = kibanaTimezone === 'Browser' ? moment.tz.guess() : kibanaTimezone;
    const immediateJobParams: JobParamsDownloadCSV = {
      searchSource,
      columns,
      browserTimezone,
      title: savedSearch.title,
    };

    const body = JSON.stringify(immediateJobParams);
    this.isDownloading = true;

    this.core.notifications.toasts.addSuccess({
      title: `Excel Download Started`,
      text: `Your Excel will download momentarily.`,
      'data-test-subj': 'csvDownloadStarted',
    });

    await this.core.http
      .post(`${API_GENERATE_IMMEDIATE}`, { body })
      .then((rawResponse: string) => {
        this.isDownloading = false;
        //TODO: convert header from key to label
        // const workbook = read(rawResponse, { type: 'string', raw: true });
        // writeFile(workbook, `${embeddable.getSavedSearch().title}.xlsx`, { type: 'binary' });
        const fileName = `${embeddable.getSavedSearch().title}.xlsx`
        const mappingCols = savedSearch?.searchSource?.fields?.index?.fieldAttrs
        console.log('mappingCols>>>', mappingCols)
        ExcelHelpers.downloadExcelFromCsv(rawResponse, fileName, mappingCols, () => {

        }, () => {
          this.onGenerationFail.bind(this)
        })
      })
      .catch(this.onGenerationFail.bind(this));
  };

  private onGenerationFail(error: Error) {
    this.isDownloading = false;
    this.core.notifications.toasts.addDanger({
      title: `Excel download failed`,
      text: `We couldn't generate your Excel at this time.`,
      'data-test-subj': 'downloadExcelFail',
    });
  }
}
