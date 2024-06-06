import { ScriptProps } from './scriptProps';
export class GasProps {
  private static _instance: GasProps | null = null;

  private constructor() {}

  public static get instance(): GasProps {
    if (!this._instance) {
      this._instance = new GasProps();
    }
    return this._instance;
  }
  private SETTING_SHEET_NAME: string = 'Settings';
  private CASH_BOOK_SHEET_NAME: string = 'CashBook';
  private MAPPING_SHEET_NAME: string = 'DensukeMapping';
  private EVENT_DATA_SHEET_NAME: string = 'EventData';
  private PERSONAL_TOTAL_SHEET_NAME: string = 'Total';

  public get settingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.SETTING_SHEET_NAME);
    if (!cashBook) {
      throw new Error('settingSheet was not found.');
    }
    return cashBook;
  }

  public get cashBookSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.CASH_BOOK_SHEET_NAME);
    if (!cashBook) {
      throw new Error('cashBookSheet was not found.');
    }
    return cashBook;
  }

  public get mappingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.MAPPING_SHEET_NAME);
    if (!cashBook) {
      throw new Error('mappingSheet was not found.');
    }
    return cashBook;
  }

  public get eventResultheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const eventResultsSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    const eventData: GoogleAppsScript.Spreadsheet.Sheet | null = eventResultsSS.getSheetByName(this.EVENT_DATA_SHEET_NAME);
    if (!eventData) {
      throw new Error('EventDataSheet was not found.');
    }
    return eventData;
  }

  public get personalTotalSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const eventResultsSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    const eventData: GoogleAppsScript.Spreadsheet.Sheet | null = eventResultsSS.getSheetByName(this.PERSONAL_TOTAL_SHEET_NAME);
    if (!eventData) {
      throw new Error('PersonalTotalSheet was not found.');
    }
    return eventData;
  }

  public get payNowFolder(): GoogleAppsScript.Drive.Folder {
    return DriveApp.getFolderById(ScriptProps.instance.folderId);
  }

  public get archiveFolder(): GoogleAppsScript.Drive.Folder {
    return DriveApp.getFolderById(ScriptProps.instance.archiveFolder);
  }

  public get payNowFolderUrl(): string {
    return 'https://drive.google.com/drive/folders/' + ScriptProps.instance.folderId + '?usp=sharing';
  }

  public get settingSheetUrl(): string {
    return this.generateSheetUrl(ScriptProps.instance.settingSheet);
  }

  public get reportSheetUrl(): string {
    return this.generateSheetUrl(ScriptProps.instance.reportSheet);
  }

  public get eventResultUrl(): string {
    return this.generateSheetUrl(ScriptProps.instance.eventResults);
  }

  private generateSheetUrl(prop: string): string {
    if (ScriptProps.isTesting()) {
      return 'https://docs.google.com/spreadsheets/d/' + prop + '?usp=sharing';
    }
    return 'https://docs.google.com/spreadsheets/d/' + prop + '/edit?usp=sharing&ccc=' + new Date().getTime();
  }
}
