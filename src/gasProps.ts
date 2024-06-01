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

  public get payNowFolder(): GoogleAppsScript.Drive.Folder {
    return DriveApp.getFolderById(ScriptProps.instance.folderId);
  }

  public get payNowFolderUrl(): string {
    return 'https://drive.google.com/drive/folders/' + ScriptProps.instance.folderId + '?usp=sharing';
  }

  public get settingSheetUrl(): string {
    if (ScriptProps.isTesting()) {
      return 'https://docs.google.com/spreadsheets/d/' + ScriptProps.instance.settingSheet + '?usp=sharing';
    }
    return 'https://docs.google.com/spreadsheets/d/' + ScriptProps.instance.settingSheet + '/edit?usp=sharing&ccc=' + new Date().getTime();
  }

  public get ReportSheetUrl(): string {
    if (ScriptProps.isTesting()) {
      return 'https://docs.google.com/spreadsheets/d/' + ScriptProps.instance.reportSheet + '?usp=sharing';
    }
    return 'https://docs.google.com/spreadsheets/d/' + ScriptProps.instance.reportSheet + '/edit?usp=sharing&ccc=' + new Date().getTime();
  }
}
