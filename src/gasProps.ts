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
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null =
      setting.getSheetByName(this.SETTING_SHEET_NAME);
    if (!cashBook) {
      throw new Error('settingSheet was not found.');
    }
    return cashBook;
  }

  public get cashBookSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null =
      setting.getSheetByName(this.CASH_BOOK_SHEET_NAME);
    if (!cashBook) {
      throw new Error('cashBookSheet was not found.');
    }
    return cashBook;
  }

  public get mappingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const setting: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
    const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null =
      setting.getSheetByName(this.MAPPING_SHEET_NAME);
    if (!cashBook) {
      throw new Error('mappingSheet was not found.');
    }
    return cashBook;
  }

  public getReportSheet(actDate: string): GoogleAppsScript.Spreadsheet.Sheet {
    const report: GoogleAppsScript.Spreadsheet.Spreadsheet =
      SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
    const reportSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
      report.getSheetByName(actDate);
    if (!reportSheet) {
      throw new Error('reportSheet was not found. actDate:' + actDate);
    }
    return reportSheet;
  }

  public get payNowFolder(): GoogleAppsScript.Drive.Folder {
    return DriveApp.getFolderById(ScriptProps.instance.folderId);
  }

  public get payNowFolderUrl(): string {
    return (
      'https://drive.google.com/drive/folders/' +
      ScriptProps.instance.folderId +
      '?usp=sharing'
    );
  }

  public get settingSheetUrl(): string {
    return (
      'https://docs.google.com/spreadsheets/d/' +
      ScriptProps.instance.settingSheet +
      '/edit?usp=sharing&ccc=' +
      new Date().getTime()
    );
  }

  public get ReportSheetUrl(): string {
    return (
      'https://docs.google.com/spreadsheets/d/' +
      ScriptProps.instance.reportSheet +
      '/edit?usp=sharing&ccc=' +
      new Date().getTime()
    );
  }

  public isKanji(userId: string): boolean {
    return this.getKanjiIds().includes(userId);
  }

  private getKanjiIds(): string[] {
    const kanjiIds: string[] = [];
    const mappingSheet = GasProps.instance.mappingSheet;
    const values = mappingSheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][3] === '幹事') {
        kanjiIds.push(values[i][2]);
      }
    }
    return kanjiIds;
  }

  public getUnpaid(actDate: string): string[] {
    const unpaid: string[] = [];
    const repo = GasProps.instance.getReportSheet(actDate);
    const values = repo.getDataRange().getValues();
    for (let i = 9; i < values.length; i++) {
      if (!values[i][2]) {
        unpaid.push(values[i][0]);
      }
    }
    return unpaid;
  }
}
