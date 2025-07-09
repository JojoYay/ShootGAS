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
    private G_RANKING_SHEET_NAME: string = '得点王ランキング';
    private A_RANKING_SHEET_NAME: string = 'アシスト王ランキング';
    private O_RANKING_SHEET_NAME: string = '岡本カップランキング';
    private VIDEO_SHEET: string = 'videos';
    public WEIGHT_RECORD_SHEET_NAME: string = 'WeightRecord';

    public get settingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.SETTING_SHEET_NAME);
        if (!cashBook) {
            throw new Error('settingSheet was not found.');
        }
        return cashBook;
    }

    public get videoSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const report: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const videos: GoogleAppsScript.Spreadsheet.Sheet | null = report.getSheetByName(this.VIDEO_SHEET);
        if (!videos) {
            throw new Error('videos was not found.');
        }
        return videos;
    }

    public get gRankingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const gRanking: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.G_RANKING_SHEET_NAME);
        if (!gRanking) {
            throw new Error('gRankingSheet was not found.');
        }
        return gRanking;
    }

    public get aRankingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const aRanking: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.A_RANKING_SHEET_NAME);
        if (!aRanking) {
            throw new Error('aRankingSheet was not found.');
        }
        return aRanking;
    }

    public get oRankingSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const oRanking: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.O_RANKING_SHEET_NAME);
        if (!oRanking) {
            throw new Error('oRankingSheet was not found.');
        }
        return oRanking;
    }

    public get cashBookSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
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

    public get eventResultSheet(): GoogleAppsScript.Spreadsheet.Sheet {
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

    public get expenseFolder(): GoogleAppsScript.Drive.Folder {
        return DriveApp.getFolderById(ScriptProps.instance.expenseFolder);
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

    public generateSheetUrl(prop: string): string {
        if (ScriptProps.isTesting()) {
            return 'https://docs.google.com/spreadsheets/d/' + prop + '?usp=sharing';
        }
        return 'https://docs.google.com/spreadsheets/d/' + prop + '/edit?usp=sharing&ccc=' + new Date().getTime();
    }

    public get weightRecordSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        let weightRecord: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName(this.WEIGHT_RECORD_SHEET_NAME);

        if (!weightRecord) {
            // シートが存在しない場合は新規作成
            weightRecord = setting.insertSheet(this.WEIGHT_RECORD_SHEET_NAME);

            // ヘッダーを設定
            const headers = ['id', 'userId', 'height', 'weight', 'bfp', 'date'];
            weightRecord.getRange(1, 1, 1, headers.length).setValues([headers]);
        }

        return weightRecord;
    }
}
