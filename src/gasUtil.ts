import { GasProps } from './gasProps';
import { LineUtil } from './lineUtil';
import { ScriptProps } from './scriptProps';

const lineUtil: LineUtil = new LineUtil();
export class GasUtil {
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
        const repo = this.getReportSheet(actDate, false);
        const values = repo.getDataRange().getValues();
        for (let i = 9; i < values.length; i++) {
            if (!values[i][2]) {
                unpaid.push(values[i][0]);
            }
        }
        return unpaid;
    }

    public getUnRegister(actDate: string): string[] {
        const unregister: string[] = [];
        const repo = this.getReportSheet(actDate, false);
        const values = repo.getDataRange().getValues();
        for (let i = 9; i < values.length; i++) {
            if (!values[i][1]) {
                unregister.push(values[i][0]);
            }
        }
        return unregister;
    }

    public registerMapping(lineName: string, densukeName: string, userId: string): void {
        const mappingSheet = GasProps.instance.mappingSheet;
        const values = mappingSheet.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][0] === lineName) {
                mappingSheet.deleteRow(i + 1);
                break;
            }
        }
        mappingSheet.appendRow([lineName, densukeName, userId]);
    }

    public uploadPayNowPic(densukeName: string, messageId: string, actDate: string): void {
        const fileNm = actDate + '_' + densukeName;
        lineUtil.getLineImage(messageId, fileNm, actDate);
    }

    public getReportSheet(actDate: string, isGenerate: boolean = false): GoogleAppsScript.Spreadsheet.Sheet {
        const report: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        let reportSheet: GoogleAppsScript.Spreadsheet.Sheet | null = report.getSheetByName(actDate);
        if (!reportSheet) {
            if (isGenerate) {
                reportSheet = report.insertSheet(actDate);
                reportSheet.activate();
                report.moveActiveSheet(1);
            } else {
                throw new Error('reportSheet was not found. actDate:' + actDate);
            }
        }
        return reportSheet;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public getLineUserIdRangeValue(): any[][] {
        const mappingSheet = GasProps.instance.mappingSheet;
        return mappingSheet.getDataRange().getValues();
    }

    public getLineUserId(densukeName: string): string {
        let userId = '';
        const mappingSheet = GasProps.instance.mappingSheet;
        const values = mappingSheet.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][1] === densukeName) {
                userId = values[i][2];
                break;
            }
        }
        return userId;
    }

    public getLineName(densukeName: string) {
        let lineName = null;
        const mappingSheet = GasProps.instance.mappingSheet;
        const values = mappingSheet.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][1] === densukeName) {
                lineName = values[i][0];
                break;
            }
        }
        return lineName;
    }

    // public isRegistered(userId: string): boolean {
    //     return !!this.getDensukeName(lineUtil.getLineDisplayName(userId));
    // }

    public getNickname(userId: string): string {
        let nickName = null;
        const mappingSheet = GasProps.instance.mappingSheet;
        const values = mappingSheet.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][2] === userId) {
                nickName = values[i][1];
                break;
            }
        }
        return nickName;
    }

    //depliciated
    public getDensukeName(lineName: string): string {
        let densukeName = null;
        const mappingSheet = GasProps.instance.mappingSheet;
        const values = mappingSheet.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][0] === lineName) {
                densukeName = values[i][1];
                break;
            }
        }
        return densukeName;
    }

    public updateLineNameOfLatestReport(lineName: string, densukeName: string, actDate: string): void {
        const repo = this.getReportSheet(actDate, false);
        const values = repo.getDataRange().getValues();
        for (let i = 0; i < values.length; i++) {
            if (values[i][0] === densukeName) {
                repo.getRange(i + 1, 2).setValue(lineName);
                break;
            }
        }
    }

    public updatePaymentStatus(desunekeName: string, actDate: string): void {
        const repo = this.getReportSheet(actDate, false);
        const values = repo.getDataRange().getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][0] === desunekeName) {
                const val: GoogleAppsScript.Spreadsheet.RichTextValue | null = this.getPaymentUrl(desunekeName, actDate);
                if (val) {
                    repo.getRange(i + 1, 3).setRichTextValue(val);
                }
                break;
            }
        }
    }

    public getPaymentUrl(densukeName: string, actDate: string): GoogleAppsScript.Spreadsheet.RichTextValue | null {
        const payNowOwner = this.getPaynowOwner();
        if (payNowOwner === densukeName) {
            return SpreadsheetApp.newRichTextValue().setText('PayNow口座主').build();
        }
        return this.getFileUrlInFolder(actDate, densukeName);
    }

    public getPaynowOwner(): string {
        const settingSheet = GasProps.instance.settingSheet;
        const payNowOwner = settingSheet.getRange('B6').getValue();
        return payNowOwner;
    }

    private getFileUrlInFolder(actDate: string, densukeName: string): GoogleAppsScript.Spreadsheet.RichTextValue | null {
        // const folder = GasProps.instance.payNowFolder;
        const lineUtil: LineUtil = new LineUtil();
        const folder = lineUtil.createPayNowFolder(actDate);
        if (!folder) {
            return null; //folderは必ず作られる
        }
        const fileName = actDate + '_' + densukeName;
        const files = folder.getFilesByName(fileName);
        const urls: string[] = [];
        if (!files.hasNext()) {
            return null;
        }
        while (files.hasNext()) {
            const file = files.next();
            urls.push(file.getUrl());
        }
        console.log(urls);
        const rtv = SpreadsheetApp.newRichTextValue().setText('');
        if (urls.length > 0) {
            rtv.setText(urls.join('\n'));
            let totalChar = -1;
            for (let i = 0; i < urls.length; i++) {
                const start = totalChar + 1;
                const end = start + urls[i].length;
                totalChar = end;
                rtv.setLinkUrl(start, end, urls[i]);
            }
        }
        return rtv.build();
    }

    public archiveFiles(actDate: string): void {
        const sourceFolder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder;
        const destinationFolder: GoogleAppsScript.Drive.Folder = GasProps.instance.archiveFolder;
        const files: GoogleAppsScript.Drive.FileIterator = sourceFolder.getFiles();
        const prefix: string = actDate + '_';
        while (files.hasNext()) {
            const file: GoogleAppsScript.Drive.File = files.next();
            if (!file.getName().startsWith(prefix)) {
                file.moveTo(destinationFolder);
            }
        }
    }

    public createSpreadSheet(title: string, folder: GoogleAppsScript.Drive.Folder, header: string[]): GoogleAppsScript.Spreadsheet.Spreadsheet {
        let spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
        const searchQuery2 = `title = '${title}' and '${folder.getId()}' in parents`; // より正確なファイル名検索クエリ
        const fileIt = folder.searchFiles(searchQuery2);
        if (fileIt.hasNext()) {
            const sheetFile = fileIt.next();
            spreadSheet = SpreadsheetApp.openById(sheetFile.getId());
        } else {
            spreadSheet = SpreadsheetApp.create(title); // スプレッドシートを作成
            const sheet = spreadSheet.getActiveSheet();
            sheet.appendRow(header);

            // 新しく作成したスプレッドシートを指定のフォルダに移動
            const file = DriveApp.getFileById(spreadSheet.getId());
            folder.addFile(file); // フォルダにファイルを追加
            DriveApp.getRootFolder().removeFile(file); // ルートフォルダからファイルを削除
        }
        return spreadSheet;
    }
}
