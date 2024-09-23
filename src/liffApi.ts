import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { GetEventHandler } from './getEventHandler';
import { LineUtil } from './lineUtil';
import { ScriptProps } from './scriptProps';

export class LiffApi {
    private test(getEventHandler: GetEventHandler): void {
        const value: string = getEventHandler.e.parameters['param'][0];
        getEventHandler.result = { result: value };
    }

    private getVideo(getEventHandler: GetEventHandler): void {
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        getEventHandler.result = { result: videos.getDataRange().getValues() };
    }

    private getPayNow(getEventHandler: GetEventHandler): void {
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        // getEventHandler.result = { result: members };
        getEventHandler.result.payNow = addy;
    }

    private getMembers(getEventHandler: GetEventHandler): void {
        const den: DensukeUtil = new DensukeUtil();
        const members = den.extractMembers();
        // getEventHandler.result = { result: members };
        getEventHandler.result.members = members;
    }

    private getRegisteredMembers(getEventHandler: GetEventHandler): void {
        const members: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        getEventHandler.result.members = members.getDataRange().getValues();
    }

    private getDensukeName(getEventHandler: GetEventHandler): void {
        const gasUtil: GasUtil = new GasUtil();
        const lineUtil: LineUtil = new LineUtil();
        const userId = getEventHandler.e.parameters['userId'][0];
        // getEventHandler.result = { result: gasUtil.getDensukeName(lineUtil.getLineDisplayName(userId)) };
        getEventHandler.result.densukeName = gasUtil.getDensukeName(lineUtil.getLineDisplayName(userId));
    }

    private getRanking(getEventHandler: GetEventHandler): void {
        const gRank = GasProps.instance.gRankingSheet.getDataRange().getValues();
        const aRank = GasProps.instance.aRankingSheet.getDataRange().getValues();
        const oRank = GasProps.instance.oRankingSheet.getDataRange().getValues();
        getEventHandler.result.gRank = gRank;
        getEventHandler.result.aRank = aRank;
        getEventHandler.result.oRank = oRank;
    }

    private getExpenseWithStatus(getEventHandler: GetEventHandler): void {
        const title: string = getEventHandler.e.parameters['title'][0];
        const userId: string = getEventHandler.e.parameters['userId'][0];
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.expenseFolder);
        const folderIt = rootFolder.getFoldersByName(title);
        if (!folderIt.hasNext()) {
            getEventHandler.result.statusMsg = 'no such expense folder found:' + title;
            console.log('no such expense folder found:' + title);
        }
        const expenseFolder = folderIt.next();
        const lineUtil: LineUtil = new LineUtil();
        // console.log('userId ' + userId);
        const lineName: string = lineUtil.getLineDisplayName(userId);
        const fileIt = expenseFolder.getFilesByName(title + '_' + lineName);
        if (fileIt.hasNext()) {
            const file = fileIt.next();
            getEventHandler.result.statusMsg = '支払い済み';
            const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
            getEventHandler.result.picUrl = picUrl;
        } else {
            let spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
            const fileIt2 = expenseFolder.getFilesByName(title);
            if (fileIt2.hasNext()) {
                const sheetFile = fileIt2.next();
                spreadSheet = SpreadsheetApp.openById(sheetFile.getId());
            } else {
                throw new Error('SpreadSheet is not available:' + title);
            }

            const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getActiveSheet();
            const sheetVal = sheet.getDataRange().getValues();
            const gasUtil: GasUtil = new GasUtil();
            const densukeName = gasUtil.getDensukeName(lineName);
            const userRow = sheetVal.find(item => item[0] === densukeName);
            const settingSheet = GasProps.instance.settingSheet;
            const addy = settingSheet.getRange('B2').getValue();
            if (userRow) {
                getEventHandler.result.statusMsg =
                    '支払額：$' + userRow[2] + ' PayNow先:' + addy + '\n支払い済みのスクリーンショットをこちらにアップロードして下さい';
                getEventHandler.result.picUrl = '';
            } else {
                getEventHandler.result.statusMsg = '支払い人として登録されていません。管理者にご確認下さい。';
            }
        }
    }

    private getStats(getEventHandler: GetEventHandler): void {
        const resultSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.personalTotalSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = resultSheet.getDataRange().getValues();
        getEventHandler.result.stats = resultValues;
    }

    private getUsers(getEventHandler: GetEventHandler): void {
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = mappingSheet.getDataRange().getValues();
        getEventHandler.result.users = resultValues;
    }

    private generateExReport(getEventHandler: GetEventHandler): void {
        const users: string[] = getEventHandler.e.parameters['users'];
        const price: string = getEventHandler.e.parameters['price'][0];
        const title: string = getEventHandler.e.parameters['title'][0];
        const payNow: string = getEventHandler.e.parameters['payNow'][0];

        let newSpreadsheet = null;
        const folder: GoogleAppsScript.Drive.Folder = GasProps.instance.expenseFolder;
        const folderIt = folder.getFoldersByName(title);
        let expenseFolder: GoogleAppsScript.Drive.Folder;
        if (folderIt.hasNext()) {
            expenseFolder = folderIt.next();
        } else {
            expenseFolder = folder.createFolder(title);
        }
        const fileIt = expenseFolder.getFilesByName(title);
        if (fileIt.hasNext()) {
            const file = fileIt.next();
            newSpreadsheet = SpreadsheetApp.openById(file.getId());
        } else {
            newSpreadsheet = SpreadsheetApp.create(title);
            const fileId = newSpreadsheet.getId();
            const file = DriveApp.getFileById(fileId);
            file.moveTo(expenseFolder);
        }
        const fileId = newSpreadsheet.getId();
        const sheet = newSpreadsheet.getActiveSheet();
        sheet.clear(); //まず全部消す
        sheet.appendRow(['名称', title]);
        sheet.appendRow(['人数', users.length]);
        sheet.appendRow(['合計金額', users.length * Number(price)]);
        sheet.appendRow(['PayNow先', payNow]);
        sheet.appendRow(['参加者（伝助名称）', '参加者（Line名称）', '金額', '支払い状況', '受け取り状況']);
        let index = 6;
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const mapVal = mappingSheet.getDataRange().getValues();
        const status: string[] = ['受渡済', ''];
        const statusVal = SpreadsheetApp.newDataValidation().requireValueInList(status).build();
        for (const user of users) {
            const mapRow = mapVal.find(item => item[1] === user);
            console.log(user);
            sheet.getRange(index, 1).setValue(user);
            sheet.getRange(index, 2).setValue(mapRow?.[0]);
            sheet.getRange(index, 3).setValue(price);
            // console.log('user ' + user + ' maprow1 ' + mapRow?.[1]);
            const fileIt = expenseFolder.getFilesByName(title + '_' + mapRow?.[0]);
            if (fileIt.hasNext()) {
                const file = fileIt.next();
                const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
                sheet.getRange(index, 4).setValue(picUrl);
            }
            // sheet.getRange(index, 1).setValue(lu.getLineDisplayName());
            sheet.getRange(index, 5).setDataValidation(statusVal);
            index++;
        }

        const lastCol = sheet.getLastColumn();
        const lastRow = sheet.getLastRow();
        sheet.getRange(5, 1, lastRow - 4, lastCol).setBorder(true, true, true, true, true, true);
        sheet.getRange(5, 1, 1, lastCol).setBackground('#fff2cc');

        const range = sheet.getRange(6, 3, lastRow - 5, 1);
        const formula = `=SUM(${range.getA1Notation()})`;
        sheet.getRange(3, 2).setFormula(formula);

        getEventHandler.result.folder = 'https://drive.google.com/drive/folders/' + ScriptProps.instance.folderId + '?usp=sharing';
        getEventHandler.result.sheet = GasProps.instance.generateSheetUrl(fileId);
        getEventHandler.result.url = ScriptProps.instance.liffUrl + '/expense/input?title=' + title;
    }

    private register(getEventHandler: GetEventHandler): void {
        const userId = getEventHandler.e.parameters['userId'][0];
        const lineUtil: LineUtil = new LineUtil();
        const gasUtil: GasUtil = new GasUtil();
        const densukeUtil: DensukeUtil = new DensukeUtil();
        const lineName = lineUtil.getLineDisplayName(userId);
        const lang = lineUtil.getLineLang(userId);
        const $ = densukeUtil.getDensukeCheerio();
        densukeUtil.generateSummaryBase($); //先に更新しておかないとエラーになる（伝助が更新されている場合）
        const members = densukeUtil.extractMembers($);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const densukeNameNew = getEventHandler.e.parameters['densukeName'][0];
        if (members.includes(densukeNameNew)) {
            if (densukeUtil.hasMultipleOccurrences(members, densukeNameNew)) {
                if (lang === 'ja') {
                    getEventHandler.result = {
                        result: '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。',
                    };
                } else {
                    getEventHandler.result = {
                        result:
                            "There are multiple entries with the name '" +
                            densukeNameNew +
                            "' on Densuke. Please update it to a unique name and register again.",
                    };
                }
            } else {
                gasUtil.registerMapping(lineName, densukeNameNew, userId);
                gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
                if (lang === 'ja') {
                    getEventHandler.result = {
                        result:
                            '伝助名称登録が完了しました。\n伝助上の名前：' +
                            densukeNameNew +
                            '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。',
                    };
                } else {
                    getEventHandler.result = {
                        result:
                            'The initial registration is complete.\nYour name in Densuke: ' +
                            densukeNameNew +
                            "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.",
                    };
                }
            }
        } else {
            if (lang === 'ja') {
                getEventHandler.result = {
                    result: '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' + densukeNameNew,
                };
            } else {
                getEventHandler.result = {
                    result:
                        '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
                        densukeNameNew,
                };
            }
        }
    }
}
