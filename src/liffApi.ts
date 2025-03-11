// import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { GetEventHandler } from './getEventHandler';
import { LineUtil } from './lineUtil';
import { SchedulerUtil } from './schedulerUtil';
import { ScoreBook } from './scoreBook';
import { ScriptProps } from './scriptProps';

export class LiffApi {
    private test(getEventHandler: GetEventHandler): void {
        const value: string = getEventHandler.e.parameters['param'][0];
        getEventHandler.result = { result: value };
    }

    private getAttendance(getEventHandler: GetEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const attendance: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('attendance');
        if (!attendance) {
            throw new Error('attendance sheet was not found.');
        }
        getEventHandler.result.attendance = attendance.getDataRange().getValues();
    }

    private loadCalendar(getEventHandler: GetEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const calendar: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('calendar');
        if (!calendar) {
            throw new Error('calendar sheet was not found.');
        }
        getEventHandler.result.calendar = calendar.getDataRange().getValues();
    }

    private getWinningTeam(getEventHandler: GetEventHandler): void {
        // console.log('getWinningTeam');

        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        // const den: DensukeUtil = new DensukeUtil();
        const su: SchedulerUtil = new SchedulerUtil();

        // const chee = den.getDensukeCheerio();
        const actDate = su.extractDateFromRownum();
        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            throw Error(this.getLogSheetName(actDate) + 'が存在しません！');
        }
        const matchId: string = getEventHandler.e.parameter['matchId']; // matchId をパラメータから取得
        const shootLogVals = shootLog.getDataRange().getValues();
        console.log('matchId:' + matchId);
        const teamGoals: { [teamName: string]: number } = {}; // チームごとの得点を集計するオブジェクト

        // shootLogVals をループして matchId が一致する行のチームごとの得点を集計 (1行目はヘッダー行と仮定)
        for (let i = 1; i < shootLogVals.length; i++) {
            const row = shootLogVals[i];
            const currentRowMatchId = row[1]; // 2列目 (B列) : 試合
            if (currentRowMatchId === matchId) {
                // matchId が一致する行のみ処理
                const teamName = row[2]; // 3列目 (C列) : チーム
                if (teamName) {
                    teamGoals[teamName] = (teamGoals[teamName] || 0) + 1; // チームの得点数をカウント
                }
            }
        }
        // console.log(teamGoals);
        let winningTeam: string = 'draw';
        let maxGoals = -1;
        let teamsWithMaxGoals: string[] = []; // 最大得点のチームを格納する配列

        for (const team in teamGoals) {
            if (teamGoals[team] > maxGoals) {
                maxGoals = teamGoals[team];
                winningTeam = team;
                teamsWithMaxGoals = [team]; // 新しい最大得点チームが見つかったので配列を更新
            } else if (teamGoals[team] === maxGoals) {
                teamsWithMaxGoals.push(team); // 最大得点と同点のチームを追加
            }
        }

        if (teamsWithMaxGoals.length > 1) {
            winningTeam = 'draw'; // 最大得点のチームが複数存在する場合は引分け
        }
        // console.log('win:' + winningTeam);
        // 勝者チーム名を responseObj に設定 (勝者がいない場合は null が設定される)
        getEventHandler.result.winningTeam = winningTeam;
    }

    private getVideo(getEventHandler: GetEventHandler): void {
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        getEventHandler.result = { result: videos.getDataRange().getValues() };
    }

    private getVideos(getEventHandler: GetEventHandler): void {
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        getEventHandler.result.videos = videos.getDataRange().getValues();
    }

    private getEventData(getEventHandler: GetEventHandler): void {
        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
        getEventHandler.result.events = eventDetail.getDataRange().getValues();
    }

    private getInfoOfTheDay(getEventHandler: GetEventHandler): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        let actDate: string = getEventHandler.e.parameter['actDate'];
        console.log('info', actDate);
        //ない場合は今のやつ その場合全体のVideoリスト・日付のリストも含める
        if (!actDate) {
            // const den: DensukeUtil = new DensukeUtil();
            // const chee = den.getDensukeCheerio();
            const videoVals = videos.getDataRange().getValues();
            console.log('videoVals', videoVals);
            if (videoVals.length > 1) {
                actDate = videoVals[videoVals.length - 1][0]; // videosシートの最終行の１列目の値を取得
            }
            console.log(actDate);
            getEventHandler.result.videos = videoVals;
            getEventHandler.result.actDates = [
                ...new Set(
                    videos
                        .getDataRange()
                        .getValues()
                        .map(val => {
                            if (typeof val[0] === 'string') {
                                return val[0]; // Stringの場合はそのまま返す
                            } else if (val[0] instanceof Date) {
                                return Utilities.formatDate(val[0], 'Asia/Singapore', 'yyyy/MM/dd'); // Date型の場合はシンガポール時刻でフォーマット
                            } else {
                                return ''; // その他の型の場合は空文字を返す (必要に応じて変更)
                            }
                        })
                        .reverse()
                ),
            ];
            const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
            getEventHandler.result.events = eventDetail.getDataRange().getValues();
        }

        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (shootLog) {
            getEventHandler.result.shootLogs = shootLog
                .getDataRange()
                .getValues()
                .slice(1)
                .filter(val => val[1].startsWith(actDate));
        }
    }

    private getTodayMatch(getEventHandler: GetEventHandler): void {
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        // const den: DensukeUtil = new DensukeUtil();
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();

        getEventHandler.result.match = videos
            .getDataRange()
            .getValues()
            .filter(val => val[0] === actDate && !val[10].endsWith('_g') && val[3] && val[4]);
    }

    private getPayNow(getEventHandler: GetEventHandler): void {
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        // getEventHandler.result = { result: members };
        getEventHandler.result.payNow = addy;
    }

    // private getMembers(getEventHandler: GetEventHandler): void {
    //     const den: DensukeUtil = new DensukeUtil();
    //     const members = den.extractMembers();
    //     // getEventHandler.result = { result: members };
    //     getEventHandler.result.members = members;
    // }

    //Densukeではなくてスプシから取ってくる
    private getTeams(getEventHandler: GetEventHandler): void {
        // const den: DensukeUtil = new DensukeUtil();
        const su: SchedulerUtil = new SchedulerUtil();

        const scoreBook: ScoreBook = new ScoreBook();
        // const chee = den.getDensukeCheerio();
        const actDate = su.extractDateFromRownum();
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);

        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = scoreBook.getEventDetailSheet(eventSS, actDate);
        // console.log('resultInput: ' + actDate);
        const values = eventDetail.getDataRange().getValues();
        getEventHandler.result.teams = values;

        const count = this.getMatchType(actDate);
        getEventHandler.result.matchCount = count;
    }

    private getMatchType(actDate: string) {
        const videoSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        const videoVals = videoSheet.getDataRange().getValues();
        let count = 0;
        for (let i = videoVals.length - 1; i >= 0; i--) {
            // Start from the last row and go backwards
            const val = videoVals[i];
            if (val[0] === actDate) {
                // Check if the first column matches actDate
                if (typeof val[10] !== 'string' || !val[10].endsWith('_g')) {
                    // Check the second condition
                    count++;
                }
            } else {
                if (count > 0) {
                    break; // If the first column does not match actDate, break the loop
                }
            }
        }
        return this.convertMatchCount(count);
    }

    private convertMatchCount(c: number): string {
        let result = '3';
        switch (c) {
            case 3: //3チームの場合
                result = '3';
                break;
            case 4:
                result = '4';
                break;
            case 10:
                result = '5';
                break;
        }
        return result;
    }

    private getLogSheetName(actDate: string) {
        return actDate + '_s';
    }

    private getScores(getEventHandler: GetEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);

        let shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            shootLog = eventSS.insertSheet(this.getLogSheetName(actDate));
            shootLog.activate();
            eventSS.moveActiveSheet(0);
            // shootLog.insertRows(shootLog.getDataRange().getLastRow(), 1);
            shootLog.getRange(1, 1).setValue('No');
            shootLog.getRange(1, 2).setValue('試合');
            shootLog.getRange(1, 3).setValue('チーム');
            shootLog.getRange(1, 4).setValue('アシスト');
            shootLog.getRange(1, 5).setValue('ゴール');
        }

        getEventHandler.result.scores = shootLog.getDataRange().getValues();
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
            // const settingSheet = GasProps.instance.settingSheet;
            // const addy = settingSheet.getRange('B2').getValue();
            const addy = sheet.getRange('B4').getValue();
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

    private getComments(getEventHandler: GetEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const comments: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('comments');
        if (!comments) {
            throw new Error('comments Sheet was not found.');
        }
        const componentId: string = getEventHandler.e.parameters['component_id'][0];
        const category: string = getEventHandler.e.parameters['category'][0];
        getEventHandler.result.comments = comments
            .getDataRange()
            .getValues()
            .filter(data => data[1] === componentId && data[2] === category)
            .reverse()
            .slice(0, 100);
    }

    private generateExReport(getEventHandler: GetEventHandler): void {
        const users: string[] = getEventHandler.e.parameters['users'];
        const price: string = getEventHandler.e.parameters['price'][0];
        const title: string = getEventHandler.e.parameters['title'][0];
        const payNow: string = getEventHandler.e.parameters['payNow'][0];
        const receiveColumn: string = getEventHandler.e.parameters['receiveColumn'][0];

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
        let statusVal = null;
        if (receiveColumn === 'true') {
            sheet.appendRow(['参加者（伝助名称）', '参加者（Line名称）', '金額', '支払い状況', '受け取り状況']);
            const status: string[] = ['受渡済', ''];
            statusVal = SpreadsheetApp.newDataValidation().requireValueInList(status).build();
        } else {
            sheet.appendRow(['参加者（伝助名称）', '参加者（Line名称）', '金額', '支払い状況']);
        }
        let index = 6;
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const mapVal = mappingSheet.getDataRange().getValues();

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
            if (statusVal) {
                // sheet.getRange(index, 1).setValue(lu.getLineDisplayName());
                sheet.getRange(index, 5).setDataValidation(statusVal);
            }
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

    // private register(getEventHandler: GetEventHandler): void {
    //     const userId = getEventHandler.e.parameters['userId'][0];
    //     const lineUtil: LineUtil = new LineUtil();
    //     const gasUtil: GasUtil = new GasUtil();
    //     const su:SchedulerUtil = new SchedulerUtil();
    //     const lineName = lineUtil.getLineDisplayName(userId);
    //     const lang = lineUtil.getLineLang(userId);
    //     // const $ = densukeUtil.getDensukeCheerio();
    //     su.generateSummaryBase(); //先に更新しておかないとエラーになる（伝助が更新されている場合）
    //     // const members = densukeUtil.extractMembers($);
    //     const actDate = su.extractDateFromRownum();
    //     const densukeNameNew = getEventHandler.e.parameters['densukeName'][0];
    //     if (members.includes(densukeNameNew)) {
    //         if (densukeUtil.hasMultipleOccurrences(members, densukeNameNew)) {
    //             if (lang === 'ja') {
    //                 getEventHandler.result = {
    //                     result: '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。',
    //                 };
    //             } else {
    //                 getEventHandler.result = {
    //                     result:
    //                         "There are multiple entries with the name '" +
    //                         densukeNameNew +
    //                         "' on Densuke. Please update it to a unique name and register again.",
    //                 };
    //             }
    //         } else {
    //             gasUtil.registerMapping(lineName, densukeNameNew, userId);
    //             gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
    //             if (lang === 'ja') {
    //                 getEventHandler.result = {
    //                     result:
    //                         '伝助名称登録が完了しました。\n伝助上の名前：' +
    //                         densukeNameNew +
    //                         '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。',
    //                 };
    //             } else {
    //                 getEventHandler.result = {
    //                     result:
    //                         'The initial registration is complete.\nYour name in Densuke: ' +
    //                         densukeNameNew +
    //                         "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.",
    //                 };
    //             }
    //         }
    //     } else {
    //         if (lang === 'ja') {
    //             getEventHandler.result = {
    //                 result: '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' + densukeNameNew,
    //             };
    //         } else {
    //             getEventHandler.result = {
    //                 result:
    //                     '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
    //                     densukeNameNew,
    //             };
    //         }
    //     }
    // }
}
