// import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasTestSuite } from './gasTestSuite';
import { GasUtil } from './gasUtil';
import { LineUtil } from './lineUtil';
import { PostEventHandler } from './postEventHandler';
import { SchedulerUtil } from './schedulerUtil';
import { ScoreBook, Title } from './scoreBook';
import { ScriptProps } from './scriptProps';

// const densukeUtil: DensukeUtil = new DensukeUtil();
const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();

export class RequestExecuter {
    public createCalendar(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = su.calendarSheet;
        const id: string = Utilities.getUuid();
        const eventType: string = postEventHander.parameter['event_type'];
        const eventName: string = postEventHander.parameter['event_name'];
        const sDate: string = postEventHander.parameter['start_datetime'];
        const eDate: string = postEventHander.parameter['end_datetime'];
        const place: string = postEventHander.parameter['place'];
        const remark: string = postEventHander.parameter['remark'];
        const recursiveType: number = 0;
        const headerRow = calendarSheet.getDataRange().getValues()[0]; // ヘッダー行を取得
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const newRowData: any[] = [];
        headerRow.forEach(header => {
            switch (header) {
                case 'ID':
                    newRowData.push(id);
                    break;
                case 'event_type':
                    newRowData.push(eventType);
                    break;
                case 'event_name':
                    newRowData.push(eventName);
                    break;
                case 'start_datetime':
                    newRowData.push(sDate);
                    break;
                case 'end_datetime':
                    newRowData.push(eDate);
                    break;
                case 'place':
                    newRowData.push(place);
                    break;
                case 'remark':
                    newRowData.push(remark);
                    break;
                case 'event_status':
                    newRowData.push(recursiveType);
                    break;
                // ID は自動で振られる想定 or スプレッドシート側で設定
                default:
                    newRowData.push(''); // その他のヘッダーの場合は空文字をセット
            }
        });
        calendarSheet.appendRow(newRowData);
    }

    public registrationFromApp(postEventHander: PostEventHandler): void {
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const userId: string = postEventHander.parameter['userId'];
        const nickname: string = postEventHander.parameter['nickname'];
        const lineName: string = postEventHander.parameter['line_name'];
        const picUrl: string = postEventHander.parameter['pic_url'];
        const headerRow = mappingSheet.getDataRange().getValues()[0]; // ヘッダー行を取得
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const newRowData: any[] = [];
        headerRow.forEach(header => {
            switch (header) {
                case 'LINE ID':
                    newRowData.push(userId);
                    break;
                case '伝助上の名前':
                    newRowData.push(nickname);
                    break;
                case 'ライン上の名前':
                    newRowData.push(lineName);
                    break;
                case 'Picture':
                    newRowData.push(picUrl);
                    break;
                case '幹事フラグ': // 幹事フラグはパラメータにないため空文字
                    newRowData.push('');
                    break;
                default:
                    newRowData.push(''); // その他のヘッダーの場合は空文字をセット
            }
        });
        mappingSheet.appendRow(newRowData);
    }

    public updateEventStatus(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = su.calendarSheet;
        const values = calendarSheet.getDataRange().getValues();
        const id: string = postEventHander.parameter['id'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const eventType: string = postEventHander.parameter['new_status'];
        let rowNumberToUpdate: number | null = null;
        // データの行をループして 'id' に一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = 1; i < values.length; i++) {
            if (values[i][0].toString() === id.toString()) {
                rowNumberToUpdate = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // 'id' が見つかったのでループを抜ける
            }
        }
        if (rowNumberToUpdate) {
            calendarSheet.getRange(rowNumberToUpdate, 8).setValue(eventType);
        } else {
            console.error(`No row found with id: ${id}.`);
            throw new Error(`No row found with id: ${id}.`); // ID が見つからない場合はエラーをthrow
        }
    }

    public updateCalendar(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = su.calendarSheet;
        // id パラメータから更新対象のIDを取得
        const id: string = postEventHander.parameter['id'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const eventType: string = postEventHander.parameter['event_type'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const eventName: string = postEventHander.parameter['event_name'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const sDate: string = postEventHander.parameter['start_datetime'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const eDate: string = postEventHander.parameter['end_datetime'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const place: string = postEventHander.parameter['place'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const remark: string = postEventHander.parameter['remark'];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const recursiveType: number = postEventHander.parameter['event_status']; // default value
        const values = calendarSheet.getDataRange().getValues();
        const headerRow = values[0]; // ヘッダー行を取得

        let rowNumberToUpdate: number | null = null;
        // データの行をループして 'id' に一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = 1; i < values.length; i++) {
            if (values[i][0].toString() === id.toString()) {
                rowNumberToUpdate = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // 'id' が見つかったのでループを抜ける
            }
        }

        if (rowNumberToUpdate) {
            // 'id' に一致する行が見つかった場合、データを更新
            console.log(`id: ${id} の行を更新`);
            const row = rowNumberToUpdate;
            // 各パラメータを該当の列に更新 (列位置はheaderRowからcolumnIndexを検索して特定)
            ['event_type', 'event_name', 'start_datetime', 'end_datetime', 'place', 'remark', 'event_status'].forEach(paramName => {
                if (postEventHander.parameter[paramName]) {
                    const colIndex = headerRow.indexOf(paramName); // ヘッダー行から列番号を取得
                    if (colIndex > -1) {
                        calendarSheet.getRange(row, colIndex + 1).setValue(postEventHander.parameter[paramName]);
                    }
                }
            });
        } else {
            console.error(`No row found with id: ${id}.`);
            throw new Error(`No row found with id: ${id}.`); // ID が見つからない場合はエラーをthrow
        }
    }

    public deleteCalendar(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = su.calendarSheet;

        // id パラメータから削除対象のIDを取得
        const id: string = postEventHander.parameter['id'];
        const values = calendarSheet.getDataRange().getValues();

        let rowNumberToDelete: number | null = null;
        // データの行をループして 'id' に一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = 1; i < values.length; i++) {
            if (values[i][0].toString() === id.toString()) {
                rowNumberToDelete = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // 'id' が見つかったのでループを抜ける
            }
        }
        if (rowNumberToDelete) {
            // 'id' に一致する行が見つかった場合、行を削除
            console.log(`id: ${id} の行を削除`);
            calendarSheet.deleteRow(rowNumberToDelete);
        } else {
            console.error(`No row found with id: ${id}.`);
            throw new Error(`No row found with id: ${id}.`); // ID が見つからない場合はエラーをthrow
        }
    }

    public updateParticipation(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const attendanceSheet = su.attendanceSheet;
        const attendanceValues = attendanceSheet.getDataRange().getValues(); // 出席シートのデータを取得
        const headerRow = attendanceValues[0]; // ヘッダー行を保持
        const param = postEventHander.parameter;
        const updates: Record<string, Record<string, string>> = {};

        // パラメータを処理して updates オブジェクトに整理
        for (const key in param) {
            const lastUnderscoreIndex = key.lastIndexOf('_');
            if (lastUnderscoreIndex !== -1) {
                const paramName = key.substring(0, lastUnderscoreIndex);
                const index = key.substring(lastUnderscoreIndex + 1);
                if (!updates[index]) {
                    updates[index] = {};
                }
                updates[index][paramName] = param[key];
            }
        }

        console.log(updates);
        for (const index in updates) {
            const updateData = updates[index];
            let rowNumberToUpdate: number | null = null;
            if (updateData['attendance_id']) {
                // attendance_id が存在する場合、更新
                const attendanceId = updateData['attendance_id'];
                for (let i = 1; i < attendanceValues.length; i++) {
                    // ヘッダー行を skip
                    if (attendanceValues[i][0] === attendanceId) {
                        // 0列目が attendance_id 列と仮定
                        rowNumberToUpdate = i + 1;
                        break;
                    }
                }
            }

            if (rowNumberToUpdate) {
                // 既存の行を更新
                console.log(`attendance_id: ${updateData['attendance_id']} の行を更新`);
                const row = rowNumberToUpdate;
                // 各パラメータを該当の列に更新 (列位置はheaderRowからcolumnIndexを検索して特定)
                ['user_id', 'year', 'month', 'date', 'status', 'calendar_id'].forEach(paramName => {
                    if (updateData[paramName]) {
                        const colIndex = headerRow.indexOf(paramName); // ヘッダー行から列番号を取得
                        if (colIndex > -1) {
                            attendanceSheet.getRange(row, colIndex + 1).setValue(updateData[paramName]);
                        }
                    }
                });
            } else {
                // 新規追加
                console.log('新規行を追加');
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const newRowData: any[] = [];
                // ヘッダー行に基づいて新しい行データを作成
                headerRow.forEach(header => {
                    if (header === 'attendance_id') {
                        newRowData.push(Utilities.getUuid()); // attendance_id がない場合は新規にUUIDを生成
                    } else if (updateData[header]) {
                        newRowData.push(updateData[header]);
                    } else {
                        newRowData.push(''); // データがない場合は空文字
                    }
                });
                console.log(newRowData);
                attendanceSheet.appendRow(newRowData);
            }
        }
    }

    //毎回全部集計してアシストと得点を入れなおす
    public closeGame(postEventHander: PostEventHandler): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            throw Error(this.getLogSheetName(actDate) + 'が存在しません！');
        }
        const shootLogVals = shootLog.getDataRange().getValues();
        const matchId: string = postEventHander.parameter['matchId'];
        const winner: string = postEventHander.parameter['winningTeam'];
        const scoreBook: ScoreBook = new ScoreBook();
        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = scoreBook.getEventDetailSheet(eventSS, actDate);
        const eventDetailVals = eventDetail.getDataRange().getValues();

        // プレイヤーごとの得点とアシストを集計するオブジェクト
        const playerStats: { [playerName: string]: { goals: number; assists: number } } = {};
        // shootLogVals をループして得点とアシストを集計
        for (let i = 1; i < shootLogVals.length; i++) {
            const row = shootLogVals[i];
            const scorer = row[4]; // ゴール (D列)
            const assister = row[3]; // アシスト (E列)

            // 得点者の集計
            if (scorer) {
                playerStats[scorer] = playerStats[scorer] || { goals: 0, assists: 0 };
                playerStats[scorer].goals++;
            }
            // アシスト者の集計
            if (assister) {
                playerStats[assister] = playerStats[assister] || { goals: 0, assists: 0 };
                playerStats[assister].assists++;
            }
        }

        for (let i = 1; i < eventDetailVals.length; i++) {
            const row = eventDetailVals[i];
            const playerName = row[0]; // 名前 (A列)
            // console.log(playerName);
            if (playerName in playerStats) {
                const stats = playerStats[playerName];
                // 得点を書き込み (0点の場合は空文字にする)
                eventDetail.getRange(i + 1, 3).setValue(stats.goals > 0 ? stats.goals : ''); // 3列目 (C列) : 得点
                // アシストを書き込み (0アシストの場合は空文字にする)
                eventDetail.getRange(i + 1, 4).setValue(stats.assists > 0 ? stats.assists : ''); // 4列目 (D列) : アシスト
            } else {
                // playerStats にデータがないプレイヤーは得点、アシストをクリア (念のため)
                eventDetail.getRange(i + 1, 3).clearContent();
                eventDetail.getRange(i + 1, 4).clearContent();
            }
        }

        // videoSheet の更新処理
        const videoSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        const videoSheetVals = videoSheet.getDataRange().getValues();
        let targetRow: number | null = null;

        // videoSheetVals をループして matchId が一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = videoSheetVals.length - 1; i >= 1; i--) {
            if (videoSheetVals[i][10] === matchId) {
                // 11列目 (K列) が matchId
                targetRow = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // matchId が見つかったのでループを抜ける
            }
        }

        if (targetRow) {
            // matchId に一致する行が見つかった場合、データを更新
            const team1Name: string = videoSheetVals[targetRow - 1][3]; // 4列目 (D列) : チーム1名
            const team2Name: string = videoSheetVals[targetRow - 1][4]; // 5列目 (E列) : チーム2名
            let team1Score: number = 0;
            let team2Score: number = 0;

            // 該当の matchId に基づいて shootLogVals をループ処理
            for (let i = 1; i < shootLogVals.length; i++) {
                // 1行目はヘッダー行をスキップ
                const log = shootLogVals[i];
                const logMatchId = log[1]; // 2列目の値を取得

                if (logMatchId === matchId) {
                    const logTeamName = log[2]; // ゴールを決めたプレイヤー名 (D列)
                    if (team1Name === logTeamName) {
                        team1Score++;
                    } else if (team2Name === logTeamName) {
                        team2Score++;
                    }
                }
            }

            videoSheet.getRange(targetRow, 8).setValue(team1Score); // 8列目 (H列) : チーム1得点
            videoSheet.getRange(targetRow, 9).setValue(team2Score); // 9列目 (I列) : チーム2得点
            videoSheet.getRange(targetRow, 10).setValue(winner); // 10列目 (J列) : 勝者

            const lastHyphenIndex = matchId.lastIndexOf('-');
            let matchType = null;
            if (lastHyphenIndex !== -1) {
                matchType = matchId.substring(lastHyphenIndex + 1);
            }

            console.log('matchType', matchType);
            if (matchType?.startsWith('4_1') || matchType?.startsWith('4_2')) {
                //今のところ４人の場合のみトーナメント
                let flg1 = false;
                let flg2 = false;
                for (let i = videoSheetVals.length - 1; i >= 1; i--) {
                    if (videoSheetVals[i][0] === actDate && videoSheetVals[i][1] === '３位決定戦') {
                        console.log('3rd ', videoSheetVals[i][10]);
                        flg1 = true;
                        const looser = winner === team1Name ? team2Name : team1Name; // ３位決定戦は勝者じゃない方のチームをセット
                        const lostMembers = winner === team1Name ? videoSheetVals[targetRow - 1][6] : videoSheetVals[targetRow - 1][5];
                        if (!videoSheetVals[i][3]) {
                            videoSheet.getRange(i + 1, 4).setValue(looser);
                            videoSheet.getRange(i + 1, 6).setValue(lostMembers);
                        } else if (!videoSheetVals[i][4]) {
                            videoSheet.getRange(i + 1, 5).setValue(looser);
                            videoSheet.getRange(i + 1, 7).setValue(lostMembers);
                        }
                    } else if (videoSheetVals[i][0] === actDate && videoSheetVals[i][1] === '決勝') {
                        console.log('1st ', videoSheetVals[i][10]);
                        flg2 = true;
                        const winMembers = winner === team1Name ? videoSheetVals[targetRow - 1][5] : videoSheetVals[targetRow - 1][6];
                        if (!videoSheetVals[i][3]) {
                            videoSheet.getRange(i + 1, 4).setValue(winner);
                            videoSheet.getRange(i + 1, 6).setValue(winMembers);
                        } else if (!videoSheetVals[i][4]) {
                            videoSheet.getRange(i + 1, 5).setValue(winner);
                            videoSheet.getRange(i + 1, 7).setValue(winMembers);
                        }
                    }
                    if (flg1 && flg2) {
                        break;
                    }
                }
            }
        } else {
            console.warn(`No row found in videoSheet with matchId: ${matchId}.`);
        }
        postEventHander.reponseObj = { success: true };
    }

    private matchTeamName(playerTeam: string, teamName: string): boolean {
        let result = false;
        switch (playerTeam) {
            case 'チーム1':
                result = 'Team1' === teamName;
                break;
            case 'チーム2':
                result = 'Team2' === teamName;
                break;
            case 'チーム3':
                result = 'Team3' === teamName;
                break;
            case 'チーム4':
                result = 'Team4' === teamName;
                break;
            case 'チーム5':
                result = 'Team5' === teamName;
                break;
            case 'チーム6':
                result = 'Team6' === teamName;
                break;
            case 'チーム7':
                result = 'Team7' === teamName;
                break;
            case 'チーム8':
                result = 'Team8' === teamName;
                break;
            case 'チーム9':
                result = 'Team9' === teamName;
                break;
            case 'チーム10':
                result = 'Team10' === teamName;
                break;
        }
        return result;
    }

    public recordGoal(postEventHander: PostEventHandler): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            throw Error(this.getLogSheetName(actDate) + 'が存在しません！');
        }
        let no: number = shootLog.getRange(shootLog.getLastRow(), 1).getValue();
        if (!Number.isInteger(no)) {
            no = 0;
        }

        // const lastRow = shootLog.getLastRow();
        const matchId: string = postEventHander.parameter['matchId'];
        const team: string = postEventHander.parameter['team'];
        const scorer: string = postEventHander.parameter['scorer'];
        const assister: string | null = postEventHander.parameter['assister'];

        shootLog.appendRow([no + 1, matchId, team, assister ? assister : '', scorer]);

        postEventHander.reponseObj = { success: true };
    }

    public updateGoal(postEventHander: PostEventHandler): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        // const den: DensukeUtil = new DensukeUtil();
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            throw Error(this.getLogSheetName(actDate) + 'が存在しません！');
        }
        const no: string = postEventHander.parameter['scoreId'];
        const matchId: string = postEventHander.parameter['matchId'];
        const team: string = postEventHander.parameter['team'];
        const scorer: string = postEventHander.parameter['scorer'];
        const assister: string | null = postEventHander.parameter['assister'];
        const values: string[][] = shootLog.getDataRange().getValues();

        let rowNumberToUpdate: number | null = null;
        // データの行をループして 'no' に一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = 1; i < values.length; i++) {
            if (values[i][0].toString() === no.toString()) {
                rowNumberToUpdate = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // 'no' が見つかったのでループを抜ける
            }
        }

        if (rowNumberToUpdate) {
            // 'no' に一致する行が見つかった場合、データを更新
            shootLog.getRange(rowNumberToUpdate, 2).setValue(matchId); // 2列目 (B列) : 試合
            shootLog.getRange(rowNumberToUpdate, 3).setValue(team); // 3列目 (C列) : チーム
            shootLog.getRange(rowNumberToUpdate, 5).setValue(scorer); // 4列目 (D列) : ゴール
            shootLog.getRange(rowNumberToUpdate, 4).setValue(assister ? assister : ''); // 5列目 (E列) : アシスト
        } else {
            console.error(`No row found with No: ${no}. Appending new row instead.`);
            shootLog.appendRow([no, matchId, team, scorer, assister ? assister : '']); // No はそのまま parameter の no を使用
        }
        postEventHander.reponseObj = { success: true };
    }

    public deleteGoal(postEventHander: PostEventHandler): void {
        console.log('exec deleteGoal method');
        console.log(postEventHander.parameter);
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const shootLog: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(this.getLogSheetName(actDate));
        if (!shootLog) {
            throw Error(this.getLogSheetName(actDate) + 'が存在しません！');
        }
        const no: string = postEventHander.parameter['scoreId'];
        const values: string[][] = shootLog.getDataRange().getValues();

        let rowNumberToUpdate: number | null = null;
        // データの行をループして 'no' に一致する行を探す (1行目はヘッダー行と仮定)
        for (let i = 1; i < values.length; i++) {
            if (values[i][0].toString() === no.toString()) {
                rowNumberToUpdate = i + 1; // スプレッドシートの行番号は1から始まるので +1
                break; // 'no' が見つかったのでループを抜ける
            }
        }

        if (rowNumberToUpdate) {
            shootLog.deleteRow(rowNumberToUpdate);
        } else {
            console.error(`No row found with No: ${no}. `);
        }
        postEventHander.reponseObj = { success: true };
    }

    public updateTeams(postEventHander: PostEventHandler): void {
        console.log(postEventHander.parameter);
        const su: SchedulerUtil = new SchedulerUtil();
        const scoreBook: ScoreBook = new ScoreBook();
        const actDate = su.extractDateFromRownum();
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);

        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = scoreBook.getEventDetailSheet(eventSS, actDate);
        // console.log('resultInput: ' + actDate);
        const values = eventDetail.getDataRange().getValues();
        const param: { [key: string]: string } = postEventHander.parameter;
        // const headerRow = values[0]; // ヘッダー行を取得
        for (const k in param) {
            if (k === 'func') {
                continue;
            }
            // valuesの1列目(columnIndex=0)にkと同じ名前があるか検索
            for (let i = 0; i < values.length; i++) {
                if (values[i][0] === k) {
                    // 同じ名前が見つかった場合、該当する行の2列目(columnIndex=1)にparam[k]を入力
                    if (param[k] === '0') {
                        eventDetail.getRange(i + 1, 2).clearContent(); // param[k]が'0'の場合はclearContent()を実行
                    } else {
                        console.log('key: ' + k + ' value: ' + this.convertVal(param[k]));

                        eventDetail.getRange(i + 1, 2).setValue(this.convertVal(param[k])); // それ以外の場合はsetValue()を実行
                    }
                    break; // 同じ名前が見つかったら、それ以降の行は検索しない
                }
            }
        }

        // videoSheet のチームメンバー更新処理
        const videoSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        const videoSheetVals = videoSheet.getDataRange().getValues();
        const eventDetailVals = eventDetail.getDataRange().getValues(); // eventDetail の値を取得

        for (let i = 1; i < videoSheetVals.length; i++) {
            // 1行目ヘッダー行をskip
            if (!videoSheetVals[i][9] && videoSheetVals[i][0] === actDate) {
                // 10列目(J列, index 9)が空の場合のみ処理
                const team1Name = videoSheetVals[i][3]; // 4列目(D列, index 3) : チーム1名
                const team2Name = videoSheetVals[i][4]; // 5列目(E列, index 4) : チーム2名

                // チーム1のメンバーを eventDetailVals から取得
                const team1Members = eventDetailVals
                    .slice(1) // ヘッダー行をskip
                    .filter(row => this.matchTeamName(row[1], team1Name)) // チーム名でfilter
                    .map(row => row[0]) // プレイヤー名(A列)のみ抽出
                    .join(', '); // カンマ区切りでjoin

                // チーム2のメンバーを eventDetailVals から取得
                const team2Members = eventDetailVals
                    .slice(1) // ヘッダー行をskip
                    .filter(row => this.matchTeamName(row[1], team2Name)) // チーム名でfilter
                    .map(row => row[0]) // プレイヤー名(A列)のみ抽出
                    .join(', '); // カンマ区切りでjoin

                // videoSheet の該当行にチームメンバーを書き込み
                videoSheet.getRange(i + 1, 6).setValue(team1Members); // 6列目(F列, index 5) : チーム1メンバー
                videoSheet.getRange(i + 1, 7).setValue(team2Members); // 7列目(G列, index 6) : チーム2メンバー
            }
        }

        // this.createShootLog(actDate, eventDetail.getDataRange().getValues());
        postEventHander.reponseObj = { success: true };
    }

    private getLogSheetName(actDate: string) {
        return actDate + '_s';
    }

    // private getTeamCount(memberCount: number) {
    //     if (memberCount < 18) {
    //         return 3;
    //     } else if (memberCount < 19) {
    //         return 4;
    //     } else {
    //         return 5;
    //     }
    // }

    public createShootLog(postEventHander: PostEventHandler) {
        const teamCount: string = postEventHander.parameter['teamCount'];
        const su: SchedulerUtil = new SchedulerUtil();
        const scoreBook: ScoreBook = new ScoreBook();
        const actDate = su.extractDateFromRownum();
        console.log('ac', actDate);
        const activitySS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const eventDetails = scoreBook.getEventDetailSheet(eventSS, actDate).getDataRange().getValues();

        const videoSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        videoSheet.activate();
        activitySS.moveActiveSheet(0);
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
        } else {
            const records = videoSheet.getDataRange().getValues();
            for (let i = records.length - 1; i >= 0; i--) {
                if (records[i][0] === actDate && !records[i][9]) {
                    videoSheet.deleteRows(i + 1, 1); // Delete the row (i+1 because spreadsheet rows are 1-indexed)
                } else {
                    continue;
                    // break;
                }
            }
        }
        const lastRow = videoSheet.getLastRow();
        console.log('last Row Value' + videoSheet.getRange(lastRow, 1).getValue());
        console.log('team count', teamCount);
        // if (videoSheet.getRange(lastRow + 1, 1).getValue() === actDate) {
        //     return;
        // }
        // const teamCount: number = this.getTeamCount(eventDetails.length - 1);
        //アップのひな形を作る（ついでにここを見る）
        switch (teamCount) {
            case '3': //3チームの場合
                videoSheet.insertRows(lastRow + 1, 4);
                this.addRow(videoSheet, lastRow + 1, actDate, eventDetails, '#1 Team1 vs Team2', 'Team1', 'Team2', '-3_1');
                this.addRow(videoSheet, lastRow + 3, actDate, eventDetails, '#3 Team1 vs Team3', 'Team1', 'Team3', '-3_2');
                this.addRow(videoSheet, lastRow + 2, actDate, eventDetails, '#2 Team2 vs Team3', 'Team2', 'Team3', '-3_3');
                this.addRow(videoSheet, lastRow + 4, actDate, eventDetails, 'ゴール集', '', '', '-3_g');
                break;
            case '4': //4チームの場合
                videoSheet.insertRows(lastRow + 1, 5);
                this.addRow(videoSheet, lastRow + 1, actDate, eventDetails, 'Team1 vs Team2', 'Team1', 'Team2', '-4_1');
                this.addRow(videoSheet, lastRow + 2, actDate, eventDetails, 'Team3 vs Team4', 'Team3', 'Team4', '-4_2');
                this.addRow(videoSheet, lastRow + 3, actDate, eventDetails, '３位決定戦', '', '', '-4_3');
                this.addRow(videoSheet, lastRow + 4, actDate, eventDetails, '決勝', '', '', '-4_4');
                this.addRow(videoSheet, lastRow + 5, actDate, eventDetails, 'ゴール集', '', '', '-4_g');
                break;
            case '5': //5チームの場合(2ピッチ前提)
                videoSheet.insertRows(lastRow + 1, 11);
                this.addRow(videoSheet, lastRow + 1, actDate, eventDetails, '#1 Pitch1 Team1 vs Team2', 'Team1', 'Team2', '-5_1_1');
                this.addRow(videoSheet, lastRow + 2, actDate, eventDetails, '#1 Pitch2 Team3 vs Team4', 'Team3', 'Team4', '-5_1_2');
                this.addRow(videoSheet, lastRow + 3, actDate, eventDetails, '#2 Pitch1 Team1 vs Team3', 'Team1', 'Team3', '-5_2_1');
                this.addRow(videoSheet, lastRow + 4, actDate, eventDetails, '#2 Pitch2 Team2 vs Team5', 'Team2', 'Team5', '-5_2_2');
                this.addRow(videoSheet, lastRow + 5, actDate, eventDetails, '#3 Pitch2 Team2 vs Team4', 'Team2', 'Team4', '-5_3_1');
                this.addRow(videoSheet, lastRow + 6, actDate, eventDetails, '#3 Pitch2 Team1 vs Team5', 'Team1', 'Team5', '-5_3_2');
                this.addRow(videoSheet, lastRow + 7, actDate, eventDetails, '#4 Pitch2 Team3 vs Team5', 'Team3', 'Team5', '-5_4_1');
                this.addRow(videoSheet, lastRow + 8, actDate, eventDetails, '#4 Pitch2 Team1 vs Team4', 'Team1', 'Team4', '-5_4_2');
                this.addRow(videoSheet, lastRow + 9, actDate, eventDetails, '#5 Pitch2 Team4 vs Team5', 'Team4', 'Team5', '-5_5_1');
                this.addRow(videoSheet, lastRow + 10, actDate, eventDetails, '#5 Pitch2 Team2 vs Team3', 'Team2', 'Team3', '-5_5_2');
                this.addRow(videoSheet, lastRow + 11, actDate, eventDetails, 'ゴール集 pitch1', '', '', '-5_1_g');
                this.addRow(videoSheet, lastRow + 12, actDate, eventDetails, 'ゴール集 pitch2', '', '', '-5_2_g');
                break;
        }
    }

    private addRow(
        videoSheet: GoogleAppsScript.Spreadsheet.Sheet,
        row: number,
        actDate: string,
        eventDetails: string[][],
        title: string,
        right: string,
        left: string,
        count: string
    ) {
        videoSheet.getRange(row, 1).setValue(actDate);
        videoSheet.getRange(row, 2).setValue(title);
        videoSheet.getRange(row, 4).setValue(right);
        videoSheet.getRange(row, 5).setValue(left);
        switch (right) {
            case 'Team1':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム1')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team2':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム2')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team3':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム3')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team4':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム4')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team5':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム5')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team6':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム6')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team7':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム7')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team8':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム8')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team9':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム9')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team10':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム10')
                        .map(val => val[0])
                        .join(', ')
                );
                break;

            default:
                break;
        }
        switch (left) {
            case 'Team1':
                videoSheet.getRange(row, 7).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム1')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team2':
                videoSheet.getRange(row, 7).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム2')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team3':
                videoSheet.getRange(row, 7).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム3')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team4':
                videoSheet.getRange(row, 7).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム4')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team5':
                videoSheet.getRange(row, 7).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム5')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team6':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム6')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team7':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム7')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team8':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム8')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team9':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム9')
                        .map(val => val[0])
                        .join(', ')
                );
                break;
            case 'Team10':
                videoSheet.getRange(row, 6).setValue(
                    eventDetails
                        .slice(1)
                        .filter(val => val[1] === 'チーム10')
                        .map(val => val[0])
                        .join(', ')
                );
                break;

            default:
                break;
        }
        videoSheet.getRange(row, 11).setValue(actDate + count);
    }

    private convertVal(val: string): string {
        if (val === '1') {
            return 'チーム1';
        } else if (val === '2') {
            return 'チーム2';
        } else if (val === '3') {
            return 'チーム3';
        } else if (val === '4') {
            return 'チーム4';
        } else if (val === '5') {
            return 'チーム5';
        } else if (val === '6') {
            return 'チーム6';
        } else if (val === '7') {
            return 'チーム7';
        } else if (val === '8') {
            return 'チーム8';
        } else if (val === '9') {
            return 'チーム9';
        } else if (val === '10') {
            return 'チーム10';
        }
        return '';
    }

    public deleteEx(postEventHander: PostEventHandler): void {
        console.log('execute deleteEx');
        const title: string = postEventHander.parameter.title;
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.expenseFolder);
        const titleFolderIt: GoogleAppsScript.Drive.FolderIterator = rootFolder.getFoldersByName(title);
        // const results = [];
        while (titleFolderIt.hasNext()) {
            const expenseFolder: GoogleAppsScript.Drive.Folder = titleFolderIt.next();
            expenseFolder.setTrashed(true);
        }
        postEventHander.reponseObj = { msg: title };
    }

    public loadExList(postEventHander: PostEventHandler): void {
        console.log('execute loadExList');
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.expenseFolder);
        const titleFolderIt: GoogleAppsScript.Drive.FolderIterator = rootFolder.getFolders();
        const results = [];
        while (titleFolderIt.hasNext()) {
            const expenseFolder: GoogleAppsScript.Drive.Folder = titleFolderIt.next();
            const title = expenseFolder.getName();
            const url = expenseFolder.getFilesByName(title).next().getUrl();
            results.push({ title: title, url: url });
        }
        postEventHander.reponseObj = { resultList: results };
    }

    public upload(postEventHander: PostEventHandler): void {
        console.log('execute upload');
        const decodedFile = Utilities.base64Decode(postEventHander.parameter.file);
        const lu: LineUtil = new LineUtil();
        const lineName = lu.getLineDisplayName(postEventHander.parameter.userId);
        const gu: GasUtil = new GasUtil();
        const densukeName = gu.getDensukeName(lineName);
        const title: string = postEventHander.parameter.title;
        const blob = Utilities.newBlob(decodedFile, 'application/octet-stream', title + '_' + lineName);
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.expenseFolder);

        const folderIt = rootFolder.getFoldersByName(title);
        if (!folderIt.hasNext()) {
            console.log('no expense folder found:' + title);
        }
        const expenseFolder = folderIt.next();
        const oldFileIt = expenseFolder.getFilesByName(title + '_' + lineName);
        while (oldFileIt.hasNext()) {
            oldFileIt.next().setTrashed(true);
        }
        const file = expenseFolder.createFile(blob);
        console.log('File uploaded to Google Drive with ID:', file.getId());

        let spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
        const fileIt = expenseFolder.getFilesByName(title);
        if (fileIt.hasNext()) {
            const sheetFile = fileIt.next();
            spreadSheet = SpreadsheetApp.openById(sheetFile.getId());
        } else {
            throw new Error('SpreadSheet is not available:' + title);
        }
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getActiveSheet();
        const sheetVal = sheet.getDataRange().getValues();
        let index = 1;
        const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
        for (const row of sheetVal) {
            if (index > 4) {
                if (row[0] === densukeName) {
                    sheet.getRange(index, 4).setValue(picUrl);
                }
            }
            index++;
        }
        postEventHander.reponseObj = { picUrl: picUrl, sheetUrl: GasProps.instance.generateSheetUrl(spreadSheet.getId()) };
    }

    public video(postEventHander: PostEventHandler): void {
        postEventHander.isFlex = true;
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        const videoValues = videos.getDataRange().getValues();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const flexMsg: any = lineUtil.getCarouselBase();
        postEventHander.messageJson = flexMsg;
        for (let i = videoValues.length - 1; i >= videoValues.length - 10; i--) {
            if (!videoValues[i] || !videoValues[i][2] || videoValues[i][2] === 'URL') {
                break;
            }
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const card: any = lineUtil.getYoutubeCard();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            flexMsg.contents.push(card);
            card.body.contents[0].url = this.getPicUrl(videoValues[i][2]);
            card.body.contents[2].text = videoValues[i][1];
            card.body.contents[3].text = Utilities.formatDate(videoValues[i][0], 'GMT+0800', 'yyyy/MM/dd');
            card.body.action.uri = videoValues[i][2];
            console.log(Utilities.formatDate(videoValues[i][0], 'GMT+0800', 'yyyy/MM/dd'));
        }
    }

    private getPicUrl(url: string): string {
        // https://youtu.be/kNuUeydJZ8I?si=tvBltuqVCilNhnng
        // http://img.youtube.com/vi/kNuUeydJZ8I/maxresdefault.jpg
        const videoIdMatch = url.match(/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/watch\?v=|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
        if (!videoIdMatch) {
            throw new Error('Invalid YouTube URL ' + url);
        }
        const videoId = videoIdMatch[1] || videoIdMatch[2];
        // Construct the thumbnail URL
        const thumbnailUrl = `https://img.youtube.com/vi/${videoId}/maxresdefault.jpg`;
        return thumbnailUrl;
    }

    public intro(postEventHander: PostEventHandler): void {
        postEventHander.resultMessage = ScriptProps.instance.channelUrl;
        postEventHander.resultImage = ScriptProps.instance.channelQr;
    }

    // public register(postEventHander: PostEventHandler): void {
    //     const lineName = lineUtil.getLineDisplayName(postEventHander.userId);
    //     const $ = densukeUtil.getDensukeCheerio();
    //     const members = densukeUtil.extractMembers($);
    //     const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
    //     const densukeNameNew = postEventHander.messageText.split('@@register@@')[1];
    //     if (members.includes(densukeNameNew)) {
    //         if (densukeUtil.hasMultipleOccurrences(members, densukeNameNew)) {
    //             if (postEventHander.lang === 'ja') {
    //                 postEventHander.resultMessage =
    //                     '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。';
    //             } else {
    //                 postEventHander.resultMessage =
    //                     "There are multiple entries with the name '" +
    //                     densukeNameNew +
    //                     "' on Densuke. Please update it to a unique name and register again.";
    //             }
    //         } else {
    //             gasUtil.registerMapping(lineName, densukeNameNew, postEventHander.userId);
    //             gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
    //             this.updateProfilePic();
    //             if (postEventHander.lang === 'ja') {
    //                 postEventHander.resultMessage =
    //                     '伝助名称登録が完了しました。\n伝助上の名前：' +
    //                     densukeNameNew +
    //                     '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。\n' +
    //                     postEventHander.userId;
    //             } else {
    //                 postEventHander.resultMessage =
    //                     'The initial registration is complete.\nYour name in Densuke: ' +
    //                     densukeNameNew +
    //                     "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.\n" +
    //                     postEventHander.userId;
    //             }
    //         }
    //     } else {
    //         if (postEventHander.lang === 'ja') {
    //             postEventHander.resultMessage =
    //                 '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' + densukeNameNew;
    //         } else {
    //             postEventHander.resultMessage =
    //                 '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
    //                 densukeNameNew;
    //         }
    //     }
    // }

    // private updateProfilePic() {
    //     // const lineUtil: LineUtil = new LineUtil();
    //     const densukeMappingVals = GasProps.instance.mappingSheet.getDataRange().getValues();
    //     let index: number = 0;
    //     for (const userRow of densukeMappingVals) {
    //         if (userRow[0] !== 'ライン上の名前') {
    //             const userId: string = userRow[2];
    //             try {
    //                 const prof = lineUtil.getLineUserProfile(userId);
    //                 if (prof) {
    //                     // console.log(userRow[0] + ': ' + prof.pictureUrl);
    //                     GasProps.instance.mappingSheet.getRange(index + 1, 5).setValue(prof.pictureUrl);
    //                 }
    //                 // eslint-disable-next-line @typescript-eslint/no-unused-vars
    //             } catch (e) {
    //                 console.log(userRow[0] + ': invalid UserId' + userId);
    //             }
    //         }
    //         index++;
    //     }
    //     return;
    // }

    public payNow(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();

        const attendees = su.extractAttendees('〇');
        const actDate = su.extractDateFromRownum();
        const messageId = postEventHander.messageId;
        const userId = postEventHander.userId;
        const lineName = lineUtil.getLineDisplayName(userId);
        const densukeName = gasUtil.getDensukeName(lineName);
        console.log(densukeName);
        if (densukeName) {
            if (attendees.includes(densukeName)) {
                gasUtil.uploadPayNowPic(densukeName, messageId, actDate);
                gasUtil.updatePaymentStatus(densukeName, actDate);
                if (postEventHander.lang === 'ja') {
                    postEventHander.resultMessage = actDate + 'の支払いを登録しました。ありがとうございます！\n' + GasProps.instance.reportSheetUrl;
                } else {
                    postEventHander.resultMessage =
                        'Payment for ' + actDate + ' has been registered. Thank you!\n' + GasProps.instance.reportSheetUrl;
                }
            } else {
                if (postEventHander.lang === 'ja') {
                    postEventHander.resultMessage =
                        '【エラー】' +
                        actDate +
                        'のスケジューラーの出席が〇になっていませんでした。スケジューラーを更新して、「伝助更新」と入力してください。\n' +
                        su.schedulerUrl;
                } else {
                    postEventHander.resultMessage =
                        '【Error】Your attendance on ' +
                        actDate +
                        " in Densuke has not been marked as 〇.\nPlease update Densuke and type 'update'.\n" +
                        su.schedulerUrl;
                }
            }
        } else {
            if (postEventHander.lang === 'ja') {
                postEventHander.resultMessage =
                    '【エラー】スケジューラー名称登録が完了していません。\n登録を完了させて、再度PayNow画像をアップロードして下さい。\n登録は「登録」と入力してください。\n' +
                    su.schedulerUrl;
            } else {
                postEventHander.resultMessage =
                    "【Error】The initial registration is not complete.\nPlease complete the initial registration and upload the PayNow photo again.\nFor the initial registration, please type 'how to register'.\n" +
                    su.schedulerUrl;
            }
        }
    }

    public myResult(postEventHander: PostEventHandler): void {
        if (!postEventHander.userId && !gasUtil.getDensukeName(lineUtil.getLineDisplayName(postEventHander.userId))) {
            postEventHander.resultMessage = '初回登録が終わっていません。"登録"と入力し、初回登録を完了させてください。';
        }
        postEventHander.isFlex = true;
        const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const jsonStr: string = ss.getSheetByName('MemberCardLayout')?.getRange(1, 1).getValue();
        const messageJson: JSON = JSON.parse(jsonStr);
        postEventHander.messageJson = messageJson;
        this.reflectOwnResult(messageJson, postEventHander.userId, postEventHander.lang);
        // postEventHander.resultMessage = jsonStr;
    }

    private translatePlace(place: string, lang: string): string {
        if (place === '1') {
            return lang !== 'ja' ? '1st' : '1位';
        } else if (place === '2') {
            return lang !== 'ja' ? '2nd' : '2位';
        } else if (place === '3') {
            return lang !== 'ja' ? '3rd' : '3位';
        } else {
            return lang !== 'ja' ? place + 'th' : place + '位';
        }
    }

    private chooseMedal(place: number): string {
        if (place === 1) {
            return 'https://lh3.googleusercontent.com/d/1ishdfKxuj1fuz7kU6HOZ0NXh7jrZAr0H';
        } else if (place === 2) {
            return 'https://lh3.googleusercontent.com/d/1KKI0m8X3iR6nk1KC0eLbMHvY3QgWxUjz';
        } else if (place === 3) {
            return 'https://lh3.googleusercontent.com/d/1iqWrPdjUDe66MguqAjAiR08pYEAFL-u4';
        } else {
            return 'https://lh3.googleusercontent.com/d/1wMh5Ofoxq89EBIuijDhM-CG52kzUwP1g';
        }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private reflectOwnResult(jsonMessage: any, userId: string, lang: string): void {
        const resultSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.personalTotalSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = resultSheet.getDataRange().getValues();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultRow: any[] | undefined = resultValues.find(row => row[0] === userId);
        if (resultRow) {
            //個人戦績
            if (lang !== 'ja') {
                jsonMessage.contents[0].body.contents[2].contents[0].contents[0].text = String('Match Attendance'); //参加数
                jsonMessage.contents[0].body.contents[2].contents[1].contents[0].text = String('Total Goals No'); //通算ゴール数
                jsonMessage.contents[0].body.contents[2].contents[2].contents[0].text = String('Total Assists No'); //通算アシスト数
                jsonMessage.contents[0].body.contents[2].contents[3].contents[0].text = String('Top Scorers Rnk'); //得点王ランキング
                jsonMessage.contents[0].body.contents[2].contents[4].contents[0].text = String('Top Assist Rnk'); //アシスト王ランキング
                jsonMessage.contents[0].body.contents[2].contents[5].contents[0].text = String('Okamoto Cup Rnk'); //岡本カップランキング

                jsonMessage.contents[0].body.contents[3].text = 'Okamoto Cup Result'; //１位獲得数
                jsonMessage.contents[0].body.contents[4].contents[0].contents[0].text = 'No of Championship'; //１位獲得数
                jsonMessage.contents[0].body.contents[4].contents[1].contents[0].text = 'No of Last-place'; //最下位獲得数
                jsonMessage.contents[0].body.contents[4].contents[2].contents[0].text = 'Okamoto Cup points'; //チームポイント獲得数
            }
            jsonMessage.contents[0].body.contents[0].contents[0].text = String(resultRow[1]); //名前
            jsonMessage.contents[0].body.contents[2].contents[0].contents[1].text = String(resultRow[2]); //参加数
            jsonMessage.contents[0].body.contents[2].contents[1].contents[1].text = String(resultRow[5]); //通算ゴール数
            jsonMessage.contents[0].body.contents[2].contents[2].contents[1].text = String(resultRow[6]); //通算アシスト数
            jsonMessage.contents[0].body.contents[2].contents[3].contents[1].text = String(this.translatePlace(resultRow[11], lang)); //得点王ランキング
            jsonMessage.contents[0].body.contents[2].contents[4].contents[1].text = String(this.translatePlace(resultRow[12], lang)); //アシスト王ランキング
            jsonMessage.contents[0].body.contents[2].contents[5].contents[1].text = String(this.translatePlace(resultRow[13], lang)); //岡本カップランキング

            jsonMessage.contents[0].body.contents[4].contents[0].contents[1].text = String(resultRow[9]); //１位獲得数
            jsonMessage.contents[0].body.contents[4].contents[1].contents[1].text = String(resultRow[10]); //最下位獲得数
            jsonMessage.contents[0].body.contents[4].contents[2].contents[1].text = String(resultRow[8]); //チームポイント獲得数

            if (resultRow[14] === 1) {
                jsonMessage.contents[0].body.contents[0].contents[1] = {};
                jsonMessage.contents[0].body.contents[0].contents[1].type = 'image';
                jsonMessage.contents[0].body.contents[0].contents[1].url = 'https://lh3.googleusercontent.com/d/1fAy83HzkttX06Vm-wt5oRPWlB-JOWcC0';
                jsonMessage.contents[0].body.contents[0].contents[1].size = 'xxs';
                jsonMessage.contents[0].body.contents[0].contents[1].align = 'end';
            }
        }
        //ランキング
        let ten: string = '点';
        if (lang !== 'ja') {
            ten = '';
        }
        const gRankingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.gRankingSheet;
        const gRankValues = gRankingSheet.getDataRange().getValues();
        this.writeRankingContents(gRankValues, jsonMessage, lang, ten, 1);

        const aRankingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.aRankingSheet;
        const aRankValues = aRankingSheet.getDataRange().getValues();
        this.writeRankingContents(aRankValues, jsonMessage, lang, ten, 2);

        const oRankingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.oRankingSheet;
        const oRankValues = oRankingSheet.getDataRange().getValues();
        this.writeRankingContents(oRankValues, jsonMessage, lang, 'pt', 3);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private writeRankingContents(aRankValues: any[][], jsonMessage: any, lang: string, ten: string, contentsIndex: number) {
        // const densukeVals = GasProps.instance.mappingSheet.getDataRange().getValues();
        for (const ranking of aRankValues) {
            if (ranking[0] !== '' && ranking[0] !== '伝助名称' && ranking[1] < 4 && ranking[3] > 0) {
                // if (ranking[1] === '1') {
                jsonMessage.contents[contentsIndex].body.contents.push({
                    type: 'box',
                    layout: 'baseline',
                    spacing: 'sm',
                    contents: [
                        {
                            type: 'icon',
                            url: this.chooseMedal(ranking[1]),
                        },
                        {
                            type: 'text',
                            text: this.translatePlace(ranking[1], lang),
                            wrap: true,

                            flex: 1,
                        },
                        // {
                        //     aspectMode: 'cover',
                        //     size: '20px',
                        //     type: 'image',
                        //     url: this.rankingPic(ranking[0], densukeVals),
                        // },
                        {
                            type: 'text',
                            text: ranking[0],
                            flex: 4,
                        },
                        {
                            type: 'text',
                            text: ranking[3] + ten,
                            flex: 1,
                        },
                        {
                            type: 'icon',
                            url: this.rankingArrow(ranking[1], ranking[2]),
                            margin: 'none',
                            offsetTop: '2px',
                        },
                    ],
                });
            }
        }
    }

    // // eslint-disable-next-line @typescript-eslint/no-explicit-any
    // private rankingPic(densukeNm: string, densukeVals: any[][]): string {
    //     const userId = gasUtil.getLineUserId(densukeNm);
    //     const row = densukeVals.find(item => item[2] === userId);
    //     let url = 'https://lh3.googleusercontent.com/d/1wMh5Ofoxq89EBIuijDhM-CG52kzUwP1g';
    //     if (row && row[4]) {
    //         url = row[4];
    //     }
    //     return url;
    // }

    private rankingArrow(place: number, past: number): string {
        if (!past) {
            return 'https://lh3.googleusercontent.com/d/1KsKJg9LNZOS0pMGq4Yqzv10ZfBGDsEKB';
        } else if (place < past) {
            return 'https://lh3.googleusercontent.com/d/1h8FcN6ESmMc4gKKGpRvi2x3Nk_ss9eIZ';
        } else if (place > past) {
            return 'https://lh3.googleusercontent.com/d/1fmHGmCjYTlmEoElnh-S441K3r0zmoCXt';
        } else if (place === past) {
            return 'https://lh3.googleusercontent.com/d/1KjbGAgb9Cid7Osoj7UZwY-V8fp5or5sa';
        }
        return '';
    }

    public aggregate(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const attendees = su.extractAttendees('〇');
        const actDate = su.extractDateFromRownum();
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        console.log('actDate' + actDate);
        console.log('attendees' + attendees);
        su.generateSummaryBase(attendees, actDate);
        postEventHander.resultMessage = su.getSummaryStr(attendees, actDate, addy);
    }

    // public unpaid(postEventHander: PostEventHandler): void {
    //     const $ = densukeUtil.getDensukeCheerio();
    //     const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
    //     const unpaid = gasUtil.getUnpaid(actDate);
    //     postEventHander.resultMessage = '未払いの人 (' + unpaid.length + '名): ' + unpaid.join(', ');
    // }

    public remind(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        postEventHander.resultMessage = su.generateRemind();
    }

    // public densukeUpd(postEventHander: PostEventHandler): void {
    //     const $ = densukeUtil.getDensukeCheerio();
    //     const lineName = lineUtil.getLineDisplayName(postEventHander.userId);
    //     const members = densukeUtil.extractMembers($);
    //     const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '〇', members);
    //     const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
    //     const settingSheet = GasProps.instance.settingSheet;
    //     const addy = settingSheet.getRange('B2').getValue();
    //     densukeUtil.generateSummaryBase($);
    //     postEventHander.paynowOwnerMsg = '【' + lineName + 'さんにより更新されました】\n' + densukeUtil.getSummaryStr(attendees, actDate, addy);
    //     // this.sendMessageToPaynowOwner(ownerMessage);
    //     if (postEventHander.lang === 'ja') {
    //         postEventHander.resultMessage = '伝助の更新ありがとうございました！PayNowのスクリーンショットを再度こちらへ送って下さい。';
    //     } else {
    //         postEventHander.resultMessage = 'Thank you for updating Densuke! Please send PayNow screenshot here again.';
    //     }
    // }

    // public regInfo(postEventHander: PostEventHandler): void {
    //     const su: SchedulerUtil = new SchedulerUtil();
    //     if (postEventHander.lang === 'ja') {
    //         postEventHander.resultMessage =
    //             '伝助名称の登録を行います。\n伝助のアカウント名を以下のフォーマットで入力してください。\n@@register@@伝助名前\n例）@@register@@やまだじょ\n' +
    //             su.schedulerUrl;
    //     } else {
    //         postEventHander.resultMessage =
    //             'We will perform the densuke name registration.\nPlease enter your Densuke account name in the following format:\n@@register@@XXXXX\nExample)@@register@@Sahim\n' +
    //             su.schedulerUrl;
    //     }
    // }

    public managerInfo(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        if (gasUtil.isKanji(postEventHander.userId)) {
            postEventHander.resultMessage =
                '設定：' +
                GasProps.instance.settingSheetUrl +
                '\nPayNow：' +
                GasProps.instance.payNowFolderUrl +
                '\nReport URL:' +
                GasProps.instance.reportSheetUrl +
                '\nEvent Result URL:' +
                GasProps.instance.eventResultUrl +
                '\nスケジューラー：' +
                su.schedulerUrl +
                '\nチャット状況：' +
                ScriptProps.instance.chat +
                '\nメッセージ利用状況：' +
                ScriptProps.instance.messageUsage +
                '\n' +
                '\nAppScript：' +
                'https://script.google.com/home/projects/1K0K--vXLzq1p97HHwCbdynsASyjFBndjbkz5YDj9_PN_yG4K1qS4pBok/executions' +
                '\n' +
                postEventHander.generateCommandList();
            // '\n 利用可能コマンド:集計, aggregate, 紹介, introduce, 登録, how to register, リマインド, remind, 伝助更新, update, 未払い, unpaid, 未登録参加者, unregister, @@register@@名前 ';
        } else {
            postEventHander.resultMessage = 'えっ！？このコマンドは平民のキミには内緒だよ！';
        }
    }

    // public unRegister(postEventHander: PostEventHandler) {
    //     this.aggregate(postEventHander);
    //     const $ = densukeUtil.getDensukeCheerio();
    //     const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
    //     const unRegister = gasUtil.getUnRegister(actDate);
    //     postEventHander.resultMessage = '現在未登録の参加者 (' + unRegister.length + '名): ' + unRegister.join(', ');
    // }

    public ranking(postEventHander: PostEventHandler): void {
        const scoreBook: ScoreBook = new ScoreBook();
        const su: SchedulerUtil = new SchedulerUtil();

        // const $ = densukeUtil.getDensukeCheerio();
        const actDate = su.extractDateFromRownum();
        const attendees = su.extractAttendees('〇');
        scoreBook.makeEventFormat(actDate, attendees);
        scoreBook.generateScoreBook(actDate, attendees, Title.ASSIST);
        scoreBook.generateScoreBook(actDate, attendees, Title.TOKUTEN);
        scoreBook.generateOkamotoBook(actDate, attendees);
        scoreBook.aggregateScore();
        postEventHander.resultMessage = 'ランキングが更新されました\n' + GasProps.instance.eventResultUrl;
    }

    public systemTest(postEventHander: PostEventHandler): void {
        try {
            ScriptProps.startTest();
            this.managerInfo(postEventHander);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const gasTest: any = new GasTestSuite();
            if (postEventHander.messageText.startsWith('システムテスト@')) {
                const testCommand: string = postEventHander.messageText.split('システムテスト@')[1];
                if (typeof gasTest[testCommand] === 'function') {
                    gasTest[testCommand](postEventHander, this);
                }
            } else {
                const methodNames: string[] = Object.getOwnPropertyNames(GasTestSuite.prototype).filter(
                    name => name !== 'constructor' && name.startsWith('test')
                );
                methodNames.forEach(methodName => {
                    if (typeof gasTest[methodName] === 'function') {
                        gasTest[methodName](postEventHander, this);
                    }
                });
            }
            postEventHander.resultMessage = postEventHander.testResult.join('\n');
            postEventHander.resultImage = '';
        } finally {
            ScriptProps.endTest();
        }
    }
}
