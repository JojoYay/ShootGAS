// import { DensukeUtil } from './densukeUtil';
import { TotalScore } from './totalScore';
import { GasProps } from './gasProps';
import { ScriptProps } from './scriptProps';
import { GasUtil } from './gasUtil';

export enum Title {
    TOKUTEN = '得点王ランキング',
    ASSIST = 'アシスト王ランキング',
    OKAMOTO = '岡本カップランキング',
    ATTEND = '出席ランキング',
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function testaaaa() {
    console.log('test');
    // const scoreBook: ScoreBook = new ScoreBook();
    // scoreBook.makeEventFormat();
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function calcAllResult() {
    const scoreBook: ScoreBook = new ScoreBook();
    const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
    // const reportSheets: GoogleAppsScript.Spreadsheet.Sheet[] = reportSS.getSheets();
    const eventSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
    const eventSummaryVal = eventSheet.getDataRange().getValues();
    for (let i = eventSummaryVal.length - 1; i >= 1; i--) {
        const actDate: string = eventSummaryVal[i][1];
        const reportSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(actDate);
        if (reportSheet) {
            console.log(actDate);
            const attendees = scoreBook.getAttendeesFromRecord(reportSheet);
            scoreBook.generateOkamotoBook(actDate, attendees);
            scoreBook.generateScoreBook(actDate, attendees, Title.TOKUTEN);
            scoreBook.generateScoreBook(actDate, attendees, Title.ASSIST);
        }
    }
    // scoreBook.makeEventFormat();
    scoreBook.aggregateScore();
}

export class ScoreBook {
    public getAttendeesFromRecord(report: GoogleAppsScript.Spreadsheet.Sheet): string[] {
        const result: string[] = [];
        const repoVals = report.getDataRange().getValues();
        for (let i = 9; i < repoVals.length; i++) {
            result.push(repoVals[i][0]);
        }
        return result;
    }

    public aggregateScore(): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const eventDetails: GoogleAppsScript.Spreadsheet.Sheet[] = eventSS.getSheets();
        const totalResult: GoogleAppsScript.Spreadsheet.Sheet = this.getTotalSheet(eventDetails);
        const eventSheet: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDataSheet(eventDetails);
        const dataList: TotalScore[] = this.exstractTotalScores(eventSheet, eventDetails);
        // console.log(dataList);
        this.writeTotalRecord(totalResult, dataList);
    }

    public makeEventFormat(actDate: string, attendees: string[]): void {
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        this.updateEventSheet(actDate, attendees);

        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate);
        this.updateAttendeeName(eventDetail, attendees);
    }

    public getEventDetailSheet(eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet, actDate: string): GoogleAppsScript.Spreadsheet.Sheet {
        let eventDetail: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(actDate);

        if (!eventDetail) {
            eventDetail = eventSS.insertSheet(actDate);
            eventDetail.appendRow(['名前', 'チーム', '得点', 'アシスト']);
            this.moveSheetToHead(eventDetail, eventSS);
        }
        return eventDetail;
    }

    private exstractTotalScores(eventSheet: GoogleAppsScript.Spreadsheet.Sheet, eventDetails: GoogleAppsScript.Spreadsheet.Sheet[]) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const eventSheetVal: any[][] = eventSheet.getDataRange().getValues();
        const dataList: TotalScore[] = [];
        for (const sheet of eventDetails) {
            if (sheet.getSheetName() === 'Total' || sheet.getSheetName() === 'EventData' || sheet.getSheetName().endsWith('_s')) {
                continue;
            }
            const allValues = sheet.getDataRange().getValues();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const eventRow: any[] | undefined = eventSheetVal.find(item => item[1] === sheet.getSheetName());
            if (!eventRow) {
                continue;
                // throw new Error('以下のデータが見つかりません actDate SheetName:' + sheet.getSheetName());
            }
            const g: GasUtil = new GasUtil();
            const densukeMappingValue = g.getLineUserIdRangeValue();
            // console.log(eventRow);
            const totalMatchs = eventSheetVal.length - 1;
            const topOkamotoPoint = this.getTopPoint(eventRow);
            const bottomOkamotoPoint = 1; //最下位は１という事にする

            for (const allValueRow of allValues) {
                // console.log(allValueRow);
                if (allValueRow[0] === '名前') {
                    continue;
                }
                let totalScore: TotalScore | null = null;
                for (const t of dataList) {
                    if (t.name === allValueRow[0]) {
                        totalScore = t;
                        break;
                    }
                }
                if (!totalScore) {
                    totalScore = new TotalScore();
                    dataList.push(totalScore);
                    totalScore.name = allValueRow[0];
                    totalScore.totalMatchs = totalMatchs;
                }

                totalScore.playTime++;
                const userRow = densukeMappingValue.find(item => item[1] === allValueRow[0]);
                if (userRow && userRow[2]) {
                    totalScore.userId = userRow[2];
                }
                if (eventRow[4] === '晴れ') {
                    totalScore.sunnyPlay++;
                } else if (eventRow[4] === '雨') {
                    totalScore.rainyPlay++;
                }
                if (
                    eventRow[5] === totalScore.name ||
                    eventRow[17] === totalScore.name ||
                    eventRow[18] === totalScore.name ||
                    eventRow[19] === totalScore.name ||
                    eventRow[20] === totalScore.name
                ) {
                    totalScore.mipCount++;
                }
                if (allValueRow[1]) {
                    const teamPt: number = Number(totalScore.fetchTeamPoint(eventRow, allValueRow[1]));
                    totalScore.teamPoint += teamPt;
                    // if (totalScore.isTopTeam(eventRow, allValueRow[1])) {
                    if (teamPt === topOkamotoPoint) {
                        totalScore.winCount++;
                    }
                    // if (totalScore.fetchTeamPoint(eventRow, allValueRow[1]) === 0) {
                    if (teamPt === bottomOkamotoPoint) {
                        totalScore.loseCount++;
                    }
                }
                if (allValueRow[2]) {
                    totalScore.goalCount += allValueRow[2];
                }
                if (allValueRow[3]) {
                    totalScore.assistCount += allValueRow[3];
                }
            }
        }
        return dataList;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private getTopPoint(eventRow: any[]) {
        let max: number = 0;
        for (let i = 7; i <= 16; i++) {
            const num: number = eventRow[i];
            if (eventRow[i] && num > max) {
                max = eventRow[i];
            }
        }
        return max;
    }

    // private writeTotalRecord(totalResult: GoogleAppsScript.Spreadsheet.Sheet, dataList: TotalScore[]) {
    //     let lastRow: number = totalResult.getLastRow();
    //     if (lastRow > 2) {
    //         totalResult.deleteRows(2, lastRow - 1);
    //     }

    //     // データを一度に追加するための配列を作成
    //     const rowsToAdd = dataList.map(score => [
    //         score.userId,
    //         score.name,
    //         score.playTime,
    //         score.sunnyPlay,
    //         score.rainyPlay,
    //         score.goalCount,
    //         score.assistCount,
    //         score.mipCount,
    //         score.teamPoint,
    //         score.winCount,
    //         score.loseCount,
    //         score.totalMatchs,
    //     ]);

    //     // 一度にデータを追加
    //     totalResult.getRange(2, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);

    //     lastRow = totalResult.getLastRow();
    //     const lastCol = totalResult.getLastColumn();
    //     if (lastRow > 1) {
    //         totalResult.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true);
    //         totalResult.getRange(2, 1, lastRow - 1, lastCol).sort({ column: 6, ascending: false });

    //         let rank = 1;
    //         let prevScore = null;
    //         let prevRank = 1;
    //         let rangeVals = totalResult.getDataRange().getValues();

    //         // 得点王
    //         const grankingUpdates = [];
    //         const orankingUpdates = [];
    //         for (let i = 1; i < rangeVals.length; i++) {
    //             const currentScore = rangeVals[i][5];
    //             if (currentScore !== prevScore) {
    //                 prevRank = rank;
    //             }
    //             grankingUpdates.push([prevRank]);
    //             if (currentScore !== prevScore) {
    //                 rank++;
    //             }
    //             prevScore = currentScore;
    //         }
    //         totalResult.getRange(2, 13, grankingUpdates.length, 1).setValues(grankingUpdates);

    //         // 追加のランクを設定
    //         for (let i = 1; i < rangeVals.length; i++) {
    //             orankingUpdates.push([rank - 1]);
    //         }
    //         totalResult.getRange(2, 17, orankingUpdates.length, 1).setValues(orankingUpdates);

    //         // okamoto
    //         rank = 1;
    //         prevScore = null;
    //         prevRank = 1;
    //         totalResult.getRange(2, 1, lastRow - 1, lastCol).sort({ column: 9, ascending: false });
    //         rangeVals = totalResult.getDataRange().getValues();
    //         const okamotoUpdates = [];
    //         for (let i = 1; i < rangeVals.length; i++) {
    //             const currentScore = rangeVals[i][8];
    //             if (currentScore !== prevScore) {
    //                 prevRank = rank;
    //             }
    //             okamotoUpdates.push([prevRank]);
    //             if (currentScore !== prevScore) {
    //                 rank++;
    //             }
    //             prevScore = currentScore;
    //         }
    //         totalResult.getRange(2, 15, okamotoUpdates.length, 1).setValues(okamotoUpdates);

    //         // assist
    //         rank = 1;
    //         prevScore = null;
    //         prevRank = 1;
    //         const assistUpdates = [];
    //         for (let i = 1; i < rangeVals.length; i++) {
    //             const currentScore = rangeVals[i][6];
    //             if (currentScore !== prevScore) {
    //                 prevRank = rank;
    //             }
    //             assistUpdates.push([prevRank]);
    //             if (currentScore !== prevScore) {
    //                 rank++;
    //             }
    //             prevScore = currentScore;
    //         }
    //         totalResult.getRange(2, 14, assistUpdates.length, 1).setValues(assistUpdates);

    //         // 特別な条件を満たす場合の処理
    //         const mipNames: string[] = this.checkMip(GasProps.instance.eventResultSheet.getDataRange().getValues());
    //         const specialUpdates = [];
    //         for (let i = 1; i < rangeVals.length; i++) {
    //             const currentName = rangeVals[i][1];
    //             const currentGranking = rangeVals[i][12];
    //             const currentAranking = rangeVals[i][13];
    //             const currentOranking = rangeVals[i][14];

    //             if (currentGranking === 1 || currentAranking === 1 || currentOranking === 1 || mipNames.includes(currentName)) {
    //                 specialUpdates.push([1]);
    //             } else {
    //                 specialUpdates.push([0]);
    //             }
    //         }
    //         totalResult.getRange(2, 16, specialUpdates.length, 1).setValues(specialUpdates);
    //     }
    // }

    private writeTotalRecord(totalResult: GoogleAppsScript.Spreadsheet.Sheet, dataList: TotalScore[]) {
        let lastRow: number = totalResult.getLastRow();
        if (lastRow > 2) {
            // ヘッダーを除く全行をクリア
            totalResult.getRange(2, 1, lastRow - 1, totalResult.getLastColumn()).clearContent();
            totalResult.getRange(2, 1, lastRow - 1, totalResult.getLastColumn()).setBorder(false, false, false, false, false, false);
        }
        for (const score of dataList) {
            totalResult.appendRow([
                score.userId,
                score.name,
                score.playTime,
                score.sunnyPlay,
                score.rainyPlay,
                score.goalCount,
                score.assistCount,
                score.mipCount,
                score.teamPoint,
                score.winCount,
                score.loseCount,
                score.totalMatchs,
            ]);
        }
        lastRow = totalResult.getLastRow();
        const lastCol = totalResult.getLastColumn();
        if (lastRow > 1) {
            totalResult.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true);
            totalResult.getRange(2, 1, lastRow, lastCol).sort({ column: 6, ascending: false });
            let rank = 1;
            let prevScore = null;
            let prevRank = 1;
            let rangeVals = totalResult.getDataRange().getValues();
            //得点王
            for (let i = 1; i < rangeVals.length; i++) {
                const currentScore = rangeVals[i][5];
                if (currentScore !== prevScore) {
                    prevRank = rank;
                }
                totalResult.getRange(i + 1, 13).setValue(prevRank);

                if (currentScore !== prevScore) {
                    rank++;
                }
                prevScore = currentScore;
            }
            for (let i = 1; i < rangeVals.length; i++) {
                totalResult.getRange(i + 1, 17).setValue(rank - 1);
            }

            rank = 1;
            prevScore = null;
            prevRank = 1;
            totalResult.getRange(2, 1, lastRow, lastCol).sort({ column: 9, ascending: false });
            rangeVals = totalResult.getDataRange().getValues();
            //okamoto
            for (let i = 1; i < rangeVals.length; i++) {
                const currentScore = rangeVals[i][8];
                if (currentScore !== prevScore) {
                    prevRank = rank;
                }
                totalResult.getRange(i + 1, 15).setValue(prevRank);
                if (currentScore !== prevScore) {
                    rank++;
                }
                prevScore = currentScore;
            }
            for (let i = 1; i < rangeVals.length; i++) {
                totalResult.getRange(i + 1, 19).setValue(rank - 1);
            }

            totalResult.getRange(2, 1, lastRow, lastCol).sort({ column: 7, ascending: false });
            rangeVals = totalResult.getDataRange().getValues();
            rank = 1;
            prevScore = null;
            prevRank = 1;

            const eventData: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
            const mipNames: string[] = this.checkMip(eventData.getDataRange().getValues());
            //assist
            for (let i = 1; i < rangeVals.length; i++) {
                const currentScore = rangeVals[i][6];
                if (currentScore !== prevScore) {
                    prevRank = rank;
                }
                totalResult.getRange(i + 1, 14).setValue(prevRank);
                if (currentScore !== prevScore) {
                    rank++;
                }
                prevScore = currentScore;
            }
            for (let i = 1; i < rangeVals.length; i++) {
                totalResult.getRange(i + 1, 18).setValue(rank - 1);
            }

            rangeVals = totalResult.getDataRange().getValues();
            for (let i = 1; i < rangeVals.length; i++) {
                const currentName = rangeVals[i][1];
                const currentGranking = rangeVals[i][12];
                const currentAranking = rangeVals[i][13];
                const currentOranking = rangeVals[i][14];

                // if (currentGranking === '1位' || currentAranking === '1位' || mipNames.includes(currentName)) {
                if (currentGranking === 1 || currentAranking === 1 || currentOranking === 1 || mipNames.includes(currentName)) {
                    // console.log(currentName + ':' + currentAranking + ' ' + rangeVals[i]);
                    totalResult.getRange(i + 1, 16).setValue(1);
                }
            }
        }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private checkMip(eventDataVal: any[][]): string[] {
        const resultNames: string[] = [];
        if (eventDataVal.length > 0) {
            for (let i = 0; i < 3; i++) {
                if (eventDataVal[i]) {
                    resultNames.push(eventDataVal[i][5]);
                }
            }
        }
        return resultNames;
    }

    private getTotalSheet(sheets: GoogleAppsScript.Spreadsheet.Sheet[]): GoogleAppsScript.Spreadsheet.Sheet {
        for (const sheet of sheets) {
            if (sheet.getSheetName() === 'Total') {
                return sheet;
            }
        }
        throw new Error('Total Sheet was not found');
    }

    private getEventDataSheet(sheets: GoogleAppsScript.Spreadsheet.Sheet[]): GoogleAppsScript.Spreadsheet.Sheet {
        for (const sheet of sheets) {
            if (sheet.getSheetName() === 'EventData') {
                return sheet;
            }
        }
        throw new Error('EventData Sheet was not found');
    }

    public updateAttendeeName(eventDetail: GoogleAppsScript.Spreadsheet.Sheet, attendees: string[]): void {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const allDetails: any[][] = eventDetail.getDataRange().getValues();
        const attendeesArray = Array.isArray(attendees) ? attendees : [attendees]; // attendees が配列でない場合の安全策

        const lastRow: number = eventDetail.getLastRow();
        // 条件に合うデータを格納する配列
        const filteredData: string[][] = [];
        //あとで存在してるやつかどうか判定で使うSet
        const loggedAttendeesSet = new Set<string>();

        if (attendees.length > 0) {
            if (lastRow > 1) {
                // allDetails を逆順にループ
                for (let i = 0; i < allDetails.length; i++) {
                    const name = allDetails[i][0];
                    const isEmptyRow = allDetails[i].slice(1, 4).every(cell => !cell); // B列からD列がすべて空かチェック
                    if (attendeesArray.includes(name) && !isEmptyRow) {
                        // 条件に合うデータを filteredData に追加
                        filteredData.push(allDetails[i]);
                        loggedAttendeesSet.add(allDetails[i][0] as string);
                        // console.log(allDetails[i]);
                    }
                }
                if (lastRow > 1) {
                    eventDetail.getRange(2, 1, lastRow - 1, eventDetail.getLastColumn()).clearContent(); // ヘッダーを除く全行をクリア
                    // eventDetail.getRange(2, 1, lastRow - 1, eventDetail.getLastColumn()).clearDataValidations(); // ヘッダーを除く全行をクリア
                }
                // filteredData を一括でシートに書き込む
                if (filteredData.length > 0) {
                    eventDetail.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
                }
            }
            const teamName: string[] = [
                'チーム1',
                'チーム2',
                'チーム3',
                'チーム4',
                'チーム5',
                'チーム6',
                'チーム7',
                'チーム8',
                'チーム9',
                'チーム10',
            ];

            // 新しく追加した行のインデックスを取得
            let newRowIndex = eventDetail.getLastRow(); // appendRow の後に行数を取得
            // データバリデーションを設定
            const teamNameVal = SpreadsheetApp.newDataValidation().requireValueInList(teamName).build();
            const goalCount: string[] = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
            const goalCountVal = SpreadsheetApp.newDataValidation().requireValueInList(goalCount).build();

            const validations = {
                2: teamNameVal, // 2列目 (B列) : チーム名
                3: goalCountVal, // 3列目 (C列) : 得点
                4: goalCountVal, // 4列目 (D列) : 得点
            };

            for (let i = 0; i < attendees.length; i++) {
                if (!loggedAttendeesSet.has(attendees[i])) {
                    eventDetail.appendRow([attendees[i]]); // 新しい行を追加
                    newRowIndex++;
                    for (const [column, validation] of Object.entries(validations)) {
                        eventDetail.getRange(newRowIndex, Number(column)).setDataValidation(validation);
                    }
                }
            }
        }
    }

    private moveSheetToHead(sheet: GoogleAppsScript.Spreadsheet.Sheet, eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
        sheet.activate();
        eventSS.moveActiveSheet(3);
    }

    public updateEventSheet(actDate: string, attendees: string[]): void {
        const eventData: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rangeValues: any[][] = eventData.getDataRange().getValues();
        const now: Date = new Date();
        const rowNumber: number = this.getTargetCol(rangeValues, actDate);
        if (rowNumber === -1) {
            this.createInitialEvent(attendees, eventData, now, actDate);
        } else {
            const valuesToSet = [now, actDate, attendees.length, attendees.join(', ')];
            eventData.getRange(rowNumber + 1, 1, 1, valuesToSet.length).setValues([valuesToSet]);
            // eventData.getRange(rowNumber + 1, 1).setValue(now);
            // eventData.getRange(rowNumber + 1, 2).setValue(actDate);
            // eventData.getRange(rowNumber + 1, 3).setValue(attendees.length);
            // eventData.getRange(rowNumber + 1, 4).setValue(attendees.join(', '));
            const attendVal = SpreadsheetApp.newDataValidation().requireValueInList(attendees).build();
            eventData.getRange(rowNumber + 1, 6).setDataValidation(attendVal);
        }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private getTargetCol(rangeValues: any[][], actDate: string): number {
        if (rangeValues.length > 1) {
            for (let i = 1; i < rangeValues.length; i++) {
                const actDateInSheet: string = rangeValues[i][1];
                if (actDateInSheet === actDate) {
                    return i;
                }
            }
        }
        return -1;
    }

    // private createInitialEvent(attendees: string[], eventData: GoogleAppsScript.Spreadsheet.Sheet, now: Date, actDate: string): void {
    //     const teamPoint: string[] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
    //     const teamPointVal = SpreadsheetApp.newDataValidation().requireValueInList(teamPoint).build();
    //     const weather: string[] = ['晴れ', '曇り', '雨'];
    //     const weatherVal = SpreadsheetApp.newDataValidation().requireValueInList(weather).build();
    //     const attendVal = SpreadsheetApp.newDataValidation().requireValueInList(attendees).build();
    //     const lastCol: number = eventData.getLastColumn();
    //     eventData.appendRow([now, actDate, attendees.length, attendees.join(',')]);
    //     const lastRow: number = eventData.getLastRow();

    //     eventData.getRange(lastRow, 5).setDataValidation(weatherVal);
    //     eventData.getRange(lastRow, 6).setDataValidation(attendVal);

    //     eventData.getRange(lastRow, 8).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 9).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 10).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 11).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 12).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 13).setDataValidation(teamPointVal);
    //     //追加
    //     eventData.getRange(lastRow, 14).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 15).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 16).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 17).setDataValidation(teamPointVal);
    //     eventData.getRange(lastRow, 18).setDataValidation(attendVal);
    //     eventData.getRange(lastRow, 19).setDataValidation(attendVal);
    //     eventData.getRange(lastRow, 20).setDataValidation(attendVal);
    //     eventData.getRange(lastRow, 21).setDataValidation(attendVal);

    //     if (lastRow > 2) {
    //         const newRowRange: GoogleAppsScript.Spreadsheet.Range = eventData.getRange(lastRow, 1, 1, lastCol - 1);
    //         eventData.moveRows(newRowRange, 2);
    //     }
    // }
    private createInitialEvent(attendees: string[], eventData: GoogleAppsScript.Spreadsheet.Sheet, now: Date, actDate: string): void {
        const teamPoint: string[] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
        const teamPointVal = SpreadsheetApp.newDataValidation().requireValueInList(teamPoint).build();
        const weather: string[] = ['晴れ', '曇り', '雨'];
        const weatherVal = SpreadsheetApp.newDataValidation().requireValueInList(weather).build();
        const attendVal = SpreadsheetApp.newDataValidation().requireValueInList(attendees).build();
        const lastCol: number = eventData.getLastColumn();

        // 新しい行を追加
        eventData.appendRow([now, actDate, attendees.length, attendees.join(',')]);
        const lastRow: number = eventData.getLastRow();

        // データバリデーションを一度に設定するための配列を作成
        const validations = Array(lastCol).fill(null); // lastColの数だけnullで初期化
        validations[4] = weatherVal; // 5列目
        validations[5] = attendVal; // 6列目

        // 8列目から21列目までのデータバリデーションを設定
        for (let i = 7; i <= 20; i++) {
            validations[i] = i <= 12 ? teamPointVal : attendVal; // 8-12列目はteamPointVal、13-21列目はattendVal
        }

        // 一度に範囲を設定
        const range = eventData.getRange(lastRow, 1, 1, lastCol);
        range.setDataValidations([validations]);

        if (lastRow > 2) {
            const newRowRange: GoogleAppsScript.Spreadsheet.Range = eventData.getRange(lastRow, 1, 1, lastCol - 1);
            eventData.moveRows(newRowRange, 2);
        }
    }

    // public generateOkamotoBook(actDate: string, attendees: string[]) {
    //     const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
    //     let scoreSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(Title.OKAMOTO);

    //     if (!scoreSheet) {
    //         scoreSheet = reportSS.insertSheet(Title.OKAMOTO);
    //         scoreSheet.appendRow(['伝助名称', '順位', '前回順位', '合計ポイント']);
    //         scoreSheet.insertRowBefore(1);
    //     }

    //     if (!this.isActDateExists(actDate, scoreSheet)) {
    //         scoreSheet.insertColumnBefore(5);
    //         scoreSheet.getRange('E2').setValue(actDate);
    //         if (scoreSheet.getLastColumn() > 5) {
    //             scoreSheet.getRange(3, 2, scoreSheet.getLastRow() - 1, 1).copyTo(scoreSheet.getRange(3, 3, scoreSheet.getLastRow() - 1, 1));
    //         }
    //     }

    //     this.addAttendee(scoreSheet, attendees, true);
    //     const scoreValues = scoreSheet.getDataRange().getValues();
    //     const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    //     const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate);
    //     // eslint-disable-next-line @typescript-eslint/no-explicit-any
    //     const detailValues: any[][] = eventDetail.getDataRange().getValues();
    //     const eventSummary: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDataSheet(eventSS.getSheets());
    //     const eventRow = eventSummary
    //         .getDataRange()
    //         .getValues()
    //         .find(item => item[1].toString() === actDate);

    //     if (!eventRow) {
    //         throw new Error(actDate + ' event is not found in EventData Sheet');
    //     }

    //     const resultPoints = eventRow.slice(7, 17).map(point => point || 0); // 7-16列目のポイントを取得
    //     const lastCol = scoreSheet.getLastColumn();
    //     const updates = []; // 更新するデータを格納する配列

    //     for (const score of scoreValues) {
    //         if (score[0] === '伝助名称' || score[0] === '') {
    //             continue;
    //         }

    //         const resultRow = detailValues.find(item => !!item[0] && item[0] === score[0]);
    //         if (resultRow) {
    //             const teamIndex = parseInt(resultRow[1].replace('チーム', '')) - 1; // チーム名からインデックスを取得
    //             const point = resultPoints[teamIndex] || 0; // ポイントを取得
    //             updates.push([point]); // ポイントを更新リストに追加
    //         } else {
    //             updates.push([0]); // 該当チームがない場合は0を追加
    //         }
    //     }

    //     // 一度にポイントを設定
    //     scoreSheet.getRange(3, 5, updates.length, 1).setValues(updates);

    //     // 合計ポイントの計算式を設定
    //     for (let index = 0; index < updates.length; index++) {
    //         const formula = `=SUM(${scoreSheet.getRange(3 + index, 5, 1, lastCol - 4).getA1Notation()})`;
    //         scoreSheet.getRange(3 + index, 4).setFormula(formula);
    //     }

    //     const finalRow = scoreSheet.getLastRow();
    //     const finalCol = scoreSheet.getLastColumn();
    //     scoreSheet.getRange(2, 1, finalRow - 1, finalCol).setBorder(true, true, true, true, true, true);
    //     scoreSheet.getRange(3, 1, finalRow - 1, finalCol).sort({ column: 4, ascending: false });
    //     scoreSheet.getRange(2, 1, 1, finalCol).setBackground('#fff2cc');
    //     scoreSheet.activate();
    //     reportSS.moveActiveSheet(1);

    //     let rank = 1;
    //     let prevScore = null;
    //     let prevRank = 1;
    //     const rangeVals = scoreSheet.getDataRange().getValues();
    //     const rankUpdates = []; // ランクを格納する配列

    //     for (let i = 2; i < rangeVals.length; i++) {
    //         const currentScore = rangeVals[i][3];
    //         if (currentScore !== prevScore) {
    //             prevRank = rank;
    //         }
    //         rankUpdates.push([prevRank]); // ランクを更新リストに追加
    //         if (currentScore !== prevScore) {
    //             rank++;
    //         }
    //         prevScore = currentScore;
    //     }

    //     // 一度にランクを設定
    //     scoreSheet.getRange(3, 2, rankUpdates.length, 1).setValues(rankUpdates);
    // }

    public generateOkamotoBook(actDate: string, attendees: string[]) {
        const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        let scoreSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(Title.OKAMOTO);
        if (!scoreSheet) {
            scoreSheet = reportSS.insertSheet(Title.OKAMOTO);
            // scoreSheet.appendRow(['伝助名称', 'line名称', '合計得点', actDate]);
            scoreSheet.appendRow(['伝助名称', '順位', '前回順位', '合計ポイント']);
            scoreSheet.insertRowBefore(1);
        }
        if (!this.isActDateExists(actDate, scoreSheet)) {
            scoreSheet.insertColumnBefore(5);
            scoreSheet.getRange('E2').setValue(actDate);
            if (scoreSheet.getLastColumn() > 5) {
                scoreSheet.getRange(3, 2, scoreSheet.getLastRow() - 1, 1).copyTo(scoreSheet.getRange(3, 3, scoreSheet.getLastRow() - 1, 1));
            }
        }
        this.addAttendee(scoreSheet, attendees, true);
        const scoreValues = scoreSheet.getDataRange().getValues();
        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate); //ちょっとこのメソッドは危ない（順序によっては新規で作ってる）
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const detailValues: any[][] = eventDetail.getDataRange().getValues(); //こっちが入力したシート
        const eventSummary: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDataSheet(eventSS.getSheets());
        const eventRow = eventSummary
            .getDataRange()
            .getValues()
            .find(item => item[1].toString() === actDate);
        if (!eventRow) {
            throw new Error(actDate + ' event is not found in EventData Sheet');
        }
        const resultPoints = [];
        for (let i = 7; i < 17; i++) {
            if (eventRow[i]) {
                resultPoints.push(eventRow[i]);
            } else {
                resultPoints.push(0);
            }
        }
        let index = 3;
        const lastCol = scoreSheet.getLastColumn();
        for (const score of scoreValues) {
            if (score[0] === '伝助名称' || score[0] === '') {
                continue;
            }
            const resultRow = detailValues.find(item => !!item[0] && item[0] === score[0]); //無い場合もある
            console.log(resultRow);
            if (resultRow) {
                let point: number = 0;
                const team: string = resultRow[1];
                if (team === 'チーム1') {
                    point = resultPoints[0];
                } else if (team === 'チーム2') {
                    point = resultPoints[1];
                } else if (team === 'チーム3') {
                    point = resultPoints[2];
                } else if (team === 'チーム4') {
                    point = resultPoints[3];
                } else if (team === 'チーム5') {
                    point = resultPoints[4];
                } else if (team === 'チーム6') {
                    point = resultPoints[5];
                } else if (team === 'チーム7') {
                    point = resultPoints[6];
                } else if (team === 'チーム8') {
                    point = resultPoints[7];
                } else if (team === 'チーム9') {
                    point = resultPoints[8];
                } else if (team === 'チーム10') {
                    point = resultPoints[9];
                }
                scoreSheet.getRange(index, 5).setValue(point);
            }
            const range = scoreSheet.getRange(index, 5, 1, lastCol - 4);
            const formula = `=SUM(${range.getA1Notation()})`;
            scoreSheet.getRange(index, 4).setFormula(formula);

            index++;
        }
        const finalRow = scoreSheet.getLastRow();
        const finalCol = scoreSheet.getLastColumn();
        scoreSheet.getRange(2, 1, finalRow - 1, finalCol).setBorder(true, true, true, true, true, true);
        scoreSheet.getRange(3, 1, finalRow - 1, finalCol).sort({ column: 4, ascending: false });
        scoreSheet.getRange(2, 1, 1, finalCol).setBackground('#fff2cc');
        scoreSheet.activate();
        reportSS.moveActiveSheet(1);

        let rank = 1;
        let prevScore = null;
        let prevRank = 1;
        const rangeVals = scoreSheet.getDataRange().getValues();
        for (let i = 2; i < rangeVals.length; i++) {
            const currentScore = rangeVals[i][3];
            if (currentScore !== prevScore) {
                prevRank = rank;
            }
            // scoreSheet.getRange(i + 1, 2).setValue(prevRank + '位');
            scoreSheet.getRange(i + 1, 2).setValue(prevRank);
            if (currentScore !== prevScore) {
                rank++;
            }
            prevScore = currentScore;
        }
    }

    // public generateScoreBook(actDate: string, attendees: string[], title: Title): void {
    //     const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
    //     let scoreSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(title);

    //     if (!scoreSheet) {
    //         scoreSheet = reportSS.insertSheet(title);
    //         scoreSheet.appendRow(['伝助名称', '順位', '前回順位', '合計得点']);
    //         scoreSheet.insertRowBefore(1);
    //     }

    //     if (!this.isActDateExists(actDate, scoreSheet)) {
    //         scoreSheet.insertColumnBefore(5);
    //         scoreSheet.getRange('E2').setValue(actDate);
    //         if (scoreSheet.getLastColumn() > 5) {
    //             scoreSheet.getRange(3, 2, scoreSheet.getLastRow() - 1, 1).copyTo(scoreSheet.getRange(3, 3, scoreSheet.getLastRow() - 1, 1));
    //         }
    //     }

    //     this.addAttendee(scoreSheet, attendees, true);

    //     const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    //     const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate);
    //     // eslint-disable-next-line @typescript-eslint/no-explicit-any
    //     const eventValues: any[][] = eventDetail.getDataRange().getValues();
    //     const allVals = scoreSheet.getDataRange().getValues();
    //     const lastCol = scoreSheet.getLastColumn();
    //     const updates = []; // 更新するデータを格納する配列

    //     for (const allRow of allVals) {
    //         if (allRow[0] === '伝助名称' || allRow[0] === '') {
    //             continue;
    //         }
    //         for (const eventRow of eventValues) {
    //             if (eventRow[0] === '名前') {
    //                 continue;
    //             }
    //             if (eventRow[0] === allRow[0]) {
    //                 const score = title === Title.ASSIST ? eventRow[3] : eventRow[2];
    //                 updates.push([score]); // スコアを更新リストに追加
    //                 break; // 一致したら次の allRow に進む
    //             }
    //         }
    //     }

    //     // 一度にスコアを設定
    //     scoreSheet.getRange(3, 5, updates.length, 1).setValues(updates);

    //     // 合計ポイントの計算式を設定
    //     for (let index = 0; index < updates.length; index++) {
    //         const formula = `=SUM(${scoreSheet.getRange(3 + index, 5, 1, lastCol - 3).getA1Notation()})`;
    //         scoreSheet.getRange(3 + index, 4).setFormula(formula);
    //     }

    //     const finalRow = scoreSheet.getLastRow();
    //     const finalCol = scoreSheet.getLastColumn();
    //     scoreSheet.getRange(2, 1, finalRow - 1, finalCol).setBorder(true, true, true, true, true, true);
    //     scoreSheet.getRange(3, 1, finalRow - 1, finalCol).sort({ column: 4, ascending: false });
    //     scoreSheet.getRange(2, 1, 1, finalCol).setBackground('#fff2cc');
    //     scoreSheet.activate();
    //     reportSS.moveActiveSheet(1);

    //     let rank = 1;
    //     let prevScore = null;
    //     let prevRank = 1;
    //     const rangeVals = scoreSheet.getDataRange().getValues();
    //     const rankUpdates = []; // ランクを格納する配列

    //     for (let i = 2; i < rangeVals.length; i++) {
    //         const currentScore = rangeVals[i][3];
    //         if (currentScore !== prevScore) {
    //             prevRank = rank;
    //         }
    //         rankUpdates.push([prevRank]); // ランクを更新リストに追加
    //         if (currentScore !== prevScore) {
    //             rank++;
    //         }
    //         prevScore = currentScore;
    //     }

    //     // 一度にランクを設定
    //     scoreSheet.getRange(3, 2, rankUpdates.length, 1).setValues(rankUpdates);
    // }

    public generateScoreBook(actDate: string, attendees: string[], title: Title): void {
        const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        let scoreSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(title);
        if (!scoreSheet) {
            scoreSheet = reportSS.insertSheet(title);
            // scoreSheet.appendRow(['伝助名称', 'line名称', '合計得点', actDate]);
            scoreSheet.appendRow(['伝助名称', '順位', '前回順位', '合計得点']);
            scoreSheet.insertRowBefore(1);
        }

        if (!this.isActDateExists(actDate, scoreSheet)) {
            scoreSheet.insertColumnBefore(5);
            scoreSheet.getRange('E2').setValue(actDate);
            if (scoreSheet.getLastColumn() > 5) {
                scoreSheet.getRange(3, 2, scoreSheet.getLastRow() - 1, 1).copyTo(scoreSheet.getRange(3, 3, scoreSheet.getLastRow() - 1, 1));
            }
        }
        this.addAttendee(scoreSheet, attendees, true);

        const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
        const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate); //ちょっとこのメソッドは危ない（順序によっては新規で作ってる）
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const eventValues: any[][] = eventDetail.getDataRange().getValues(); //こっちが入力したシート
        const allVals = scoreSheet.getDataRange().getValues(); //こっちがランキングのシート
        // console.log(eventValues);
        const lastCol = scoreSheet.getLastColumn();
        let index: number = 3;
        for (const allRow of allVals) {
            if (allRow[0] === '伝助名称' || allRow[0] === '') {
                continue;
            }
            for (const eventRow of eventValues) {
                if (eventRow[0] === '名前') {
                    continue;
                }
                // console.log('allRow[0]:' + allRow[0]);
                // console.log('resultRow[0]:' + eventRow[0]);
                if (eventRow[0] === allRow[0]) {
                    if (title === Title.ASSIST) {
                        scoreSheet.getRange(index, 5).setValue(eventRow[3]);
                    } else if (title === Title.TOKUTEN) {
                        scoreSheet.getRange(index, 5).setValue(eventRow[2]);
                    }
                }
                const range = scoreSheet.getRange(index, 5, 1, lastCol - 3);
                const formula = `=SUM(${range.getA1Notation()})`;
                scoreSheet.getRange(index, 4).setFormula(formula);
            }
            index++;
        }

        const finalRow = scoreSheet.getLastRow();
        const finalCol = scoreSheet.getLastColumn();
        scoreSheet.getRange(2, 1, finalRow - 1, finalCol).setBorder(true, true, true, true, true, true);
        scoreSheet.getRange(3, 1, finalRow - 1, finalCol).sort({ column: 4, ascending: false });
        scoreSheet.getRange(2, 1, 1, finalCol).setBackground('#fff2cc');
        scoreSheet.activate();
        reportSS.moveActiveSheet(1);

        let rank = 1;
        let prevScore = null;
        let prevRank = 1;
        const rangeVals = scoreSheet.getDataRange().getValues();
        for (let i = 2; i < rangeVals.length; i++) {
            const currentScore = rangeVals[i][3];
            if (currentScore !== prevScore) {
                prevRank = rank;
            }
            // scoreSheet.getRange(i + 1, 2).setValue(prevRank + '位');
            scoreSheet.getRange(i + 1, 2).setValue(prevRank);
            if (currentScore !== prevScore) {
                rank++;
            }
            prevScore = currentScore;
        }
    }

    private addAttendee(scoreSheet: GoogleAppsScript.Spreadsheet.Sheet, attendees: string[], removeZero: boolean): void {
        if (removeZero) {
            this.removeZeroPpl(scoreSheet);
        }
        //シートがある場合は
        // const lastRow = scoreSheet.getLastRow();
        const allDataValues = scoreSheet.getDataRange().getValues();
        // if (allDataValues.length > 2) {
        // console.log(allDataValues);
        for (let i = 0; i < attendees.length; i++) {
            let find = false;
            for (let j: number = 2; j < allDataValues.length; j++) {
                // console.log(allDataValues[j][0] + ' : ' + attendees[i]);
                if (allDataValues[j][0] === attendees[i]) {
                    // console.log(allDataValues[j][0] + ' : ' + attendees[i] + ' matched!');
                    find = true;
                    break;
                }
            }
            if (!find) {
                //本日の参加者をぶっこむ(名前のみ) すでにある人は入れない
                scoreSheet.appendRow([attendees[i]]);
            }
        }
    }

    private removeZeroPpl(scoreSheet: GoogleAppsScript.Spreadsheet.Sheet): void {
        const values = scoreSheet.getDataRange().getValues();
        // console.log(values);
        for (let i = values.length - 1; i >= 1; i--) {
            if (values[i][3] === 0) {
                scoreSheet.deleteRow(i + 1);
            }
        }
    }

    private isActDateExists(actDate: string, scoreSheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
        return actDate === scoreSheet.getRange('E2').getValue();
    }
}
