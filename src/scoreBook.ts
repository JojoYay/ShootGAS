import { DensukeUtil } from './densukeUtil';
import { TotalScore } from './totalScore';
import { GasProps } from './gasProps';
import { ScriptProps } from './scriptProps';

export enum Title {
  TOKUTEN = '得点王ランキング',
  ASSIST = 'アシスト王ランキング',
  OKAMOTO = '岡本カップ',
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function testaaaaa() {
  const scoreBook: ScoreBook = new ScoreBook();
  // scoreBook.makeEventFormat();
  scoreBook.aggregateScore();
}

export class ScoreBook {
  public aggregateScore(): void {
    const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    const eventDetails: GoogleAppsScript.Spreadsheet.Sheet[] = eventSS.getSheets();
    const totalResult: GoogleAppsScript.Spreadsheet.Sheet = this.getTotalSheet(eventDetails);
    const eventSheet: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDataSheet(eventDetails);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const eventSheetVal: any[][] = eventSheet.getDataRange().getValues();
    const dataList: TotalScore[] = [];
    for (const sheet of eventDetails) {
      if (sheet.getSheetName() === 'Total' || sheet.getSheetName() === 'EventData') {
        continue;
      }
      const allValues = sheet.getDataRange().getValues();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const eventRow: any[] | undefined = eventSheetVal.find(item => item[1] === sheet.getSheetName());
      if (!eventRow) {
        throw new Error('eventがないなんてことはない');
      }
      // console.log(eventRow);
      for (const allValueRow of allValues) {
        console.log(allValueRow);
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
        }
        totalScore.playTime++;
        if (eventRow[4] === '晴れ') {
          totalScore.sunnyPlay++;
        } else if (eventRow[4] === '雨') {
          totalScore.rainyPlay++;
        }
        if (eventRow[5] === totalScore.name) {
          totalScore.mipCount++;
        }
        if (allValueRow[1]) {
          totalScore.teamPoint += totalScore.fetchTeamPoint(eventRow, allValueRow[1]);
          if (totalScore.isTopTeam(eventRow, allValueRow[1])) {
            totalScore.winCount++;
          }
          if (totalScore.fetchTeamPoint(eventRow, allValueRow[1]) === 0) {
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
    // console.log(dataList);
    const lastRow: number = totalResult.getLastRow();
    if (lastRow > 2) {
      totalResult.deleteRows(2, lastRow - 1);
    }
    for (const score of dataList) {
      totalResult.appendRow([
        '',
        score.name,
        score.playTime,
        score.sunnyPlay,
        score.rainyPlay,
        '',
        score.goalCount,
        '',
        score.assistCount,
        score.mipCount,
        score.teamPoint,
        score.winCount,
        score.loseCount,
      ]);
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private getEventRow(eventSheetVal: any[][], actDate: string) {
    return eventSheetVal[
      eventSheetVal.findIndex(item => {
        item[1] === actDate;
      })
    ];
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

  public makeEventFormat(): void {
    const densukeUtil: DensukeUtil = new DensukeUtil();
    const $ = densukeUtil.getDensukeCheerio();
    const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
    const members = densukeUtil.extractMembers($);
    const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);

    const eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.eventResults);
    this.updateEventSheet(actDate, attendees);

    const eventDetail: GoogleAppsScript.Spreadsheet.Sheet = this.getEventDetailSheet(eventSS, actDate);
    this.updateAttendeeName(eventDetail, attendees);
    this.updateEventDetails(eventDetail);
  }

  private getEventDetailSheet(eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet, actDate: string): GoogleAppsScript.Spreadsheet.Sheet {
    let eventDetail: GoogleAppsScript.Spreadsheet.Sheet | null = eventSS.getSheetByName(actDate);

    if (!eventDetail) {
      eventDetail = eventSS.insertSheet(actDate);
      eventDetail.appendRow(['名前', 'チーム', '得点', 'アシスト']);
      this.moveSheetToHead(eventDetail, eventSS);
    }
    return eventDetail;
  }

  private updateAttendeeName(eventDetail: GoogleAppsScript.Spreadsheet.Sheet, attendees: string[]) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allDetails: any[][] = eventDetail.getDataRange().getValues();
    for (let i = allDetails.length - 1; i >= 1; i--) {
      const name = allDetails[i][0];
      if (!name.includes(attendees) && !allDetails[i][1] && !allDetails[i][2] && !allDetails[i][3] && !allDetails[i][4]) {
        eventDetail.deleteRow(i + 1);
      }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allLoggedAttendees: any[][] = eventDetail.getDataRange().getValues();
    for (let i = 0; i < attendees.length; i++) {
      let find = false;
      for (let j = 1; j < allLoggedAttendees.length; j++) {
        if (allLoggedAttendees[j][0] === attendees[i]) {
          find = true;
          break;
        }
      }
      if (!find) {
        eventDetail.appendRow([attendees[i]]);
      }
    }
  }

  private updateEventDetails(eventDetail: GoogleAppsScript.Spreadsheet.Sheet) {
    const teamName: string[] = ['チーム1', 'チーム2', 'チーム3', 'チーム4', 'チーム5', 'チーム6'];
    const teamNameVal = SpreadsheetApp.newDataValidation().requireValueInList(teamName).build();
    // const teamPoint: string[] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
    // const teamPointVal = SpreadsheetApp.newDataValidation().requireValueInList(teamPoint).build();
    const goalCount: string[] = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
    const goalCountVal = SpreadsheetApp.newDataValidation().requireValueInList(goalCount).build();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allNewAttendees: any[][] = eventDetail.getDataRange().getValues();
    for (let i = 1; i < allNewAttendees.length; i++) {
      eventDetail.getRange(i + 1, 2).setDataValidation(teamNameVal);
      eventDetail.getRange(i + 1, 3).setDataValidation(goalCountVal);
      eventDetail.getRange(i + 1, 4).setDataValidation(goalCountVal);
    }
  }

  private moveSheetToHead(sheet: GoogleAppsScript.Spreadsheet.Sheet, eventSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    sheet.activate();
    eventSS.moveActiveSheet(3);
  }

  private updateEventSheet(actDate: string, attendees: string[]) {
    const eventData: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventDataSheet;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rangeValues: any[][] = eventData.getDataRange().getValues();
    const now: Date = new Date();
    const rowNumber: number = this.getTargetCol(rangeValues, actDate);
    if (rowNumber === -1) {
      this.createInitialEvent(attendees, eventData, now, actDate);
    } else {
      eventData.getRange(rowNumber + 1, 1).setValue(now);
      eventData.getRange(rowNumber + 1, 3).setValue(attendees.length);
      eventData.getRange(rowNumber + 1, 4).setValue(attendees.join(', '));
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

  private createInitialEvent(attendees: string[], eventData: GoogleAppsScript.Spreadsheet.Sheet, now: Date, actDate: string) {
    const teamPoint: string[] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
    const teamPointVal = SpreadsheetApp.newDataValidation().requireValueInList(teamPoint).build();
    const weather: string[] = ['晴れ', '曇り', '雨'];
    const weatherVal = SpreadsheetApp.newDataValidation().requireValueInList(weather).build();
    const attendVal = SpreadsheetApp.newDataValidation().requireValueInList(attendees).build();
    const lastCol: number = eventData.getLastColumn();
    eventData.appendRow([now, actDate, attendees.length, attendees.join(',')]);
    const lastRow: number = eventData.getLastRow();

    eventData.getRange(lastRow, 5).setDataValidation(weatherVal);
    eventData.getRange(lastRow, 6).setDataValidation(attendVal);

    eventData.getRange(lastRow, 8).setDataValidation(teamPointVal);
    eventData.getRange(lastRow, 9).setDataValidation(teamPointVal);
    eventData.getRange(lastRow, 10).setDataValidation(teamPointVal);
    eventData.getRange(lastRow, 11).setDataValidation(teamPointVal);
    eventData.getRange(lastRow, 12).setDataValidation(teamPointVal);
    eventData.getRange(lastRow, 13).setDataValidation(teamPointVal);

    if (lastRow > 2) {
      const newRowRange: GoogleAppsScript.Spreadsheet.Range = eventData.getRange(lastRow, 1, 1, lastCol - 1);
      eventData.moveRows(newRowRange, 2);
    }
  }

  public generateScoreBook(actDate: string, attendees: string[], title: Title): void {
    const reportSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
    const goalCount: string[] = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
    const goalCountVal = SpreadsheetApp.newDataValidation().requireValueInList(goalCount).build();
    let scoreSheet: GoogleAppsScript.Spreadsheet.Sheet | null = reportSS.getSheetByName(title);
    if (!scoreSheet) {
      scoreSheet = reportSS.insertSheet(title);
      // scoreSheet.appendRow(['伝助名称', 'line名称', '合計得点', actDate]);
      scoreSheet.appendRow(['伝助名称', '合計得点']);
      scoreSheet.insertRowBefore(1);
    }
    if (!this.isActDateExists(actDate, scoreSheet)) {
      scoreSheet.insertColumnBefore(3);
      scoreSheet.getRange('C2').setValue(actDate);
    }
    this.addAttendee(scoreSheet, attendees, true);

    const newLocal_1 = scoreSheet.getLastRow();
    const lastCol = scoreSheet.getLastColumn();
    const allVals = scoreSheet.getDataRange().getValues();
    console.log(newLocal_1);
    for (let i = 3; i <= newLocal_1; i++) {
      console.log('excel上:' + allVals[i - 1][0]);
      console.log(attendees);
      scoreSheet.getRange(i, 3).setDataValidation(goalCountVal);
      const range = scoreSheet.getRange(i, 3, 1, lastCol - 2);
      const formula = `=SUM(${range.getA1Notation()})`;
      scoreSheet.getRange(i, 2).setFormula(formula);
    }

    const finalRow = scoreSheet.getLastRow();
    const finalCol = scoreSheet.getLastColumn();
    scoreSheet.getRange(2, 1, finalRow - 1, finalCol).setBorder(true, true, true, true, true, true);
    scoreSheet.getRange(3, 1, finalRow - 1, finalCol).sort({ column: 2, ascending: false });

    scoreSheet.activate();
    reportSS.moveActiveSheet(1);
  }

  private addAttendee(scoreSheet: GoogleAppsScript.Spreadsheet.Sheet, attendees: string[], removeZero: boolean) {
    if (removeZero) {
      this.removeZeroPpl(scoreSheet);
    }
    //シートがある場合は
    // const lastRow = scoreSheet.getLastRow();
    const allDataValues = scoreSheet.getDataRange().getValues();
    console.log(allDataValues);
    for (let i = 0; i < attendees.length; i++) {
      let find = false;
      for (let j: number = 0; j < allDataValues.length; j++) {
        if (allDataValues[j][1] === attendees[i]) {
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

  private removeZeroPpl(scoreSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const values = scoreSheet.getDataRange().getValues();
    console.log(values);
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][1] === 0) {
        scoreSheet.deleteRow(i + 1);
      }
    }
  }

  private isActDateExists(actDate: string, scoreSheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
    return actDate === scoreSheet.getRange('C2').getValue();
  }
}