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

  public uploadPayNowPic(lineName: string, messageId: string, actDate: string): void {
    const fileNm = actDate + '_' + lineName;
    const folder = GasProps.instance.payNowFolder;
    const files = folder.getFilesByName(fileNm);
    if (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
    }
    lineUtil.getLineImage(messageId, fileNm);
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

  public updatePaymentStatus(lineName: string, actDate: string): void {
    const repo = this.getReportSheet(actDate, false);
    const values = repo.getDataRange().getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][1] === lineName) {
        repo.getRange(i + 1, 3).setValue(this.getPaymentUrl(lineName, actDate));
        break;
      }
    }
  }

  public getPaymentUrl(lineName: string, actDate: string) {
    const payNowOwner = this.getPaynowOwner();
    if (payNowOwner === lineName) {
      return 'PayNow口座主';
    }
    return this.getFileUrlInFolder(actDate, lineName);
  }

  public getPaynowOwner(): string {
    const settingSheet = GasProps.instance.settingSheet;
    const payNowOwner = settingSheet.getRange('B6').getValue();
    return payNowOwner;
  }

  private getFileUrlInFolder(actDate: string, lineName: string) {
    if (!lineName) {
      return '';
    }
    const folderProp = ScriptProps.instance.folderId;
    const folder = DriveApp.getFolderById(folderProp);
    const fileName = actDate + '_' + lineName;
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      return file.getUrl();
    } else {
      return '';
    }
  }
}
