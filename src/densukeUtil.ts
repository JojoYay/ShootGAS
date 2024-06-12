import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
// import { ScoreBook, Title } from './scoreBook';
import { ScriptProps } from './scriptProps';
const gasUtil: GasUtil = new GasUtil();

export class DensukeUtil {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public getDensukeCheerio(): any {
        const url: string = this.getDensukeUrl();
        const html: string = UrlFetchApp.fetch(url).getContentText();
        // @ts-ignore
        const $ = Cheerio.load(html);
        return $;
    }

    public getDensukeUrl(): string {
        const settingSheet = GasProps.instance.settingSheet;
        let url: string = '';
        if (settingSheet) {
            url = settingSheet.getRange('B1').getValue();
        }
        return url;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public extractMembers($: any = this.getDensukeCheerio()): string[] {
        const data: string[] = [];
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        $('td a').each((i: number, element: unknown) => {
            const text = $(element).text();
            const href = $(element).attr('href');
            if (href && href.startsWith('javascript:memberdata(')) {
                data.push(text.trim());
            }
        });
        return data;
    }

    public extractAttendees(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        $: any = this.getDensukeCheerio(),
        rowNum: number,
        symbol: string,
        members: string[]
    ): string[] {
        const row = $(`#listtable tr`).eq(rowNum);
        const attend: string[] = [];
        row.find('td')
            .slice(4)
            .each((i: number, element: unknown) => {
                const text = $(element).text();
                if (text === symbol) {
                    attend.push(members[i]);
                }
            });
        return attend;
    }

    public extractDateFromRownum(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        $: any = this.getDensukeCheerio(),
        rowNum: number
    ): string {
        const row = $(`#listtable tr`).eq(rowNum);
        const cell = row.find('td[nowrap]').first(); // 最初の<td nowrap="">を取得
        return cell.text();
    }

    public generateRemind($ = this.getDensukeCheerio()): string {
        const members: string[] = this.extractMembers($);
        const attendees: string[] = this.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);
        const unknown: string[] = this.extractAttendees($, ScriptProps.instance.ROWNUM, '△', members);
        const actDate: string = this.extractDateFromRownum($, ScriptProps.instance.ROWNUM);

        let remindStr: string =
            '次回予定' +
            actDate +
            'リマインドです！\n伝助の更新お忘れなく！\nThis is gentle reminder of ' +
            actDate +
            '.\nPlease update your Densuke schedule.\n';
        if (attendees.length < 10) {
            remindStr =
                '次回予定' +
                actDate +
                'がピンチです！\n参加できる方、ぜひ伝助で参加表明お願いします！！！\nThis is gentle reminder of ' +
                actDate +
                '.\nDue to the low number of participants, there is a possibility of cancellation.\nPlease join us!\n';
        }
        const summary: string =
            remindStr +
            '〇(' +
            attendees.length +
            '名): ' +
            attendees.join(', ') +
            '\n' +
            '△(' +
            unknown.length +
            '名): ' +
            unknown.join(', ') +
            '\n' +
            '伝助URL：' +
            this.getDensukeUrl();
        return summary;
    }

    public generateSummaryBase($ = this.getDensukeCheerio()): void {
        // スプレッドシートとシートを取得
        const settingSheet = GasProps.instance.settingSheet;
        const cashBook = GasProps.instance.cashBookSheet;
        const members: string[] = this.extractMembers($);
        const attendees: string[] = this.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);
        const actDate: string = this.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        // データの範囲を取得
        const cRangeValues = cashBook.getDataRange().getValues();
        // 2カラム目（B列）に指定した値がある行を逆順に削除
        for (let i = cRangeValues.length - 1; i >= 0; i--) {
            if (cRangeValues[i][1] === actDate) {
                // B列はインデックス1
                cashBook.deleteRow(i + 1);
            }
        }
        const lastRow: number = cashBook.getLastRow();
        const orgPrice: number = cashBook.getRange(lastRow, 5).getValue();
        const rentalFee: number = settingSheet.getRange('B3').getValue();
        const attendFee: number = settingSheet.getRange('B4').getValue();
        gasUtil.archiveFiles(actDate);
        const attendFeeTotal: number = attendFee * attendees.length;
        const report: GoogleAppsScript.Spreadsheet.Sheet = gasUtil.getReportSheet(actDate, true); //ない場合作る
        const dd: string = new Date().toLocaleString();
        report.getRange('A1').setValue('日付');
        report.getRange('B1').setValue(actDate);
        report.getRange('A2').setValue('更新日付');
        report.getRange('B2').setValue(dd);
        report.getRange('A3').setValue('繰り越し残高(SGD)');
        report.getRange('B3').setValue('' + orgPrice);
        report.getRange('A4').setValue('参加人数(人)');
        report.getRange('B4').setValue('' + attendees.length);
        report.getRange('A5').setValue('参加費合計(SGD))');
        report.getRange('B5').setValue('' + attendFeeTotal);
        report.getRange('A6').setValue('ピッチ使用料金(SGD)');
        report.getRange('B6').setValue('' + rentalFee);
        report.getRange('A7').setValue('余剰金残高(SGD)');
        report.getRange('B7').setValue('' + (orgPrice - rentalFee + attendFeeTotal));

        report.getRange('A9').setValue('参加者（伝助名称）');
        report.getRange('B9').setValue('参加者（Line名称）');
        report.getRange('C9').setValue('支払い状況');

        const values = report.getDataRange().getValues();
        for (let i = values.length; i >= 10; i--) {
            report.deleteRow(i);
        }
        for (let i = 0; i < attendees.length; i++) {
            const lineName = gasUtil.getLineName(attendees[i]);
            report.appendRow([attendees[i], lineName]);
            const paymentUrl: GoogleAppsScript.Spreadsheet.RichTextValue | null = gasUtil.getPaymentUrl(attendees[i], actDate);
            const lastRow = report.getLastRow();
            if (paymentUrl) {
                report.getRange(lastRow, 3).setRichTextValue(paymentUrl);
            }
            // report.appendRow([attendees[i], lineName, gasUtil.getPaymentUrl(attendees[i], actDate)]);
        }
        report.setColumnWidth(1, 170);
        report.setColumnWidth(2, 200);
        report.getRange(1, 1, 7, 2).setBorder(true, true, true, true, true, true);
        report.getRange(1, 1, 7, 1).setBackground('#fff2cc');

        const rlastRow = report.getLastRow();
        report.getRange(9, 1, rlastRow - 8, 3).setBorder(true, true, true, true, true, true);
        report.getRange(9, 1, 1, 3).setBackground('#fff2cc');

        const attendOrg = orgPrice + attendFeeTotal;
        if (cashBook) {
            cashBook.appendRow([dd, actDate, '参加費(' + attendees.length + '名)', '' + attendFeeTotal, '' + attendOrg]);
            cashBook.appendRow([dd, actDate, 'ピッチ使用料金', '-' + rentalFee, '' + (orgPrice - rentalFee + attendFeeTotal)]);
            const clastRow = cashBook.getLastRow();
            // console.log('lastRow:' + clastRow);
            cashBook.getRange(1, 1, 1, 5).setBackground('#fff2cc');
            cashBook.getRange(1, 1, clastRow, 5).setBorder(true, true, true, true, true, true);
        } else {
            throw new Error('Cash Book not found');
        }
        this.copySheetInSpreadsheet();

        // const scoreBook: ScoreBook = new ScoreBook();
        // scoreBook.generateScoreBook(actDate, attendees, Title.ASSIST);
        // scoreBook.generateScoreBook(actDate, attendees, Title.TOKUTEN);
    }

    private copySheetInSpreadsheet(): void {
        const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.cashBookSheet;

        const oldCashBook = spreadsheet.getSheetByName('CashBook');
        if (oldCashBook) {
            spreadsheet.deleteSheet(oldCashBook);
        }

        if (sheet) {
            const newSheet = sheet.copyTo(spreadsheet);
            newSheet.setName('CashBook');
            newSheet.activate();
            spreadsheet.moveActiveSheet(1);
        }
    }

    public getSummaryStr(attendees: string[], actDate: string, payNowAddy: string): string {
        let paynowStr = '';
        if (gasUtil.getUnpaid(actDate).length === 0) {
            paynowStr =
                '入金ありがとうございました。今回のレポートになります。詳細はリンクをご確認下さい。\nThank you for your payment.\nPlease find the report for this transaction below.\nFor more details, please check the provided link.\n';
        } else {
            paynowStr =
                'みなさま、ご参加ありがとうございました。\n入金後PayNowのスクリーンショットをSundayShootちゃんねるに送信して下さい。\nThank you all for your paticipation! After making the payment, please send the PayNow screenshot to Sunday Shoot Line Channel.\n';
        }
        const summary =
            paynowStr +
            '[' +
            actDate +
            ']のReport\n' +
            '参加者 participants (' +
            attendees.length +
            '名): ' +
            attendees.join(', ') +
            '\n' +
            'Report URL:' +
            GasProps.instance.reportSheetUrl +
            '\nPayNow先: ' +
            payNowAddy;
        return summary;
    }

    // private archiveFiles(actDate: string): void {
    //   const folderProp = ScriptProps.instance.folderId;
    //   const archiveProp = ScriptProps.instance.archiveFolder;
    //   const sourceFolder = DriveApp.getFolderById(folderProp);
    //   const destinationFolder = DriveApp.getFolderById(archiveProp);
    //   const files = sourceFolder.getFiles();
    //   const prefix = actDate + '_';
    //   while (files.hasNext()) {
    //     const file = files.next();
    //     if (!file.getName().startsWith(prefix)) {
    //       file.moveTo(destinationFolder);
    //     }
    //   }
    // }

    // private getLineName(densukeName: string) {
    //   let lineName = null;
    //   const mappingSheet = GasProps.instance.mappingSheet;
    //   const values = mappingSheet.getDataRange().getValues();
    //   for (let i = values.length - 1; i >= 0; i--) {
    //     if (values[i][1] === densukeName) {
    //       lineName = values[i][0];
    //       break;
    //     }
    //   }
    //   return lineName;
    // }
}
