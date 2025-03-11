import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { ScriptProps } from './scriptProps';
const gasUtil: GasUtil = new GasUtil();

export class SchedulerUtil {
    public get schedulerUrl(): string {
        return 'https://shootsundayfront.web.app/calendar';
    }

    public get calendarSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('calendar');
        if (!calendarSheet) {
            console.error('シート "calendar" が見つかりません。');
            throw new Error('シート "calendar" が見つかりません。');
        }
        return calendarSheet;
    }

    public get attendanceSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const attendanceSheet: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('attendance');
        if (!attendanceSheet) {
            console.error('シート "attendance" が見つかりません。');
            throw new Error('シート "attendance" が見つかりません。');
        }
        return attendanceSheet;
    }

    //集計対象イベントのAttendeesを取っている
    public extractAttendees(symbol: string): string[] {
        const attendanceSheet: GoogleAppsScript.Spreadsheet.Sheet = this.attendanceSheet;
        const aValues = attendanceSheet.getDataRange().getValues();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const cValues = calendarSheet.getDataRange().getValues();

        const attend: string[] = [];
        // event_status=20 のイベントを探す
        for (let i = 1; i < cValues.length; i++) {
            // 1行目はヘッダーなのでスキップ
            const event = cValues[i];
            if (!event || event.length < 8 || event[7] !== 20) continue; // データが不足している or event_status が 20 でない場合はスキップ

            const targetCalendarId = event[0]; // calendar_id (1列目)
            // attendanceSheetから該当calendar_idとsymbolに一致するuser_idを抽出
            for (let j = 1; j < aValues.length; j++) {
                // 1行目はヘッダーなのでスキップ
                const attendance = aValues[j];
                if (!attendance || attendance.length < 7) continue; // データが不足している場合はスキップ

                const aCalendarId = attendance[6]; // calendar_id (7列目)
                const status = attendance[5]; // status (6列目)
                const userId = attendance[1]; // user_id (2列目)
                console.log(aCalendarId === targetCalendarId);
                console.log(status === symbol);
                if (aCalendarId === targetCalendarId && status === symbol) {
                    attend.push(userId);
                }
            }
            // event_status=20 のイベントは複数存在しない前提なので、最初に見つかった時点でループを抜ける
            break;
        }
        // mappingSheetを利用してuserIdの配列を伝助上の名称の配列に変換
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const mappingValues = mappingSheet.getDataRange().getValues();
        const userIdToDensukeNameMap: { [key: string]: string } = {};
        // mappingSheetからuserIdと伝助上の名前のマッピングを作成
        for (let i = 1; i < mappingValues.length; i++) {
            const row = mappingValues[i];
            const userId = row[2]; // LINE ID (3列目)
            const densukeName = row[1]; // 伝助上の名前 (2列目)
            if (userId && densukeName) {
                userIdToDensukeNameMap[userId] = densukeName;
            }
        }
        // userIdの配列を伝助上の名称の配列に変換
        const densukeNames = attend.map(userId => userIdToDensukeNameMap[userId] || userId);
        // console.log(densukeNames);
        return densukeNames;
    }

    public extractDateFromRownum(): string {
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const cValues = calendarSheet.getDataRange().getValues();
        let dateStr = '';
        // event_status=20 のイベントを探す
        for (let i = 1; i < cValues.length; i++) {
            // 1行目はヘッダーなのでスキップ
            const event = cValues[i];
            if (!event || event.length < 8 || event[7] !== 20) continue; // データが不足している or event_status が 20 でない場合はスキップ
            const date = new Date(event[3]);
            dateStr = event[2] + '(' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd MMM') + ')'; // calendar_id (1列目)
            // event_status=20 のイベントは複数存在しない前提なので、最初に見つかった時点でループを抜ける
            break;
        }
        if (dateStr === '') {
            throw new Error('event_status=20のデータが見つかりません');
        }
        return dateStr;
    }

    public generateRemind(): string {
        const attendees: string[] = this.extractAttendees('〇');
        const unknown: string[] = this.extractAttendees('△');
        const actDate: string = this.extractDateFromRownum();

        let remindStr: string =
            '次回予定' +
            actDate +
            'リマインドです！\nスケジューラーの更新お忘れなく！\nThis is gentle reminder of ' +
            actDate +
            '.\nPlease update your schedule.\n';
        if (attendees.length < 10) {
            remindStr =
                '次回予定' +
                actDate +
                'がピンチです！\n参加できる方、ぜひ参加表明お願いします！！！\nThis is gentle reminder of ' +
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
            this.schedulerUrl;
        return summary;
    }

    public generateSummaryBase(attendees: string[], actDate: string): void {
        // スプレッドシートとシートを取得
        const settingSheet = GasProps.instance.settingSheet;
        const cashBook = GasProps.instance.cashBookSheet;
        // const attendees: string[] = this.extractAttendees('〇');
        // const actDate: string = this.extractDateFromRownum();
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
        this.reCalcTotalVal(cashBook);
        for (let i = 0; i < attendees.length; i++) {
            const lineName = gasUtil.getLineName(attendees[i]);
            report.appendRow([attendees[i], lineName]);
            const paymentUrl: GoogleAppsScript.Spreadsheet.RichTextValue | null = gasUtil.getPaymentUrl(attendees[i], actDate);
            const lastRow = report.getLastRow();
            if (paymentUrl) {
                report.getRange(lastRow, 3).setRichTextValue(paymentUrl);
            }
        }
        report.setColumnWidth(1, 170);
        report.setColumnWidth(2, 200);
        report.getRange(1, 1, 7, 2).setBorder(true, true, true, true, true, true);
        report.getRange(1, 1, 7, 1).setBackground('#fff2cc');

        const rlastRow = report.getLastRow();
        report.getRange(9, 1, rlastRow - 8, 3).setBorder(true, true, true, true, true, true);
        report.getRange(9, 1, 1, 3).setBackground('#fff2cc');

        cashBook.appendRow([dd, actDate, '参加費(' + attendees.length + '名)', '' + attendFeeTotal, '=' + 'E' + lastRow + '+D' + (lastRow + 1)]);
        cashBook.appendRow([dd, actDate, 'ピッチ使用料金', '-' + rentalFee, '=' + 'E' + (lastRow + 1) + '+D' + (lastRow + 2)]);
        const clastRow = cashBook.getLastRow();
        cashBook.getRange(1, 1, clastRow, 5).setBorder(true, true, true, true, true, true);
        this.reCalcTotalVal(cashBook);
    }

    private reCalcTotalVal(cashBook: GoogleAppsScript.Spreadsheet.Sheet) {
        const allData = cashBook.getDataRange().getValues();
        let index = 1;
        for (const allRow of allData) {
            if (allRow[3] && allRow[3] !== '金額(SGD)') {
                const formula = `E${index - 1}+D${index}`;
                cashBook.getRange(index, 5).setFormula(formula);
            }
            index++;
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

    public hasMultipleOccurrences(array: string[], searchString: string): boolean {
        let count = 0;
        for (const item of array) {
            if (item === searchString) {
                count++;
            }
            if (count >= 2) {
                return true;
            }
        }
        return false;
    }
}
