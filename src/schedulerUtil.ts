import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { ScriptProps } from './scriptProps';
import { AttendeeInfo } from './Types/AttendeeInfo';
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

    public extractAttendeeInfo(symbol: string): AttendeeInfo[] {
        const attendanceSheet: GoogleAppsScript.Spreadsheet.Sheet = this.attendanceSheet;
        const aValues = attendanceSheet.getDataRange().getValues();
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const cValues = calendarSheet.getDataRange().getValues();

        // mappingSheetを利用してuserIdと伝助上の名前のマッピングを作成
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const mappingValues = mappingSheet.getDataRange().getValues();
        const userIdToNameMap: { [key: string]: { densukeName: string; lineName: string } } = {};
        for (let i = 1; i < mappingValues.length; i++) {
            const row = mappingValues[i];
            const userId = row[2]; // LINE ID (3列目)
            const densukeName = row[1]; // 伝助上の名前 (2列目)
            const lineName = row[0]; // LINE名称 (1列目)
            if (userId && densukeName) {
                userIdToNameMap[userId] = { densukeName, lineName };
            }
        }

        const attend: AttendeeInfo[] = [];
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
                // console.log('attendance', attendance);
                if (!attendance || attendance.length < 7) continue; // データが不足している場合はスキップ

                const aCalendarId = attendance[6]; // calendar_id (7列目)
                const status = attendance[5]; // status (6列目)
                const userId = attendance[1]; // user_id (2列目)
                if (aCalendarId === targetCalendarId && status === symbol) {
                    // マッピングから名前情報を取得
                    const nameInfo = userIdToNameMap[userId] || { densukeName: userId, lineName: userId };

                    // オブジェクトとして返す
                    attend.push({
                        userId: userId,
                        userName: nameInfo,
                        adult: attendance[7] || 1, // 実際のデータから取得する必要があります
                        child: attendance[8] || 0, // 実際のデータから取得する必要があります
                    });
                }
            }
            // event_status=20 のイベントは複数存在しない前提なので、最初に見つかった時点でループを抜ける
            break;
        }
        return attend;
    }

    // public extractAttendeeUserIds(symbol: string): string[] {
    //     const attendanceSheet: GoogleAppsScript.Spreadsheet.Sheet = this.attendanceSheet;
    //     const aValues = attendanceSheet.getDataRange().getValues();
    //     const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
    //     const cValues = calendarSheet.getDataRange().getValues();

    //     const attend: string[] = [];
    //     // event_status=20 のイベントを探す
    //     for (let i = 1; i < cValues.length; i++) {
    //         // 1行目はヘッダーなのでスキップ
    //         const event = cValues[i];
    //         if (!event || event.length < 8 || event[7] !== 20) continue; // データが不足している or event_status が 20 でない場合はスキップ

    //         const targetCalendarId = event[0]; // calendar_id (1列目)
    //         // attendanceSheetから該当calendar_idとsymbolに一致するuser_idを抽出
    //         for (let j = 1; j < aValues.length; j++) {
    //             // 1行目はヘッダーなのでスキップ
    //             const attendance = aValues[j];
    //             if (!attendance || attendance.length < 7) continue; // データが不足している場合はスキップ

    //             const aCalendarId = attendance[6]; // calendar_id (7列目)
    //             const status = attendance[5]; // status (6列目)
    //             const userId = attendance[1]; // user_id (2列目)
    //             console.log(aCalendarId === targetCalendarId);
    //             console.log(status === symbol);
    //             if (aCalendarId === targetCalendarId && status === symbol) {
    //                 attend.push(userId);
    //             }
    //         }
    //         // event_status=20 のイベントは複数存在しない前提なので、最初に見つかった時点でループを抜ける
    //         break;
    //     }
    //     return attend;
    // }

    //集計対象イベントのAttendeesを取っている
    public extractAttendees(symbol: string): string[] {
        const attendeeInfos: AttendeeInfo[] = this.extractAttendeeInfo(symbol);
        // userNameのdensukeNameを直接取得
        const densukeNames = attendeeInfos.map(info => info.userName.densukeName);
        return densukeNames;
    }

    // 参加者を大人と子供を含めて抽出するメソッド
    public extractPlayers(noGuest: boolean = false): string[] {
        const attendeeInfos: AttendeeInfo[] = this.extractAttendeeInfo('〇');
        const players: string[] = [];

        for (let i = 0; i < attendeeInfos.length; i++) {
            const attendeeInfo = attendeeInfos[i];
            const densukeName = attendeeInfo.userName.densukeName;

            // 大人の処理
            if (attendeeInfo.adult === 1) {
                // 大人が1人の場合はそのまま追加
                players.push(densukeName);
            } else if (attendeeInfo.adult >= 2) {
                // 大人が2人以上の場合は1人目はそのまま、2人目以降はゲスト表記
                players.push(densukeName); // 1人目
                if (!noGuest) {
                    for (let j = 1; j < attendeeInfo.adult; j++) {
                        players.push(densukeName + '_Guest' + j);
                    }
                }
            }

            // 子供の処理
            if (attendeeInfo.child >= 1) {
                for (let k = 0; k < attendeeInfo.child; k++) {
                    players.push(densukeName + '_Child' + (k + 1));
                }
            }
        }

        return players;
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
        const attendeeInfos: AttendeeInfo[] = this.extractAttendeeInfo('〇');
        const attendeeUserIds: string[] = attendeeInfos.map(info => info.userId);
        const actDate: string = this.extractDateFromRownum();

        // 大人の総数とゲスト数を計算
        const totalAdults: number = attendeeInfos.reduce((sum, info) => sum + info.adult, 0);
        const uniqueUserIds = new Set(attendeeUserIds).size;
        const guestCount = totalAdults - uniqueUserIds;

        // 参加者名の配列を作成（大人が2人以上の場合はゲスト表記を追加）
        const attendeeNames: string[] = [];
        for (let i = 0; i < attendeeInfos.length; i++) {
            const attendeeInfo = attendeeInfos[i];
            const attendeeName = attendeeInfo.userName.densukeName;

            if (attendeeInfo.adult === 1) {
                attendeeNames.push(attendeeName);
            } else {
                // 大人が2人以上の場合はゲスト表記を追加
                for (let j = 0; j < attendeeInfo.adult; j++) {
                    if (j === 0) {
                        attendeeNames.push(attendeeName);
                    } else {
                        attendeeNames.push(attendeeName + '_Guest' + j);
                    }
                }
            }
        }

        // △（未定）の参加者も同様に処理
        const unknownInfos: AttendeeInfo[] = this.extractAttendeeInfo('△');
        // const unknownUserIds: string[] = unknownInfos.map(info => info.userId);
        const unknownTotalAdults: number = unknownInfos.reduce((sum, info) => sum + info.adult, 0);

        const unknownNames: string[] = [];
        for (let i = 0; i < unknownInfos.length; i++) {
            const attendeeInfo = unknownInfos[i];
            const attendeeName = attendeeInfo.userName.densukeName;

            if (attendeeInfo.adult === 1) {
                unknownNames.push(attendeeName);
            } else {
                // 大人が2人以上の場合はゲスト表記を追加
                for (let j = 0; j < attendeeInfo.adult; j++) {
                    if (j === 0) {
                        unknownNames.push(attendeeName);
                    } else {
                        unknownNames.push(attendeeName + '_Guest' + j);
                    }
                }
            }
        }

        let remindStr: string =
            '次回予定' +
            actDate +
            'リマインドです！\nスケジューラーの更新お忘れなく！\nThis is gentle reminder of ' +
            actDate +
            '.\nPlease update your schedule.\n';
        if (totalAdults < 10) {
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
            totalAdults +
            '名、うちゲスト' +
            guestCount +
            '名): ' +
            attendeeNames.join(', ') +
            '\n' +
            '△(' +
            unknownTotalAdults +
            '名): ' +
            unknownNames.join(', ') +
            '\n' +
            'スケジューラーURL：' +
            this.schedulerUrl;
        return summary;
    }

    public generateSummaryBase(): void {
        const cashBook = GasProps.instance.cashBookSheet;
        // const attendees: string[] = this.extractAttendees('〇');
        const attendeeInfos: AttendeeInfo[] = this.extractAttendeeInfo('〇');
        console.log('attendeeInfos', attendeeInfos);
        const attendeeUserIds: string[] = attendeeInfos.map(info => info.userId);
        const actDate: string = this.extractDateFromRownum();
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
        const rentalFee: number = this.getPitchFee();
        const attendFee: number = this.getPaticipationFee();
        gasUtil.archiveFiles(actDate);

        // adultのカウント分で参加費合計を計算
        const totalAdults: number = attendeeInfos.reduce((sum, info) => sum + info.adult, 0);
        const attendFeeTotal: number = attendFee * totalAdults;

        const report: GoogleAppsScript.Spreadsheet.Sheet = gasUtil.getReportSheet(actDate, true); //ない場合作る
        const dd: string = new Date().toLocaleString();
        report.getRange('A1').setValue('日付');
        report.getRange('B1').setValue(actDate);
        report.getRange('A2').setValue('更新日付');
        report.getRange('B2').setValue(dd);
        report.getRange('A3').setValue('繰り越し残高(SGD)');
        report.getRange('B3').setValue('' + orgPrice);
        report.getRange('A4').setValue('参加人数(人)');
        report.getRange('B4').setValue('' + totalAdults);
        report.getRange('A5').setValue('参加費合計(SGD))');
        report.getRange('B5').setValue('' + attendFeeTotal);
        report.getRange('A6').setValue('ピッチ使用料金(SGD)');
        report.getRange('B6').setValue('' + rentalFee);
        report.getRange('A7').setValue('余剰金残高(SGD)');
        report.getRange('B7').setValue('' + (orgPrice - rentalFee + attendFeeTotal));

        report.getRange('A9').setValue('参加者（スケジューラ名称）');
        report.getRange('B9').setValue('参加者（Line名称）');
        report.getRange('C9').setValue('支払い状況');
        report.getRange('D9').setValue('金額(SGD)');

        const values = report.getDataRange().getValues();
        for (let i = values.length; i >= 10; i--) {
            report.deleteRow(i);
        }
        // this.reCalcTotalVal(cashBook);
        const mappingSheet = GasProps.instance.mappingSheet;
        const mapValues = mappingSheet.getDataRange().getValues();
        const userIdToDensukeNameMap: { [key: string]: [string, string] } = {};

        for (let i = 1; i < mapValues.length; i++) {
            const row = mapValues[i];
            const userId = row[2]; // LINE ID (3列目)
            const densukeName = row[1]; // 伝助上の名前 (2列目)
            const lineName = row[0]; // 伝助上の名前 (2列目)
            if (userId && densukeName) {
                userIdToDensukeNameMap[userId] = [densukeName, lineName];
            }
        }

        // attendeeInfosを直接使用して、userName情報を取得
        for (let i = 0; i < attendeeInfos.length; i++) {
            const attendeeInfo = attendeeInfos[i];
            const individualFee = attendFee * attendeeInfo.adult;

            report.appendRow([attendeeInfo.userName.densukeName, attendeeInfo.userName.lineName, '', individualFee]);
            const paymentUrl: GoogleAppsScript.Spreadsheet.RichTextValue | null = gasUtil.getPaymentUrl(attendeeInfo.userName.densukeName, actDate);
            const lastRow = report.getLastRow();
            if (paymentUrl) {
                report.getRange(lastRow, 3).setRichTextValue(paymentUrl);
            }
        }
        report.setColumnWidth(1, 170);
        report.setColumnWidth(2, 200);
        report.setColumnWidth(4, 100);
        report.getRange(1, 1, 7, 2).setBorder(true, true, true, true, true, true);
        report.getRange(1, 1, 7, 1).setBackground('#fff2cc');

        const rlastRow = report.getLastRow();
        report.getRange(9, 1, rlastRow - 8, 4).setBorder(true, true, true, true, true, true);
        report.getRange(9, 1, 1, 4).setBackground('#fff2cc');

        // ゲスト数の計算（大人の総数 - ユニークなuserIdの数）
        const uniqueUserIds = new Set(attendeeUserIds).size;
        const guestCount = totalAdults - uniqueUserIds;

        cashBook.appendRow([
            dd,
            actDate,
            '参加費 ' + totalAdults + '人分 （うちゲスト' + guestCount + '名）',
            '' + attendFeeTotal,
            '=IF(ROW()=5, INDEX(D:D, ROW()), INDEX(E:E, ROW()-1) + INDEX(D:D, ROW()))',
            // '=' + 'E' + lastRow + '+D' + (lastRow + 1),
        ]);
        // cashBook.appendRow([dd, actDate, 'ピッチ使用料金', '-' + rentalFee, '=' + 'E' + (lastRow + 1) + '+D' + (lastRow + 2)]);
        cashBook.appendRow([
            dd,
            actDate,
            'ピッチ使用料金',
            '-' + rentalFee,
            '=IF(ROW()=5, INDEX(D:D, ROW()), INDEX(E:E, ROW()-1) + INDEX(D:D, ROW()))',
        ]);
        const clastRow = cashBook.getLastRow();
        cashBook.getRange(1, 1, clastRow, 5).setBorder(true, true, true, true, true, true);
        // this.reCalcTotalVal(cashBook);
    }

    // private reCalcTotalVal(cashBook: GoogleAppsScript.Spreadsheet.Sheet) {
    //     const allData = cashBook.getDataRange().getValues();
    //     let index = 1;
    //     for (const allRow of allData) {
    //         if (allRow[3] && allRow[3] !== '金額(SGD)') {
    //             const formula = `E${index - 1}+D${index}`;
    //             cashBook.getRange(index, 5).setFormula(formula);
    //         }
    //         index++;
    //     }
    // }

    public getSummaryStr(): string {
        const attendeeInfos: AttendeeInfo[] = this.extractAttendeeInfo('〇');
        const attendeeUserIds: string[] = attendeeInfos.map(info => info.userId);
        const actDate: string = this.extractDateFromRownum();
        const payNowAddy: string = this.getPayNowAddress();
        const paticipationFee: string = this.getPaticipationFee();

        // 大人の総数とゲスト数を計算
        const totalAdults: number = attendeeInfos.reduce((sum, info) => sum + info.adult, 0);
        const uniqueUserIds = new Set(attendeeUserIds).size;
        const guestCount = totalAdults - uniqueUserIds;

        // 参加者名の配列を作成（大人が2人以上の場合はゲスト表記を追加）
        const attendeeNames: string[] = [];
        for (let i = 0; i < attendeeInfos.length; i++) {
            const attendeeInfo = attendeeInfos[i];
            const attendeeName = attendeeInfo.userName.densukeName;

            if (attendeeInfo.adult === 1) {
                attendeeNames.push(attendeeName);
            } else {
                // 大人が2人以上の場合はゲスト表記を追加
                for (let j = 0; j < attendeeInfo.adult; j++) {
                    if (j === 0) {
                        attendeeNames.push(attendeeName);
                    } else {
                        attendeeNames.push(attendeeName + '_Guest' + j);
                    }
                }
            }
        }

        let paynowStr = '';
        if (gasUtil.getUnpaid(actDate).length === 0) {
            paynowStr =
                '入金ありがとうございました。今回のレポートになります。詳細はリンクをご確認下さい。\nThank you for your payment.\nPlease find the report for this transaction below.\nFor more details, please check the provided link.\n';
        } else {
            paynowStr =
                'みなさま、ご参加ありがとうございました。\n$' +
                paticipationFee +
                '入金後PayNowのスクリーンショットをSundayShootちゃんねるに送信して下さい。\nThank you all for your paticipation! After making the payment, please send the PayNow screenshot to Sunday Shoot Line Channel.\n';
        }
        const summary =
            paynowStr +
            '[' +
            actDate +
            ']のReport\n' +
            '参加者 participants (' +
            totalAdults +
            '名、うちゲスト' +
            guestCount +
            '名): ' +
            attendeeNames.join(', ') +
            '\n' +
            'Report URL:' +
            GasProps.instance.reportSheetUrl +
            '\nPayNow先: ' +
            payNowAddy +
            '\n参加費: $' +
            paticipationFee;
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

    public getPitchFee() {
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const calendarVals: string[][] = calendarSheet.getDataRange().getValues();
        for (let i = 1; i < calendarVals.length; i++) {
            const event = calendarVals[i];
            // if (!event || event.length < 8) continue; // データが不足している場合はスキップ
            if (event[7].toString() === '20') {
                // event_status が 20 の場合
                const pitchFee = event[8]; // pitch_fee (9列目)
                if (pitchFee) {
                    return String(pitchFee); // pitch_fee を文字列に変換して返す
                } else {
                    //データがない場合デフォルト
                    const settingSheet = GasProps.instance.settingSheet;
                    return settingSheet.getRange('B3').getValue();
                }
            }
        }
        throw new Error('event_status=20 のイベントが見つかりません');
    }

    public getPaticipationFee() {
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const calendarVals: string[][] = calendarSheet.getDataRange().getValues();
        for (let i = 1; i < calendarVals.length; i++) {
            const event = calendarVals[i];
            // if (!event || event.length < 8) continue; // データが不足している場合はスキップ
            if (event[7].toString() === '20') {
                // event_status が 20 の場合
                const paticipationFee = event[10]; // pitch_fee (9列目)
                if (paticipationFee) {
                    return String(paticipationFee); // pitch_fee を文字列に変換して返す
                } else {
                    //データがない場合デフォルト
                    const settingSheet = GasProps.instance.settingSheet;
                    return settingSheet.getRange('B4').getValue();
                }
            }
        }
        throw new Error('event_status=20 のイベントが見つかりません');
    }

    public getPayNowAddress() {
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet = this.calendarSheet;
        const calendarVals: string[][] = calendarSheet.getDataRange().getValues();
        for (let i = 1; i < calendarVals.length; i++) {
            const event = calendarVals[i];
            // if (!event || event.length < 8) continue; // データが不足している場合はスキップ
            if (event[7].toString() === '20') {
                // event_status が 20 の場合
                const payNow = event[9]; // pitch_fee (9列目)
                if (payNow) {
                    return String(payNow); // pitch_fee を文字列に変換して返す
                } else {
                    //データがない場合デフォルト
                    const settingSheet = GasProps.instance.settingSheet;
                    return settingSheet.getRange('B2').getValue();
                }
            }
        }
        throw new Error('event_status=20 のイベントが見つかりません');
    }
}
