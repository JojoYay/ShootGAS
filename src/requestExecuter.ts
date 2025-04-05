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
    public insertCashBook(postEventHandler: PostEventHandler): void {
        const memo: string = postEventHandler.parameter['memo'];
        const title: string = postEventHandler.parameter['title'];
        const updateUser: string = postEventHandler.parameter['updateUser'] || ''; // Assuming you have this parameter
        const createUser: string = postEventHandler.parameter['createUser'] || ''; // Assuming you have this parameter
        const payee: string = postEventHandler.parameter['payee'] || ''; // Assuming you have this parameter, default to empty if not provided
        const amount: string = postEventHandler.parameter['amount'];
        const invoiceId: string = ''; //隊費直接入力なのでInvoiceは無い
        const calendarId: string = postEventHandler.parameter['calendarId'];

        const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = this.insertCashBookData(
            title,
            memo,
            payee,
            amount,
            invoiceId,
            calendarId,
            updateUser,
            createUser
        ); // 6 is the index for the Balance column
        // Optionally, you can return the updated cashBook data
        postEventHandler.reponseObj.cashBook = cashBook.getDataRange().getValues();
    }

    public deleteCashBook(postEventHandler: PostEventHandler): void {
        const bookId: string = postEventHandler.parameter['bookId'];
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('cashBook2');

        if (cashBook) {
            const values = cashBook.getDataRange().getValues();
            let rowToDelete: number | null = null;
            for (let i = 1; i < values.length; i++) {
                // Assuming the first row is a header
                if (values[i][0] === bookId) {
                    // Assuming bookId is in the first column
                    rowToDelete = i + 1; // Spreadsheet rows are 1-indexed
                    break;
                }
            }
            if (rowToDelete) {
                cashBook.deleteRow(rowToDelete);
            } else {
                postEventHandler.reponseObj.err = '削除するデータが見つかりませんでした BookId:' + bookId;
            }
            postEventHandler.reponseObj.cashBook = cashBook.getDataRange().getValues();
        }
    }

    public insertCashBookData(
        title: string,
        memo: string,
        payee: string,
        amount: string,
        invoiceId: string,
        calendarId: string,
        updateUser: string,
        createUser: string
    ) {
        const bookId: string = Utilities.getUuid();
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('cashBook2');
        if (!cashBook) {
            throw new Error('cashBook sheet was not found.');
        }

        const now: Date = new Date();
        const create: Date = now;
        const lastUpdate: Date = now;

        // Prepare the row data to be added
        const newRow = [
            bookId, // BookId
            title, // Title
            memo, // Memo
            payee, // Payee
            amount, // Amount
            '', // Balance
            invoiceId, // InvoiceId (assuming this is empty for now)
            calendarId, // CalendarId
            lastUpdate, // LastUpdate
            updateUser, // UpdateUser
            create, // Create
            createUser, // CreateUser
        ];
        // Append the new row to the cashBook sheet
        cashBook.appendRow(newRow);
        // Get the last row number to set the formula for balance
        const lastRow = cashBook.getLastRow();

        // Set the formula for the balance column (assuming it's the 6th column)
        const balanceFormula = `=IF(ROW()=2, INDEX(E:E, ROW()), INDEX(F:F, ROW()-1) + INDEX(E:E, ROW()))`;
        cashBook.getRange(lastRow, 6).setFormula(balanceFormula); // 6 is the index for the Balance column
        return cashBook;
    }

    public uploadToYoutube(postEventHander: PostEventHandler): void {
        console.log('uploadToYoutube');
        const fileName: string = postEventHander.parameter['fileName'];
        const fileType: string = postEventHander.parameter['fileType'];
        const fileSize: string = postEventHander.parameter['fileSize'];
        const actDate: string = postEventHander.parameter['actDate'];
        const title: string = actDate + ' ' + fileName;
        console.log('fileName', fileName);
        console.log('title', title);
        console.log('fileSize', fileSize);
        const accessToken = ScriptApp.getOAuthToken();
        // YouTube API の Resumable Upload URL を取得
        const options = {
            method: 'post',
            headers: {
                'Authorization': 'Bearer ' + accessToken,
                'X-Upload-Content-Type': fileType,
                'X-Upload-Content-Length': fileSize,
                'Content-Type': 'application/json; charset=UTF-8',
            },
            // muteHttpExceptions: true,
            payload: JSON.stringify({
                snippet: {
                    // title: fileName,
                    title: title,
                    description: 'YouTube API Upload via GAS',
                    tags: ['ShootSunday', 'YouTube API'],
                    // categoryId: '21',
                },
                status: {
                    privacyStatus: 'unlisted',
                    madeForKids: false, // 子供向けではない設定を追加
                },
            }),
        };
        // console.log('payload:', options.payload); // payload をログ出力
        const response = UrlFetchApp.fetch(
            'https://www.googleapis.com/upload/youtube/v3/videos?uploadType=resumable&part=snippet,status',
            // @ts-ignore
            options
        );

        const headers = response.getAllHeaders();
        // @ts-ignore
        const uploadUrl = headers.Location;
        console.log(uploadUrl);
        postEventHander.reponseObj.uploadUrl = uploadUrl;
        // const videoId = this.getVideoIdByTitle(title);
        // // const videoId = 'hogehoge';
        // // 動画URLを生成
        // const videoUrl = `https://www.youtube.com/watch?v=${videoId}`;
        // console.log('YouTube Video URL:', videoUrl);
        // postEventHander.reponseObj.videoUrl = videoUrl;
        // const accessToken2 = ScriptApp.getOAuthToken();
        // console.log(accessToken2);
        // postEventHander.reponseObj.token = accessToken2;
    }

    private getVideoIdByTitle(videoTitle: string): string | null {
        try {
            const response = YouTube.Search?.list('id,snippet', {
                forMine: true,
                type: 'video',
                q: videoTitle, // 検索クエリに動画タイトルを指定
            });
            if (response && response.items && response.items.length > 0) {
                // 検索結果が複数件の場合、最初の動画をvideoIdとする (より厳密な絞り込みが必要な場合あり)
                const video: GoogleAppsScript.YouTube.Schema.SearchResult = response.items[0];
                if (video.id?.videoId) {
                    return video.id.videoId;
                }
            }
            console.log('動画が見つかりませんでした。タイトル:', videoTitle);
            return null;
        } catch (error) {
            console.error('Videos: get API エラー:', error);
            return null;
        }
    }

    public updateEventData(postEventHander: PostEventHandler): void {
        const title: string = postEventHander.parameter['title']; //こいつで一意
        const weather: string = postEventHander.parameter['weather'];
        const mip1: string = postEventHander.parameter['mip1'];
        const reason: string = postEventHander.parameter['reason'];
        const team1: string = postEventHander.parameter['team1'];
        const team2: string = postEventHander.parameter['team2'];
        const team3: string = postEventHander.parameter['team3'];
        const team4: string = postEventHander.parameter['team4'];
        const team5: string = postEventHander.parameter['team5'];
        const team6: string = postEventHander.parameter['team6'];
        const team7: string = postEventHander.parameter['team7'];
        const team8: string = postEventHander.parameter['team8'];
        const team9: string = postEventHander.parameter['team9'];
        const team10: string = postEventHander.parameter['team10'];
        const mip2: string = postEventHander.parameter['mip2'];
        const mip3: string = postEventHander.parameter['mip3'];
        const mip4: string = postEventHander.parameter['mip4'];
        const mip5: string = postEventHander.parameter['mip5'];

        const eventDetailSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.eventResultSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const eventDetailValues: any[][] = eventDetailSheet.getDataRange().getValues();

        let targetRowIndex = -1;
        // 2列目のタイトルをキーにデータを検索
        for (let i = 1; i < eventDetailValues.length; i++) {
            // 1行目はヘッダー行と仮定
            if (eventDetailValues[i][1] === title) {
                // 2列目（インデックス1）がタイトル
                targetRowIndex = i;
                break; // タイトルが一致する行が見つかったらループを抜ける
            }
        }

        if (targetRowIndex !== -1) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const updateValues: any[] = eventDetailValues[targetRowIndex]; // 更新前の行データをコピー

            // 現在の更新日時をmm/dd/yyyy hh24:mi:ss形式で入力
            const now = new Date();
            const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
            updateValues[0] = formattedDate; // 1列目に更新日時を設定

            updateValues[4] = weather !== undefined ? weather : updateValues[4]; // 天気 (5列目)
            updateValues[5] = mip1 !== undefined ? mip1 : updateValues[5]; // MIP1 (6列目)
            updateValues[6] = reason !== undefined ? reason : updateValues[6]; // 選出理由 (7列目)
            updateValues[7] = team1 !== undefined ? team1 : updateValues[7]; // team1 (8列目)
            updateValues[8] = team2 !== undefined ? team2 : updateValues[8]; // team2 (9列目)
            updateValues[9] = team3 !== undefined ? team3 : updateValues[9]; // team3 (10列目)
            updateValues[10] = team4 !== undefined ? team4 : updateValues[10]; // team4 (11列目)
            updateValues[11] = team5 !== undefined ? team5 : updateValues[11]; // team5 (12列目)
            updateValues[12] = team6 !== undefined ? team6 : updateValues[12]; // team6 (13列目)
            updateValues[13] = team7 !== undefined ? team7 : updateValues[13]; // team7 (14列目)
            updateValues[14] = team8 !== undefined ? team8 : updateValues[14]; // team8 (15列目)
            updateValues[15] = team9 !== undefined ? team9 : updateValues[15]; // team9 (16列目)
            updateValues[16] = team10 !== undefined ? team10 : updateValues[16]; // team10 (17列目)
            updateValues[17] = mip2 !== undefined ? mip2 : updateValues[17]; // MIP2 (18列目)
            updateValues[18] = mip3 !== undefined ? mip3 : updateValues[18]; // MIP3 (19列目)
            updateValues[19] = mip4 !== undefined ? mip4 : updateValues[19]; // MIP4 (20列目)
            updateValues[20] = mip5 !== undefined ? mip5 : updateValues[20]; // MIP5 (21列目)

            // シートに書き戻し
            eventDetailSheet.getRange(targetRowIndex + 1, 1, 1, updateValues.length).setValues([updateValues]);
        } else {
            Logger.log(`No event found with title: ${title}`);
        }
    }

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
        const payNow: string = postEventHander.parameter['paynow_link'];
        const pitch: string = postEventHander.parameter['pitch_fee'];
        const paticipation: string = postEventHander.parameter['paticipation_fee'];
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
                case 'pitch_fee':
                    newRowData.push(pitch);
                    break;
                case 'paynow_link':
                    newRowData.push(payNow);
                    break;
                case 'paticipation_fee':
                    newRowData.push(paticipation);
                    break;
                default:
                    newRowData.push(''); // その他のヘッダーの場合は空文字をセット
            }
        });
        console.log(payNow);
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
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const eventType: string = postEventHander.parameter['event_type'];
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const eventName: string = postEventHander.parameter['event_name'];
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const sDate: string = postEventHander.parameter['start_datetime'];
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const place: string = postEventHander.parameter['place'];
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const remark: string = postEventHander.parameter['remark'];
        // // eslint-disable-next-line @typescript-eslint/no-unused-vars
        // const recursiveType: number = postEventHander.parameter['event_status']; // default value
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
            [
                'event_type',
                'event_name',
                'start_datetime',
                'end_datetime',
                'place',
                'remark',
                'event_status',
                'pitch_fee',
                'paynow_link',
                'paticipation_fee',
            ].forEach(paramName => {
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

    private insertComments(postEventHander: PostEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const comments: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('comments');
        if (!comments) {
            throw new Error('comments Sheet was not found.');
        }
        const componentId: string = postEventHander.parameter['component_id'];
        const category: string = postEventHander.parameter['category'];
        const content: string = postEventHander.parameter['content'];
        const createUser: string = postEventHander.parameter['create_user'];
        const headerRow = comments.getDataRange().getValues()[0]; // ヘッダー行を取得
        // console.log(componentId);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const newRowData: any[] = [];
        headerRow.forEach(header => {
            switch (header) {
                case 'id':
                    newRowData.push(Utilities.getUuid());
                    break;
                case 'component_id':
                    newRowData.push(componentId);
                    break;
                case 'category':
                    newRowData.push(category);
                    break;
                case 'user_id':
                    newRowData.push(createUser);
                    break;
                case 'content':
                    newRowData.push(content);
                    break;
                case 'created': // 幹事フラグはパラメータにないため空文字
                    newRowData.push(new Date());
                    break;
                default:
                    newRowData.push(''); // その他のヘッダーの場合は空文字をセット
            }
        });
        comments.appendRow(newRowData);
    }

    private deleteComments(postEventHander: PostEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const comments: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('comments');
        if (!comments) {
            throw new Error('comments Sheet was not found.');
        }
        const id: string = postEventHander.parameter['id'];
        const values = comments.getDataRange().getValues();

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
            comments.deleteRow(rowNumberToDelete);
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
        // console.log('execute loadExList');
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.expenseFolder);
        const titleFolderIt: GoogleAppsScript.Drive.FolderIterator = rootFolder.getFolders();
        const results = [];
        while (titleFolderIt.hasNext()) {
            const expenseFolder: GoogleAppsScript.Drive.Folder = titleFolderIt.next();
            const title = expenseFolder.getName();
            // const url = expenseFolder.getFilesByName(title).next().getUrl();
            const url = expenseFolder.getUrl(); // フォルダのURLを取得
            results.push({ title: title, url: url });
        }
        postEventHander.reponseObj = { resultList: results };
    }

    public uploadPaticipationPayNow(postEventHander: PostEventHandler): void {
        console.log('execute uploadPaticipationPayNow');
        const decodedFile = Utilities.base64Decode(postEventHander.parameter.file);
        const userId: string = postEventHander.parameter.userId;
        const mappingSheet = GasProps.instance.mappingSheet;
        const mapVals = mappingSheet.getDataRange().getValues();
        const userVal = mapVals.filter(row => row[2] === userId)[0]; // 1列目が calendarId

        const densukeName: string = userVal[1].toString();
        // const lineName: string = userVal[0].toString();
        const actDate: string = postEventHander.parameter.actDate;
        // const title: string = postEventHander.parameter.title;
        const blob = Utilities.newBlob(decodedFile, 'application/octet-stream', actDate + '_' + densukeName);
        const lineUtil: LineUtil = new LineUtil();
        const payNowFolder = lineUtil.createPayNowFolder(actDate);
        if (!payNowFolder) {
            return; //folderは必ず作られる
        }
        const fileNameToSearch = actDate + '_' + densukeName;
        const searchQuery = `title = '${fileNameToSearch}' and '${payNowFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
        const oldFileIt = payNowFolder.searchFiles(searchQuery); // searchFiles を使用

        while (oldFileIt.hasNext()) {
            oldFileIt.next().setTrashed(true);
        }
        const file = payNowFolder.createFile(blob);
        console.log(densukeName + ' uploaded ' + file.getName() + ' in ' + actDate);
        gasUtil.updatePaymentStatus(densukeName, actDate);
        const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
        postEventHander.reponseObj = { picUrl: picUrl };
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

    public uploadInvoice(postEventHander: PostEventHandler): void {
        console.log('execute uploadinvoice');
        const decodedFile = Utilities.base64Decode(postEventHander.parameter.file);
        const userId: string = postEventHander.parameter.userId;
        // const calendarId: string = postEventHander.parameter.calendarId;
        const actDate: string = postEventHander.parameter.actDate;
        const amount: string = postEventHander.parameter.amount;
        const remarks: string = postEventHander.parameter.remarks;

        const mappingSheet = GasProps.instance.mappingSheet;
        const mapVals = mappingSheet.getDataRange().getValues();
        const userVal = mapVals.filter(row => row[2] === userId)[0]; // 1列目が calendarId

        const densukeName: string = userVal[1].toString();
        const blob = Utilities.newBlob(decodedFile, 'application/octet-stream', actDate + '_' + densukeName);
        const lineUtil: LineUtil = new LineUtil();
        const payNowFolder = lineUtil.createPayNowFolder(actDate);
        if (!payNowFolder) {
            return; //folderは必ず作られる
        }
        const file = payNowFolder.createFile(blob);
        console.log(densukeName + ' uploaded ' + file.getName() + ' in ' + actDate);
        // gasUtil.updatePaymentStatus(densukeName, actDate);
        const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();

        const paymentSummary: GoogleAppsScript.Spreadsheet.Spreadsheet = gasUtil.createSpreadSheet(actDate, payNowFolder, [
            'id',
            'upload日付',
            'ユーザー名',
            '金額',
            'メモ',
            '画像',
            '状態',
        ]);
        console.log('spreadSheet' + paymentSummary.getUrl());
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = paymentSummary.getActiveSheet();

        const newId = Utilities.getUuid(); // 例: ID
        const newUploadDate = new Date().toLocaleDateString(); // 現在の日付を取得
        const newUserName = densukeName; // 例: ユーザー名
        const newAmount = amount; // 例: 金額
        const newMemo = remarks; // 例: メモ
        const newImageUrl = picUrl; // 例: 画像URL
        sheet.appendRow([newId, newUploadDate, newUserName, newAmount, newMemo, newImageUrl, '未清算']);
        postEventHander.reponseObj = { picUrl: picUrl, sheetUrl: GasProps.instance.generateSheetUrl(paymentSummary.getId()) };
    }

    public deleteInvoice(postEventHander: PostEventHandler): void {
        console.log('execute deleteInvoice');
        const actDate: string = postEventHander.parameter.actDate;
        const invoiceId: string = postEventHander.parameter.invoiceId;

        const payNowFolder = lineUtil.createPayNowFolder(actDate);
        if (!payNowFolder) {
            return; //folderは必ず作られる
        }
        const paymentSummary: GoogleAppsScript.Spreadsheet.Spreadsheet = gasUtil.createSpreadSheet(actDate, payNowFolder, [
            'id',
            'upload日付',
            'ユーザー名',
            '金額',
            'メモ',
            '画像',
            '状態',
        ]);
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = paymentSummary.getActiveSheet();
        const data = sheet.getDataRange().getValues(); // Get all data from the sheet

        // Find the row with the matching invoiceId
        for (let i = 1; i < data.length; i++) {
            // Start from 1 to skip header
            if (data[i][0] === invoiceId) {
                // Assuming 'id' is in the first column (index 0)
                const picUrl = data[i][5]; // Assuming '画像' is in the sixth column (index 5)
                const fileId = picUrl.split('/d/')[1].split('/')[0]; // Extract the file ID from the URL

                // Move the file to trash
                const file = DriveApp.getFileById(fileId);
                file.setTrashed(true); // Move the file to trash

                // Delete the row
                sheet.deleteRow(i + 1); // +1 because sheet.deleteRow is 1-based index
                break; // Exit the loop after deleting the row
            }
        }
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

    public payNow(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();

        const attendees = su.extractAttendees('〇');
        const actDate = su.extractDateFromRownum();
        const messageId = postEventHander.messageId;
        const userId = postEventHander.userId;
        const densukeName = gasUtil.getNickname(userId);
        // console.log(densukeName);
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
                        'のスケジューラーの出席が〇になっていませんでした。スケジューラーを更新してください。\n' +
                        su.schedulerUrl;
                } else {
                    postEventHander.resultMessage =
                        '【Error】Your attendance on ' +
                        actDate +
                        ' in the scheduler has not been marked as 〇.\nPlease update scheduler.\n' +
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
        if (!postEventHander.userId && !gasUtil.getNickname(postEventHander.userId)) {
            postEventHander.resultMessage = '初回登録が終わっていません。スケジューラーへアクセスし、初回登録を完了させてください。';
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
        su.generateSummaryBase();
        postEventHander.resultMessage = su.getSummaryStr();
    }

    public unpaid(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();
        const unpaid = gasUtil.getUnpaid(actDate);
        postEventHander.resultMessage = '未払いの人 (' + unpaid.length + '名): ' + unpaid.join(', ');
    }

    public remind(postEventHander: PostEventHandler): void {
        const su: SchedulerUtil = new SchedulerUtil();
        postEventHander.resultMessage = su.generateRemind();
    }

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
