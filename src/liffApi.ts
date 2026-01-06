// import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { GetEventHandler } from './getEventHandler';
import { LineUtil } from './lineUtil';
import { RequestExecuter } from './requestExecuter';
import { SchedulerUtil } from './schedulerUtil';
import { ScoreBook } from './scoreBook';
import { ScriptProps } from './scriptProps';

export class LiffApi {
    private getWeightRecord(getEventHandler: GetEventHandler): void {
        getEventHandler.e.parameter['sheetName'] = 'WeightRecord';
        getEventHandler.e.parameter['type'] = 'setting';
        this.getSheetData(getEventHandler);
        getEventHandler.result.weightRecord = getEventHandler.result.data;
        delete getEventHandler.result.data;
    }

    private saveWeightRecord(getEventHandler: GetEventHandler): void {
        getEventHandler.e.parameter['sheetName'] = 'WeightRecord';
        getEventHandler.e.parameter['type'] = 'setting';
        const userId = getEventHandler.e.parameter['userId'];
        const height = getEventHandler.e.parameter['height'];
        const weight = getEventHandler.e.parameter['weight'];
        const bfp = getEventHandler.e.parameter['bfp'];
        const date = getEventHandler.e.parameter['date'];

        // データをJSON文字列として設定
        getEventHandler.e.parameter['data'] = JSON.stringify([userId, height, weight, bfp, date]);

        // this.saveSheetData(getEventHandler);
    }

    private deleteWeightRecord(getEventHandler: GetEventHandler): void {
        getEventHandler.e.parameter['sheetName'] = 'WeightRecord';
        getEventHandler.e.parameter['type'] = 'setting';

        this.deleteSheetData(getEventHandler);
    }

    private test(getEventHandler: GetEventHandler): void {
        const value: string = getEventHandler.e.parameters['param'][0];
        getEventHandler.result = { result: value };
    }

    private updateYTVideo(getEventHandler: GetEventHandler): void {
        try {
            const videoTitle: string = getEventHandler.e.parameter['videoTitle'];
            const actDate: string = getEventHandler.e.parameter['actDate'];
            const fileName: string = getEventHandler.e.parameter['fileName'];
            const response = YouTube.Search?.list('id,snippet', {
                forMine: true,
                type: 'video',
                q: videoTitle, // 検索クエリに動画タイトルを指定
            });
            if (response && response.items && response.items.length > 0) {
                // 検索結果が複数件の場合、最初の動画をvideoIdとする (より厳密な絞り込みが必要な場合あり)
                const video: GoogleAppsScript.YouTube.Schema.SearchResult = response.items[0];
                if (video.id?.videoId) {
                    const videoUrl: string = `https://www.youtube.com/watch?v=${video.id.videoId}`;

                    const videoSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
                    const videoSheetVals = videoSheet.getDataRange().getValues();
                    let matchedRowIndex = -1;
                    // ２行目からデータ行を検索 (１行目、２行目は検索条件として使用するためスキップ)
                    for (let i = 2; i < videoSheetVals.length; i++) {
                        const row = videoSheetVals[i];
                        if (row[0] === actDate && row[1] === fileName) {
                            matchedRowIndex = i;
                            break; // 最初に見つかった行で処理を終える
                        }
                    }

                    if (matchedRowIndex !== -1) {
                        // マッチする行が見つかった場合、３列目（C列）にvideoUrlを書き込む
                        videoSheet.getRange(matchedRowIndex + 1, 3).setValue(videoUrl);
                        console.log(`Matched row found at index ${matchedRowIndex + 1}. videoUrl updated.`);
                    } else {
                        console.log(`No matching row found for actDate: ${actDate} and Title: ${fileName}`);
                    }
                } else {
                    console.log('動画が見つかりませんでした。タイトル:', videoTitle);
                }
            }
        } catch (error) {
            console.error('Videos: get API エラー:', error);
        }
    }

    private getYTVideoIdByTitle(getEventHandler: GetEventHandler): void {
        try {
            const videoTitle: string = getEventHandler.e.parameter['videoTitle'];
            const response = YouTube.Search?.list('id,snippet', {
                forMine: true,
                type: 'video',
                q: videoTitle, // 検索クエリに動画タイトルを指定
            });
            if (response && response.items && response.items.length > 0) {
                // 検索結果が複数件の場合、最初の動画をvideoIdとする (より厳密な絞り込みが必要な場合あり)
                const video: GoogleAppsScript.YouTube.Schema.SearchResult = response.items[0];
                if (video.id?.videoId) {
                    getEventHandler.result.videoId = `https://www.youtube.com/watch?v=${video.id.videoId}`;
                }
            }
            console.log('動画が見つかりませんでした。タイトル:', videoTitle);
            // return '';
        } catch (error) {
            console.error('Videos: get API エラー:', error);
            // return '';
        }
    }

    private getCalendar(getEventHandler: GetEventHandler): void {
        const calId: string = getEventHandler.e.parameter['calendarId'];
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('calendar');
        if (!calendarSheet) {
            throw new Error('calendar sheet was not found.');
        }
        const calendarValues = calendarSheet.getDataRange().getValues();
        const filteredCalendar = calendarValues.filter(row => row[0] === calId); // 1列目が calendarId

        getEventHandler.result.event = filteredCalendar[0];
    }

    private getAttendance(getEventHandler: GetEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const attendance: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('attendance');
        if (!attendance) {
            throw new Error('attendance sheet was not found.');
        }
        getEventHandler.result.attendance = attendance.getDataRange().getValues();
        getEventHandler.result.jsonAttendance = this.convertSheetDataToJson(getEventHandler.result.attendance);
    }

    private getAttendees(getEventHandler: GetEventHandler): void {
        const calendarId: string = getEventHandler.e.parameter['calendarId'];
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const attendance: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('attendance');
        if (!attendance) {
            throw new Error('attendance sheet was not found.');
        }
        const attendanceValues = attendance.getDataRange().getValues();
        const filteredAttendees = attendanceValues.filter(row => row[6] === calendarId && row[5] === '〇'); // 7列目が calendar_id
        // フィルターしたデータのuserIdを抽出
        const attendeeUserIds = filteredAttendees.map(row => row[1]); // 2列目が user_id

        // mappingSheetからデータを取得
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        const mappingValues = mappingSheet.getDataRange().getValues();

        const matchedMappingData = [];

        // mappingSheetの3列目とattendeeUserIdsをマッチング
        for (let i = 1; i < mappingValues.length; i++) {
            // 1行目はヘッダー行と仮定
            const mappingRow = mappingValues[i];
            const mappingUserId = mappingRow[2]; // 3列目が user_id (LINE ID) と仮定

            if (attendeeUserIds.includes(mappingUserId)) {
                matchedMappingData.push(mappingRow);
            }
        }
        getEventHandler.result.attendees = matchedMappingData;
    }

    private loadCashBook(getEventHandler: GetEventHandler): void {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const cashBook: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('cashBook2');
        if (!cashBook) {
            throw new Error('cashBook sheet was not found.');
        }
        getEventHandler.result.cashBook = cashBook.getDataRange().getValues();
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
        // console.log('info', actDate);
        //ない場合は今のやつ その場合全体のVideoリスト・日付のリストも含める
        if (!actDate) {
            // const den: DensukeUtil = new DensukeUtil();
            // const chee = den.getDensukeCheerio();
            const videoVals = videos.getDataRange().getValues();
            // console.log('videoVals', videoVals);
            if (videoVals.length > 1) {
                actDate = videoVals[videoVals.length - 1][0]; // videosシートの最終行の１列目の値を取得
            }
            // console.log(actDate);
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

        const su: SchedulerUtil = new SchedulerUtil();
        const actDate = su.extractDateFromRownum();

        getEventHandler.result.match = videos
            .getDataRange()
            .getValues()
            .filter(val => val[0] === actDate && !val[10].endsWith('_g') && !val[10].endsWith('d') && val[3] && val[4]);
    }

    private getPayNow(getEventHandler: GetEventHandler): void {
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        // getEventHandler.result = { result: members };
        getEventHandler.result.payNow = addy;
    }

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
            // ヘッダーを一度に設定するための配列を作成
            const headers = [['No', '試合', 'チーム', 'アシスト', 'ゴール']];
            // 一度に範囲を設定
            shootLog.getRange(1, 1, 1, headers[0].length).setValues(headers);
        }

        getEventHandler.result.scores = shootLog.getDataRange().getValues();
    }

    private getRegisteredMembers(getEventHandler: GetEventHandler): void {
        const members: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        getEventHandler.result.members = members.getDataRange().getValues();
    }

    private getDensukeName(getEventHandler: GetEventHandler): void {
        const gasUtil: GasUtil = new GasUtil();
        // const lineUtil: LineUtil = new LineUtil();
        const userId = getEventHandler.e.parameters['userId'][0];
        // getEventHandler.result = { result: gasUtil.getDensukeName(lineUtil.getLineDisplayName(userId)) };
        getEventHandler.result.densukeName = gasUtil.getNickname(userId);
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
        // const fileIt = expenseFolder.getFilesByName(title + '_' + lineName);

        const fileNameToSearch = title + '_' + lineName;
        const searchQuery = `title = '${fileNameToSearch}' and '${expenseFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
        const fileIt = expenseFolder.searchFiles(searchQuery); // searchFiles を使用

        if (fileIt.hasNext()) {
            const file = fileIt.next();
            getEventHandler.result.statusMsg = '支払い済み';
            const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
            getEventHandler.result.picUrl = picUrl;
        } else {
            let spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
            // const fileIt2 = expenseFolder.getFilesByName(title);
            const searchQuery = `title = '${title}' and '${expenseFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
            const fileIt2 = expenseFolder.searchFiles(searchQuery); // searchFiles を使用

            if (fileIt2.hasNext()) {
                const sheetFile = fileIt2.next();
                spreadSheet = SpreadsheetApp.openById(sheetFile.getId());
            } else {
                throw new Error('SpreadSheet is not available:' + title);
            }

            const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getActiveSheet();
            const sheetVal = sheet.getDataRange().getValues();
            const gasUtil: GasUtil = new GasUtil();
            // const densukeName = gasUtil.getDensukeName(lineName);
            const densukeName = gasUtil.getNickname(userId);
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

    private updateOpenPaymentStatus(getEventHandler: GetEventHandler): void {
        const id: string = getEventHandler.e.parameters['id'][0];
        const folderName: string = getEventHandler.e.parameters['folderName'][0];
        const userId: string = getEventHandler.e.parameters['userId'][0];

        const parentFolder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder; // 親フォルダを取得
        const folders = parentFolder.getFolders();
        const mappingSheetVal = GasProps.instance.mappingSheet.getDataRange().getValues();
        const su: SchedulerUtil = new SchedulerUtil();
        const calVal = su.calendarSheet.getDataRange().getValues();

        while (folders.hasNext()) {
            const folder = folders.next();
            const searchQuery = `title = '${folderName}' and '${folder.getId()}' in parents`;
            const files = folder.searchFiles(searchQuery);

            if (files.hasNext()) {
                const file = files.next();
                const spreadsheet = SpreadsheetApp.openById(file.getId()); // スプレッドシートを開く
                const sheet = spreadsheet.getActiveSheet(); // アクティブなシートを取得
                const data = sheet.getDataRange().getValues(); // シートのデータを取得

                for (let i = 1; i < data.length; i++) {
                    if (data[i][0] === id) {
                        data[i][6] = '清算済';
                        sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]); // Update the specific row

                        const re: RequestExecuter = new RequestExecuter();
                        const user: string[] | undefined = mappingSheetVal.find(user => user[1] === data[i][2]);
                        const payeeId = user ? user[2] : '';
                        const lastUnderscoreIndex = folderName.lastIndexOf('(');
                        const calendarTitle = lastUnderscoreIndex !== -1 ? folderName.substring(0, lastUnderscoreIndex) : folderName;
                        console.log(calendarTitle);
                        const calendar: string[] | undefined = calVal.find(cal => cal[2] === calendarTitle);
                        const calendarId = calendar ? calendar[0] : '';
                        console.log(calendarId);
                        re.insertCashBookData(folderName, data[i][4], payeeId, data[i][3], data[i][0], calendarId, userId, userId);
                        break;
                    }
                }
            }
        }

        this.loadOpenPayment(getEventHandler);
        // getEventHandler.result.openPayment = allData;
    }

    //現在終了していない、かつ、フォルダが存在している支払いを取得する
    private loadOpenPayment(getEventHandler: GetEventHandler): void {
        const parentFolder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder; // 親フォルダを取得
        const folders = parentFolder.getFolders();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const allData: any[] = []; // データを格納する配列

        while (folders.hasNext()) {
            const folder = folders.next();
            const folderName = folder.getName(); // フォルダ名を取得

            // スプレッドシートを検索
            const searchQuery = `title = '${folderName}' and '${folder.getId()}' in parents`;
            const files = folder.searchFiles(searchQuery);

            if (files.hasNext()) {
                const file = files.next();
                const spreadsheet = SpreadsheetApp.openById(file.getId()); // スプレッドシートを開く
                const sheet = spreadsheet.getActiveSheet(); // アクティブなシートを取得
                const data = sheet.getDataRange().getValues(); // シートのデータを取得

                // データを連想配列に追加（ヘッダーを除く）
                for (let i = 1; i < data.length; i++) {
                    // 1行目はヘッダーなのでスキップ
                    const rowData = data[i];
                    // フォルダ名を追加して連想配列に格納
                    allData.push({
                        id: rowData[0], // Assuming 'id' is in the first column
                        uploadDate: rowData[1], // Assuming 'upload日付' is in the second column
                        userName: rowData[2], // Assuming 'ユーザー名' is in the third column
                        amount: rowData[3], // Assuming '金額' is in the fourth column
                        memo: rowData[4], // Assuming 'メモ' is in the fifth column
                        image: rowData[5], // Assuming '画像' is in the sixth column
                        status: rowData[6], // Assuming '状態' is in the seventh column
                        folderName: folderName,
                    });
                }
            }
        }

        // 取得したデータを結果に格納
        getEventHandler.result.openPayment = allData;
    }

    //そのユーザーのInvoiceを取得
    private getInvoices(getEventHandler: GetEventHandler): void {
        const gasUtil: GasUtil = new GasUtil();
        const userId: string = getEventHandler.e.parameter['userId'];
        // const lang: string = getEventHandler.e.parameter['lang'];
        this.getCalendar(getEventHandler);
        const calendarVals = getEventHandler.result.event;
        const mappingSheet = GasProps.instance.mappingSheet;
        const mapVals = mappingSheet.getDataRange().getValues();
        const userVal = mapVals.filter(row => row[2] === userId)[0];
        const densukeName: string = userVal[1].toString();
        const date = new Date(calendarVals[3]);
        const actDate: string = calendarVals[2] + '(' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd MMM') + ')'; // calendar_id (1列目)
        const lineUtil: LineUtil = new LineUtil();
        const payNowFolder = lineUtil.createPayNowFolder(actDate, false); //ない場合は作らない
        if (payNowFolder) {
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

            const allData = sheet.getDataRange().getValues(); // シートの全データを取得
            // const header = allData[0]; // ヘッダー行
            const filteredData = allData.filter((row, index) => index === 0 || row[2] === densukeName); // ヘッダー行を除外してフィルタリング
            getEventHandler.result.invoices = filteredData;
        } else {
            getEventHandler.result.invoices = [];
        }
        getEventHandler.result.actDate = actDate;
        getEventHandler.result.statusMsg = '立て替えを行ったレシートをアップロードして下さい。';
        getEventHandler.result.statusMsg2 = '金額はSGDで入力して下さい（日本円不可）';
        getEventHandler.result.statusMsg3 = '日本円のレートは $1 = 110円 です';
    }

    private getPaticipationFeeWithStatus(getEventHandler: GetEventHandler): void {
        const userId: string = getEventHandler.e.parameter['userId'];
        const lang: string = getEventHandler.e.parameter['lang'];
        this.getCalendar(getEventHandler);
        const calendarVals = getEventHandler.result.event;
        const mappingSheet = GasProps.instance.mappingSheet;
        const mapVals = mappingSheet.getDataRange().getValues();
        const userVal = mapVals.filter(row => row[2] === userId)[0];
        const densukeName: string = userVal[1].toString();
        const date = new Date(calendarVals[3]);
        const actDate: string = calendarVals[2] + '(' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd MMM') + ')'; // calendar_id (1列目)
        // const actDate: string = calendarVals[2].toString();
        const lineUtil: LineUtil = new LineUtil();
        const payNowFolder = lineUtil.createPayNowFolder(actDate);
        if (payNowFolder) {
            console.log(payNowFolder.getId());
            getEventHandler.result.actDate = actDate;

            const fileNameToSearch = actDate + '_' + densukeName;
            const searchQuery = `title = '${fileNameToSearch}' and '${payNowFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
            const fileIt = payNowFolder.searchFiles(searchQuery); // searchFiles を使用
            // const fileIt = payNowFolder.getFilesByName(actDate + '_' + densukeName);
            console.log('fileIt.hasNext', fileIt.hasNext());
            if (fileIt.hasNext()) {
                const file = fileIt.next();
                getEventHandler.result.statusMsg = lang === 'ja-JP' ? '支払い済み' : 'Payment Received';
                getEventHandler.result.statusMsg2 = '';
                getEventHandler.result.statusMsg3 =
                    lang === 'ja-JP' ? '写真を変更する場合は再度写真を選択して下さい。' : 'Please re-select the photo if you want to change it';
                const picUrl: string = 'https://lh3.googleusercontent.com/d/' + file.getId();
                getEventHandler.result.picUrl = picUrl;
            } else {
                let paticipationFee = calendarVals[10];
                if (!paticipationFee) {
                    if (paticipationFee.toString() === '0') {
                        getEventHandler.result.statusMsg =
                            lang === 'ja-JP' ? '参加費無料。お支払いの必要はありません。' : 'Participation fee is free. No payment is required.';
                        getEventHandler.result.statusMsg2 = '';
                        getEventHandler.result.statusMsg3 = '';
                        getEventHandler.result.picUrl = '';
                        return;
                    } else {
                        const settingSheet = GasProps.instance.settingSheet;
                        paticipationFee = settingSheet.getRange('B4').getValue();
                    }
                }
                let addy = calendarVals[9];
                if (!addy) {
                    const settingSheet = GasProps.instance.settingSheet;
                    addy = settingSheet.getRange('B2').getValue();
                }
                getEventHandler.result.statusMsg = lang === 'ja-JP' ? '支払額：$' + paticipationFee : 'Payment amount: $' + paticipationFee;
                getEventHandler.result.statusMsg2 = lang === 'ja-JP' ? 'PayNow先:' + addy : 'PayNow to:' + addy;
                getEventHandler.result.statusMsg3 =
                    lang === 'ja-JP'
                        ? '支払い済みのスクリーンショットをこちらにアップロードして下さい'
                        : 'Please upload a screenshot of your payment here';

                getEventHandler.result.picUrl = '';
            }
        }
    }

    private getStats(getEventHandler: GetEventHandler): void {
        const resultSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.personalTotalSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = resultSheet.getDataRange().getValues();
        getEventHandler.result.stats = resultValues;
        getEventHandler.result.jsonStats = this.convertSheetDataToJson(resultValues);
    }

    /**
     * スプレッドシートの2次元配列を、ヘッダー行をキーとしたJSONオブジェクトの配列に変換する
     * @param values スプレッドシートから取得した2次元配列（1行目がヘッダー）
     * @returns ヘッダーをキーとしたJSONオブジェクトの配列
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private convertSheetDataToJson(values: any[][]): any[] {
        if (!values || values.length === 0) {
            return [];
        }
        const headers: string[] = values[0].map(header => String(header || ''));
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const jsonArray: any[] = [];
        for (let i = 1; i < values.length; i++) {
            const row = values[i];
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const jsonObject: any = {};
            for (let j = 0; j < headers.length; j++) {
                const header = headers[j];
                if (header) {
                    jsonObject[header] = row[j];
                }
            }
            jsonArray.push(jsonObject);
        }
        return jsonArray;
    }

    public getUsers(getEventHandler: GetEventHandler): void {
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = mappingSheet.getDataRange().getValues();
        getEventHandler.result.users = resultValues;
        getEventHandler.result.jsonUsers = this.convertSheetDataToJson(resultValues);
        console.log(getEventHandler.result.jsonUsers);
    }

    public getQuizData(getEventHandler: GetEventHandler): void {
        const tabName: string = getEventHandler.e.parameter['tabName'];

        if (!tabName) {
            // tabNameが指定されていない場合は、getUsersと同じ動作をする
            this.getUsers(getEventHandler);
            getEventHandler.result.quizData = getEventHandler.result.users;
            return;
        }

        // tabNameに基づいてシートからデータを取得
        // tabNameをシート名として使用するか、または設定シートから取得
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        let quizSheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;

        // まず、tabNameをシート名として検索
        quizSheet = setting.getSheetByName(tabName);

        if (!quizSheet) {
            // シートが見つからない場合は、mappingSheetを使用（デフォルト動作）
            const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const resultValues: any[][] = mappingSheet.getDataRange().getValues();
            getEventHandler.result.quizData = resultValues;
            getEventHandler.result.jsonQuizData = this.convertSheetDataToJson(resultValues);
            console.log('Quiz sheet not found, using mappingSheet. tabName:', tabName);
            return;
        }

        // シートからデータを取得
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = quizSheet.getDataRange().getValues();
        getEventHandler.result.quizData = resultValues;
        getEventHandler.result.jsonQuizData = this.convertSheetDataToJson(resultValues);
        console.log('Quiz data loaded from sheet:', tabName);
    }

    private getYTComments(getEventHandler: GetEventHandler): void {
        const videoId: string = this.getYouTubeVideoId(getEventHandler.e.parameters['url'][0]);

        try {
            if (YouTube.CommentThreads && videoId) {
                const response = YouTube.CommentThreads.list('snippet,replies', {
                    videoId: videoId,
                    maxResults: 20, // 取得するコメントの最大数 (調整可能)
                    order: 'time', // コメントの並び順 (time: 新しい順, relevance: 関連性の高い順)
                });
                if (response.items) {
                    const comments = response.items.map(item => {
                        const snippet = item.snippet?.topLevelComment?.snippet;
                        return {
                            author: snippet?.authorDisplayName || '',
                            // comment: snippet?.textDisplay || '',
                            comment: snippet?.textDisplay ? this.convertYouTubeLink(snippet?.textDisplay) : '',
                            publishedAt: snippet?.publishedAt || '',
                        };
                    });
                    getEventHandler.result.comments = comments; // 結果を getEventHandler.result に格納
                }
            } else {
                getEventHandler.result.comments = []; // コメントがない場合は空の配列を格納
            }
        } catch (error) {
            console.error('Error fetching YouTube comments:', error);
            getEventHandler.result.comments = []; // エラーが発生した場合は空の配列を格納
            getEventHandler.result.error = 'Failed to fetch comments.'; // エラーメッセージを格納
        }
    }

    private convertYouTubeLink(inputString: string): string {
        const urlRegex = /href="([^"]*watch\?v=([a-zA-Z0-9_-]+)[^"]*)"/;
        const timeRegex = /&amp;t=(\d+)/;
        const urlMatch = inputString.match(urlRegex);
        const timeMatch = inputString.match(timeRegex);
        if (!urlMatch || !timeMatch) {
            return inputString;
        }
        const videoId = urlMatch[2]; // ビデオID
        const time = timeMatch[1]; // 時間
        const newUrl = `https://youtu.be/${videoId}?t=${time}`;
        const replacedString = inputString.replace(/<a href="([^"]*)"/, `<a href="${newUrl}" target="_blank"`);
        return replacedString;
    }

    private getYouTubeVideoId(url: string): string {
        if (url) {
            // 短縮URL (youtu.be/xxxx)
            let match = url.match(/^https:\/\/youtu\.be\/([a-zA-Z0-9_-]+)/);
            if (match) {
                return match[1];
            }

            // 埋め込みURL (youtube.com/embed/xxxx)
            match = url.match(/^https:\/\/www\.youtube\.com\/embed\/([a-zA-Z0-9_-]+)/);
            if (match) {
                return match[1];
            }

            // 通常のURL (youtube.com/watch?v=xxxx&...)
            match = url.match(/^https:\/\/www\.youtube\.com\/watch\?v=([a-zA-Z0-9_-]+)/);
            if (match) {
                return match[1];
            }

            // Shorts URL (youtube.com/shorts/xxxx)
            match = url.match(/^https:\/\/www\.youtube\.com\/shorts\/([a-zA-Z0-9_-]+)/);
            if (match) {
                return match[1];
            }
        }
        return ''; // ビデオIDが見つからない場合
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
        // const fileIt = expenseFolder.getFilesByName(title);
        const searchQuery = `title = '${title}' and '${expenseFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
        const fileIt = expenseFolder.searchFiles(searchQuery); // searchFiles を使用

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
            // const fileIt = expenseFolder.getFilesByName(title + '_' + mapRow?.[0]);
            const search = title + '_' + mapRow?.[0];
            const searchQuery = `title = '${search}' and '${expenseFolder.getId()}' in parents`; // より正確なファイル名検索クエリ
            const fileIt = expenseFolder.searchFiles(searchQuery); // searchFiles を使用
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

    // 汎用的なシート操作メソッド
    private getSheetData(getEventHandler: GetEventHandler): void {
        const sheetName: string = getEventHandler.e.parameter['sheetName'];
        const type: string = getEventHandler.e.parameter['type'];

        const sheet: GoogleAppsScript.Spreadsheet.Sheet = this.getSheetByName(sheetName, type);
        getEventHandler.result[sheetName] = sheet.getDataRange().getValues();
    }

    private deleteSheetData(getEventHandler: GetEventHandler): void {
        const sheetName: string = getEventHandler.e.parameter['sheetName'];
        const type: string = getEventHandler.e.parameter['type'];
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = this.getSheetByName(sheetName, type);
        const sheetValues = sheet.getDataRange().getValues();
        const id = getEventHandler.e.parameter['id'];

        const rowIndex = sheetValues.findIndex(row => row[0] === id);
        if (rowIndex !== -1) {
            sheet.deleteRow(rowIndex + 1);
        } else {
            throw new Error(`Data Not Found in ${sheetName}, id: ${id}`);
        }
    }

    private getSheetByName(sheetName: string, type: string): GoogleAppsScript.Spreadsheet.Sheet {
        let ss: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
        let sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
        if (type === 'setting') {
            ss = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        } else if (type === 'report') {
            ss = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        } else {
            ss = SpreadsheetApp.openById(ScriptProps.instance.reportSheet);
        }
        sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            throw new Error(`Sheet '${sheetName}' was not found. type: ${type}`);
        }
        return sheet;
    }

    public loadVideoFolders(getEventHander: GetEventHandler): void {
        const rootFolder = DriveApp.getFolderById(ScriptProps.instance.videoFolder);
        const titleFolderIt: GoogleAppsScript.Drive.FolderIterator = rootFolder.getFolders();
        const results = [];
        while (titleFolderIt.hasNext()) {
            const folder: GoogleAppsScript.Drive.Folder = titleFolderIt.next();
            const title = folder.getName();
            const url = folder.getUrl(); // フォルダのURLを取得
            const folderId = folder.getId(); // フォルダIDを取得

            // フォルダ内のファイル情報を取得
            const filesIterator = folder.getFiles();
            const fileNames: string[] = [];
            let fileCount = 0;

            while (filesIterator.hasNext()) {
                const file = filesIterator.next();
                fileNames.push(file.getName());
                fileCount++;
            }

            // サブフォルダの情報を取得
            const subFoldersIterator = folder.getFolders();
            const subFolderNames: string[] = [];
            let hasResultFolder = false;
            let hasYouTubeFolder = false;

            while (subFoldersIterator.hasNext()) {
                const subFolder = subFoldersIterator.next();
                const subFolderName = subFolder.getName();
                subFolderNames.push(subFolderName);

                if (subFolderName === 'result') {
                    hasResultFolder = true;
                }
                if (subFolderName === 'YouTube') {
                    hasYouTubeFolder = true;
                }
            }

            // フォルダの作成日時と更新日時を取得
            const createdTime = folder.getDateCreated().toISOString();
            const modifiedTime = folder.getLastUpdated().toISOString();

            results.push({
                id: folderId,
                title: title,
                url: url,
                createdTime: createdTime,
                modifiedTime: modifiedTime,
                fileCount: fileCount,
                fileNames: fileNames,
                subFolderNames: subFolderNames,
                hasResultFolder: hasResultFolder,
                hasYouTubeFolder: hasYouTubeFolder,
            });
        }
        getEventHander.result = { videoFolders: results };
    }

    public executeVideoTask(getEventHandler: GetEventHandler): void {
        const folderId: string = getEventHandler.e.parameter['folderId'];
        const folderName: string = getEventHandler.e.parameter['folderName'];
        const taskType: string = getEventHandler.e.parameter['taskType'];

        try {
            // Google Colabのノートブックファイルを取得
            const colabFolderId = '1Yr4NsedItfew0cQSeG2vZl8kJFdVZL4i';
            const colabFolder = DriveApp.getFolderById(colabFolderId);

            // タスクタイプに応じたnotebookファイル名とメッセージを設定
            const taskConfig = this.getTaskConfig(taskType);

            // 指定されたnotebookを検索
            const notebookFiles = colabFolder.getFilesByName(taskConfig.notebookName);

            if (notebookFiles.hasNext()) {
                const notebookFile = notebookFiles.next();
                const notebookId = notebookFile.getId();

                // サービスアカウントの認証情報をColabに渡すためのパラメータ
                const serviceAccountEmail = 'shoot-sunday-sg@colablogics.iam.gserviceaccount.com';
                // const encodedFolderId = encodeURIComponent(folderId);
                // const encodedFolderName = encodeURIComponent(folderName);

                // Colabで開くためのURLを生成（必要なパラメータを含む）
                const colabUrl = `https://colab.research.google.com/drive/${notebookId}`;

                // 処理対象フォルダの情報をログに記録
                console.log(`${taskConfig.logMessage} started for folder: ${folderName} (${folderId})`);

                getEventHandler.result = {
                    success: true,
                    colabUrl: colabUrl,
                    message: `${taskConfig.successMessage}: ${folderName}`,
                    folderName: folderName,
                    folderId: folderId,
                    taskType: taskType,
                    serviceAccountEmail: serviceAccountEmail,
                };
            } else {
                throw new Error(`${taskConfig.errorMessage}: ${taskConfig.notebookName}`);
            }
        } catch (error) {
            console.error(`Video task (${taskType}) error:`, error);
            getEventHandler.result = {
                success: false,
                error: error instanceof Error ? error.message : String(error),
                message: `${taskType}処理でエラーが発生しました: ${folderName}`,
                taskType: taskType,
            };
        }
    }

    private getTaskConfig(taskType: string): {
        notebookName: string;
        logMessage: string;
        successMessage: string;
        errorMessage: string;
    } {
        switch (taskType) {
            case 'goal':
                return {
                    notebookName: 'videoGoal.ipynb',
                    logMessage: 'Goal video creation',
                    successMessage: 'ゴール集作成を開始しました',
                    errorMessage: 'ゴール集作成用のnotebookが見つかりません',
                };
            case 'merge':
                return {
                    notebookName: 'videoMerge.ipynb',
                    logMessage: 'Video merge',
                    successMessage: '動画統合を開始しました',
                    errorMessage: '動画統合用のnotebookが見つかりません',
                };
            // case 'upload':
            //     return {
            //         notebookName: 'youtube_uploader.ipynb',
            //         logMessage: 'YouTube upload',
            //         successMessage: 'YouTubeアップロードを開始しました',
            //         errorMessage: 'YouTubeアップロード用のnotebookが見つかりません'
            //     };
            default:
                throw Error('task typeが指定されていません');
        }
    }
}
