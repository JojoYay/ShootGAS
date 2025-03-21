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
                }
            }
            console.log('動画が見つかりませんでした。タイトル:', videoTitle);
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
            .filter(val => val[0] === actDate && (!val[10].endsWith('_g') || !val[10].endsWith('d')) && val[3] && val[4]);
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
                if (typeof val[10] !== 'string' || !val[10].endsWith('_g') || !val[10].endsWith('d')) {
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
            getEventHandler.result.statusMst3 =
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
                    getEventHandler.result.statusMst3 = '';
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
            getEventHandler.result.statusMst3 =
                lang === 'ja-JP'
                    ? '支払い済みのスクリーンショットをこちらにアップロードして下さい'
                    : 'Please upload a screenshot of your payment here';

            getEventHandler.result.picUrl = '';
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
        // 1. URLとt=XXXXの部分を抽出
        const urlRegex = /href="([^"]*watch\?v=([a-zA-Z0-9_-]+)[^"]*)"/;
        const timeRegex = /&amp;t=(\d+)/;

        const urlMatch = inputString.match(urlRegex);
        const timeMatch = inputString.match(timeRegex);

        if (!urlMatch || !timeMatch) {
            return inputString; // URLまたは時間が抽出できなかった場合は元の文字列を返す
        }

        // const fullUrl = urlMatch[1]; // URL全体
        const videoId = urlMatch[2]; // ビデオID
        const time = timeMatch[1]; // 時間

        // 2. 新しいURLを作成
        const newUrl = `https://youtu.be/${videoId}?t=${time}`;
        // const newUrl = `youtube://watch?v=${videoId}&t=${time}`;
        // 3. 元のURLを新しいURLで置き換える
        // const replacedString = inputString.replace(fullUrl, newUrl);
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
