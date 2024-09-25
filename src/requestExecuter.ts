import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasTestSuite } from './gasTestSuite';
import { GasUtil } from './gasUtil';
import { LineUtil } from './lineUtil';
import { PostEventHandler } from './postEventHandler';
import { ScoreBook, Title } from './scoreBook';
import { ScriptProps } from './scriptProps';

const densukeUtil: DensukeUtil = new DensukeUtil();
const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();

export class RequestExecuter {
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

    public register(postEventHander: PostEventHandler): void {
        const lineName = lineUtil.getLineDisplayName(postEventHander.userId);
        const $ = densukeUtil.getDensukeCheerio();
        const members = densukeUtil.extractMembers($);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const densukeNameNew = postEventHander.messageText.split('@@register@@')[1];
        if (members.includes(densukeNameNew)) {
            if (densukeUtil.hasMultipleOccurrences(members, densukeNameNew)) {
                if (postEventHander.lang === 'ja') {
                    postEventHander.resultMessage =
                        '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。';
                } else {
                    postEventHander.resultMessage =
                        "There are multiple entries with the name '" +
                        densukeNameNew +
                        "' on Densuke. Please update it to a unique name and register again.";
                }
            } else {
                gasUtil.registerMapping(lineName, densukeNameNew, postEventHander.userId);
                gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
                this.updateProfilePic();
                if (postEventHander.lang === 'ja') {
                    postEventHander.resultMessage =
                        '伝助名称登録が完了しました。\n伝助上の名前：' +
                        densukeNameNew +
                        '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。\n' +
                        postEventHander.userId;
                } else {
                    postEventHander.resultMessage =
                        'The initial registration is complete.\nYour name in Densuke: ' +
                        densukeNameNew +
                        "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.\n" +
                        postEventHander.userId;
                }
            }
        } else {
            if (postEventHander.lang === 'ja') {
                postEventHander.resultMessage =
                    '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' + densukeNameNew;
            } else {
                postEventHander.resultMessage =
                    '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
                    densukeNameNew;
            }
        }
    }

    private updateProfilePic() {
        // const lineUtil: LineUtil = new LineUtil();
        const densukeMappingVals = GasProps.instance.mappingSheet.getDataRange().getValues();
        let index: number = 0;
        for (const userRow of densukeMappingVals) {
            if (userRow[0] !== 'ライン上の名前') {
                const userId: string = userRow[2];
                try {
                    const prof = lineUtil.getLineUserProfile(userId);
                    if (prof) {
                        // console.log(userRow[0] + ': ' + prof.pictureUrl);
                        GasProps.instance.mappingSheet.getRange(index + 1, 5).setValue(prof.pictureUrl);
                    }
                    // eslint-disable-next-line @typescript-eslint/no-unused-vars
                } catch (e) {
                    console.log(userRow[0] + ': invalid UserId' + userId);
                }
            }
            index++;
        }
        return;
    }

    public payNow(postEventHander: PostEventHandler): void {
        const $ = densukeUtil.getDensukeCheerio();
        const members = densukeUtil.extractMembers($);
        const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
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
                        'の伝助の出席が〇になっていませんでした。伝助を更新して、「伝助更新」と入力してください。\n' +
                        densukeUtil.getDensukeUrl();
                } else {
                    postEventHander.resultMessage =
                        '【Error】Your attendance on ' +
                        actDate +
                        " in Densuke has not been marked as 〇.\nPlease update Densuke and type 'update'.\n" +
                        densukeUtil.getDensukeUrl();
                }
            }
        } else {
            if (postEventHander.lang === 'ja') {
                postEventHander.resultMessage =
                    '【エラー】伝助名称登録が完了していません。\n登録を完了させて、再度PayNow画像をアップロードして下さい。\n登録は「登録」と入力してください。\n' +
                    densukeUtil.getDensukeUrl();
            } else {
                postEventHander.resultMessage =
                    "【Error】The initial registration is not complete.\nPlease complete the initial registration and upload the PayNow photo again.\nFor the initial registration, please type 'how to register'.\n" +
                    densukeUtil.getDensukeUrl();
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
        let $ = densukeUtil.getDensukeCheerio();
        if (postEventHander.mockDensukeCheerio) {
            $ = postEventHander.mockDensukeCheerio;
        }
        const members = densukeUtil.extractMembers($);
        const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        densukeUtil.generateSummaryBase($);
        postEventHander.resultMessage = densukeUtil.getSummaryStr(attendees, actDate, addy);
    }

    public unpaid(postEventHander: PostEventHandler): void {
        const $ = densukeUtil.getDensukeCheerio();
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const unpaid = gasUtil.getUnpaid(actDate);
        postEventHander.resultMessage = '未払いの人 (' + unpaid.length + '名): ' + unpaid.join(', ');
    }

    public remind(postEventHander: PostEventHandler): void {
        postEventHander.resultMessage = densukeUtil.generateRemind();
    }

    public densukeUpd(postEventHander: PostEventHandler): void {
        const $ = densukeUtil.getDensukeCheerio();
        const lineName = lineUtil.getLineDisplayName(postEventHander.userId);
        const members = densukeUtil.extractMembers($);
        const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const settingSheet = GasProps.instance.settingSheet;
        const addy = settingSheet.getRange('B2').getValue();
        densukeUtil.generateSummaryBase($);
        postEventHander.paynowOwnerMsg = '【' + lineName + 'さんにより更新されました】\n' + densukeUtil.getSummaryStr(attendees, actDate, addy);
        // this.sendMessageToPaynowOwner(ownerMessage);
        if (postEventHander.lang === 'ja') {
            postEventHander.resultMessage = '伝助の更新ありがとうございました！PayNowのスクリーンショットを再度こちらへ送って下さい。';
        } else {
            postEventHander.resultMessage = 'Thank you for updating Densuke! Please send PayNow screenshot here again.';
        }
    }

    public regInfo(postEventHander: PostEventHandler): void {
        if (postEventHander.lang === 'ja') {
            postEventHander.resultMessage =
                '伝助名称の登録を行います。\n伝助のアカウント名を以下のフォーマットで入力してください。\n@@register@@伝助名前\n例）@@register@@やまだじょ\n' +
                densukeUtil.getDensukeUrl();
        } else {
            postEventHander.resultMessage =
                'We will perform the densuke name registration.\nPlease enter your Densuke account name in the following format:\n@@register@@XXXXX\nExample)@@register@@Sahim\n' +
                densukeUtil.getDensukeUrl();
        }
    }

    public managerInfo(postEventHander: PostEventHandler): void {
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
                '\n伝助：' +
                densukeUtil.getDensukeUrl() +
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

    public unRegister(postEventHander: PostEventHandler) {
        this.aggregate(postEventHander);
        const $ = densukeUtil.getDensukeCheerio();
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const unRegister = gasUtil.getUnRegister(actDate);
        postEventHander.resultMessage = '現在未登録の参加者 (' + unRegister.length + '名): ' + unRegister.join(', ');
    }

    public ranking(postEventHander: PostEventHandler): void {
        const scoreBook: ScoreBook = new ScoreBook();
        const $ = densukeUtil.getDensukeCheerio();
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const members = densukeUtil.extractMembers($);
        const attendees = densukeUtil.extractAttendees($, ScriptProps.instance.ROWNUM, '○', members);

        scoreBook.makeEventFormat();
        scoreBook.aggregateScore();

        scoreBook.generateScoreBook(actDate, attendees, Title.ASSIST);
        scoreBook.generateScoreBook(actDate, attendees, Title.TOKUTEN);
        scoreBook.generateOkamotoBook(actDate, attendees);

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
