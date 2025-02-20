import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { GetEventHandler } from './getEventHandler';
import { LiffApi } from './liffApi';
import { LineUtil } from './lineUtil';
import { COMMAND_MAP, PostEventHandler } from './postEventHandler';
import { RequestExecuter } from './requestExecuter';

// const lineUtil: LineUtil = new LineUtil();
// const gasUtil: GasUtil = new GasUtil();

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateProfilePic() {
    const lineUtil: LineUtil = new LineUtil();
    const densukeMappingVals = GasProps.instance.mappingSheet.getDataRange().getValues();
    let index: number = 0;
    for (const userRow of densukeMappingVals) {
        if (userRow[0] !== 'ライン上の名前') {
            const userId: string = userRow[2];
            try {
                const prof = lineUtil.getLineUserProfile(userId);
                if (prof) {
                    console.log(userRow[0] + ': ' + prof.pictureUrl);
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

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doGet(e: GoogleAppsScript.Events.DoGet): GoogleAppsScript.Content.TextOutput {
    // console.log(e);
    const getEventHandler: GetEventHandler = new GetEventHandler(e);
    for (const methodName of getEventHandler.funcs) {
        executeMethod(new LiffApi(), methodName, getEventHandler);
    }
    return ContentService.createTextOutput(JSON.stringify(getEventHandler.result));
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
    // console.log('sasdsdsadsafsadd');

    const lineUtil: LineUtil = new LineUtil();
    const gasUtil: GasUtil = new GasUtil();
    const requestExecuter: RequestExecuter = new RequestExecuter();
    const postEventHander: PostEventHandler = new PostEventHandler(e);
    try {
        for (const item of COMMAND_MAP) {
            if (item.condition(postEventHander)) {
                executeMethod(requestExecuter, item.func, postEventHander);
            }
        }
        if (postEventHander.type) {
            if (postEventHander.isFlex) {
                lineUtil.sendFlexReply(postEventHander.replyToken, postEventHander.messageJson);
            } else {
                lineUtil.sendLineReply(postEventHander.replyToken, postEventHander.resultMessage, postEventHander.resultImage);
            }
            if (postEventHander.paynowOwnerMsg) {
                lineUtil.sendLineMessage(gasUtil.getLineUserId(gasUtil.getDensukeName(gasUtil.getPaynowOwner())), postEventHander.paynowOwnerMsg);
            }
        }
    } catch (err) {
        postEventHander.resultMessage = '[Error] ' + (err as Error).message + '\n' + (err as Error).stack;
        lineUtil.sendLineReply(postEventHander.replyToken, postEventHander.resultMessage, null);
        throw err;
    }
    // const allowedOrigins = [
    //     'https://shootsundayfront.web.app',
    //     'https://test-8120f.web.app',
    //     'https://shootsundayfront.web.app',
    //     'https://test-8120f.web.app',
    // ];
    //   const origin = e.origin || e.requestHeaders.origin; // リクエストの Origin を取得 (GAS 環境によって取得方法が異なる場合あり)

    return ContentService.createTextOutput(JSON.stringify(postEventHander.reponseObj)).setMimeType(ContentService.MimeType.JSON);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function executeMethod(obj: any, methodName: string, args: any) {
    console.log('Execute ' + methodName);
    if (typeof obj[methodName] === 'function') {
        return obj[methodName](args);
    } else {
        //こいつは基本的にCOMMAND_MAPに指定したメソッド名がRequestExecuterに存在する場合は発生しない（ので無視してよい）
        throw new Error(`Method ${methodName} does not exist on the object ${obj}.`);
    }
}

// function errorMessage(postEventHander: PostEventHandler): void {
//   const userId: string = postEventHander.userId;
//   const lang: string = lineUtil.getLineLang(userId);
//   const replyToken: string = postEventHander.replyToken;
//   let reply: string = '';
//   if (lang === 'ja') {
//     reply =
//       '【エラー】申し訳ありません、理解できませんでした。再度正しく入力してください。';
//   } else {
//     reply =
//       "【Error】I'm sorry, I didn't understand. Please enter the correct input again.";
//   }
//   lineUtil.sendLineReply(replyToken, reply, '');
// }
