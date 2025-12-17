import { GasProps } from './gasProps';
import { GetEventHandler } from './getEventHandler';
// import { GoogleCalendar } from './googleCalendar';
import { LiffApi } from './liffApi';
import { LineUtil } from './lineUtil';
import { COMMAND_MAP, PostEventHandler } from './postEventHandler';
import { RequestExecuter } from './requestExecuter';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function calendarWork() {
    const re: RequestExecuter = new RequestExecuter();
    re.work1();
}
// function syncAllToCalendar() {
//     const g: GoogleCalendar = new GoogleCalendar();
//     g.syncAllToCalendar();
// }

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
                    if (prof.pictureUrl) {
                        GasProps.instance.mappingSheet.getRange(index + 1, 5).setValue(prof.pictureUrl);
                    } else if (prof.userName === 'yagisho') {
                        GasProps.instance.mappingSheet
                            .getRange(index + 1, 5)
                            .setValue('https://lh3.googleusercontent.com/d/1Qc2YASdbIoikuT5bNcdnObhznf-WD5rK');
                    }
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
    let getEventHandler: GetEventHandler;
    try {
        getEventHandler = new GetEventHandler(e);
        const liffApi = new LiffApi();
        for (const methodName of getEventHandler.funcs) {
            executeMethod(liffApi, methodName, getEventHandler);
        }
        // jsonで始まるキーの値はすでにJSONオブジェクトなので、そのまま使用（stringify時に正しく処理される）
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const processedResult: any = {};
        for (const key in getEventHandler.result) {
            if (key.startsWith('json')) {
                // jsonで始まるキーはすでにJSONオブジェクト（またはJSON文字列）なので、そのまま使用
                if (typeof getEventHandler.result[key] === 'string') {
                    // 文字列の場合はJSONオブジェクトに変換
                    try {
                        processedResult[key] = JSON.parse(getEventHandler.result[key]);
                        // eslint-disable-next-line @typescript-eslint/no-unused-vars
                    } catch (e) {
                        // JSON文字列として解析できない場合はそのまま
                        processedResult[key] = getEventHandler.result[key];
                    }
                } else {
                    // オブジェクト/配列の場合はそのまま使用
                    processedResult[key] = getEventHandler.result[key];
                }
            } else {
                processedResult[key] = getEventHandler.result[key];
            }
        }
        return ContentService.createTextOutput(JSON.stringify(processedResult)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        console.log(err);
        return ContentService.createTextOutput(JSON.stringify({ err: (err as Error).message, stacktrace: (err as Error).stack })).setMimeType(
            ContentService.MimeType.JSON
        );
    }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
    const lineUtil: LineUtil = new LineUtil();
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
        }
    } catch (err) {
        console.log(err);
        postEventHander.resultMessage = '[Error] ' + (err as Error).message + '\n' + (err as Error).stack;
        postEventHander.reponseObj.err = '[Error] ' + (err as Error).message + '\n' + (err as Error).stack;
        if (postEventHander.replyToken) {
            lineUtil.sendLineReply(postEventHander.replyToken, postEventHander.resultMessage, null);
        }
        return ContentService.createTextOutput(JSON.stringify(postEventHander.reponseObj)).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(postEventHander.reponseObj)).setMimeType(ContentService.MimeType.JSON);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function executeMethod(obj: any, methodName: string, args: any) {
    // console.log('Execute ' + methodName);
    if (typeof obj[methodName] === 'function') {
        return obj[methodName](args);
    } else {
        //こいつは基本的にCOMMAND_MAPに指定したメソッド名がRequestExecuterに存在する場合は発生しない（ので無視してよい）
        throw new Error(`Method ${methodName} does not exist on the object ${obj}.`);
    }
}
