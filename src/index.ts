import { GasUtil } from './gasUtil';
import { LineUtil } from './lineUtil';
import { COMMAND_MAP, PostEventHandler } from './postEventHandler';
import { RequestExecuter } from './requestExecuter';

const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doGet(e: GoogleAppsScript.Events.DoGet): GoogleAppsScript.Content.TextOutput {
  return ContentService.createTextOutput('Hello World');
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
  const requestExecuter: RequestExecuter = new RequestExecuter();
  const postEventHander: PostEventHandler = new PostEventHandler(e);
  for (const item of COMMAND_MAP) {
    if (item.condition(postEventHander)) {
      executeMethod(requestExecuter, item.func, postEventHander);
    }
  }
  if (postEventHander.isFlex) {
    lineUtil.sendFlexReply(postEventHander.replyToken, postEventHander.messageJson);
  } else {
    lineUtil.sendLineReply(postEventHander.replyToken, postEventHander.resultMessage, postEventHander.resultImage);
  }
  if (postEventHander.paynowOwnerMsg) {
    lineUtil.sendLineMessage(gasUtil.getLineUserId(gasUtil.getDensukeName(gasUtil.getPaynowOwner())), postEventHander.paynowOwnerMsg);
  }
  return ContentService.createTextOutput(JSON.stringify({ content: 'post ok' })).setMimeType(ContentService.MimeType.JSON);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function executeMethod(obj: any, methodName: string, args: any) {
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
