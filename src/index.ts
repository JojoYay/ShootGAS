import { GasUtil } from './gasUtil';
import { LineUtil } from './lineUtil';
import { PostEventHandler } from './postEventHandler';
import { RequestExecuter } from './requestExecuter';

const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doGet(e: GoogleAppsScript.Events.DoGet): GoogleAppsScript.Content.TextOutput {
  return ContentService.createTextOutput('Hello World');
}

type Command = {
  func: string;
  condition: (postEventHander: PostEventHandler) => boolean;
};

const COMMAND_MAP: Command[] = [
  {
    func: 'payNow',
    condition: (postEventHander: PostEventHandler) => postEventHander.type === 'message' && postEventHander.messageType === 'image',
  },
  {
    func: 'aggregate',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '集計' || postEventHander.messageText === 'aggregate'),
  },
  {
    func: 'unpaid',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '未払い' || postEventHander.messageText === 'unpaid'),
  },
  {
    func: 'densukeUpd',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '伝助更新' || postEventHander.messageText === 'update'),
  },
  {
    func: 'remind',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === 'リマインド' || postEventHander.messageText === 'remind'),
  },
  {
    func: 'intro',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '紹介' || postEventHander.messageText === 'introduce'),
  },
  {
    func: 'regInfo',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '登録' || postEventHander.messageText === '@@register@@' || postEventHander.messageText === 'how to register'),
  },
  {
    func: 'managerInfo',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' &&
      postEventHander.messageType === 'text' &&
      (postEventHander.messageText === '管理' || postEventHander.messageText === 'manage'),
  },
  {
    func: 'register',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' && postEventHander.messageType === 'text' && postEventHander.messageText.startsWith('@@register@@'),
  },
  {
    func: 'systemTest',
    condition: (postEventHander: PostEventHandler) =>
      postEventHander.type === 'message' && postEventHander.messageType === 'text' && postEventHander.messageText.startsWith('システムテスト'),
  },
];

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
  const requestExecuter: RequestExecuter = new RequestExecuter();
  const postEventHander: PostEventHandler = new PostEventHandler(e);
  for (const item of COMMAND_MAP) {
    if (item.condition(postEventHander)) {
      executeMethod(requestExecuter, item.func, postEventHander);
    }
  }
  lineUtil.sendLineReply(postEventHander.replyToken, postEventHander.resultMessage, postEventHander.resultImage);
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
