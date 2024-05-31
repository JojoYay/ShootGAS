import { ScriptProps } from './scriptProps';
import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { LineUtil } from './lineUtil';
import { GasUtil } from './gasUtil';

const densukeUtil: DensukeUtil = new DensukeUtil();
const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();
const ROWNUM: number = ScriptProps.instance.ROWNUM; //とりあえず一番上からデータとってくる運用

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function testeeee(): void {
  // console.log(FOLDER_ID);
  // console.log(ARCHIVE_FOLDER);
  // console.log(CHANNEL_QR);
  // console.log(CHANNEL_URL);
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doGet(
  e: GoogleAppsScript.Events.DoGet
): GoogleAppsScript.Content.TextOutput {
  console.log(e);
  return ContentService.createTextOutput('Hello World');
}

interface Event {
  type: string;
  message: {
    type: string;
    text?: string;
  };
}

type Command = {
  func: string;
  condition: (event: Event) => boolean;
};

const COMMAND_MAP: Command[] = [
  {
    func: 'payNow',
    condition: (event: Event) =>
      event.type === 'message' && event.message.type === 'image',
  },
  {
    func: 'aggregate',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '集計' || event.message.text === 'aggregate'),
  },
  {
    func: 'unpaid',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '未払い' || event.message.text === 'unpaid'),
  },
  {
    func: 'densukeUpd',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '伝助更新' || event.message.text === 'update'),
  },
  {
    func: 'remind',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === 'リマインド' || event.message.text === 'remind'),
  },
  {
    func: 'intro',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '紹介' || event.message.text === 'introduce'),
  },
  {
    func: 'regInfo',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '登録' ||
        event.message.text === '@@register@@' ||
        event.message.text === 'how to register'),
  },
  {
    func: 'managerInfo',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '管理' || event.message.text === 'manage'),
  },
  {
    func: 'register',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text ?? '').startsWith('@@register@@'),
  },
];

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(
  e: GoogleAppsScript.Events.DoPost
): GoogleAppsScript.Content.TextOutput {
  const requestExecuter = new RequestExecuter();
  const json = JSON.parse(e.postData.contents);
  const event = json.events[0]; //CHATGPTがコレでよいと言いやがったけどいいのかな
  let done = false;
  for (const item of COMMAND_MAP) {
    if (item.condition(event)) {
      done = true;
      executeMethod(requestExecuter, item.func, json);
    }
  }
  if (!done) {
    errorMessage(json);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ content: 'post ok' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function executeMethod(obj: any, methodName: string, args: any) {
  if (typeof obj[methodName] === 'function') {
    return obj[methodName](args);
  } else {
    //こいつは基本的にCOMMAND_MAPに指定したメソッド名がRequestExecuterに存在する場合は発生しない（ので無視してよい）
    throw new Error(
      `Method ${methodName} does not exist on the object ${obj}.`
    );
  }
}

function sendMessageToPaynowOwner(message: string): void {
  lineUtil.sendLineMessage(
    gasUtil.getLineUserId(gasUtil.getDensukeName(gasUtil.getPaynowOwner())),
    message
  );
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function errorMessage(json: any): void {
  const userId: string = json.events[0].source.userId;
  const lang: string = lineUtil.getLineLang(userId);
  const replyToken: string = json.events[0].replyToken;
  let reply: string = '';
  if (lang === 'ja') {
    reply =
      '【エラー】申し訳ありません、理解できませんでした。再度正しく入力してください。';
  } else {
    reply =
      "【Error】I'm sorry, I didn't understand. Please enter the correct input again.";
  }
  lineUtil.sendLineReply(replyToken, reply, '');
}

class RequestExecuter {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  intro(json: any): void {
    const replyToken = json.events[0].replyToken;
    lineUtil.sendLineReply(
      replyToken,
      ScriptProps.instance.channelUrl,
      ScriptProps.instance.channelQr
    );
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  register(json: any): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const event: any = json.events[0];
    const replyToken = event.replyToken;
    const userMessage = event.message.text;
    const userId = event.source.userId;
    const lang = lineUtil.getLineLang(userId);
    const lineName = lineUtil.getLineDisplayName(userId);
    const $ = densukeUtil.getDensukeCheerio();
    const members = densukeUtil.extractMembers($);
    const actDate = densukeUtil.extractDateFromRownum($, ROWNUM);
    const densukeNameNew = userMessage.split('@@register@@')[1];
    let replyMessage = null;
    if (members.includes(densukeNameNew)) {
      if (hasMultipleOccurrences(members, densukeNameNew)) {
        if (lang === 'ja') {
          replyMessage =
            '伝助上で"' +
            densukeNameNew +
            '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。';
        } else {
          replyMessage =
            "There are multiple entries with the name '" +
            densukeNameNew +
            "' on Densuke. Please update it to a unique name and register again.";
        }
      } else {
        gasUtil.registerMapping(lineName, densukeNameNew, userId);
        gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
        if (lang === 'ja') {
          replyMessage =
            '伝助名称登録が完了しました。\n伝助上の名前：' +
            densukeNameNew +
            '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。\n' +
            userId;
        } else {
          replyMessage =
            'The initial registration is complete.\nYour name in Densuke: ' +
            densukeNameNew +
            "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.\n" +
            userId;
        }
      }
    } else {
      if (lang === 'ja') {
        replyMessage =
          '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' +
          densukeNameNew;
      } else {
        replyMessage =
          '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
          densukeNameNew;
      }
    }
    lineUtil.sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  payNow(json: any): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const event: any = json.events[0];
    const replyToken = event.replyToken;
    const $ = densukeUtil.getDensukeCheerio();
    const members = densukeUtil.extractMembers($);
    const attendees = densukeUtil.extractAttendees($, ROWNUM, '○', members);
    const actDate = densukeUtil.extractDateFromRownum($, ROWNUM);
    const messageId = event.message.id;
    const userId = event.source.userId;
    const lineName = lineUtil.getLineDisplayName(userId);
    const lang = lineUtil.getLineLang(userId);
    const densukeName = gasUtil.getDensukeName(lineName);
    let replyMessage = null;
    if (densukeName) {
      if (attendees.includes(densukeName)) {
        gasUtil.uploadPayNowPic(lineName, messageId, actDate);
        gasUtil.updatePaymentStatus(lineName, actDate);
        if (lang === 'ja') {
          replyMessage =
            actDate +
            'の支払いを登録しました。ありがとうございます！\n' +
            GasProps.instance.ReportSheetUrl;
        } else {
          replyMessage =
            'Payment for ' +
            actDate +
            ' has been registered. Thank you!\n' +
            GasProps.instance.ReportSheetUrl;
        }
      } else {
        if (lang === 'ja') {
          replyMessage =
            actDate +
            '【エラー】当日の伝助の出席が〇になっていませんでした。伝助を更新して、「伝助更新」と入力してください。\n' +
            densukeUtil.getDensukeUrl();
        } else {
          replyMessage =
            '【Error】Your attendance on ' +
            actDate +
            " in Densuke has not been marked as 〇.\nPlease update Densuke and type 'update'.\n" +
            densukeUtil.getDensukeUrl();
        }
      }
    } else {
      if (lang === 'ja') {
        replyMessage =
          '【エラー】伝助名称登録が完了していません。\n登録を完了させて、再度PayNow画像をアップロードして下さい。\n登録は「登録」と入力してください。\n' +
          densukeUtil.getDensukeUrl();
      } else {
        replyMessage =
          "【Error】The initial registration is not complete.\nPlease complete the initial registration and upload the PayNow photo again.\nFor the initial registration, please type 'how to register'.\n" +
          densukeUtil.getDensukeUrl();
      }
    }
    lineUtil.sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  aggregate(json: any): void {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = densukeUtil.getDensukeCheerio();
    const members = densukeUtil.extractMembers($);
    const attendees = densukeUtil.extractAttendees($, ROWNUM, '○', members);
    const actDate = densukeUtil.extractDateFromRownum($, ROWNUM);
    const settingSheet = GasProps.instance.settingSheet;
    const addy = settingSheet.getRange('B2').getValue();
    densukeUtil.generateSummaryBase($);
    lineUtil.sendLineReply(
      replyToken,
      densukeUtil.getSummaryStr(attendees, actDate, addy),
      ''
    );
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  unpaid(json: any): void {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = densukeUtil.getDensukeCheerio();
    const actDate = densukeUtil.extractDateFromRownum($, ROWNUM);
    const unpaid = gasUtil.getUnpaid(actDate);
    lineUtil.sendLineReply(
      replyToken,
      '未払いの人 (' + unpaid.length + '名): ' + unpaid.join(', ') + '\n',
      ''
    );
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  remind(json: any): void {
    const replyToken = json.events[0].replyToken;
    lineUtil.sendLineReply(replyToken, densukeUtil.generateRemind(), '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  densukeUpd(json: any): void {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = densukeUtil.getDensukeCheerio();
    const userId = event.source.userId;
    const lang = lineUtil.getLineLang(userId);
    const lineName = lineUtil.getLineDisplayName(userId);
    const members = densukeUtil.extractMembers($);
    const attendees = densukeUtil.extractAttendees($, ROWNUM, '○', members);
    const actDate = densukeUtil.extractDateFromRownum($, ROWNUM);
    const settingSheet = GasProps.instance.settingSheet;
    const addy = settingSheet.getRange('B2').getValue();

    densukeUtil.generateSummaryBase($);
    const ownerMessage =
      '【' +
      lineName +
      'さんにより更新されました】\n' +
      densukeUtil.getSummaryStr(attendees, actDate, addy);
    sendMessageToPaynowOwner(ownerMessage);
    let replyMessage = null;
    if (lang === 'ja') {
      replyMessage =
        '伝助の更新ありがとうございました！PayNowのスクリーンショットを再度こちらへ送って下さい。';
    } else {
      replyMessage =
        'Thank you for updating Densuke! Please send PayNow screenshot here again.';
    }
    lineUtil.sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  regInfo(json: any): void {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;
    const lang = lineUtil.getLineLang(userId);
    let replyMessage = null;
    if (lang === 'ja') {
      replyMessage =
        '伝助名称の登録を行います。\n伝助のアカウント名を以下のフォーマットで入力してください。\n@@register@@伝助名前\n例）@@register@@やまだじょ\n' +
        densukeUtil.getDensukeUrl();
    } else {
      replyMessage =
        'We will perform the densuke name registration.\nPlease enter your Densuke account name in the following format:\n@@register@@XXXXX\nExample)@@register@@Sahim\n' +
        densukeUtil.getDensukeUrl();
    }
    lineUtil.sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  managerInfo(json: any): void {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;
    let replyMessage = null;
    if (gasUtil.isKanji(userId)) {
      replyMessage =
        '設定：' +
        GasProps.instance.settingSheetUrl +
        '\nPayNow：' +
        GasProps.instance.payNowFolderUrl;
      '\nReport URL:' + GasProps.instance.ReportSheetUrl;
      '\n伝助：' +
        densukeUtil.getDensukeUrl() +
        '\nチャット状況：' +
        ScriptProps.instance.chat;
      '\nメッセージ利用状況：' +
        ScriptProps.instance.messageUsage +
        '\n 利用可能コマンド:集計, 紹介, 登録, リマインド, 伝助更新, 未払い, @@register@@名前 ';
    } else {
      replyMessage = 'えっ！？このコマンドは平民のキミには内緒だよ！';
    }
    lineUtil.sendLineReply(replyToken, replyMessage, '');
  }
}

function hasMultipleOccurrences(
  array: string[],
  searchString: string
): boolean {
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
