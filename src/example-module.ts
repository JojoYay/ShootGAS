const ROWNUM = 1; //とりあえず一番上からデータとってくる運用
const REPORT_SHEET = '1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU';
const SETTING_SHEET = '1RymgqowExhXB3p2N9i1OYdypIhCtMZkRWr1Wpo1CTyI';
const LINE_ACCESS_TOKEN =
  'glhgotcUd27uledjA6p02Rv4tkbaqJeUJzAVLeiGNB8Us7tomv7ZdZZ8cz1q+PmVch28XYoh0SlKqPZpwk7+SJkAjZz/CmELk4CNP/DNiKjRUrIQUGDEmhLL40rXDJ8e1lBubmGGRJfkhE0BkXH2CwdB04t89/1O/w1cDnyilFU=';
const FOLDER_ID = '1fcBWvaoAlWNJVmmH4s2spVi5_EJdE3aC';
const ARCHIVE_FOLDER = '1IdRNjbo752JhlATsDic3bGUZDEY5fk9B';
const CHANNEL_QR = 'https://qr-official.line.me/sid/L/164wqxlu.png';
const CHANNEL_URL = 'https://lin.ee/1O9VC3Q';
const SETTING_SHEET_NAME = 'Settings';
const CASH_BOOK_SHEET_NAME = 'CashBook';
const MAPPING_SHEET_NAME = 'DensukeMapping';

function generateRemind($ = getDensukeCheerio()): string {
  const members: string[] = exstractMembers($);
  const attendees: string[] = exstractAttendees($, ROWNUM, '○', members);
  const unknown: string[] = exstractAttendees($, ROWNUM, '△', members);
  const actDate: string = extractDateFromRownum($, ROWNUM);

  let remindStr: string =
    '次回予定' +
    actDate +
    'リマインドです！\n伝助の更新お忘れなく！\nThis is gentle reminder of ' +
    actDate +
    '.\nPlease update your Densuke schedule.\n';
  if (attendees.length < 10) {
    remindStr =
      '次回予定' +
      actDate +
      'がピンチです！\n参加できる方、ぜひ伝助で参加表明お願いします！！！\nThis is gentle reminder of ' +
      actDate +
      '.\nDue to the low number of participants, there is a possibility of cancellation.\nPlease join us!\n';
  }
  const summary: string =
    remindStr +
    '〇(' +
    attendees.length +
    '名): ' +
    attendees.join(', ') +
    '\n' +
    '△(' +
    unknown.length +
    '名): ' +
    unknown.join(', ') +
    '\n' +
    '伝助URL：' +
    getDensukeUrl();
  return summary;
}

function generateSummaryBase($ = getDensukeCheerio()) {
  // スプレッドシートとシートを取得
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
  const cashBook = ss.getSheetByName(CASH_BOOK_SHEET_NAME);
  const members: string[] = exstractMembers($);
  const attendees: string[] = exstractAttendees($, ROWNUM, '○', members);
  const actDate: string = extractDateFromRownum($, ROWNUM);
  if (!cashBook || !settingSheet) {
    return;
  }
  // データの範囲を取得
  const range = cashBook.getDataRange();
  const values = range.getValues();
  // 2カラム目（B列）に指定した値がある行を逆順に削除
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][1] === actDate) {
      // B列はインデックス1
      cashBook.deleteRow(i + 1);
    }
  }
  const lastRow: number = cashBook.getLastRow();
  const orgPrice: number = cashBook.getRange(lastRow, 5).getValue();
  const rentalFee: number = settingSheet.getRange('B3').getValue();
  const attendFee: number = settingSheet.getRange('B4').getValue();

  generateSummarySheet(orgPrice, rentalFee, attendFee, actDate, attendees);
}

function getDensukeUrl(): string {
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
  let url: string = '';
  if (settingSheet) {
    url = settingSheet.getRange('B1').getValue();
  }
  return url;
}
function getDensukeCheerio() {
  const url: string = getDensukeUrl();
  const html: string = UrlFetchApp.fetch(url).getContentText();
  const $ = Cheerio.load(html);
  return $;
}

function generateSummarySheet(
  orgPrice: number,
  rentalFee: number,
  attendFee: number,
  actDate: string,
  attendees: string[]
) {
  const attendFeeTotal: number = attendFee * attendees.length;
  const report: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(REPORT_SHEET);
  let logSheet: GoogleAppsScript.Spreadsheet.Sheet | null = report.getSheetByName(actDate);
  if (!logSheet) {
    logSheet = report.insertSheet(actDate);
  }
  archiveFiles(actDate);
  logSheet.activate();
  report.moveActiveSheet(1);
  logSheet.getRange('A1').setValue('日付');
  logSheet.getRange('B1').setValue(actDate);
  logSheet.getRange('A2').setValue('更新日付');
  logSheet.getRange('B2').setValue(getCurrentTime());
  logSheet.getRange('A3').setValue('繰り越し残高(SGD)');
  logSheet.getRange('B3').setValue('' + orgPrice);
  logSheet.getRange('A4').setValue('参加人数(人)');
  logSheet.getRange('B4').setValue('' + attendees.length);
  logSheet.getRange('A5').setValue('参加費合計(SGD))');
  logSheet.getRange('B5').setValue('' + attendFeeTotal);
  logSheet.getRange('A6').setValue('ピッチ使用料金(SGD)');
  logSheet.getRange('B6').setValue('' + rentalFee);
  logSheet.getRange('A7').setValue('余剰金残高(SGD)');
  logSheet
    .getRange('B7')
    .setValue('' + (orgPrice - rentalFee + attendFeeTotal));

  logSheet.getRange('A9').setValue('参加者（伝助名称）');
  logSheet.getRange('B9').setValue('参加者（Line名称）');
  logSheet.getRange('C9').setValue('支払い状況');

  const values = logSheet.getDataRange().getValues();
  for (let i = values.length; i >= 10; i--) {
    logSheet.deleteRow(i);
  }
  for (let i = 0; i < attendees.length; i++) {
    logSheet.appendRow([
      attendees[i],
      getLineName(attendees[i]),
      getPaymentUrl(getLineName(attendees[i]), actDate),
    ]);
  }
  logSheet.setColumnWidth(1, 170);
  logSheet.setColumnWidth(2, 200);
  const setting = SpreadsheetApp.openById(SETTING_SHEET);
  const cashBook = setting.getSheetByName(CASH_BOOK_SHEET_NAME);
  const date = new Date();
  const attendOrg = orgPrice + attendFeeTotal;
  if (cashBook) {
    cashBook.appendRow([
      date,
      actDate,
      '参加費(' + attendees.length + '名)',
      '' + attendFeeTotal,
      '' + attendOrg,
    ]);
    cashBook.appendRow([
      date,
      actDate,
      'ピッチ使用料金',
      '▲' + rentalFee,
      '' + (orgPrice - rentalFee + attendFeeTotal),
    ]);
  }
}

function getSummaryStr(
  attendees: string[],
  actDate: string,
  payNowAddy: string
): string {
  let paynowStr = '';
  if (getUnpaid(actDate).length === 0) {
    paynowStr =
      '入金ありがとうございました。今回のレポートになります。詳細はリンクをご確認下さい。\nThank you for your payment.\nPlease find the report for this transaction below.\nFor more details, please check the provided link.\n';
  } else {
    paynowStr =
      'みなさま、ご参加ありがとうございました。\n入金後PayNowのスクリーンショットをSundayShootちゃんねるに送信して下さい。\nThank you all for your paticipation! After making the payment, please send the PayNow screenshot to Sunday Shoot Line Channel.\n';
  }

  const summary =
    paynowStr +
    '[' +
    actDate +
    ']のReport\n' +
    '参加者 participants (' +
    attendees.length +
    '名): ' +
    attendees.join(', ') +
    '\n' +
    'Report URL:https://docs.google.com/spreadsheets/d/1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU/edit?usp=sharing&cache=' +
    new Date().getTime() +
    ' \nPayNow先:' +
    payNowAddy;
  return summary;
}

function archiveFiles(actDate: string) {
  try {
    const sourceFolder = DriveApp.getFolderById(FOLDER_ID);
    const destinationFolder = DriveApp.getFolderById(ARCHIVE_FOLDER);
    const files = sourceFolder.getFiles();
    const prefix = actDate + '_';
    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().startsWith(prefix)) {
        file.moveTo(destinationFolder);
      }
    }
  } catch (e: unknown) {
    Logger.log('Error: ' + (e as Error).message);
  }
}

// function doGet(
//   e: GoogleAppsScript.Events.DoGet
// ): GoogleAppsScript.Content.TextOutput {
//   return ContentService.createTextOutput('Hello World');
// }

// interface ICommand {
//   func: string;
//   condition: (event: any) => boolean;
// }

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
      event.message.text === '集計',
  },
  {
    func: 'unpaid',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      event.message.text === '未払い',
  },
  {
    func: 'densukeUpd',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      event.message.text === '伝助更新',
  },
  {
    func: 'remind',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      event.message.text === 'リマインド',
  },
  {
    func: 'intro',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      event.message.text === '紹介',
  },
  {
    func: 'regInfo',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      (event.message.text === '登録' || event.message.text === '@@register@@'),
  },
  {
    func: 'managerInfo',
    condition: (event: Event) =>
      event.type === 'message' &&
      event.message.type === 'text' &&
      event.message.text === '管理',
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

  for (const item of COMMAND_MAP) {
    if (item.condition(event)) {
      executeMethod(requestExecuter, item.func, json);
    } else {
      errorMessage(json);
    }
  }
  return ContentService.createTextOutput(
    JSON.stringify({ content: 'post ok' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function executeMethod(obj: any, methodName: string, args: []) {
  if (typeof obj[methodName] === 'function') {
    return obj[methodName](...args);
  } else {
    //こいつは基本的にCOMMAND_MAPに指定したメソッド名がRequestExecuterに存在する場合は発生しない（ので無視してよい）
    throw new Error(
      `Method ${methodName} does not exist on the object ${obj}.`
    );
  }
}

function sendMessageToPaynowOwner(message: string) {
  sendLineMessage(getLineUserId(getDensukeName(getPaynowOwner())), message);
}

function getPaynowOwner() {
  const setting = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = setting.getSheetByName(SETTING_SHEET_NAME);
  if (!settingSheet) {
    return;
  }
  const payNowOwner = settingSheet.getRange('B6').getValue();
  return payNowOwner;
}

function sendLineMessage(userId: string, message: string) {
  if (userId) {
    const url = 'https://api.line.me/v2/bot/message/push';
    const headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
    };
    const postData = {
      to: userId,
      messages: [
        {
          type: 'text',
          text: message,
        },
      ],
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      headers: headers,
      payload: JSON.stringify(postData),
    };
    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log(response.getContentText());
    } catch (e) {
      Logger.log('Error: ' + e);
    }
  }
}
function sendLineReply(
  replyToken: string,
  messageText: string,
  imageUrl: string
) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
  };
  const postData = {
    replyToken: replyToken,
    messages: [
      {
        type: 'text',
        text: messageText,
      },
    ],
  };
  if (imageUrl) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const imgObj: any = {
      type: 'image',
      originalContentUrl: imageUrl,
      previewImageUrl: imageUrl,
    };
    postData.messages.push(imgObj);
  }
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(postData),
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText());
  } catch (e) {
    Logger.log('Error: ' + e);
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function errorMessage(json: any) {
  const userId: string = json.events[0].source.userId;
  const lang: string = getLineLang(userId);
  const replyToken: string = json.events[0].replyToken;
  let reply: string = '';
  if (lang === 'ja') {
    reply =
      '【エラー】申し訳ありません、理解できませんでした。再度正しく入力してください。';
  } else {
    reply =
      "【Error】I'm sorry, I didn't understand. Please enter the correct input again.";
  }
  sendLineReply(replyToken, reply, '');
}

class RequestExecuter {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  intro(json: any) {
    const replyToken = json.events[0].replyToken;
    sendLineReply(replyToken, CHANNEL_URL, CHANNEL_QR);
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  register(json: any) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const event: any = json.events[0];
    const replyToken = event.replyToken;
    const userMessage = event.message.text;
    const userId = event.source.userId;
    const lang = getLineLang(userId);
    const lineName = getLineDisplayName(userId);
    const $ = getDensukeCheerio();
    const members = exstractMembers($);
    const actDate = extractDateFromRownum($, ROWNUM);
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
        registerMapping(lineName, densukeNameNew, userId);
        updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
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
    sendLineReply(replyToken, replyMessage, '');

    return [replyMessage, null];
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  payNow(json: any) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const event: any = json.events[0];
    const replyToken = event.replyToken;
    const $ = getDensukeCheerio();
    const members = exstractMembers($);
    const attendees = exstractAttendees($, ROWNUM, '○', members);
    const actDate = extractDateFromRownum($, ROWNUM);
    const messageId = event.message.id;
    const userId = event.source.userId;
    const lineName = getLineDisplayName(userId);
    const lang = getLineLang(userId);
    const densukeName = getDensukeName(lineName);
    let replyMessage = null;
    if (densukeName) {
      if (attendees.includes(densukeName)) {
        uploadPayNowPic(lineName, messageId, actDate);
        updatePaymentStatus(lineName, actDate);
        if (lang === 'ja') {
          replyMessage =
            actDate +
            'の支払いを登録しました。ありがとうございます！\nhttps://docs.google.com/spreadsheets/d/1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU/edit?usp=sharing&ccc=' +
            new Date().getTime();
        } else {
          replyMessage =
            'Payment for ' +
            actDate +
            ' has been registered. Thank you!\nhttps://docs.google.com/spreadsheets/d/1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU/edit?usp=sharing&ccc=' +
            new Date().getTime();
        }
      } else {
        if (lang === 'ja') {
          replyMessage =
            actDate +
            '当日の伝助の出席が〇になっていませんでした。伝助を更新して、「伝助更新」と入力してください。';
        } else {
          replyMessage =
            'Your attendance on ' +
            actDate +
            " in Densuke has not been marked as 〇.\nPlease update Densuke and type '伝助更新'.";
        }
      }
    } else {
      if (lang === 'ja') {
        replyMessage =
          '【エラー】伝助名称登録が完了していません。\n登録を完了させて、再度PayNow画像をアップロードして下さい。\n登録は「登録」と入力してください。';
      } else {
        replyMessage =
          "【Error】The initial registration is not complete.\nPlease complete the initial registration and upload the PayNow photo again.\nFor the initial registration, please type '登録'.";
      }
    }
    sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  aggregate(json: any) {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = getDensukeCheerio();
    const members = exstractMembers($);
    const attendees = exstractAttendees($, ROWNUM, '○', members);
    const actDate = extractDateFromRownum($, ROWNUM);

    const ss = SpreadsheetApp.openById(SETTING_SHEET);
    const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
    if (!settingSheet) {
      return;
    }
    const addy = settingSheet.getRange('B2').getValue();
    generateSummaryBase($);
    sendLineReply(replyToken, getSummaryStr(attendees, actDate, addy), '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  unpaid(json: any) {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = getDensukeCheerio();
    const actDate = extractDateFromRownum($, ROWNUM);
    const unpaid = getUnpaid(actDate);
    sendLineReply(
      replyToken,
      '未払いの人 (' + unpaid.length + '名): ' + unpaid.join(', ') + '\n',
      ''
    );
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  remind(json: any) {
    const replyToken = json.events[0].replyToken;
    sendLineReply(replyToken, generateRemind(), '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  densukeUpd(json: any) {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const $ = getDensukeCheerio();
    const userId = event.source.userId;
    const lang = getLineLang(userId);
    const lineName = getLineDisplayName(userId);
    const members = exstractMembers($);
    const attendees = exstractAttendees($, ROWNUM, '○', members);
    const actDate = extractDateFromRownum($, ROWNUM);

    const ss = SpreadsheetApp.openById(SETTING_SHEET);
    const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
    if (!settingSheet) {
      return;
    }
    const addy = settingSheet.getRange('B2').getValue();

    generateSummaryBase($);
    const ownerMessage =
      '【' +
      lineName +
      'さんにより更新されました】\n以下再送お願いします\n' +
      getSummaryStr(attendees, actDate, addy);
    sendMessageToPaynowOwner(ownerMessage);
    let replyMessage = null;
    if (lang === 'ja') {
      replyMessage =
        '伝助の更新ありがとうございました！PayNowのスクリーンショットを再度こちらへ送って下さい。';
    } else {
      replyMessage =
        'Thank you for updating Densuke! Please send PayNow screenshot here again.';
    }
    sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  regInfo(json: any) {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;
    const lang = getLineLang(userId);
    let replyMessage = null;
    if (lang === 'ja') {
      replyMessage =
        '伝助名称の登録を行います。\n伝助のアカウント名を以下のフォーマットで入力してください。\n@@register@@伝助名前\n例）@@register@@やまだじょ\n' +
        getDensukeUrl();
    } else {
      replyMessage =
        'We will perform the densuke name registration.\nPlease enter your Densuke account name in the following format:\n@@register@@XXXXX\nExample)@@register@@Sahim';
    }
    sendLineReply(replyToken, replyMessage, '');
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  managerInfo(json: any) {
    const event = json.events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;
    let replyMessage = null;
    if (isKanji(userId)) {
      replyMessage =
        '設定：https://docs.google.com/spreadsheets/d/1RymgqowExhXB3p2N9i1OYdypIhCtMZkRWr1Wpo1CTyI/edit?usp=sharing \nPayNow：https://drive.google.com/drive/folders/1fcBWvaoAlWNJVmmH4s2spVi5_EJdE3aC?usp=sharin \nReport URL:https://docs.google.com/spreadsheets/d/1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU/edit?usp=sharing&ccc=' +
        new Date().getTime() +
        '\n伝助：' +
        getDensukeUrl() +
        '\nメッセージ利用状況：https://manager.line.biz/account/@164wqxlu/purchase \n 利用可能コマンド:集計, 紹介, 登録, リマインド, 伝助更新, 未払い, @@register@@名前, ';
    } else {
      replyMessage = 'えっ！？このコマンドは平民のキミには内緒だよ！';
    }
    sendLineReply(replyToken, replyMessage, '');
  }
}

function isKanji(userId: string) {
  return getKanjiIds().includes(userId);
}

function getKanjiIds() {
  const kanjiIds: string[] = [];
  const report = SpreadsheetApp.openById(SETTING_SHEET);
  const mappingSheet = report.getSheetByName(MAPPING_SHEET_NAME);
  if (!mappingSheet) {
    return kanjiIds;
  }
  const values = mappingSheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][3] === '幹事') {
      kanjiIds.push(values[i][2]);
    }
  }
  return kanjiIds;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function exstractMembers($: any): string[] {
  const members: string[] = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  $('#mymember td.member').each((element: any) => {
    members.push($(element).text());
  });
  return members;
}

function exstractAttendees(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  $: any,
  rowNum: number,
  symbol: string,
  members: string[]
): string[] {
  const attendees: string[] = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  $(`#row${rowNum} td.attendee:contains("${symbol}")`).each((element: any) => {
    const idx = $(element).index();
    attendees.push(members[idx]);
  });
  return attendees;
}
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractDateFromRownum($: any, rowNum: number): string {
  return $(`#row${rowNum} td.date`).text();
}

function getLineName(member: string): string {
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const sheet = ss.getSheetByName('LineNames');
  if (!sheet) {
    return '';
  }
  const range = sheet.getDataRange();
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === member) {
      return values[i][1];
    }
  }
  return '';
}

function getPaymentUrl(lineName: string, actDate: string): string {
  return `https://payment.url/${lineName}/${actDate}`;
}

function getUnpaid(actDate: string): string[] {
  const unpaid: string[] = [];
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const paymentsSheet = ss.getSheetByName('Payments');
  if (!paymentsSheet) {
    return [];
  }
  const range = paymentsSheet.getDataRange();
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][1] !== actDate) {
      unpaid.push(values[i][0]);
    }
  }
  return unpaid;
}

function getCurrentTime(): string {
  return new Date().toLocaleString();
}

function getLineUserId(densukeName: string): string {
  let userId = '';
  const report = SpreadsheetApp.openById(SETTING_SHEET);
  const mappingSheet = report.getSheetByName(MAPPING_SHEET_NAME);
  if (!mappingSheet) {
    return userId;
  }
  const values = mappingSheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][1] === densukeName) {
      userId = values[i][2];
      break;
    }
  }
  return userId;
}

function getLineUserProfile(userId: string) {
  const url = `https://api.line.me/v2/bot/profile/${userId}`;
  const headers = {
    Authorization: 'Bearer ' + LINE_ACCESS_TOKEN,
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: headers,
  };
  const response = UrlFetchApp.fetch(url, options);
  const userProfile = JSON.parse(response.getContentText());
  return userProfile;
}

function getLineDisplayName(userId: string) {
  return getLineUserProfile(userId).displayName;
}

function getLineLang(userId: string) {
  return getLineUserProfile(userId).language;
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

function registerMapping(
  lineName: string,
  densukeName: string,
  userId: string
) {
  const report = SpreadsheetApp.openById(SETTING_SHEET);
  const mappingSheet = report.getSheetByName('DensukeMapping');
  if (!mappingSheet) {
    return;
  }
  const values = mappingSheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === lineName) {
      mappingSheet.deleteRow(i + 1);
      break;
    }
  }
  mappingSheet.appendRow([lineName, densukeName, userId]);
}

function updateLineNameOfLatestReport(
  lineName: string,
  densukeName: string,
  actDate: string
) {
  const report = SpreadsheetApp.openById(REPORT_SHEET);
  const repo = report.getSheetByName(actDate);
  if (!repo) {
    return;
  }
  const values = repo.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === densukeName) {
      repo.getRange(i + 1, 2).setValue(lineName);
      break;
    }
  }
}

function getDensukeName(lineName: string): string {
  let densukeName = null;
  const report = SpreadsheetApp.openById(SETTING_SHEET);
  const mappingSheet = report.getSheetByName('DensukeMapping');
  if (!mappingSheet) {
    return '';
  }
  const values = mappingSheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === lineName) {
      densukeName = values[i][1];
      break;
    }
  }
  return densukeName;
}

function uploadPayNowPic(
  lineName: string,
  messageId: string,
  actDate: string
): string {
  const fileNm = actDate + '_' + lineName;
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByName(fileNm);
  if (files.hasNext()) {
    const file = files.next();
    file.setTrashed(true);
  }
  const imageUrl = getLineImage(messageId, fileNm);
  return imageUrl;
}

function getLineImage(messageId: string, fileName: string): string {
  const folder = DriveApp.getFolderById(FOLDER_ID);

  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const headers = {
    Authorization: 'Bearer ' + LINE_ACCESS_TOKEN,
  };
  const response = UrlFetchApp.fetch(url, { headers: headers });
  const blob = response.getBlob().setName(fileName);
  const file = folder.createFile(blob);
  return file.getUrl();
}

function updatePaymentStatus(lineName: string, actDate: string) {
  const report = SpreadsheetApp.openById(REPORT_SHEET);
  const repo = report.getSheetByName(actDate);
  if (!repo) {
    return;
  }
  const values = repo.getDataRange().getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][1] === lineName) {
      repo.getRange(i + 1, 3).setValue(getPaymentUrl(lineName, actDate));
      break;
    }
  }
}
