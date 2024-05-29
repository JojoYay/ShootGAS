export function hello() {
  return 'Hello Apps Script!';
}

import * as Cheerio from 'cheerio';

const ROWNUM= 1; //とりあえず一番上からデータとってくる運用
const REPORT_SHEET = '1wyYYlpHbGCmeX4v3G2FjuQ4kuQpz9d3j2DveSPTbMbU';
const SETTING_SHEET = '1RymgqowExhXB3p2N9i1OYdypIhCtMZkRWr1Wpo1CTyI';
const LINE_ACCESS_TOKEN = 'glhgotcUd27uledjA6p02Rv4tkbaqJeUJzAVLeiGNB8Us7tomv7ZdZZ8cz1q+PmVch28XYoh0SlKqPZpwk7+SJkAjZz/CmELk4CNP/DNiKjRUrIQUGDEmhLL40rXDJ8e1lBubmGGRJfkhE0BkXH2CwdB04t89/1O/w1cDnyilFU=';
const FOLDER_ID = '1fcBWvaoAlWNJVmmH4s2spVi5_EJdE3aC';
const ARCHIVE_FOLDER = '1IdRNjbo752JhlATsDic3bGUZDEY5fk9B';
const CHANNEL_QR = 'https://qr-official.line.me/sid/L/164wqxlu.png';
const CHANNEL_URL = 'https://lin.ee/1O9VC3Q';

function generateRemind($: any= getDensukeCheerio()): string {
  const members = exstractMembers($);
  const attendees = exstractAttendees($, ROWNUM, '○', members);
  const unknown = exstractAttendees($, ROWNUM, '△', members);
  const actDate = extractDateFromRownum($, ROWNUM);

  let remindStr = '次回予定' + actDate + 'リマインドです！\n伝助の更新お忘れなく！\nThis is gentle reminder of ' + actDate + ".\nPlease update your Densuke schedule.\n";
  if (attendees.length < 10) {
    remindStr = '次回予定' + actDate + 'がピンチです！\n参加できる方、ぜひ伝助で参加表明お願いします！！！\nThis is gentle reminder of ' + actDate + ".\nDue to the low number of participants, there is a possibility of cancellation.\nPlease join us!\n";
  }
  let summary = remindStr +
    '〇(' + attendees.length + '名): ' + attendees.join(', ') + '\n' +
    '△(' + unknown.length + '名): ' + unknown.join(', ') + '\n' +
    '伝助URL：' + getDensukeUrl();
  return summary;
}

function generateSummaryBase($: any = getDensukeCheerio(), lang: string = 'ja') {
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = ss.getSheetByName("Settings");
  const cashBook = ss.getSheetByName('CashBook');

  const members = exstractMembers($);
  const attendees = exstractAttendees($, ROWNUM, '○', members);
  const actDate = extractDateFromRownum($, ROWNUM);
  if (!!cashBook && !!settingSheet){
    const range = cashBook.getDataRange();
    const values = range.getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][1] === actDate) {
        cashBook.deleteRow(i + 1);
      }
    }
  
    const lastRow = cashBook.getLastRow();
    const orgPrice = cashBook.getRange(lastRow, 5).getValue();
    const rentalFee = settingSheet.getRange('B3').getValue();
    const attendFee = settingSheet.getRange('B4').getValue();
  
    generateSummarySheet(orgPrice, rentalFee, attendFee, actDate, attendees);
  }
}

function getDensukeUrl(): string {
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = ss.getSheetByName('Settings');
  let url = null;
  if(!!settingSheet){
    url = settingSheet.getRange('B1').getValue();
  }
  return url;
}

function getDensukeCheerio(): any {
  const url = getDensukeUrl();
  const html = UrlFetchApp.fetch(url).getContentText();
  const $ = Cheerio.load(html);
  return $;
}

function generateSummarySheet(orgPrice: number, rentalFee: number, attendFee: number, actDate: string, attendees: string[]) {
  const attendFeeTotal = Number(attendFee * attendees.length);
  const report = SpreadsheetApp.openById(REPORT_SHEET);
  let logSheet = report.getSheetByName(actDate);
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
  logSheet.getRange('B7').setValue('' + (orgPrice - rentalFee + attendFeeTotal));

  logSheet.getRange('A9').setValue('参加者（伝助名称）');
  logSheet.getRange('B9').setValue('参加者（Line名称）');
  logSheet.getRange('C9').setValue('支払い状況');

  const values = logSheet.getDataRange().getValues();
  for (let i = values.length; i >= 10; i--) {
    logSheet.deleteRow(i);
  }

  for (let i = 0; i < attendees.length; i++) {
    logSheet.appendRow([attendees[i], getLineName(attendees[i]), getPaymentUrl(getLineName(attendees[i]), actDate)]);
  }
  logSheet.setColumnWidth(1, 170);
  logSheet.setColumnWidth(2, 200);
  const setting = SpreadsheetApp.openById(SETTING_SHEET);
  const cashBook = setting.getSheetByName('CashBook');
  const date = new Date();
  const attendOrg = orgPrice + attendFeeTotal;
  cashBook.appendRow([date, actDate, '参加費(' + attendees.length + '名)', '' + attendFeeTotal, '' + attendOrg]);
  cashBook.appendRow([date, actDate, 'ピッチ使用料金', '▲' + rentalFee, '' + (orgPrice - rentalFee + attendFeeTotal)]);
}

function doGet(e: GoogleAppsScript.Events.DoGet): GoogleAppsScript.Content.TextOutput {
  return ContentService.createTextOutput("Hello World");
}

function test() {
  console.log(getUnpaid('6/2(日)'));
}

interface ICommand {
  func: string;
  condition: (event: any) => boolean;
}

const COMMAND_MAP: ICommand[] = [
  { func: 'payNow', condition: (event) => event.type === 'message' && event.message.type === 'image' },
  { func: 'aggregate', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === '集計' },
  { func: 'unpaid', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === '未払い' },
  { func: 'densukeUpd', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === '伝助更新' },
  { func: 'remind', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === 'リマインド' },
  { func: 'intro', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === '紹介' },
  { func: 'regInfo', condition: (event) => event.type === 'message' && event.message.type === 'text' && (event.message.text === '予定' || event.message.text === '参加表明') },
  { func: 'checkQr', condition: (event) => event.type === 'message' && event.message.type === 'text' && event.message.text === 'QR' },
  { func: 'register', condition: (event) => event.type === 'follow' },
];

function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
  const json = JSON.parse(e.postData.contents);
  const event = json.events[0];

  const command = COMMAND_MAP.find((cmd) => cmd.condition(event));

  if (command) {
    eval(command.func)(json);
  }

  const content = ContentService.createTextOutput(JSON.stringify({ 'content': 'post ok' }));
  content.setMimeType(ContentService.MimeType.JSON);
  return content;
}

function intro() {
  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': 'チーム名: Kick Off FC\n種目: サッカー\n主催: Yuta Sakurai\n活動日: 日曜日\n活動場所: 鶴見緑地公園\n'
      },
      {
        'type': 'text',
        'text': 'みなさん、是非参加してください！\n\nLINEグループのQRコードはこちら：' + CHANNEL_QR
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function register(json: any) {
  const postData = {
    'to': json.events[0].source.userId,
    'messages': [
      {
        'type': 'text',
        'text': '登録ありがとうございます！次回の活動日をお知らせします。\n活動日: ' + getActivityDate()
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function payNow(json: any) {
  const userId = json.events[0].source.userId;
  const messageId = json.events[0].message.id;
  const url = `https://api.line.me/v2/bot/message/${messageId}/content`;

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    headers: {
      'Authorization': `Bearer ${LINE_ACCESS_TOKEN}`
    },
    method: 'get'
  };

  const response = UrlFetchApp.fetch(url, options);
  const imageBlob = response.getBlob();

  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const paymentsSheet = ss.getSheetByName('Payments');

  paymentsSheet.appendRow([new Date(), userId, imageBlob]);

  const postData = {
    'to': userId,
    'messages': [
      {
        'type': 'text',
        'text': '支払いが確認されました。ありがとうございました！'
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function aggregate() {
  const $ = getDensukeCheerio();
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const sheet = ss.getSheetByName("Settings");
  const cashBook = ss.getSheetByName('CashBook');

  const members = exstractMembers($);
  const attendees = exstractAttendees($, ROWNUM, '○', members);
  const unknown = exstractAttendees($, ROWNUM, '△', members);
  const actDate = extractDateFromRownum($, ROWNUM);

  const summary = generateRemind($);
  const total = generateSummaryBase($, 'ja');

  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': summary
      },
      {
        'type': 'text',
        'text': total
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function unpaid() {
  const $ = getDensukeCheerio();
  const unpaidList = getUnpaid(getActivityDate());

  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': '未払い者リストです：\n' + unpaidList.join('\n')
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function remind() {
  const $ = getDensukeCheerio();
  const summary = generateRemind($);

  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': summary
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function densukeUpd() {
  const $ = getDensukeCheerio();
  const actDate = extractDateFromRownum($, ROWNUM);

  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': '伝助が更新されました！\n' + actDate + 'の予定をご確認ください。'
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function checkQr() {
  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'image',
        'originalContentUrl': CHANNEL_QR,
        'previewImageUrl': CHANNEL_QR
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function regInfo() {
  const $ = getDensukeCheerio();
  const summary = generateRemind($);

  const postData = {
    'to': 'C1844f1cf9a5218c508fb2b846fa5283c',
    'messages': [
      {
        'type': 'text',
        'text': summary
      }
    ]
  };

  sendPostRequest('https://api.line.me/v2/bot/message/push', postData);
}

function sendPostRequest(url: string, postData: any) {
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    'headers': {
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
      'Content-Type': 'application/json'
    },
    'method': 'post',
    'payload': JSON.stringify(postData)
  };
  UrlFetchApp.fetch(url, options);
}

function getActivityDate(): string {
  const ss = SpreadsheetApp.openById(SETTING_SHEET);
  const settingSheet = ss.getSheetByName('Settings');
  const date = settingSheet.getRange('B2').getValue();
  return date;
}

function exstractMembers($: any): string[] {
  const members: string[] = [];
  $('#mymember td.member').each((index: number, element:any) => {
    members.push($(element).text());
  });
  return members;
}

function exstractAttendees($: any, rowNum: number, symbol: string, members: string[]): string[] {
  const attendees: string[] = [];
  $(`#row${rowNum} td.attendee:contains("${symbol}")`).each((index: number, element:any) => {
    const idx = $(element).index();
    attendees.push(members[idx]);
  });
  return attendees;
}

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

// function getDensukeCheerio(): CheerioStatic {
//   const url = 'https://www.densuke.biz/';
//   const response = UrlFetchApp.fetch(url);
//   return Cheerio.load(response.getContentText());
// }

// function generateRemind($: CheerioStatic): string {
//   const attendees = exstractAttendees($, ROWNUM, '○', exstractMembers($));
//   const unknown = exstractAttendees($, ROWNUM, '△', exstractMembers($));
//   const nonAttendees = exstractAttendees($, ROWNUM, '×', exstractMembers($));

//   return `出席者:\n${attendees.join('\n')}\n\n未回答者:\n${unknown.join('\n')}\n\n欠席者:\n${nonAttendees.join('\n')}`;
// }

// function generateSummaryBase($: CheerioStatic, lang: string): string {
//   const actDate = extractDateFromRownum($, ROWNUM);
//   const members = exstractMembers($);
//   const totalMembers = members.length;
//   const attendees = exstractAttendees($, ROWNUM, '○', members).length;
//   const nonAttendees = exstractAttendees($, ROWNUM, '×', members).length;

//   if (lang === 'ja') {
//     return `活動日: ${actDate}\nメンバー数: ${totalMembers}\n出席者数: ${attendees}\n欠席者数: ${nonAttendees}`;
//   } else {
//     return `Date: ${actDate}\nTotal Members: ${totalMembers}\nAttendees: ${attendees}\nNon-attendees: ${nonAttendees}`;
//   }
// }





