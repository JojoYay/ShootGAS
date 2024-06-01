import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { PostEventHandler } from './postEventHandler';
import { RequestExecuter } from './requestExecuter';

const gasUtil: GasUtil = new GasUtil();

export class GasTestSuite {
  private initializeSheet() {
    const baseSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById('1w5PEeUvJezki1EeG2H1x_2E8uRvTe995zX-YmX0CsU4');

    const baseSetting: GoogleAppsScript.Spreadsheet.Sheet | null = baseSS.getSheetByName('Settings');
    const settingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.settingSheet;
    this.copySheetData(baseSetting, settingSheet);

    const baseMapping: GoogleAppsScript.Spreadsheet.Sheet | null = baseSS.getSheetByName('DensukeMapping');
    const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
    this.copySheetData(baseMapping, mappingSheet);

    const baseCashBook: GoogleAppsScript.Spreadsheet.Sheet | null = baseSS.getSheetByName('CashBook');
    const cashBookSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.cashBookSheet;
    this.copySheetData(baseCashBook, cashBookSheet);

    const baseRepoSS: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById('1_7vZy4DqaEzl1C3R4TNjoDPkQboddwgMsCIO8sCpCHw');
    const mayData: GoogleAppsScript.Spreadsheet.Sheet | null = baseRepoSS.getSheetByName('5/26(日)');
    const actualMay: GoogleAppsScript.Spreadsheet.Sheet = gasUtil.getReportSheet('5/26(日)');
    this.copySheetData(mayData, actualMay);

    const juneData: GoogleAppsScript.Spreadsheet.Sheet | null = baseRepoSS.getSheetByName('6/2(日)');
    const actualJune: GoogleAppsScript.Spreadsheet.Sheet = gasUtil.getReportSheet('5/26(日)');
    this.copySheetData(juneData, actualJune);
  }

  private copySheetData(sourceSheet: GoogleAppsScript.Spreadsheet.Sheet | null, targetSheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    if (!sourceSheet) {
      throw new Error('コピー元のエクセルがなぜか無い！');
    }
    const sourceData = sourceSheet.getDataRange().getValues();
    targetSheet.clear();
    targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  }

  public testIntro1(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    requestExecuter.intro(postEventHander);
    if (
      postEventHander.resultMessage === 'https://lin.ee/LIlqNmE' &&
      postEventHander.resultImage === 'https://qr-official.line.me/sid/L/848rxuwb.png'
    ) {
      postEventHander.testResult.push('testIntro1:passed');
    } else {
      postEventHander.testResult.push('testIntro1:failed');
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public testIsKanji(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    const nishimuraUserId = 'Ueee7560851b8534817a14454e89e5bbc';
    const noKanjiId = 'Uf395b2a8c82788dc3331b62f0cf96578';
    const jojoId = 'U398bbcb257b5bfdae2a86c928543ab22';
    if (gasUtil.isKanji(nishimuraUserId) && !gasUtil.isKanji(noKanjiId) && gasUtil.isKanji(jojoId) && !gasUtil.isKanji('')) {
      postEventHander.testResult.push('testIsKanji:passed');
    } else {
      postEventHander.testResult.push('testIsKanji:failed');
    }
  }

  public testRegister1(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    this.initializeSheet();
    postEventHander.messageText = '@@register@@やまだじょ';
    const jojoId = 'U398bbcb257b5bfdae2a86c928543ab22';
    postEventHander.userId = jojoId;
    requestExecuter.register(postEventHander);
    const expectMsg1 =
      '伝助名称登録が完了しました。\n伝助上の名前：やまだじょ\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。\n' +
      postEventHander.userId;
    const result1 = !gasUtil.isKanji(jojoId); //幹事の文字が消えたら合格
    const result2 = expectMsg1 === postEventHander.resultMessage;
    if (result1 && result2) {
      postEventHander.testResult.push('testRegister1:passed');
    } else {
      postEventHander.testResult.push('testRegister1:failed' + postEventHander.resultMessage);
    }
  }

  public testRegister2(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    this.initializeSheet();
    const jojoId = 'U398bbcb257b5bfdae2a86c928543ab22';
    postEventHander.userId = jojoId;

    postEventHander.messageText = '@@register@@ほげ田鼻毛太郎';
    requestExecuter.register(postEventHander);
    const expectMsg2 = '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：ほげ田鼻毛太郎';
    const result3 = expectMsg2 === postEventHander.resultMessage;

    if (result3) {
      postEventHander.testResult.push('testRegister2:passed');
    } else {
      postEventHander.testResult.push('testRegister2:failed' + postEventHander.resultMessage);
    }
  }

  public testRegister3(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    this.initializeSheet();
    const jojoId = 'U398bbcb257b5bfdae2a86c928543ab22';
    postEventHander.userId = jojoId;
    postEventHander.messageText = '@@register@@安田';
    requestExecuter.register(postEventHander);
    const expectMsg3 = '伝助上で"安田"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。';
    const result4 = expectMsg3 === postEventHander.resultMessage;

    if (result4) {
      postEventHander.testResult.push('testRegister:passed');
    } else {
      postEventHander.testResult.push('testRegister:failed' + postEventHander.resultMessage);
    }
  }

  public testPayNow1(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    //Soma(Ucb9beba3011ec9cf85c5482efa132e9b)さんで実行
    const somaId = 'Ucb9beba3011ec9cf85c5482efa132e9b';
    postEventHander.userId = somaId;
    requestExecuter.payNow(postEventHander);
    const expectation1: string = '6/2(日)の支払いを登録しました。ありがとうございます！\n' + GasProps.instance.ReportSheetUrl;
    const folder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder;
    const files = folder.getFilesByName('6/2(日)_相馬究(Kiwamu Soma)');
    if (postEventHander.resultMessage === expectation1 && files.hasNext()) {
      postEventHander.testResult.push('testPayNow1:passed');
    } else {
      postEventHander.testResult.push('testPayNow1:failed' + postEventHander.resultMessage);
    }
  }

  public testPayNow2(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    //千葉（Uf395b2a8c82788dc3331b62f0cf96578）がメッセージ送った事を再現
    const chibaId = 'Uf395b2a8c82788dc3331b62f0cf96578';
    postEventHander.userId = chibaId;
    requestExecuter.payNow(postEventHander);
    const expectation: string =
      '【エラー】6/2(日)の伝助の出席が〇になっていませんでした。伝助を更新して、「伝助更新」と入力してください。\nhttps://densuke.biz/list?cd=DTDR7Cu7rmkZy9YA';
    if (postEventHander.resultMessage === expectation) {
      postEventHander.testResult.push('testPayNow2:passed');
    } else {
      postEventHander.testResult.push('testPayNow2:failed' + postEventHander.resultMessage);
    }
  }

  public testPayNow3(postEventHander: PostEventHandler, requestExecuter: RequestExecuter): void {
    //なべ（tekitouID）がメッセージ送った事を再現(つか誰か実際わからん)
    const nabeId = 'tekitoutekitoutekitou';
    postEventHander.userId = nabeId;
    requestExecuter.payNow(postEventHander);
    const expectation: string =
      '【エラー】伝助名称登録が完了していません。\n登録を完了させて、再度PayNow画像をアップロードして下さい。\n登録は「登録」と入力してください。\nhttps://densuke.biz/list?cd=DTDR7Cu7rmkZy9YA';
    if (postEventHander.resultMessage === expectation) {
      postEventHander.testResult.push('testPayNow3:passed');
    } else {
      postEventHander.testResult.push('testPayNow3:failed' + postEventHander.resultMessage);
    }
  }

  public testAggregate() {}
}
