import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasTestSuite } from './gasTestSuite';
import { GasUtil } from './gasUtil';
import { LineUtil } from './lineUtil';
import { PostEventHandler } from './postEventHandler';
import { ScriptProps } from './scriptProps';

const densukeUtil: DensukeUtil = new DensukeUtil();
const lineUtil: LineUtil = new LineUtil();
const gasUtil: GasUtil = new GasUtil();

export class RequestExecuter {
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
      if (this.hasMultipleOccurrences(members, densukeNameNew)) {
        if (postEventHander.lang === 'ja') {
          postEventHander.resultMessage =
            '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。';
        } else {
          postEventHander.resultMessage =
            "There are multiple entries with the name '" + densukeNameNew + "' on Densuke. Please update it to a unique name and register again.";
        }
      } else {
        gasUtil.registerMapping(lineName, densukeNameNew, postEventHander.userId);
        gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
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
          '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' + densukeNameNew;
      }
    }
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
          postEventHander.resultMessage = actDate + 'の支払いを登録しました。ありがとうございます！\n' + GasProps.instance.ReportSheetUrl;
        } else {
          postEventHander.resultMessage = 'Payment for ' + actDate + ' has been registered. Thank you!\n' + GasProps.instance.ReportSheetUrl;
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
        GasProps.instance.ReportSheetUrl +
        '\n伝助：' +
        densukeUtil.getDensukeUrl() +
        '\nチャット状況：' +
        ScriptProps.instance.chat +
        '\nメッセージ利用状況：' +
        ScriptProps.instance.messageUsage +
        '\n 利用可能コマンド:集計, 紹介, 登録, リマインド, 伝助更新, 未払い, @@register@@名前 ';
    } else {
      postEventHander.resultMessage = 'えっ！？このコマンドは平民のキミには内緒だよ！';
    }
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

  private hasMultipleOccurrences(array: string[], searchString: string): boolean {
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
}
