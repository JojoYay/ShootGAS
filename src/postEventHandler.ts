import { LineUtil } from './lineUtil';

export class PostEventHandler {
  private _messageText: string;
  private _messageType: string; //message or image
  private _type: string; //messageだけ

  private _messageId: string;
  private _replyToken: string;
  private _userId: string;
  private _lang: string;

  private _resultMessage: string;
  private _resultImage: string | null;
  private _paynowOwnerMsg: string | null;
  private _testResult: string[];

  public constructor(e: GoogleAppsScript.Events.DoPost) {
    const json = JSON.parse(e.postData.contents);
    const event = json.events[0];
    if (event.message) {
      this._messageText = event.message.text;
      this._messageType = event.message.type;
    } else {
      this._messageText = '';
      this._messageType = '';
    }
    this._type = event.type;

    this._userId = event.source.userId;
    this._replyToken = event.replyToken;
    this._messageId = event.messageId;
    const lineUtil: LineUtil = new LineUtil();
    this._lang = lineUtil.getLineLang(this._userId);
    if (this._lang === 'ja') {
      this._resultMessage = '【エラー】申し訳ありません、理解できませんでした。再度正しく入力してください。';
    } else {
      this._resultMessage = "【Error】I'm sorry, I didn't understand. Please enter the correct input again.";
    }
    this._resultImage = null;
    this._paynowOwnerMsg = null;
    this._testResult = [];
  }

  public get messageId(): string {
    return this._messageId;
  }
  public set messageId(value: string) {
    this._messageId = value;
  }
  public get replyToken(): string {
    return this._replyToken;
  }
  public set replyToken(value: string) {
    this._replyToken = value;
  }
  public get userId(): string {
    return this._userId;
  }
  public set userId(value: string) {
    this._userId = value;
  }
  public get messageText(): string {
    return this._messageText;
  }
  public set messageText(value: string) {
    this._messageText = value;
  }
  public get type(): string {
    return this._type;
  }
  public set type(value: string) {
    this._type = value;
  }
  public get messageType(): string {
    return this._messageType;
  }
  public set messageType(value: string) {
    this._messageType = value;
  }

  public get resultMessage(): string {
    return this._resultMessage;
  }
  public set resultMessage(value: string) {
    this._resultMessage = value;
  }
  public get resultImage(): string | null {
    return this._resultImage;
  }
  public set resultImage(value: string) {
    this._resultImage = value;
  }
  public get paynowOwnerMsg(): string | null {
    return this._paynowOwnerMsg;
  }
  public set paynowOwnerMsg(value: string) {
    this._paynowOwnerMsg = value;
  }

  public get lang(): string {
    return this._lang;
  }
  public set lang(value: string) {
    this._lang = value;
  }

  public get testResult(): string[] {
    return this._testResult;
  }
  public set testResult(value: string[]) {
    this._testResult = value;
  }
}
