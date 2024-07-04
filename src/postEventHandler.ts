import { LineUtil } from './lineUtil';

type Command = {
    func: string;
    lineCmd: string;
    display: boolean;
    condition: (postEventHander: PostEventHandler) => boolean;
};

export const COMMAND_MAP: Command[] = [
    {
        func: 'payNow',
        lineCmd: '',
        display: false,
        condition: (postEventHander: PostEventHandler) => postEventHander.type === 'message' && postEventHander.messageType === 'image',
    },
    {
        func: 'upload',
        lineCmd: '',
        display: false,
        condition: (postEventHander: PostEventHandler) => postEventHander.parameter.func === 'upload',
    },
    {
        func: 'aggregate',
        lineCmd: '集計, aggregate',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '集計' || postEventHander.messageText.toLowerCase() === 'aggregate'),
    },
    {
        func: 'video',
        lineCmd: 'ビデオ, video',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === 'ビデオ' || postEventHander.messageText.toLowerCase() === 'video`'),
    },

    {
        func: 'remind',
        lineCmd: 'リマインド, remind',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === 'リマインド' || postEventHander.messageText.toLowerCase() === 'remind'),
    },
    {
        func: 'unpaid',
        lineCmd: '未払い, unpaid',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '未払い' || postEventHander.messageText.toLowerCase() === 'unpaid'),
    },
    {
        func: 'unRegister',
        lineCmd: '未登録参加者, unregister',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '未登録参加者' || postEventHander.messageText.toLowerCase() === 'unregister'),
    },
    {
        func: 'densukeUpd',
        lineCmd: '伝助更新, update',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '伝助更新' || postEventHander.messageText.toLowerCase() === 'update'),
    },
    {
        func: 'intro',
        lineCmd: '紹介, introduce',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '紹介' || postEventHander.messageText.toLowerCase() === 'introduce'),
    },
    {
        func: 'regInfo',
        lineCmd: '登録, how to register',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '登録' ||
                postEventHander.messageText.toLowerCase() === '@@register@@' ||
                postEventHander.messageText.toLowerCase() === 'how to register'),
    },
    {
        func: 'ranking',
        lineCmd: 'ランキング, ranking',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === 'ランキング' || postEventHander.messageText.toLowerCase() === 'calc'),
    },
    {
        func: 'myResult',
        lineCmd: '戦績, my result',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '戦績' || postEventHander.messageText.toLowerCase() === 'my result'),
    },
    {
        func: 'managerInfo',
        lineCmd: '管理, manage',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            (postEventHander.messageText === '管理' || postEventHander.messageText.toLowerCase() === 'manage'),
    },
    {
        func: 'register',
        lineCmd: '@@register@@',
        display: true,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' &&
            postEventHander.messageType === 'text' &&
            postEventHander.messageText.toLowerCase().startsWith('@@register@@'),
    },
    {
        func: 'systemTest',
        lineCmd: 'システムテスト',
        display: false,
        condition: (postEventHander: PostEventHandler) =>
            postEventHander.type === 'message' && postEventHander.messageType === 'text' && postEventHander.messageText.startsWith('システムテスト'),
    },
];

export class PostEventHandler {
    private _messageText: string = '';
    private _messageType: string = ''; //message or image
    private _type: string = ''; //messageだけ

    private _messageId: string = '';
    private _replyToken: string = '';
    private _userId: string = '';
    private _lang: string = '';

    private _resultMessage: string = '';
    private _resultImage: string | null = null;
    private _paynowOwnerMsg: string | null = null;
    private _testResult: string[] = [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private _mockDensukeCheerio: any | null;
    public isFlex: boolean = false;
    public messageJson: JSON | null = null;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public parameter: any = {};
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public reponseObj: any = {};

    public constructor(e: GoogleAppsScript.Events.DoPost) {
        //LINE APIの場合
        if (e.postData && e.postData.contents) {
            const json = JSON.parse(e.postData.contents);
            const event = json.events[0];
            console.log(event);
            if (event.message) {
                this._messageText = event.message.text;
                this._messageType = event.message.type;
                // } else {
                //     this._messageText = '';
                //     this._messageType = '';
            }
            this._type = event.type;

            this._userId = event.source.userId;
            this._replyToken = event.replyToken;
            this._messageId = event.message.id;
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
        } else {
            this.parameter = e.parameter;
            this.reponseObj = e.parameter;
            this._userId = e.parameter.userId;
        }
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
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public get mockDensukeCheerio(): any | null {
        return this._mockDensukeCheerio;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public set mockDensukeCheerio(value: any | null) {
        this._mockDensukeCheerio = value;
    }

    public generateCommandList(): string {
        let result: string = '利用可能コマンド: ';
        for (let i = 0; i < COMMAND_MAP.length; i++) {
            if (COMMAND_MAP[i].display) {
                result += COMMAND_MAP[i].lineCmd;
                if (COMMAND_MAP[i].func !== 'register') {
                    result += ', ';
                }
            }
        }
        return result;
    }
}
