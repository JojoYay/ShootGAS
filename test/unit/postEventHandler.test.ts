/**
 * Copyright 2024 JojoYay
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Unit tests for PostEventHandler and COMMAND_MAP.
 * LineUtil is mocked so LINE API calls are not made.
 */
const mockGetLineLang = jest.fn().mockReturnValue('ja');

jest.mock('../../src/lineUtil', () => ({
    LineUtil: jest.fn().mockImplementation(() => ({
        getLineLang: mockGetLineLang,
    })),
}));

import { COMMAND_MAP, PostEventHandler } from '../../src/postEventHandler';

// ─── ヘルパー ─────────────────────────────────────────────

/** COMMAND_MAP から func 名でコンディション関数を取得する */
function findCondition(funcName: string) {
    const cmd = COMMAND_MAP.find(c => c.func === funcName);
    if (!cmd) throw new Error(`COMMAND_MAP に登録されていません: ${funcName}`);
    return cmd.condition;
}

/** PostEventHandler を new せずにプロパティを直接注入したモックを返す */
function mockHandler(props: { type?: string; messageType?: string; messageText?: string; parameter?: Record<string, string> }): PostEventHandler {
    return {
        type: props.type ?? '',
        messageType: props.messageType ?? '',
        messageText: props.messageText ?? '',
        parameter: props.parameter ?? {},
    } as unknown as PostEventHandler;
}

/** LINE Webhook 形式の DoPost イベントを生成する */
function makeLineDoPost(event: object): GoogleAppsScript.Events.DoPost {
    return {
        postData: { contents: JSON.stringify({ events: [event] }) },
        parameter: {},
    } as unknown as GoogleAppsScript.Events.DoPost;
}

/** パラメータベースの DoPost イベントを生成する */
function makeParamDoPost(parameter: Record<string, string>): GoogleAppsScript.Events.DoPost {
    return { parameter } as unknown as GoogleAppsScript.Events.DoPost;
}

// ─── COMMAND_MAP condition テスト ─────────────────────────

describe('COMMAND_MAP conditions', () => {
    // LINE Bot テキストコマンド
    describe('aggregate', () => {
        it('2-1: 「集計」でconditionがtrueを返す', () => {
            expect(findCondition('aggregate')(mockHandler({ type: 'message', messageType: 'text', messageText: '集計' }))).toBe(true);
        });
        it('2-2: 「aggregate」でconditionがtrueを返す', () => {
            expect(findCondition('aggregate')(mockHandler({ type: 'message', messageType: 'text', messageText: 'aggregate' }))).toBe(true);
        });
        it('2-3: toLowerCase() で比較するため「AGGREGATE」（大文字）でもconditionがtrueを返す', () => {
            // コードは messageText.toLowerCase() === 'aggregate' で比較するため大文字でもマッチする
            expect(findCondition('aggregate')(mockHandler({ type: 'message', messageType: 'text', messageText: 'AGGREGATE' }))).toBe(true);
        });
    });

    describe('video', () => {
        it('2-4: 「ビデオ」でconditionがtrueを返す', () => {
            expect(findCondition('video')(mockHandler({ type: 'message', messageType: 'text', messageText: 'ビデオ' }))).toBe(true);
        });
    });

    describe('remind', () => {
        it('2-5: 「リマインド」でconditionがtrueを返す', () => {
            expect(findCondition('remind')(mockHandler({ type: 'message', messageType: 'text', messageText: 'リマインド' }))).toBe(true);
        });
        it('2-6: 「remind」でconditionがtrueを返す', () => {
            expect(findCondition('remind')(mockHandler({ type: 'message', messageType: 'text', messageText: 'remind' }))).toBe(true);
        });
    });

    describe('unpaid', () => {
        it('2-7: 「未払い」でconditionがtrueを返す', () => {
            expect(findCondition('unpaid')(mockHandler({ type: 'message', messageType: 'text', messageText: '未払い' }))).toBe(true);
        });
        it('2-8: 「unpaid」でconditionがtrueを返す', () => {
            expect(findCondition('unpaid')(mockHandler({ type: 'message', messageType: 'text', messageText: 'unpaid' }))).toBe(true);
        });
    });

    describe('intro', () => {
        it('2-9: 「紹介」でconditionがtrueを返す', () => {
            expect(findCondition('intro')(mockHandler({ type: 'message', messageType: 'text', messageText: '紹介' }))).toBe(true);
        });
        it('2-10: 「introduce」でconditionがtrueを返す', () => {
            expect(findCondition('intro')(mockHandler({ type: 'message', messageType: 'text', messageText: 'introduce' }))).toBe(true);
        });
    });

    describe('ranking', () => {
        it('2-11: 「ランキング」でconditionがtrueを返す', () => {
            expect(findCondition('ranking')(mockHandler({ type: 'message', messageType: 'text', messageText: 'ランキング' }))).toBe(true);
        });
        it('2-12: 「calc」でconditionがtrueを返す', () => {
            expect(findCondition('ranking')(mockHandler({ type: 'message', messageType: 'text', messageText: 'calc' }))).toBe(true);
        });
    });

    describe('myResult', () => {
        it('2-13: 「戦績」でconditionがtrueを返す', () => {
            expect(findCondition('myResult')(mockHandler({ type: 'message', messageType: 'text', messageText: '戦績' }))).toBe(true);
        });
        it('2-14: 「my result」でconditionがtrueを返す', () => {
            expect(findCondition('myResult')(mockHandler({ type: 'message', messageType: 'text', messageText: 'my result' }))).toBe(true);
        });
    });

    describe('managerInfo', () => {
        it('2-15: 「管理」でconditionがtrueを返す', () => {
            expect(findCondition('managerInfo')(mockHandler({ type: 'message', messageType: 'text', messageText: '管理' }))).toBe(true);
        });
        it('2-16: 「manage」でconditionがtrueを返す', () => {
            expect(findCondition('managerInfo')(mockHandler({ type: 'message', messageType: 'text', messageText: 'manage' }))).toBe(true);
        });
    });

    describe('systemTest', () => {
        it('2-17: 「システムテスト」でconditionがtrueを返す', () => {
            expect(findCondition('systemTest')(mockHandler({ type: 'message', messageType: 'text', messageText: 'システムテスト' }))).toBe(true);
        });
        it('2-18: 「システムテスト123」（前方一致）でconditionがtrueを返す', () => {
            expect(findCondition('systemTest')(mockHandler({ type: 'message', messageType: 'text', messageText: 'システムテスト123' }))).toBe(true);
        });
    });

    describe('payNow', () => {
        it('2-19: 画像メッセージの場合にconditionがtrueを返す', () => {
            expect(findCondition('payNow')(mockHandler({ type: 'message', messageType: 'image' }))).toBe(true);
        });
    });

    // funcパラメータ系コマンド
    const FUNC_COMMANDS = [
        'updateUser',
        'uploadInvoice',
        'deleteInvoice',
        'insertCashBook',
        'deleteCashBook',
        'uploadToYoutube',
        'uploadToYoutube2',
        'updateEventData',
        'deleteComments',
        'insertComments',
        'registrationFromApp',
        'updateEventStatus',
        'updateCalendar',
        'createCalendar',
        'deleteCalendar',
        'updateParticipation',
        'uploadPaticipationPayNow',
        'upload',
        'loadExList',
        'deleteEx',
        'updateTeams',
        'recordGoal',
        'updateGoal',
        'deleteGoal',
        'closeGame',
        'createShootLog',
        'saveSheetData',
    ];

    describe.each(FUNC_COMMANDS)('func=%s', funcName => {
        it(`2-xx: func=${funcName} のリクエストで condition がtrueを返す`, () => {
            expect(findCondition(funcName)(mockHandler({ parameter: { func: funcName } }))).toBe(true);
        });
    });
});

// ─── PostEventHandler コンストラクタ テスト ──────────────

describe('PostEventHandler', () => {
    describe('constructor (LINE Bot flow)', () => {
        beforeEach(() => {
            mockGetLineLang.mockReset();
            mockGetLineLang.mockReturnValue('ja');
        });

        it('2-20: postback イベントの場合にメッセージ解析をスキップする', () => {
            const e = makeLineDoPost({
                type: 'postback',
                source: { userId: 'Utest' },
                replyToken: 'token',
            });
            const handler = new PostEventHandler(e);
            expect(handler.messageText).toBe('');
            expect(handler.messageType).toBe('');
        });

        it('2-21: 日本語ユーザーの場合に日本語エラーメッセージが設定される', () => {
            mockGetLineLang.mockReturnValue('ja');
            const e = makeLineDoPost({
                type: 'message',
                message: { type: 'text', text: '不明なコマンド', id: 'msg1' },
                source: { userId: 'Utest' },
                replyToken: 'token',
            });
            const handler = new PostEventHandler(e);
            expect(handler.resultMessage).toContain('【エラー】');
        });

        it('2-22: 非日本語ユーザーの場合に英語エラーメッセージが設定される', () => {
            mockGetLineLang.mockReturnValue('en');
            const e = makeLineDoPost({
                type: 'message',
                message: { type: 'text', text: 'unknown', id: 'msg2' },
                source: { userId: 'Utest' },
                replyToken: 'token',
            });
            const handler = new PostEventHandler(e);
            expect(handler.resultMessage).toContain('【Error】');
        });
    });

    describe('constructor (parameter flow)', () => {
        it('パラメータベースリクエストの場合に parameter プロパティが設定される', () => {
            const e = makeParamDoPost({ func: 'updateUser', userId: 'Utest' });
            const handler = new PostEventHandler(e);
            expect(handler.parameter.func).toBe('updateUser');
            expect(handler.userId).toBe('Utest');
        });
    });

    describe('generateCommandList', () => {
        it('2-23: display=true のコマンドをすべて含む文字列を返す', () => {
            const e = makeParamDoPost({});
            const handler = new PostEventHandler(e);
            const list = handler.generateCommandList();
            const displayCmds = COMMAND_MAP.filter(c => c.display).map(c => c.lineCmd);
            for (const cmd of displayCmds) {
                expect(list).toContain(cmd);
            }
        });

        it('2-24: display=false のコマンドを含まない文字列を返す', () => {
            const e = makeParamDoPost({});
            const handler = new PostEventHandler(e);
            const list = handler.generateCommandList();
            const hiddenFuncs = COMMAND_MAP.filter(c => !c.display && c.lineCmd !== '').map(c => c.lineCmd);
            for (const cmd of hiddenFuncs) {
                expect(list).not.toContain(cmd);
            }
        });
    });
});
