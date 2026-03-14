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
 * Unit tests for GetEventHandler.
 * LiffApi is mocked so GAS dependencies are not required.
 */
jest.mock('../../src/liffApi', () => ({
    LiffApi: jest.fn().mockImplementation(() => ({
        getTeams: jest.fn(),
        getScores: jest.fn(),
        getCalendar: jest.fn(),
        getVideos: jest.fn(),
        // 意図的に関数ではないプロパティを混在させる
        nonFunctionProp: 'string-value',
    })),
}));

import { GetEventHandler } from '../../src/getEventHandler';

function makeDoGet(funcs: string[]): GoogleAppsScript.Events.DoGet {
    return {
        parameters: { func: funcs },
        parameter: {},
    } as unknown as GoogleAppsScript.Events.DoGet;
}

describe('GetEventHandler', () => {
    describe('constructor', () => {
        it('1-1: 有効な func が渡された場合にコンストラクタが正常完了する', () => {
            expect(() => new GetEventHandler(makeDoGet(['getTeams']))).not.toThrow();
        });

        it('1-2: LiffApi に存在しない func が渡された場合に Error がthrowされる', () => {
            expect(() => new GetEventHandler(makeDoGet(['nonExistentMethod']))).toThrow('Func is not registered:nonExistentMethod');
        });

        it('1-3: LiffApi に存在するが関数でないプロパティが func に指定された場合に Error がthrowされる', () => {
            expect(() => new GetEventHandler(makeDoGet(['nonFunctionProp']))).toThrow('Func is not registered:nonFunctionProp');
        });

        it('1-4: 複数の有効な func が渡された場合にコンストラクタが正常完了する', () => {
            expect(() => new GetEventHandler(makeDoGet(['getTeams', 'getScores']))).not.toThrow();
        });

        it('1-5: func パラメータが空配列の場合にコンストラクタが正常完了する', () => {
            expect(() => new GetEventHandler(makeDoGet([]))).not.toThrow();
        });
    });
});
