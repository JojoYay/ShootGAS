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
 * Unit tests for ScriptProps.
 * PropertiesService is mocked globally.
 */
import { ScriptProps } from '../../src/scriptProps';

const mockGetProperty = jest.fn();

beforeAll(() => {
    (global as unknown as { PropertiesService: unknown }).PropertiesService = {
        getScriptProperties: () => ({
            getProperty: mockGetProperty,
        }),
    };
});

beforeEach(() => {
    // シングルトンとモードをリセット
    (ScriptProps as unknown as { _instance: null })['_instance'] = null;
    (ScriptProps as unknown as { _mode: string })['_mode'] = 'normal';
    mockGetProperty.mockReset();
});

// ─────────────────────────────────────────────────────────

describe('ScriptProps', () => {
    describe('instance (singleton)', () => {
        it('5-9: 2回呼び出しで同一インスタンスを返す', () => {
            const a = ScriptProps.instance;
            const b = ScriptProps.instance;
            expect(a).toBe(b);
        });
    });

    describe('usersSheet', () => {
        it('5-1: usersSheet プロパティが設定されている場合にその値を返す', () => {
            mockGetProperty.mockImplementation((key: string) => {
                if (key === 'usersSheet') return 'users-sheet-id';
                if (key === 'settingSheet') return 'setting-sheet-id';
                return null;
            });
            expect(ScriptProps.instance.usersSheet).toBe('users-sheet-id');
        });

        it('5-2: usersSheet が未設定の場合に settingSheet の値を返す', () => {
            mockGetProperty.mockImplementation((key: string) => {
                if (key === 'usersSheet') return null;
                if (key === 'settingSheet') return 'setting-sheet-id';
                return null;
            });
            expect(ScriptProps.instance.usersSheet).toBe('setting-sheet-id');
        });
    });

    describe('calendarId', () => {
        it('5-3: calendarId が未設定の場合に Error がthrowされる', () => {
            mockGetProperty.mockReturnValue(null);
            expect(() => ScriptProps.instance.calendarId).toThrow('Script Property (calendarId) was not found');
        });
    });

    describe('eventResults', () => {
        it('5-4: eventResults が未設定の場合に Error がthrowされる', () => {
            mockGetProperty.mockReturnValue(null);
            expect(() => ScriptProps.instance.eventResults).toThrow('Script Property (eventResults) was not found');
        });
    });

    describe('reportSheet', () => {
        it('5-5: reportSheet が未設定の場合に Error がthrowされる', () => {
            mockGetProperty.mockReturnValue(null);
            expect(() => ScriptProps.instance.reportSheet).toThrow('Script Property (reportSheet) was not found');
        });

        it('5-6: テストモードの場合にテスト用スプレッドシート ID が返される', () => {
            mockGetProperty.mockReturnValue('real-sheet-id');
            ScriptProps.startTest();
            expect(ScriptProps.instance.reportSheet).toBe('1Ej-9kZIMpGW66BUm0cGS1iG1RDgUsgBd5fo5V97xirg');
            ScriptProps.endTest();
        });
    });

    describe('settingSheet', () => {
        it('5-7: settingSheet が未設定の場合に Error がthrowされる', () => {
            mockGetProperty.mockReturnValue(null);
            expect(() => ScriptProps.instance.settingSheet).toThrow('Script Property (settingProp) was not found');
        });

        it('5-8: テストモードの場合にテスト用スプレッドシート ID が返される', () => {
            mockGetProperty.mockReturnValue('real-setting-id');
            ScriptProps.startTest();
            expect(ScriptProps.instance.settingSheet).toBe('1PfBvcVqO_d-JIs6VxwSJLW0GZAS0c6xfsOMUQKlTU30');
            ScriptProps.endTest();
        });
    });

    describe('folderId', () => {
        it('5-8b: folderId が未設定の場合に Error がthrowされる', () => {
            mockGetProperty.mockReturnValue(null);
            expect(() => ScriptProps.instance.folderId).toThrow('Script Property (folderProp) was not found');
        });
    });

    describe('startTest / endTest / isTesting', () => {
        it('5-10: startTest 後に isTesting が true を返す', () => {
            ScriptProps.startTest();
            expect(ScriptProps.isTesting()).toBe(true);
            ScriptProps.endTest();
        });

        it('5-11: endTest 後に isTesting が false を返す', () => {
            ScriptProps.startTest();
            ScriptProps.endTest();
            expect(ScriptProps.isTesting()).toBe(false);
        });
    });
});
