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
 * Extended unit tests for GasProps (4-2 to 4-12).
 * ScriptProps is mocked with usersSheet != settingSheet (non-fallback configuration).
 * For the DensukeMapping fallback, see gasProps.test.ts.
 */

// ScriptProps インスタンスをミュータブルなオブジェクトとしてモックする
jest.mock('../../src/scriptProps', () => ({
    ScriptProps: {
        instance: {
            usersSheet: 'users-sheet-id',
            settingSheet: 'setting-sheet-id',
            reportSheet: 'report-sheet-id',
            eventResults: 'event-results-id',
            folderId: 'folder-id',
            archiveFolder: 'archive-folder-id',
            expenseFolder: 'expense-folder-id',
        },
        isTesting: jest.fn().mockReturnValue(false),
    },
}));

import { GasProps } from '../../src/gasProps';
import { ScriptProps } from '../../src/scriptProps';

const mockGetSheetByName = jest.fn();
const mockInsertSheet = jest.fn();

beforeAll(() => {
    (global as unknown as { SpreadsheetApp: unknown }).SpreadsheetApp = {
        openById: jest.fn().mockReturnValue({
            getSheetByName: mockGetSheetByName,
            insertSheet: mockInsertSheet,
        }),
    };
});

beforeEach(() => {
    // GasProps シングルトンをリセット
    (GasProps as unknown as { _instance: null })['_instance'] = null;
    mockGetSheetByName.mockReset();
    mockInsertSheet.mockReset();
    // ScriptProps.instance を既定値に戻す
    const inst = ScriptProps.instance as unknown as Record<string, string>;
    inst.usersSheet = 'users-sheet-id';
    inst.settingSheet = 'setting-sheet-id';
    inst.reportSheet = 'report-sheet-id';
    (ScriptProps.isTesting as jest.Mock).mockReturnValue(false);
});

// ─────────────────────────────────────────────────────────

describe('GasProps (extended)', () => {
    describe('instance (singleton)', () => {
        it('4-10: 2回呼び出しで同一インスタンスを返す', () => {
            const a = GasProps.instance;
            const b = GasProps.instance;
            expect(a).toBe(b);
        });
    });

    describe('usersSheet', () => {
        it('4-2: usersSheet が settingSheet と異なる ID の場合に Users シートを使う', () => {
            const USERS_FIXTURE = [['lineName', 'densukeName', 'userId', 'role']];
            mockGetSheetByName.mockImplementation((name: string) =>
                name === 'Users' ? { getDataRange: () => ({ getValues: () => USERS_FIXTURE }) } : null
            );

            const sheet = GasProps.instance.usersSheet;
            expect(sheet).toBeDefined();
            expect(sheet.getDataRange().getValues()).toEqual(USERS_FIXTURE);
        });

        it('4-3: 対象シートが見つからない場合に Error がthrowされる', () => {
            mockGetSheetByName.mockReturnValue(null);
            expect(() => GasProps.instance.usersSheet).toThrow('usersSheet was not found');
        });
    });

    describe('settingSheet', () => {
        it('4-4: Settings シートが正常に取得できる', () => {
            const mockSheet = { id: 'settings' };
            mockGetSheetByName.mockImplementation((name: string) => (name === 'Settings' ? mockSheet : null));
            expect(GasProps.instance.settingSheet).toBe(mockSheet);
        });

        it('4-5: Settings シートが存在しない場合に Error がthrowされる', () => {
            mockGetSheetByName.mockReturnValue(null);
            expect(() => GasProps.instance.settingSheet).toThrow('settingSheet was not found');
        });
    });

    describe('videoSheet', () => {
        it('4-6: videos シートが存在しない場合に Error がthrowされる', () => {
            mockGetSheetByName.mockReturnValue(null);
            expect(() => GasProps.instance.videoSheet).toThrow('videos was not found');
        });
    });

    describe('cashBookSheet', () => {
        it('4-7: CashBook シートが存在しない場合に Error がthrowされる', () => {
            mockGetSheetByName.mockReturnValue(null);
            expect(() => GasProps.instance.cashBookSheet).toThrow('cashBookSheet was not found');
        });
    });

    describe('generateSheetUrl', () => {
        it('4-8: 通常モードでタイムスタンプ付き URL を返す', () => {
            (ScriptProps.isTesting as jest.Mock).mockReturnValue(false);
            const url = GasProps.instance.generateSheetUrl('sheet-abc');
            expect(url).toMatch(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/sheet-abc\/edit\?usp=sharing&ccc=\d+$/);
        });

        it('4-9: テストモードでタイムスタンプなしの URL を返す', () => {
            (ScriptProps.isTesting as jest.Mock).mockReturnValue(true);
            const url = GasProps.instance.generateSheetUrl('sheet-abc');
            expect(url).toBe('https://docs.google.com/spreadsheets/d/sheet-abc?usp=sharing');
        });
    });

    describe('weightRecordSheet', () => {
        it('4-11: WeightRecord シートが存在する場合に既存シートを返す', () => {
            const mockSheet = { id: 'weightRecord' };
            mockGetSheetByName.mockImplementation((name: string) => (name === 'WeightRecord' ? mockSheet : null));

            expect(GasProps.instance.weightRecordSheet).toBe(mockSheet);
            expect(mockInsertSheet).not.toHaveBeenCalled();
        });

        it('4-12: WeightRecord シートが存在しない場合に新規作成してヘッダーを設定する', () => {
            const mockSetValues = jest.fn();
            const mockNewSheet = {
                getRange: jest.fn().mockReturnValue({ setValues: mockSetValues }),
            };
            mockGetSheetByName.mockReturnValue(null);
            mockInsertSheet.mockReturnValue(mockNewSheet);

            GasProps.instance.weightRecordSheet;

            expect(mockInsertSheet).toHaveBeenCalledWith('WeightRecord');
            expect(mockSetValues).toHaveBeenCalledWith([['id', 'userId', 'height', 'weight', 'bfp', 'date']]);
        });
    });
});
