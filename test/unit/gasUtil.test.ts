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
 * Unit tests for GasUtil with mocked GasProps.
 * GasProps.instance.usersSheet is mocked to return fixed data.
 * For "usersSheet not set, fallback to DensukeMapping" behavior, see gasProps.test.ts.
 */
jest.mock('../../src/gasProps', () => ({
    GasProps: {
        instance: {
            usersSheet: {
                getDataRange: () => ({
                    getValues: () => [
                        ['lineName', 'densukeName', 'userId', 'role'],
                        ['line1', 'name1', 'Ukanji1', '幹事'],
                        ['line2', 'name2', 'UnotKanji', 'member'],
                        ['line3', 'nick3', 'Uuser3', 'member'],
                    ],
                }),
            },
        },
    },
}));

jest.mock('../../src/scriptProps', () => ({
    ScriptProps: {
        instance: {
            reportSheet: 'mock-report-sheet-id',
        },
    },
}));

import { GasUtil } from '../../src/gasUtil';

// ─── SpreadsheetApp グローバルモック ──────────────────────
// getReportSheet が使う SpreadsheetApp を各テストで差し替えられるよう
// jest.fn() として保持する
const mockGetSheetByName = jest.fn();
const mockInsertSheet = jest.fn();
const mockMoveActiveSheet = jest.fn();
const mockActivate = jest.fn();

beforeAll(() => {
    const mockSpreadsheet = {
        getSheetByName: mockGetSheetByName,
        insertSheet: mockInsertSheet,
        moveActiveSheet: mockMoveActiveSheet,
    };
    (global as unknown as { SpreadsheetApp: unknown }).SpreadsheetApp = {
        openById: jest.fn().mockReturnValue(mockSpreadsheet),
    };
});

beforeEach(() => {
    mockGetSheetByName.mockReset();
    mockInsertSheet.mockReset();
    mockMoveActiveSheet.mockReset();
    mockActivate.mockReset();
});

// ─────────────────────────────────────────────────────────

describe('GasUtil', () => {
    describe('isKanji', () => {
        it('3-1: 幹事ユーザー ID の場合に true を返す', () => {
            expect(new GasUtil().isKanji('Ukanji1')).toBe(true);
        });

        it('3-2: 非幹事ユーザー ID の場合に false を返す', () => {
            expect(new GasUtil().isKanji('UnotKanji')).toBe(false);
        });

        it('3-3: 未登録ユーザー ID の場合に false を返す', () => {
            expect(new GasUtil().isKanji('Uunknown')).toBe(false);
        });

        it('3-4: 空文字列の場合に false を返す', () => {
            expect(new GasUtil().isKanji('')).toBe(false);
        });
    });

    describe('getLineUserId', () => {
        it('3-5: 伝助名前に対応する LINE ID を返す', () => {
            expect(new GasUtil().getLineUserId('name1')).toBe('Ukanji1');
        });

        it('3-6: 存在しない伝助名前の場合に空文字列を返す', () => {
            expect(new GasUtil().getLineUserId('nonExistent')).toBe('');
        });
    });

    describe('getLineName', () => {
        it('3-7: 伝助名前に対応する LINE 表示名を返す', () => {
            expect(new GasUtil().getLineName('name1')).toBe('line1');
        });

        it('3-8: 存在しない伝助名前の場合に null を返す', () => {
            expect(new GasUtil().getLineName('nonExistent')).toBeNull();
        });
    });

    describe('getNickname', () => {
        it('3-9: ユーザー ID に対応するニックネームを返す', () => {
            expect(new GasUtil().getNickname('Uuser3')).toBe('nick3');
        });

        it('3-10: 存在しないユーザー ID の場合に null を返す', () => {
            expect(new GasUtil().getNickname('Uunknown')).toBeNull();
        });
    });

    describe('getDensukeName', () => {
        it('3-11: LINE 名前に対応する伝助名前を返す', () => {
            expect(new GasUtil().getDensukeName('line1')).toBe('name1');
        });

        it('3-12: 存在しない LINE 名前の場合に null を返す', () => {
            expect(new GasUtil().getDensukeName('nonExistent')).toBeNull();
        });
    });

    describe('getUnpaid', () => {
        it('3-13: 支払い列が空の参加者を配列で返す', () => {
            const mockSheet = {
                getDataRange: () => ({
                    getValues: () => [
                        // 行0〜8: ヘッダー
                        ...Array.from({ length: 9 }, (_, i) => [`header${i}`, '', '']),
                        // 行9: 支払い済み
                        ['player1', 'line1', 'paid'],
                        // 行10: 未払い
                        ['player2', 'line2', ''],
                        // 行11: 未払い
                        ['player3', 'line3', ''],
                    ],
                }),
            } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

            const util = new GasUtil();
            jest.spyOn(util, 'getReportSheet').mockReturnValue(mockSheet);
            expect(util.getUnpaid('2024-01-01')).toEqual(['player2', 'player3']);
        });

        it('3-14: 全員支払い済みの場合に空配列を返す', () => {
            const mockSheet = {
                getDataRange: () => ({
                    getValues: () => [
                        ...Array.from({ length: 9 }, (_, i) => [`header${i}`, '', '']),
                        ['player1', 'line1', 'paid'],
                        ['player2', 'line2', 'paid'],
                    ],
                }),
            } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

            const util = new GasUtil();
            jest.spyOn(util, 'getReportSheet').mockReturnValue(mockSheet);
            expect(util.getUnpaid('2024-01-01')).toEqual([]);
        });
    });

    describe('getReportSheet', () => {
        it('3-15: 既存シートが存在する場合にそのシートを返す', () => {
            const mockSheet = { getDataRange: jest.fn() } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
            mockGetSheetByName.mockReturnValue(mockSheet);

            expect(new GasUtil().getReportSheet('2024-01-01', false)).toBe(mockSheet);
        });

        it('3-16: 存在しないシートで isGenerate=false の場合に Error がthrowされる', () => {
            mockGetSheetByName.mockReturnValue(null);

            expect(() => new GasUtil().getReportSheet('nonExistent', false)).toThrow('reportSheet was not found. actDate:nonExistent');
        });

        it('3-17: 存在しないシートで isGenerate=true の場合に新規シートを作成して返す', () => {
            const mockNewSheet = {
                activate: mockActivate,
            } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
            mockGetSheetByName.mockReturnValue(null);
            mockInsertSheet.mockReturnValue(mockNewSheet);

            const result = new GasUtil().getReportSheet('newDate', true);
            expect(mockInsertSheet).toHaveBeenCalledWith('newDate');
            expect(result).toBe(mockNewSheet);
        });
    });
});
