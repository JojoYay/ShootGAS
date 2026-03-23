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
 * Unit tests for LiffApi#generateExReport
 *
 * Verifies:
 * - Report spreadsheet is created with correct header rows and column headers
 * - Participant names are resolved via dynamic column index ('LINE ID' / '伝助上の名前' / 'ライン上の名前' headers)
 * - Fallback to hardcoded indices (2/1/0) when the header row has no matching column names
 * - Users whose userId is not found in the mapping sheet are skipped
 * - receiveColumn=true adds a 6th column header and data validation
 * - receiveColumn=false uses only 5 columns
 */

// ─── GasProps mock (overridden per-test via mockUsersSheetValues) ─────────────
let mockUsersSheetValues: unknown[][] = [];

jest.mock('../../src/gasProps', () => ({
    GasProps: {
        instance: {
            get usersSheet() {
                return {
                    getDataRange: () => ({
                        getValues: () => mockUsersSheetValues,
                    }),
                };
            },
            get expenseFolder() {
                return mockExpenseFolder;
            },
            generateSheetUrl: (id: string) => `https://docs.google.com/spreadsheets/d/${id}`,
        },
    },
}));

jest.mock('../../src/scriptProps', () => ({
    ScriptProps: {
        instance: {
            folderId: 'mock-folder-id',
            liffUrl: 'https://liff.example.com',
        },
    },
}));

// LiffApi imports GasUtil and others – stub them out so GAS globals are not needed
jest.mock('../../src/gasUtil', () => ({ GasUtil: jest.fn() }));
jest.mock('../../src/lineUtil', () => ({ LineUtil: jest.fn() }));
jest.mock('../../src/schedulerUtil', () => ({ SchedulerUtil: jest.fn() }));
jest.mock('../../src/scoreBook', () => ({ ScoreBook: jest.fn() }));
jest.mock('../../src/requestExecuter', () => ({ RequestExecuter: jest.fn() }));

import { LiffApi } from '../../src/liffApi';
import { GetEventHandler } from '../../src/getEventHandler';

// ─── Shared mocks ─────────────────────────────────────────────────────────────

// Track calls to sheet.appendRow / getRange.setValue / getRange.setFormula
const mockAppendRow = jest.fn();
const mockGetRange = jest.fn();
const mockClear = jest.fn();
const mockSetBorder = jest.fn();
const mockSetBackground = jest.fn();
const mockSetFormula = jest.fn();
const mockSetValue = jest.fn();
const mockSetDataValidation = jest.fn();
const mockGetLastColumn = jest.fn().mockReturnValue(5);
const mockGetLastRow = jest.fn().mockReturnValue(10);

// range object returned by sheet.getRange(...)
const mockRange = {
    setValue: mockSetValue,
    setFormula: mockSetFormula,
    setBorder: mockSetBorder,
    setBackground: mockSetBackground,
    setDataValidation: mockSetDataValidation,
    getA1Notation: jest.fn().mockReturnValue('D6:D10'),
};

const mockSheet = {
    clear: mockClear,
    appendRow: mockAppendRow,
    getRange: mockGetRange,
    getLastColumn: mockGetLastColumn,
    getLastRow: mockGetLastRow,
};
mockGetRange.mockReturnValue(mockRange);

// newSpreadsheet
const mockNewSpreadsheet = {
    getActiveSheet: jest.fn().mockReturnValue(mockSheet),
    getId: jest.fn().mockReturnValue('mock-spreadsheet-id'),
};

// DriveApp mock
const mockMoveTo = jest.fn();
const mockFileForMove = { moveTo: mockMoveTo };
const mockDriveGetFileById = jest.fn().mockReturnValue(mockFileForMove);

// expenseFolder mock
const mockExpenseFolderSearchFiles = jest.fn();
const mockExpenseFolderGetId = jest.fn().mockReturnValue('mock-expense-folder-id');
const mockExpenseFolderCreateFolder = jest.fn();
let mockExpenseFolder: unknown;

// SpreadsheetApp mock
const mockSpreadsheetCreate = jest.fn().mockReturnValue(mockNewSpreadsheet);
const mockSpreadsheetOpenById = jest.fn().mockReturnValue(mockNewSpreadsheet);
const mockDataValidationBuild = jest.fn().mockReturnValue('mock-validation');
const mockRequireValueInList = jest.fn().mockReturnValue({ build: mockDataValidationBuild });
const mockNewDataValidation = jest.fn().mockReturnValue({ requireValueInList: mockRequireValueInList });

beforeAll(() => {
    (global as unknown as Record<string, unknown>).SpreadsheetApp = {
        create: mockSpreadsheetCreate,
        openById: mockSpreadsheetOpenById,
        newDataValidation: mockNewDataValidation,
    };
    (global as unknown as Record<string, unknown>).DriveApp = {
        getFileById: mockDriveGetFileById,
    };
});

beforeEach(() => {
    jest.clearAllMocks();
    mockGetRange.mockReturnValue(mockRange);
    mockGetLastColumn.mockReturnValue(5);
    mockGetLastRow.mockReturnValue(10);
    mockNewSpreadsheet.getActiveSheet.mockReturnValue(mockSheet);
    mockNewSpreadsheet.getId.mockReturnValue('mock-spreadsheet-id');
    mockNewDataValidation.mockReturnValue({ requireValueInList: mockRequireValueInList });
    mockRequireValueInList.mockReturnValue({ build: mockDataValidationBuild });
    mockDataValidationBuild.mockReturnValue('mock-validation');

    // Default: folder has no existing sub-folder or file -> create new
    const mockSubFolderIt = { hasNext: () => false };
    const mockFileIt = { hasNext: () => false };
    mockExpenseFolderCreateFolder.mockReturnValue({
        getId: mockExpenseFolderGetId,
        searchFiles: mockExpenseFolderSearchFiles,
    });
    mockExpenseFolderSearchFiles.mockReturnValue(mockFileIt);
    mockExpenseFolder = {
        getFoldersByName: jest.fn().mockReturnValue(mockSubFolderIt),
        createFolder: mockExpenseFolderCreateFolder,
        getId: mockExpenseFolderGetId,
        searchFiles: mockExpenseFolderSearchFiles,
    };
});

// ─── Helper to call the private generateExReport method directly ─────────────

function callGenerateExReport(params: { users: string[]; price: string; title: string; payNow: string; receiveColumn: string }): GetEventHandler {
    const e = {
        parameters: {
            func: ['generateExReport'],
            users: params.users,
            price: [params.price],
            title: [params.title],
            payNow: [params.payNow],
            receiveColumn: [params.receiveColumn],
        },
        parameter: {} as Record<string, string>,
    } as unknown as GoogleAppsScript.Events.DoGet;
    const handler = new GetEventHandler(e);
    const liffApi = new LiffApi();
    // generateExReport is private; call via bracket notation
    (liffApi as unknown as Record<string, (_h: GetEventHandler) => void>)['generateExReport'](handler);
    return handler;
}

// ─── Tests ───────────────────────────────────────────────────────────────────

describe('LiffApi#generateExReport', () => {
    // Users sheet with named headers matching the new 'Users' sheet structure
    const newUsersSheetData = [
        ['ライン上の名前', '伝助上の名前', 'LINE ID', '幹事フラグ', 'Picture'],
        ['LineAlice', 'Alice', 'U001', '', ''],
        ['LineBob', 'Bob', 'U002', '', ''],
        ['LineCarol', 'Carol', 'U003', '', ''],
    ];

    describe('ヘッダー行の書き込み', () => {
        beforeEach(() => {
            mockUsersSheetValues = newUsersSheetData;
        });

        it('9-1: appendRow で 名称・人数・合計金額・PayNow先 の順に書き込まれる', () => {
            callGenerateExReport({
                users: ['U001', 'U002'],
                price: '1000',
                title: 'テスト清算',
                payNow: 'PayNow@example',
                receiveColumn: 'false',
            });

            expect(mockAppendRow).toHaveBeenNthCalledWith(1, ['名称', 'テスト清算']);
            expect(mockAppendRow).toHaveBeenNthCalledWith(2, ['人数', 2]);
            expect(mockAppendRow).toHaveBeenNthCalledWith(3, ['合計金額', 2000]);
            expect(mockAppendRow).toHaveBeenNthCalledWith(4, ['PayNow先', 'PayNow@example']);
        });

        it('9-2: receiveColumn=false のとき 5列のカラムヘッダーが書き込まれる', () => {
            callGenerateExReport({
                users: ['U001'],
                price: '500',
                title: 'テスト',
                payNow: 'pay@example',
                receiveColumn: 'false',
            });

            expect(mockAppendRow).toHaveBeenCalledWith(['参加者（伝助名称）', '参加者（Line名称）', 'LINE_ID', '金額', '支払い状況']);
        });

        it('9-3: receiveColumn=true のとき 6列のカラムヘッダーが書き込まれる', () => {
            callGenerateExReport({
                users: ['U001'],
                price: '500',
                title: 'テスト',
                payNow: 'pay@example',
                receiveColumn: 'true',
            });

            expect(mockAppendRow).toHaveBeenCalledWith(['参加者（伝助名称）', '参加者（Line名称）', 'LINE_ID', '金額', '支払い状況', '受け取り状況']);
        });

        it('9-4: 合計金額は users.length × price の積になる', () => {
            callGenerateExReport({
                users: ['U001', 'U002', 'U003'],
                price: '3000',
                title: '合計テスト',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            expect(mockAppendRow).toHaveBeenCalledWith(['合計金額', 9000]);
        });
    });

    describe('参加者名の解決（動的カラムインデックス）', () => {
        beforeEach(() => {
            mockUsersSheetValues = newUsersSheetData;
        });

        it('9-5: LINE ID ヘッダーが存在するとき、該当ユーザーの伝助名称とライン名称が正しくセットされる', () => {
            callGenerateExReport({
                users: ['U002'],
                price: '500',
                title: 'テスト清算',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            // row index=6, col1=伝助名称 (Bob), col2=ライン名称 (LineBob), col3=LINE ID (U002)
            expect(mockGetRange).toHaveBeenCalledWith(6, 1);
            expect(mockGetRange).toHaveBeenCalledWith(6, 2);
            expect(mockGetRange).toHaveBeenCalledWith(6, 3);

            // setValue の順番: col1=Bob, col2=LineBob, col3=U002, col4=price
            const setValueCalls = mockSetValue.mock.calls;
            expect(setValueCalls[0][0]).toBe('Bob'); // col1: 伝助上の名前
            expect(setValueCalls[1][0]).toBe('LineBob'); // col2: ライン上の名前
            expect(setValueCalls[2][0]).toBe('U002'); // col3: LINE ID
            expect(setValueCalls[3][0]).toBe('500'); // col4: price
        });

        it('9-6: 複数ユーザーの場合、行インデックスが 6 から順に増加する', () => {
            callGenerateExReport({
                users: ['U001', 'U003'],
                price: '1000',
                title: 'テスト',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            // 1人目: row6, 2人目: row7
            const getRangeCalls = mockGetRange.mock.calls.filter((c: unknown[]) => typeof c[0] === 'number' && typeof c[1] === 'number' && c[0] >= 6);
            const rows = getRangeCalls.map((c: unknown[]) => c[0] as number);
            expect(rows).toContain(6);
            expect(rows).toContain(7);
        });

        it('9-7: マッピングシートに存在しない userId はスキップされ、行が書き込まれない', () => {
            callGenerateExReport({
                users: ['U001', 'U_NOT_EXIST', 'U003'],
                price: '1000',
                title: 'スキップテスト',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            // U001→row6, U003→row7 (U_NOT_EXISTはskip)
            const getRangeCalls = mockGetRange.mock.calls.filter((c: unknown[]) => typeof c[0] === 'number' && typeof c[1] === 'number' && c[0] >= 6);
            const rows = [...new Set(getRangeCalls.map((c: unknown[]) => c[0] as number))];
            expect(rows).toContain(6);
            expect(rows).toContain(7);
            expect(rows).not.toContain(8); // 3行目は書かれない
        });

        it('9-8: 全ユーザーがマッピングシートに存在しない場合、参加者行は書き込まれない', () => {
            callGenerateExReport({
                users: ['U_NONE'],
                price: '1000',
                title: 'スキップ全員',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            // 参加者データが存在しないので setValue は一切呼ばれない
            expect(mockSetValue).not.toHaveBeenCalled();
        });
    });

    describe('フォールバック（レガシー DensukeMapping 形式）', () => {
        it('9-9: LINE ID ヘッダーがない場合、index 2 がユーザーID列として使われる', () => {
            // Legacy format: no header names, LINE ID is at index 2
            mockUsersSheetValues = [
                ['ライン名称', '伝助名称', 'U_LEGACY_ID'], // no 'LINE ID' in header
                ['LineX', 'DensukeX', 'ULEGACY001'],
            ];

            callGenerateExReport({
                users: ['ULEGACY001'],
                price: '500',
                title: 'レガシーテスト',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            // Should find the user at index 2 (fallback)
            const setValueCalls = mockSetValue.mock.calls;
            expect(setValueCalls[2][0]).toBe('ULEGACY001'); // col3: userId written
        });

        it('9-10: 伝助上の名前 ヘッダーがない場合、index 1 が名前列として使われる', () => {
            mockUsersSheetValues = [
                ['lineName', 'densukeName', 'LINE ID'], // '伝助上の名前' not present
                ['LineAlice', 'Alice', 'U001'],
            ];

            callGenerateExReport({
                users: ['U001'],
                price: '500',
                title: 'フォールバック名前',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            const setValueCalls = mockSetValue.mock.calls;
            // col1: index 1 = 'Alice', col2: index 0 = 'LineAlice', col3: 'U001'
            expect(setValueCalls[0][0]).toBe('Alice');
            expect(setValueCalls[1][0]).toBe('LineAlice');
        });
    });

    describe('receiveColumn の挙動', () => {
        beforeEach(() => {
            mockUsersSheetValues = newUsersSheetData;
        });

        it('9-11: receiveColumn=true のとき、参加者行の第6列にデータバリデーションがセットされる', () => {
            callGenerateExReport({
                users: ['U001'],
                price: '1000',
                title: 'バリデーションテスト',
                payNow: 'pay',
                receiveColumn: 'true',
            });

            expect(mockSetDataValidation).toHaveBeenCalledWith('mock-validation');
        });

        it('9-12: receiveColumn=false のとき、データバリデーションはセットされない', () => {
            callGenerateExReport({
                users: ['U001'],
                price: '1000',
                title: 'バリデーションなし',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            expect(mockSetDataValidation).not.toHaveBeenCalled();
        });
    });

    describe('結果オブジェクト', () => {
        beforeEach(() => {
            mockUsersSheetValues = newUsersSheetData;
        });

        it('9-13: result.sheet, result.folder, result.url が設定される', () => {
            const handler = callGenerateExReport({
                users: ['U001'],
                price: '500',
                title: 'URL テスト',
                payNow: 'pay',
                receiveColumn: 'false',
            });

            expect(handler.result.sheet).toContain('mock-spreadsheet-id');
            expect(handler.result.folder).toContain('mock-folder-id');
            expect(handler.result.url).toContain('/expense/input?title=URL テスト');
        });
    });
});
