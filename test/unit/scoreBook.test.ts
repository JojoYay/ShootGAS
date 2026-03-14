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
import { ScoreBook } from '../../src/scoreBook';

// ヘッダー9行 + 参加者データを持つシートを生成するヘルパー
function makeReportSheet(playerRows: unknown[][]): GoogleAppsScript.Spreadsheet.Sheet {
    const headerRows = Array.from({ length: 9 }, (_, i) => [`header${i}`, '', '']);
    const allValues = [...headerRows, ...playerRows];
    return {
        getDataRange: () => ({
            getValues: () => allValues,
        }),
    } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
}

// Spreadsheet モックを生成するヘルパー
function makeSpreadsheet(existingSheet: GoogleAppsScript.Spreadsheet.Sheet | null) {
    const mockInsertedSheet = {
        appendRow: jest.fn(),
        activate: jest.fn(),
    } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

    const mockSS = {
        getSheetByName: jest.fn().mockReturnValue(existingSheet),
        insertSheet: jest.fn().mockReturnValue(mockInsertedSheet),
        moveActiveSheet: jest.fn(),
    } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

    return { mockSS, mockInsertedSheet };
}

describe('ScoreBook', () => {
    describe('getAttendeesFromRecord', () => {
        it('7-1: 9行目以降の1列目を参加者として返す', () => {
            const sheet = makeReportSheet([['player1'], ['player2'], ['player3']]);
            const book = new ScoreBook();
            expect(book.getAttendeesFromRecord(sheet)).toEqual(['player1', 'player2', 'player3']);
        });

        it('7-2: ヘッダー行（0〜8行目）が参加者に含まれない', () => {
            const sheet = makeReportSheet([['player1']]);
            const book = new ScoreBook();
            const result = book.getAttendeesFromRecord(sheet);
            for (let i = 0; i < 9; i++) {
                expect(result).not.toContain(`header${i}`);
            }
        });

        it('7-3: データ行が 0 件の場合に空配列を返す', () => {
            const sheet = makeReportSheet([]);
            const book = new ScoreBook();
            expect(book.getAttendeesFromRecord(sheet)).toEqual([]);
        });
    });

    describe('getEventDetailSheet', () => {
        it('7-4: 既存シートが存在する場合に既存シートを返す', () => {
            const existingSheet = {
                getDataRange: jest.fn(),
            } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
            const { mockSS } = makeSpreadsheet(existingSheet);

            const book = new ScoreBook();
            const result = book.getEventDetailSheet(mockSS, '2024-01-01');
            expect(result).toBe(existingSheet);
            expect(mockSS.insertSheet).not.toHaveBeenCalled();
        });

        it('7-5: 存在しないシートの場合に新規作成して返す', () => {
            const { mockSS, mockInsertedSheet } = makeSpreadsheet(null);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            jest.spyOn(ScoreBook.prototype as any, 'moveSheetToHead').mockImplementation(() => {});

            const book = new ScoreBook();
            const result = book.getEventDetailSheet(mockSS, '2024-01-01');
            expect(mockSS.insertSheet).toHaveBeenCalledWith('2024-01-01');
            expect(result).toBe(mockInsertedSheet);
        });

        it("7-6: 新規作成時に ['名前', 'チーム', '得点', 'アシスト'] のヘッダー行を設定する", () => {
            const { mockSS, mockInsertedSheet } = makeSpreadsheet(null);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            jest.spyOn(ScoreBook.prototype as any, 'moveSheetToHead').mockImplementation(() => {});

            const book = new ScoreBook();
            book.getEventDetailSheet(mockSS, '2024-01-01');
            expect(mockInsertedSheet.appendRow).toHaveBeenCalledWith(['名前', 'チーム', '得点', 'アシスト']);
        });
    });

    describe('getTopPoint (private)', () => {
        it('7-7: 複数チームポイントの中から最大値を返す', () => {
            const eventRow = new Array(17).fill(0);
            eventRow[7] = 10; // チーム1
            eventRow[8] = 8; // チーム2
            eventRow[9] = 6; // チーム3

            const book = new ScoreBook();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            expect((book as any).getTopPoint(eventRow)).toBe(10);
        });

        it('7-8: 全チームポイントが 0 の場合に 0 を返す', () => {
            const eventRow = new Array(17).fill(0);

            const book = new ScoreBook();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            expect((book as any).getTopPoint(eventRow)).toBe(0);
        });
    });
});
