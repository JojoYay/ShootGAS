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
import { TotalScore } from '../../src/totalScore';

// チーム名→インデックスのマッピング
const TEAM_INDEX: Record<string, number> = {
    チーム1: 7,
    チーム2: 8,
    チーム3: 9,
    チーム4: 10,
    チーム5: 11,
    チーム6: 12,
    チーム7: 13,
    チーム8: 14,
    チーム9: 15,
    チーム10: 16,
};

/** eventDataRow を生成する。指定インデックスにだけ値を入れ、残りは 0 */
function makeEventRow(pointIndex: number, pointValue: number): unknown[] {
    const row = new Array(17).fill(0);
    row[pointIndex] = pointValue;
    return row;
}

describe('TotalScore', () => {
    describe('constructor', () => {
        it('6-12: 全フィールドが初期値で生成される', () => {
            const score = new TotalScore();
            expect(score.userId).toBe('');
            expect(score.name).toBe('');
            expect(score.playTime).toBe(0);
            expect(score.sunnyPlay).toBe(0);
            expect(score.rainyPlay).toBe(0);
            expect(score.goalCount).toBe(0);
            expect(score.assistCount).toBe(0);
            expect(score.mipCount).toBe(0);
            expect(score.teamPoint).toBe(0);
            expect(score.winCount).toBe(0);
            expect(score.loseCount).toBe(0);
            expect(score.totalMatchs).toBe(0);
            expect(score.totalOrank).toBe(0);
            expect(score.totalGrank).toBe(0);
            expect(score.totalArank).toBe(0);
        });
    });

    describe('fetchTeamPoint', () => {
        it.each(Object.entries(TEAM_INDEX))('6-%#: fetchTeamPoint が %s のポイントを eventDataRow[%i] から返す', (teamName, idx) => {
            const score = new TotalScore();
            const eventRow = makeEventRow(idx, idx * 10);
            expect(score.fetchTeamPoint(eventRow, teamName)).toBe(idx * 10);
        });

        it('6-11: 未定義のチーム名の場合に 0 を返す', () => {
            const score = new TotalScore();
            const eventRow = makeEventRow(7, 99);
            expect(score.fetchTeamPoint(eventRow, 'チーム99')).toBe(0);
        });
    });
});
