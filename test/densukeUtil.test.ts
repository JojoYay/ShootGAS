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
import { DensukeUtil } from '../src/densukeUtil';
import * as Cheerio from 'cheerio';
import * as fs from 'fs';

describe('DensukeUtil', () => {
    let densukeUtil: DensukeUtil;
    let $: cheerio.Root;

    beforeEach(async () => {
        // 非同期処理を行うために async を追加
        densukeUtil = new DensukeUtil();
        // const mockDensuke = 'https://densuke.biz/list?cd=DTDR7Cu7rmkZy9YA';
        // const response = await fetch(mockDensuke); // fetch を使用して HTTP リクエストを送信
        // const html: string = await response.text(); // レスポンスのテキストを取得
        const html = fs.readFileSync('resouces/densukeHTml.html', 'utf-8');
        $ = Cheerio.load(html);
    });

    test('extractMembers should extract valid member names', () => {
        const members = densukeUtil.extractMembers($);
        expect(members).toEqual([
            '西村',
            '安田',
            '芦田',
            '岡本',
            '石濱',
            '西尾',
            '福田',
            '新谷',
            '松本な',
            'ロッキー',
            '渡邊',
            '四方',
            '德永',
            '森本',
            '阿部と',
            'スビ',
            '竹村',
            '濱直',
            '荒井',
            '八木',
            'おばたけ',
            '小林',
            '榎',
            'なべ',
            '磯崎',
            '望月',
            '山口',
            '塚本拓',
            '星',
            '梶原',
            'Sahim',
            'カヤバ',
            'ましも',
            'やまだじょ',
            '脇阪',
            '三田',
            '千葉',
            '松平',
            'Soma',
            '大内',
            'Warren',
            '安室',
            '豊田',
            '大里',
            'Suffian',
            '成瀬',
            '鈴木',
            '坂本',
            'さかもと',
        ]);
    });

    test('extractMembers should ignore non-member links', () => {
        const members = densukeUtil.extractMembers($);
        expect(members).not.toContain('ほげ田');
    });

    test('extractAttemdees should extract valid member names', () => {
        const members = densukeUtil.extractAttendees($, 1, '○', densukeUtil.extractMembers($));
        expect(members).toEqual([
            '西村',
            '芦田',
            '西尾',
            '新谷',
            'ロッキー',
            '德永',
            'スビ',
            '荒井',
            '小林',
            'なべ',
            '磯崎',
            '山口',
            '塚本拓',
            '梶原',
            'ましも',
            'やまだじょ',
            '脇阪',
            '松平',
            '安室',
            'Suffian',
            '鈴木',
            'さかもと',
        ]);
    });

    test('extractAttemdees should extract valid member names', () => {
        const members = densukeUtil.extractAttendees($, 1, '△', densukeUtil.extractMembers($));
        expect(members).toEqual(['石濱', '濱直', '大内', '豊田']);
    });

    test('extractMembers should extract valid member names', () => {
        const members = densukeUtil.extractDateFromRownum($, 1);
        expect(members).toEqual('6/2(日)');
    });
});
