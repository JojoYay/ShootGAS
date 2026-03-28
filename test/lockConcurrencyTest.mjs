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
 * LockService 並行処理テスト
 *
 * 目的:
 *   6チーム2面 Slot1（T1vsT2 と T3vsT4）を2台同時に送信し、
 *   LL1 と WL1 のスロットが正しく埋まるか（上書きされないか）を確認する。
 *
 * 使い方:
 *   node test/lockConcurrencyTest.mjs [dev|prod]
 *
 * 前提:
 *   - デプロイ済みの GAS エンドポイントがあること
 *   - テスト当日の VideoSheet に 6チーム2面の行が作成済みであること
 *     （createShootLog で先に作っておく）
 *   - ACTUAL_DATE を実際の actDate に書き換えること（例: 日曜定期(28 Mar)）
 */

// ─── 設定 ───────────────────────────────────────────────────
const ENDPOINTS = {
    dev: 'https://script.google.com/macros/s/AKfycbwf7HIt4NTEyJAMPACPS7BGBiq9CJ14wg0jBwjWdj6yCLz9wdWkmzN0PjdqzpnCUBWrSA/exec',
    prod: 'https://script.google.com/macros/s/AKfycbysS9XmULaZ9HtsLjWDYKrEuNZ7ws9xtK-Qmy3jbNfrCKYXU3upWYybefshFlX2ypfG/exec',
};

// ★ここを実際の actDate に変更してください（例: 日曜定期(28 Mar)）
const ACTUAL_DATE = '日曜定期(28 Mar)';

// Slot1 の2試合（同時送信するリクエスト）
const REQUESTS = [
    {
        label: 'T1vsT2（携帯A）',
        matchId: `${ACTUAL_DATE}-6_d1_1`,
        winningTeam: 'Team1', // ★実際のチーム名に変更
        team1Players: 'PlayerA, PlayerB',
        team2Players: 'PlayerC, PlayerD',
    },
    {
        label: 'T3vsT4（携帯B）',
        matchId: `${ACTUAL_DATE}-6_d1_2`,
        winningTeam: 'Team3', // ★実際のチーム名に変更
        team1Players: 'PlayerE, PlayerF',
        team2Players: 'PlayerG, PlayerH',
    },
];
// ────────────────────────────────────────────────────────────

const env = process.argv[2] || 'dev';
const url = ENDPOINTS[env];
if (!url) {
    throw new Error(`❌ 不明な環境: ${env}。 dev または prod を指定してください。`);
}

console.log(`\n🔧 テスト環境: ${env}`);
console.log(`🌐 エンドポイント: ${url}`);
console.log(`📅 actDate: ${ACTUAL_DATE}\n`);

/**
 * 1件の closeGame リクエストを送信する
 */
async function sendCloseGame(req) {
    const form = new FormData();
    form.append('matchId', req.matchId);
    form.append('winningTeam', req.winningTeam);
    form.append('team1Players', req.team1Players);
    form.append('team2Players', req.team2Players);

    const start = Date.now();
    console.log(`📤 [${req.label}] 送信開始 matchId=${req.matchId}`);

    try {
        const res = await fetch(url, {
            method: 'POST',
            body: form,
            // GAS は redirect するので follow が必要
            redirect: 'follow',
        });
        const elapsed = Date.now() - start;
        const text = await res.text();
        let json;
        try {
            json = JSON.parse(text);
        } catch {
            json = { raw: text };
        }

        console.log(`✅ [${req.label}] 完了 ${elapsed}ms →`, json);
        return { label: req.label, ok: res.ok, elapsed, json };
    } catch (err) {
        const elapsed = Date.now() - start;
        console.error(`❌ [${req.label}] エラー ${elapsed}ms →`, err.message);
        return { label: req.label, ok: false, elapsed, error: err.message };
    }
}

// ─── メイン: 2リクエストを同時送信 ───────────────────────────
console.log('🚀 2リクエストを同時送信します...\n');

const results = await Promise.all(REQUESTS.map(sendCloseGame));

console.log('\n─────────────────────────────────────────');
console.log('📊 結果サマリー');
console.log('─────────────────────────────────────────');
for (const r of results) {
    const status = r.ok ? '✅ 成功' : '❌ 失敗';
    console.log(`${status} ${r.label}  (${r.elapsed}ms)`);
}

console.log('\n─────────────────────────────────────────');
console.log('📋 次のステップ: Spreadsheet で以下を確認してください');
console.log('─────────────────────────────────────────');
console.log(`  VideoSheet の ${ACTUAL_DATE}-6_ll1 行:`);
console.log('    → team1 と team2 が別チームで正しく埋まっているか？');
console.log('    → どちらかが上書きされていないか？');
console.log(`  VideoSheet の ${ACTUAL_DATE}-6_wl1 行:`);
console.log('    → 同様に team1/team2 が正しく埋まっているか？');
console.log('');
console.log('  ✅ 期待結果: ll1/wl1 それぞれに2チームが正しく入る');
console.log('  ❌ NG結果:  同じチームが2つ入る、または片方が空');
