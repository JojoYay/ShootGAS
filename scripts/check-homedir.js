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
 * デバッグ用: os.homedir()と認証ファイルのパスを確認するスクリプト
 */
import os from 'os';
import path from 'path';
import fs from 'fs';

console.log('=== 環境情報の確認 ===');
console.log('os.homedir():', os.homedir());
console.log('process.cwd():', process.cwd());
console.log('process.env.USERPROFILE:', process.env.USERPROFILE);
console.log('process.env.HOME:', process.env.HOME);
console.log('process.env.HOMEPATH:', process.env.HOMEPATH);

// claspが使用する認証ファイルのパス
const globalAuthPath = path.join(os.homedir(), '.clasprc.json');
const localAuthPath = path.join(process.cwd(), '.clasprc.json');

console.log('\n=== 認証ファイルのパス ===');
console.log('Global auth path:', globalAuthPath);
console.log('Local auth path:', localAuthPath);

console.log('\n=== 認証ファイルの存在確認 ===');
console.log('Global auth file exists:', fs.existsSync(globalAuthPath));
console.log('Local auth file exists:', fs.existsSync(localAuthPath));

if (fs.existsSync(globalAuthPath)) {
    try {
        const content = fs.readFileSync(globalAuthPath, 'utf8');
        const parsed = JSON.parse(content);
        console.log('\n=== Global認証ファイルの内容 ===');
        console.log('Raw content (first 500 chars):', content.substring(0, 500));
        console.log('Parsed keys:', Object.keys(parsed));
        
        // 新しい形式（claspが期待する形式）の確認
        console.log('\n--- 新しい形式（claspが期待する形式）---');
        console.log('Has token:', !!parsed.token);
        console.log('Has access_token (in token):', !!(parsed.token && parsed.token.access_token));
        console.log('Has refresh_token (in token):', !!(parsed.token && parsed.token.refresh_token));
        console.log('Is local creds:', parsed.isLocalCreds);
        if (parsed.token) {
            console.log('Token keys:', Object.keys(parsed.token));
            console.log('Token expiry_date:', parsed.token.expiry_date ? new Date(parsed.token.expiry_date).toISOString() : 'N/A');
        }
        
        // 古い形式（gcloud形式: tokens.default）の確認
        if (parsed.tokens && parsed.tokens.default) {
            console.log('\n--- 古い形式（gcloud形式）を検出 ---');
            console.log('tokens.default exists');
            console.log('Has access_token (in tokens.default):', !!parsed.tokens.default.access_token);
            console.log('Has refresh_token (in tokens.default):', !!parsed.tokens.default.refresh_token);
            console.log('\n⚠️  警告: この形式は clasp が期待する形式ではありません！');
            console.log('clasp logout を実行してから clasp login を再実行してください。');
        }
        
        // ルートレベルのaccess_token（別の古い形式）の確認
        if (parsed.access_token && !parsed.token) {
            console.log('\n--- 別の古い形式を検出 ---');
            console.log('access_token exists at root level');
            console.log('refresh_token exists:', !!parsed.refresh_token);
            console.log('\n⚠️  警告: この形式は clasp が期待する形式ではありません！');
        }
        
        // 正しい形式かどうかの判定
        const isValidFormat = parsed.token && parsed.token.access_token && typeof parsed.isLocalCreds === 'boolean';
        console.log('\n=== 形式の検証 ===');
        console.log('✅ 正しい形式:', isValidFormat ? 'YES' : 'NO');
        if (!isValidFormat) {
            console.log('❌ 認証ファイルの形式が正しくありません。');
            console.log('   期待される形式: { "token": { "access_token": "...", ... }, "isLocalCreds": false }');
        }
    } catch (error) {
        console.error('Error reading global auth file:', error.message);
        console.error('Stack:', error.stack);
    }
} else {
    console.log('\n⚠️  Global認証ファイルが存在しません。');
    console.log('   clasp login を実行してください。');
}

if (fs.existsSync(localAuthPath)) {
    try {
        const content = fs.readFileSync(localAuthPath, 'utf8');
        const parsed = JSON.parse(content);
        console.log('\n=== Local認証ファイルの内容 ===');
        console.log('Has token:', !!parsed.token);
        console.log('Has access_token:', !!(parsed.token && parsed.token.access_token));
        console.log('Has refresh_token:', !!(parsed.token && parsed.token.refresh_token));
        console.log('Is local creds:', parsed.isLocalCreds);
        if (parsed.token) {
            console.log('Token expiry_date:', parsed.token.expiry_date ? new Date(parsed.token.expiry_date).toISOString() : 'N/A');
        }
    } catch (error) {
        console.error('Error reading local auth file:', error.message);
    }
}

// claspの設定を確認
console.log('\n=== Clasp設定の確認 ===');
console.log('clasp_config_auth env:', process.env.clasp_config_auth);

