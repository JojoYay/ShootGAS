import { GasProps } from './gasProps';
import { ScriptProps } from './scriptProps';

export class LineUtil {
    public sendLineMessage(userId: string, message: string): void {
        if (userId) {
            const url = 'https://api.line.me/v2/bot/message/push';
            const headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + ScriptProps.instance.lineAccessToken,
            };
            const postData = {
                to: userId,
                messages: [
                    {
                        type: 'text',
                        text: message,
                    },
                ],
            };
            const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
                method: 'post',
                headers: headers,
                payload: JSON.stringify(postData),
            };
            const response = UrlFetchApp.fetch(url, options);
            Logger.log(response.getContentText());
        }
    }

    public sendFlexReply(replyToken: string, flexJson: JSON | null): void {
        const url = 'https://api.line.me/v2/bot/message/reply';
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const postData = {
            replyToken: replyToken,
            messages: [
                {
                    type: 'flex',
                    altText: 'This is a Flex Message',
                    contents: flexJson,
                },
            ],
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            payload: JSON.stringify(postData),
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        Logger.log(response.getContentText());
    }

    public sendLineReply(replyToken: string, messageText: string, imageUrl: string | null): void {
        const url = 'https://api.line.me/v2/bot/message/reply';
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const postData = {
            replyToken: replyToken,
            messages: [
                {
                    type: 'text',
                    text: messageText,
                },
            ],
        };
        if (imageUrl) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const imgObj: any = {
                type: 'image',
                originalContentUrl: imageUrl,
                previewImageUrl: imageUrl,
            };
            postData.messages.push(imgObj);
        }
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            payload: JSON.stringify(postData),
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        Logger.log(response.getContentText());
    }

    public getLineProileLite(userId: string) {
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.usersSheet;
        const values = mappingSheet.getDataRange().getValues();
        return values.find(row => row[2] === userId);
    }

    public getLineUserProfile(userId: string) {
        const url = `https://api.line.me/v2/bot/profile/${userId}`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'get',
            headers: headers,
        };
        const response = UrlFetchApp.fetch(url, options);
        const userProfile = JSON.parse(response.getContentText());
        return userProfile;
    }

    public getLineDisplayName(userId: string): string {
        if (ScriptProps.isTesting()) {
            if (userId === 'Ucb9beba3011ec9cf85c5482efa132e9b') {
                return '相馬究(Kiwamu Soma)';
            } else if (userId === 'tekitoutekitoutekitou') {
                return 'なべLINE';
            } else if (userId === 'Uf395b2a8c82788dc3331b62f0cf96578') {
                return 'Takashi Chiba';
            }
        }
        return this.getLineUserProfile(userId).displayName;
    }

    public getLineLang(userId: string): string {
        if (ScriptProps.isTesting()) {
            if (
                userId === 'Ucb9beba3011ec9cf85c5482efa132e9b' ||
                userId === 'tekitoutekitoutekitou' ||
                userId === 'Uf395b2a8c82788dc3331b62f0cf96578'
            ) {
                return 'ja';
            }
        }
        return this.getLineUserProfile(userId).language;
    }

    public getLineImage(messageId: string, fileName: string, actDate: string): void {
        //まずフォルダが無ければ作る
        const folder: GoogleAppsScript.Drive.Folder | null = this.createPayNowFolder(actDate);
        if (!folder) {
            return; //フォルダは必ず作られる（trueなので）
        }
        if (ScriptProps.isTesting()) {
            //テストの場合はコピーする
            const orgFolder: GoogleAppsScript.Drive.Folder = DriveApp.getFolderById('14FCKvswWbQTgkfHVmiHviYDNqDurAFXc');
            const files = orgFolder.getFilesByName('payNowSample.jpg');
            const file = files.next();
            file.makeCopy(fileName, folder);
            return;
        }
        const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const response = UrlFetchApp.fetch(url, { headers: headers });
        const blob = response.getBlob().setName(fileName);
        // console.log('filename:' + fileName);
        folder.createFile(blob);
    }

    public createPayNowFolder(actDate: string, create: boolean = true): GoogleAppsScript.Drive.Folder | null {
        const parentFolder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder; // 親フォルダを取得
        let folder: GoogleAppsScript.Drive.Folder | GoogleAppsScript.Drive.FolderIterator = parentFolder.getFoldersByName(actDate); // actDate フォルダを検索
        if (!folder.hasNext()) {
            // actDate フォルダが存在しない場合
            if (create) {
                folder = parentFolder.createFolder(actDate); // actDate フォルダを作成
            } else {
                return null;
            }
        } else {
            // actDate フォルダが存在する場合
            folder = folder.next(); // 最初のフォルダを取得
        }
        return folder;
    }

    // ─── Rich Menu API ─────────────────────────────────────────────────────

    /**
     * LINE API にリッチメニューを作成し、richMenuId を返す。
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public createRichMenu(richMenuJson: any): string {
        const url = 'https://api.line.me/v2/bot/richmenu';
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            payload: JSON.stringify(richMenuJson),
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        const result = JSON.parse(response.getContentText());
        if (response.getResponseCode() !== 200) {
            throw new Error(`createRichMenu failed: ${response.getContentText()}`);
        }
        console.log('createRichMenu: ' + result.richMenuId);
        return result.richMenuId;
    }

    /**
     * リッチメニューに背景画像をアップロードする。
     * LINE の image upload エンドポイントは api-data.line.me を使う。
     */
    public uploadRichMenuImage(richMenuId: string, imageBlob: GoogleAppsScript.Base.Blob): void {
        const url = `https://api-data.line.me/v2/bot/richmenu/${richMenuId}/content`;
        const headers = {
            'Content-Type': imageBlob.getContentType() || 'image/png',
            'Authorization': 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            payload: imageBlob.getBytes(),
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() !== 200) {
            throw new Error(`uploadRichMenuImage failed: ${response.getContentText()}`);
        }
        console.log('uploadRichMenuImage success for: ' + richMenuId);
    }

    /**
     * ユーザーにリッチメニューをリンクする。
     */
    public linkRichMenuToUser(userId: string, richMenuId: string): void {
        const url = `https://api.line.me/v2/bot/user/${userId}/richmenu/${richMenuId}`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() !== 200) {
            throw new Error(`linkRichMenuToUser failed: ${response.getContentText()}`);
        }
        console.log(`linkRichMenuToUser: ${userId} → ${richMenuId}`);
    }

    /**
     * ユーザーからリッチメニューのリンクを解除する。
     */
    public unlinkRichMenuFromUser(userId: string): void {
        const url = `https://api.line.me/v2/bot/user/${userId}/richmenu`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'delete',
            headers: headers,
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        // 404 = no rich menu linked (not an error)
        if (response.getResponseCode() !== 200 && response.getResponseCode() !== 404) {
            throw new Error(`unlinkRichMenuFromUser failed: ${response.getContentText()}`);
        }
        console.log('unlinkRichMenuFromUser: ' + userId);
    }

    /**
     * リッチメニューをデフォルト（全ユーザー）に設定する。
     * 個別にリンクされていないユーザー全員にこのメニューが表示される。
     */
    public setDefaultRichMenu(richMenuId: string): void {
        const url = `https://api.line.me/v2/bot/user/all/richmenu/${richMenuId}`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: headers,
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() !== 200) {
            throw new Error(`setDefaultRichMenu failed: ${response.getContentText()}`);
        }
        console.log('setDefaultRichMenu: ' + richMenuId);
    }

    /**
     * デフォルトリッチメニューの設定を解除する。
     */
    public cancelDefaultRichMenu(): void {
        const url = 'https://api.line.me/v2/bot/user/all/richmenu';
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'delete',
            headers: headers,
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        // 404 = no default rich menu set (not an error)
        if (response.getResponseCode() !== 200 && response.getResponseCode() !== 404) {
            throw new Error(`cancelDefaultRichMenu failed: ${response.getContentText()}`);
        }
        console.log('cancelDefaultRichMenu done');
    }

    /**
     * LINE API からリッチメニューを削除する。
     */
    public deleteRichMenu(richMenuId: string): void {
        const url = `https://api.line.me/v2/bot/richmenu/${richMenuId}`;
        const headers = {
            Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'delete',
            headers: headers,
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() !== 200 && response.getResponseCode() !== 404) {
            throw new Error(`deleteRichMenu failed: ${response.getContentText()}`);
        }
        console.log('deleteRichMenu: ' + richMenuId);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public getCarouselBase(): any {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const baseObj: any = {
            type: 'carousel',
            contents: [],
        };
        return baseObj;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public getYoutubeCard(): any {
        const obj = {
            type: 'bubble',
            body: {
                type: 'box',
                layout: 'vertical',
                contents: [
                    {
                        type: 'image',
                        url: 'https://img.youtube.com/vi/kNuUeydJZ8I/maxresdefault.jpg',
                        size: 'full',
                        aspectMode: 'cover',
                        offsetTop: '0px',
                        offsetStart: '0px',
                        position: 'absolute',
                        align: 'center',
                        aspectRatio: '16:9',
                    },
                    {
                        type: 'image',
                        url: 'https://lh3.googleusercontent.com/d/1oL3dEwuroPj4rylysUwVDkz0OY8AXGZZ',
                        position: 'absolute',
                        size: 'full',
                        offsetTop: '0px',
                        offsetStart: '0px',
                        aspectMode: 'cover',
                        align: 'center',
                        aspectRatio: '16:9',
                    },
                    {
                        type: 'text',
                        text: 'text11111',
                        size: 'xl',
                        color: '#FFFFFF',
                        position: 'absolute',
                        margin: 'sm',
                        offsetStart: '20px',
                        offsetTop: '10px',
                    },
                    {
                        type: 'text',
                        text: 'Date Text',
                        position: 'absolute',
                        color: '#FFFFFF',
                        offsetTop: '45px',
                        margin: 'sm',
                        offsetStart: '20px',
                    },
                ],
                height: '180px',
                action: {
                    type: 'uri',
                    label: 'action',
                    uri: 'URLURLURL',
                },
                width: '360px',
            },
        };
        return obj;
    }
}
