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

  private getLineUserProfile(userId: string) {
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
      if (userId === 'Ucb9beba3011ec9cf85c5482efa132e9b' || userId === 'tekitoutekitoutekitou' || userId === 'Uf395b2a8c82788dc3331b62f0cf96578') {
        return 'ja';
      }
    }
    return this.getLineUserProfile(userId).language;
  }

  public getLineImage(messageId: string, fileName: string): void {
    if (ScriptProps.isTesting()) {
      //テストの場合はコピーする
      const orgFolder: GoogleAppsScript.Drive.Folder = DriveApp.getFolderById('14FCKvswWbQTgkfHVmiHviYDNqDurAFXc');
      const files = orgFolder.getFilesByName('payNowSample.jpg');
      const file = files.next();
      const folder: GoogleAppsScript.Drive.Folder = GasProps.instance.payNowFolder;
      file.makeCopy(fileName, folder);
      return;
    }
    const folder = GasProps.instance.payNowFolder;
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const headers = {
      Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
    };
    const response = UrlFetchApp.fetch(url, { headers: headers });
    const blob = response.getBlob().setName(fileName);
    console.log('filename:' + fileName);
    folder.createFile(blob);
    // return file.getUrl();
  }
}
