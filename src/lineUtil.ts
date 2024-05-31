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

  public sendLineReply(
    replyToken: string,
    messageText: string,
    imageUrl: string
  ): void {
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
    return this.getLineUserProfile(userId).displayName;
  }

  public getLineLang(userId: string): string {
    return this.getLineUserProfile(userId).language;
  }

  public getLineImage(messageId: string, fileName: string): string {
    const folder = GasProps.instance.payNowFolder;
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const headers = {
      Authorization: 'Bearer ' + ScriptProps.instance.lineAccessToken,
    };
    const response = UrlFetchApp.fetch(url, { headers: headers });
    const blob = response.getBlob().setName(fileName);
    const file = folder.createFile(blob);
    return file.getUrl();
  }
}
