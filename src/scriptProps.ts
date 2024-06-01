export class ScriptProps {
  private static _instance: ScriptProps | null = null;
  private static _mode: string = 'normal';

  private constructor() {}

  public static get instance(): ScriptProps {
    if (!this._instance) {
      this._instance = new ScriptProps();
    }
    return this._instance;
  }

  // public SETTING_SHEET_NAME: string = 'Settings';
  // public CASH_BOOK_SHEET_NAME: string = 'CashBook';
  // public MAPPING_SHEET_NAME: string = 'DensukeMapping';
  public ROWNUM: number = 1; //とりあえず一番上からデータとってくる運用

  public get reportSheet(): string {
    let reportProp: string | null = PropertiesService.getScriptProperties().getProperty('reportSheet');
    if (!reportProp) {
      throw new Error('Script Property (reportSheet) was not found');
    }
    if (ScriptProps._mode === 'test') {
      reportProp = '1Ej-9kZIMpGW66BUm0cGS1iG1RDgUsgBd5fo5V97xirg';
    }
    return reportProp;
  }

  public get settingSheet(): string {
    let settingProp: string | null = PropertiesService.getScriptProperties().getProperty('settingSheet');
    if (!settingProp) {
      throw new Error('Script Property (settingProp) was not found');
    }
    if (ScriptProps._mode === 'test') {
      settingProp = '1PfBvcVqO_d-JIs6VxwSJLW0GZAS0c6xfsOMUQKlTU30';
    }
    return settingProp;
  }

  public get lineAccessToken(): string {
    const lineAccessTokenProp: string | null = PropertiesService.getScriptProperties().getProperty('lineAccessToken');
    if (!lineAccessTokenProp) {
      throw new Error('Script Property (lineAccessToken) was not found');
    }
    return lineAccessTokenProp;
  }

  public get folderId(): string {
    let folderProp: string | null = PropertiesService.getScriptProperties().getProperty('folderId');
    if (!folderProp) {
      throw new Error('Script Property (folderProp) was not found');
    }
    if (ScriptProps._mode === 'test') {
      folderProp = '1mbldDFoy4pBEfXOONPcwb8HIyknQTzdV';
    }
    return folderProp;
  }

  public get archiveFolder(): string {
    let archiveFolderProp: string | null = PropertiesService.getScriptProperties().getProperty('archiveFolder');
    if (!archiveFolderProp) {
      throw new Error('Script Property (archiveFolder) was not found');
    }
    if (ScriptProps._mode === 'test') {
      archiveFolderProp = '10a1hJaEBPDTyl8d33hRWn2JiYZpCm6Vd';
    }
    return archiveFolderProp;
  }

  public get channelQr(): string {
    const channelQrProp: string | null = PropertiesService.getScriptProperties().getProperty('channelQr');
    if (!channelQrProp) {
      throw new Error('Script Property (channelQr) was not found');
    }
    return channelQrProp;
  }

  public get channelUrl(): string {
    const channelUrlProp: string | null = PropertiesService.getScriptProperties().getProperty('channelUrl');
    if (!channelUrlProp) {
      throw new Error('Script Property (channelUrl) was not found');
    }
    return channelUrlProp;
  }

  public get messageUsage(): string {
    const messageUsageProp: string | null = PropertiesService.getScriptProperties().getProperty('messageUsage');
    if (!messageUsageProp) {
      throw new Error('Script Property (messageUsage) was not found');
    }
    return messageUsageProp;
  }

  public get chat(): string {
    const chatProp: string | null = PropertiesService.getScriptProperties().getProperty('chat');
    if (!chatProp) {
      throw new Error('Script Property (chat) was not found');
    }
    return chatProp;
  }

  public static startTest(): void {
    console.log('++++++++ TEST START +++++++++++++');
    ScriptProps._mode = 'test';
  }

  public static endTest(): void {
    console.log('++++++++ TEST FINISH +++++++++++++');
    ScriptProps._mode = 'normal';
  }

  public static isTesting(): boolean {
    return this._mode === 'test';
  }
}
