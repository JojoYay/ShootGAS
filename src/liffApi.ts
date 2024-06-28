import { DensukeUtil } from './densukeUtil';
import { GasProps } from './gasProps';
import { GasUtil } from './gasUtil';
import { GetEventHandler } from './getEventHandler';
import { LineUtil } from './lineUtil';
import { ScriptProps } from './scriptProps';

export class LiffApi {
    private test(getEventHandler: GetEventHandler): void {
        const value: string = getEventHandler.e.parameters['param'][0];
        getEventHandler.result = { result: value };
    }

    private getVideo(getEventHandler: GetEventHandler): void {
        const videos: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.videoSheet;
        getEventHandler.result = { result: videos.getDataRange().getValues() };
    }

    private getMembers(getEventHandler: GetEventHandler): void {
        const den: DensukeUtil = new DensukeUtil();
        const members = den.extractMembers();
        // getEventHandler.result = { result: members };
        getEventHandler.result.members = members;
    }

    private getDensukeName(getEventHandler: GetEventHandler): void {
        const gasUtil: GasUtil = new GasUtil();
        const lineUtil: LineUtil = new LineUtil();
        const userId = getEventHandler.e.parameters['userId'][0];
        // getEventHandler.result = { result: gasUtil.getDensukeName(lineUtil.getLineDisplayName(userId)) };
        getEventHandler.result.densukeName = gasUtil.getDensukeName(lineUtil.getLineDisplayName(userId));
    }

    private getRanking(getEventHandler: GetEventHandler): void {
        const gRank = GasProps.instance.gRankingSheet.getDataRange().getValues();
        const aRank = GasProps.instance.aRankingSheet.getDataRange().getValues();
        const oRank = GasProps.instance.oRankingSheet.getDataRange().getValues();
        getEventHandler.result.gRank = gRank;
        getEventHandler.result.aRank = aRank;
        getEventHandler.result.oRank = oRank;
    }

    private getStats(getEventHandler: GetEventHandler): void {
        const resultSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.personalTotalSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = resultSheet.getDataRange().getValues();
        getEventHandler.result.stats = resultValues;
    }

    private getUsers(getEventHandler: GetEventHandler): void {
        const mappingSheet: GoogleAppsScript.Spreadsheet.Sheet = GasProps.instance.mappingSheet;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resultValues: any[][] = mappingSheet.getDataRange().getValues();
        getEventHandler.result.users = resultValues;
    }

    private register(getEventHandler: GetEventHandler): void {
        const userId = getEventHandler.e.parameters['userId'][0];
        const lineUtil: LineUtil = new LineUtil();
        const gasUtil: GasUtil = new GasUtil();
        const densukeUtil: DensukeUtil = new DensukeUtil();
        const lineName = lineUtil.getLineDisplayName(userId);
        const lang = lineUtil.getLineLang(userId);
        const $ = densukeUtil.getDensukeCheerio();
        densukeUtil.generateSummaryBase($); //先に更新しておかないとエラーになる（伝助が更新されている場合）
        const members = densukeUtil.extractMembers($);
        const actDate = densukeUtil.extractDateFromRownum($, ScriptProps.instance.ROWNUM);
        const densukeNameNew = getEventHandler.e.parameters['densukeName'][0];
        if (members.includes(densukeNameNew)) {
            if (densukeUtil.hasMultipleOccurrences(members, densukeNameNew)) {
                if (lang === 'ja') {
                    getEventHandler.result = {
                        result: '伝助上で"' + densukeNameNew + '"という名前が複数存在しています。重複のない名前に更新して再度登録して下さい。',
                    };
                } else {
                    getEventHandler.result = {
                        result:
                            "There are multiple entries with the name '" +
                            densukeNameNew +
                            "' on Densuke. Please update it to a unique name and register again.",
                    };
                }
            } else {
                gasUtil.registerMapping(lineName, densukeNameNew, userId);
                gasUtil.updateLineNameOfLatestReport(lineName, densukeNameNew, actDate);
                if (lang === 'ja') {
                    getEventHandler.result = {
                        result:
                            '伝助名称登録が完了しました。\n伝助上の名前：' +
                            densukeNameNew +
                            '\n伝助のスケジュールを登録の上、ご参加ください。\n参加費の支払いは、参加後にPayNowでこちらにスクリーンショットを添付してください。',
                    };
                } else {
                    getEventHandler.result = {
                        result:
                            'The initial registration is complete.\nYour name in Densuke: ' +
                            densukeNameNew +
                            "\nPlease register Densuke's schedule and attend.\nAfter attending, please make the payment via PayNow and attach a screenshot here.",
                    };
                }
            }
        } else {
            if (lang === 'ja') {
                getEventHandler.result = {
                    result: '【エラー】伝助上に指定した名前が見つかりません。再度登録を完了させてください\n伝助上の名前：' + densukeNameNew,
                };
            } else {
                getEventHandler.result = {
                    result:
                        '【Error】The specified name was not found in Densuke. Please complete the registration again.\nYour name in Densuke: ' +
                        densukeNameNew,
                };
            }
        }
    }
}
