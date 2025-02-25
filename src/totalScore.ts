export class TotalScore {
    public userId: string = '';
    public name: string = '';
    public playTime: number = 0;
    public sunnyPlay: number = 0;
    public rainyPlay: number = 0;
    public goalCount: number = 0;
    public assistCount: number = 0;
    public mipCount: number = 0;
    public teamPoint: number = 0;
    public winCount: number = 0;
    public loseCount: number = 0;
    public totalMatchs: number = 0;
    public totalOrank: number = 0;
    public totalGrank: number = 0;
    public totalArank: number = 0;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public fetchTeamPoint(eventDataRow: any[], teamValue: string): number {
        if (teamValue === 'チーム1') {
            return eventDataRow[7];
        }
        if (teamValue === 'チーム2') {
            return eventDataRow[8];
        }
        if (teamValue === 'チーム3') {
            return eventDataRow[9];
        }
        if (teamValue === 'チーム4') {
            return eventDataRow[10];
        }
        if (teamValue === 'チーム5') {
            return eventDataRow[11];
        }
        if (teamValue === 'チーム6') {
            return eventDataRow[12];
        }
        if (teamValue === 'チーム7') {
            return eventDataRow[13];
        }
        if (teamValue === 'チーム8') {
            return eventDataRow[14];
        }
        if (teamValue === 'チーム9') {
            return eventDataRow[15];
        }
        if (teamValue === 'チーム10') {
            return eventDataRow[16];
        }

        return 0;
    }
}
