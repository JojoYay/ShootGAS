export class TotalScore {
  public name: string = '';
  public playTime: number = 0;
  public sunnyPlay: number = 0;
  public rainyPlay: number = 0;
  public gRankingTitle: string = '';
  public goalCount: number = 0;
  public aRankingTitle: string = '';
  public assistCount: number = 0;
  public mipCount: number = 0;
  public teamPoint: number = 0;
  public winCount: number = 0;
  public loseCount: number = 0;
  // public drawCpimt: number = 0;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public isTopTeam(eventDataRow: any[], teamValue: string): boolean {
    const teamCount: number = this.getTeamCount(eventDataRow);
    if (teamCount === 2) {
      return this.fetchTeamPoint(eventDataRow, teamValue) === 1;
    } else if (teamCount === 3) {
      return this.fetchTeamPoint(eventDataRow, teamValue) === 2;
    } else if (teamCount === 4) {
      return this.fetchTeamPoint(eventDataRow, teamValue) === 3;
    } else if (teamCount === 5) {
      return this.fetchTeamPoint(eventDataRow, teamValue) === 4;
    } else if (teamCount === 6) {
      return this.fetchTeamPoint(eventDataRow, teamValue) === 5;
    }
    return false;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public getTeamCount(eventDataRow: any[]): number {
    if (!eventDataRow[10]) {
      return 3;
    } else if (!eventDataRow[11]) {
      return 4;
    } else if (!eventDataRow[12]) {
      return 5;
    } else if (!eventDataRow[13]) {
      return 6;
    }
    return -1;
  }
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
    return 0;

    // let teamPointIndex = 0;
    // switch (teamValue) {
    //   case 'チーム1':
    //     teamPointIndex = 8;
    //     break;
    //   case 'チーム2':
    //     teamPointIndex = 9;
    //     break;
    //   case 'チーム3':
    //     teamPointIndex = 10;
    //     break;
    //   case 'チーム4':
    //     teamPointIndex = 11;
    //     break;
    //   case 'チーム5':
    //     teamPointIndex = 12;
    //     break;
    //   case 'チーム6':
    //     teamPointIndex = 13;
    //     break;
    //   default:
    //     teamPointIndex = 0;
    //     break;
    // }
    // return eventDataRow[teamPointIndex] || 0;
  }
}
