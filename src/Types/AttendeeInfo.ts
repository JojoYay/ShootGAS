// 参加者情報の型定義
export interface AttendeeInfo {
    userId: string;
    userName: {
        densukeName: string;
        lineName: string;
    };
    adult: number;
    child: number;
}
