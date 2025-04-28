export class StopWatch {
    private startTime: number;
    private endTime: number;

    constructor() {
        this.startTime = 0;
        this.endTime = 0;
    }

    start(): void {
        this.startTime = new Date().getTime();
    }

    stop(): void {
        this.endTime = new Date().getTime();
    }

    getElapsedTime(): number {
        return this.endTime - this.startTime; // ミリ秒単位での経過時間を返す
    }
}
