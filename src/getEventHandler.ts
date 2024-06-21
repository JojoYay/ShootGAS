import { LiffApi } from './liffApi';

export class GetEventHandler {
    // public userId: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public e: GoogleAppsScript.Events.DoGet;
    public func: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public result: any = null;

    public constructor(e: GoogleAppsScript.Events.DoGet) {
        this.e = e;
        this.func = e.parameters['func'][0];
        // this.userId = e.parameters['userId'][0];
        // if (!this.userId || !gasUtil.isRegistered(this.userId)) {
        // if (!this.userId) {
        //     throw new Error('userId is not found:' + this.userId);
        // }
        const liffApi = new LiffApi();
        if (!(this.func in liffApi) || typeof liffApi[this.func as keyof LiffApi] !== 'function') {
            throw new Error('Func is not registered:' + this.func);
        }
    }
}
