import { LiffApi } from './liffApi';

export class GetEventHandler {
    // public userId: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public e: GoogleAppsScript.Events.DoGet;
    public funcs: string[];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public result: any = {};

    public constructor(e: GoogleAppsScript.Events.DoGet) {
        this.e = e;
        this.funcs = e.parameters['func'];
        // console.log('funcs: ' + this.funcs);
        // this.userId = e.parameters['userId'][0];
        // if (!this.userId || !gasUtil.isRegistered(this.userId)) {
        // if (!this.userId) {
        //     throw new Error('userId is not found:' + this.userId);
        // }
        const liffApi = new LiffApi();
        for (const func of this.funcs) {
            if (!(func in liffApi) || typeof liffApi[func as keyof LiffApi] !== 'function') {
                throw new Error('Func is not registered:' + func);
            }
        }
    }
}
