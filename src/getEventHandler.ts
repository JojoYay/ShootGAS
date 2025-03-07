import { LiffApi } from './liffApi';

export class GetEventHandler {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public e: GoogleAppsScript.Events.DoGet;
    public funcs: string[];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public result: any = {};

    public constructor(e: GoogleAppsScript.Events.DoGet) {
        this.e = e;
        this.funcs = e.parameters['func'];
        const liffApi = new LiffApi();
        for (const func of this.funcs) {
            if (!(func in liffApi) || typeof liffApi[func as keyof LiffApi] !== 'function') {
                throw new Error('Func is not registered:' + func);
            }
        }
    }
}
