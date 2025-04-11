import { InvocationContext } from "@azure/functions";
 
interface ITrackTraceArgs {
    message: string, properties: {
        [key: string]: string | undefined;
         source?: string; method?: string; details?: string ;
        siteUrl?: string; listId?: string; itemId?: string; columnName?: string;
    }, severity: string
}
interface ITrackExceptionArgs {
    exception: Error, properties: { 
        [key: string]: string | undefined;
        source?: string; method?: string; details?: string ;
        siteUrl?: string; listId?: string; itemId?: string; columnName?: string;
    }
}
 
interface ITrackEventArgs {
    name: string, properties: { source?: string; requestBody?: string; requestQuery?: string }
}
 
export interface ILogger {
    init: (context: InvocationContext) => void;
    trackTrace: (args: ITrackTraceArgs) => void;
    trackException: (args: ITrackExceptionArgs) => void;
    trackEvent: (args: ITrackEventArgs) => void;
}
 
export class Logger implements ILogger {
    private _fnContext: InvocationContext;
    public constructor() { }
 
    public init(context: InvocationContext): void {
        this._fnContext = context;
    }
 
    public trackTrace(args: ITrackTraceArgs): void {
        if (!this._fnContext) {
            throw new Error("Logger not initialized");
        }
        const { severity, message, properties } = args;
        this._fnContext.log(`${severity?.toUpperCase()}\t${message}\t${properties?.source}\t${properties?.method}\t${properties?.details || ''}`);
    }
    public trackException(args: ITrackExceptionArgs): void {
        if (!this._fnContext) {
            throw new Error("Logger not initialized");
        }
        const { exception, properties } = args;
        let errorMessage = `${exception}`;
        try {
            errorMessage = JSON.stringify(exception)?.replace(/[\r\n]/g, '')?.replace(/\t/g, ' ');
        } catch (_) { }
        this._fnContext.log(`ERROR\t${errorMessage}\t${properties?.source}\t${properties?.method}\t${properties?.details || ''}`);
    }
    public trackEvent(args: ITrackEventArgs): void {
        if (!this._fnContext) {
            throw new Error("Logger not initialized");
        }
        const { name, properties } = args;
        this._fnContext.log(`INFO\t${name}\t${properties?.source}\t${properties?.requestBody}\t${properties?.requestQuery}`);
    }
}
 
export const logger: ILogger = new Logger();