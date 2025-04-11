import { InvocationContext } from "@azure/functions";
import { FunctionListener, Logger as PnPLogger, LogLevel } from "@pnp/logging";

interface ITrackTraceArgs {
    message: string, properties: { source?: string; method?: string; request_type?:string }, severity: string
}
interface ITrackExceptionArgs {
    exception: Error, properties: { source?: string; method?: string }
}

interface ITrackEventArgs {
    name: string, properties: { source?: string; requestBody?: string; requestQuery?: string }
}

export interface ILogger {
    init: (context: InvocationContext) => void;
    trackTrace: (args:ITrackTraceArgs) => void;
    trackException: (args:ITrackExceptionArgs) => void;
    trackEvent: (args:ITrackEventArgs) => void;
    initWithConsole: (loglevel:LogLevel, console:Console) => void;

}

export class Logger implements ILogger {
    private _fnContext: InvocationContext;
    private _console: Console;
    public constructor() { }

    public init(context: InvocationContext): void {
        this._fnContext = context;
        PnPLogger.activeLogLevel = LogLevel.Info;
        PnPLogger.subscribe(FunctionListener((entry)=>{
            this.trackTrace({
                message: `${entry.message}. Data: '${JSON.stringify(entry.data)?.replace(/[\r\n]/g, '')?.replace(/\t/g, ' ')}'`,
                properties: { source: "PnP"},
                severity: LogLevel[entry.level]
            });
        }));
        
    }

    public initWithConsole(loglevel:LogLevel, console:Console): void {
        PnPLogger.activeLogLevel = loglevel;
        PnPLogger.subscribe(FunctionListener((entry)=>{
            this.trackTrace({
                message: `${entry.message}. Data: '${JSON.stringify(entry.data)?.replace(/[\r\n]/g, '')?.replace(/\t/g, ' ')}'`,
                properties: { source: "PnP"},
                severity: LogLevel[entry.level]
            });
        }));
        this._console = console;
    }

    public trackTrace(args:ITrackTraceArgs): void {
        if(!this._console && !this._fnContext){
            throw new Error("Logger not initialized");
        }
        const { severity, message, properties } = args;
        if(this._console){
            this._console.log(`${severity?.toUpperCase()}\t${(properties.request_type || '')}${message}\t${properties?.source}\t${properties?.method}`);
            return;
        }
        this._fnContext.log(`${severity?.toUpperCase()}\t${(properties.request_type || '')}${message}\t${properties?.source}\t${properties?.method}`);
    }
    
    public trackException(args:ITrackExceptionArgs): void {
        if(!this._console && !this._fnContext){
            throw new Error("Logger not initialized");
        }
        const { exception, properties } = args;
        let errorMessage = `${exception}`;
        if(exception?.message){
            errorMessage = `${exception.message}`;
        }
        if(exception?.stack){
            errorMessage += ` Stack: ${exception.stack.replace(/[\r\n]/g, '')}`;
        }
        if(this._console){
            this._console.log(`ERROR\t${errorMessage}\t${properties?.source}\t${properties?.method}`);
            return;
        }
        this._fnContext.log(`ERROR\t${errorMessage}\t${properties?.source}\t${properties?.method}`);
    }
    public trackEvent(args:ITrackEventArgs): void {
        if(!this._console && !this._fnContext){
            throw new Error("Logger not initialized");
        }
        const { name, properties } = args;
        if(this._console){
            this._console.log(`INFO\t${name}\t${properties?.source}\t${properties?.requestBody}\t${properties?.requestQuery}`);
            return;
        }
        this._fnContext.log(`INFO\t${name}\t${properties?.source}\t${properties?.requestBody}\t${properties?.requestQuery}`);
    }
}

export const logger: ILogger = new Logger();