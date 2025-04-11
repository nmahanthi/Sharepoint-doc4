import { IImportJob } from "../../../interfaces/mass-import";
import { AadHttpClient } from "@microsoft/sp-http";

export class MassImportService {
    private static _instance: MassImportService;
    private _jobIdQueue: { id: string; resolve: (value: IImportJob) => void; reject: (reason: unknown) => void }[] = [];
    private _delayedGetIds?: NodeJS.Timeout;

    private constructor(private _dmsClient: AadHttpClient, private _massImportEndpointUrl: string, private _siteUrl: string) { }
    public static get instance(): MassImportService {
        if (!MassImportService._instance) {
            throw new Error("MassImportService is not initialized");
        }
        return MassImportService._instance;
    }

    public static initialize(
        dmsClient: AadHttpClient,
        massImportEndpointUrl: string,
        siteUrl: string
    ): void {
        if (!this._instance) {
            this._instance = new MassImportService(dmsClient, massImportEndpointUrl, siteUrl);
        }
    }

    public async getMassImportJob(id: string): Promise<IImportJob> {
        if (this._delayedGetIds) {
            clearTimeout(this._delayedGetIds);
            this._delayedGetIds = undefined;
        }
        let pResolve: ((value: IImportJob | PromiseLike<IImportJob>) => void) | undefined = undefined;
        let pReject: ((reason?: unknown) => void) | undefined = undefined;
        const result = new Promise<IImportJob>((resolve, reject) => {
            pResolve = resolve;
            pReject = reject;
        });
        if (!pResolve || !pReject) {
            throw new Error("Promise not initialized");
        }
        this._jobIdQueue.push({ id, resolve: pResolve, reject: pReject });
        this._delayedGetIds = setTimeout(() => {
            this._dmsClient.get(
                `${this._massImportEndpointUrl}&siteUrl=${encodeURIComponent(this._siteUrl)}&id=${encodeURIComponent(this._jobIdQueue.map(i => i.id).join(","))}`,
                AadHttpClient.configurations.v1)
                .then(result => {
                    result.json().then(data => {
                        this._jobIdQueue.forEach(job => {
                            const jobData = data.find((d: IImportJob) => d.id === job.id);
                            if (jobData) {
                                job.resolve(jobData);
                            } else {
                                job.reject(new Error(`Job with id ${job.id} not found`));
                            }
                            this._jobIdQueue = this._jobIdQueue.filter(j => j.id !== job.id);
                        });
                    }).catch(err => {
                        this._jobIdQueue.forEach(job => {
                            job.reject(err);
                            this._jobIdQueue = this._jobIdQueue.filter(j => j.id !== job.id);
                        });
                    });
                })
                .catch(err => {
                    this._jobIdQueue.forEach(job => {
                        job.reject(err);
                        this._jobIdQueue = this._jobIdQueue.filter(j => j.id !== job.id);
                    });
                });
        }, 100);
        return result;
    }
}