import { GraphFI } from "az-common/node_modules/@pnp/graph/index.js";
import { SPFI } from "az-common/node_modules/@pnp/sp/index.js";
import { IListInfo } from "az-common/node_modules/@pnp/sp/lists/index.js";
import { IFileInfo } from "az-common/node_modules/@pnp/sp/files/index.js";
import archiver from 'archiver';
import axios from "axios";
import { logger, pnpjs } from "az-common/dist/index.mjs";
import { IBaselineEvent, IBaseLineExportEvent } from "../model/IBaselineRequest";
import { IBaselineItem } from "../model/index.js";
import AppSettings from "../AppSettings.js";
import ExcelJS from 'exceljs';
import { createReadStream, statSync } from "fs";

export class FileHelper {
    private _graph: GraphFI;
    private _sp: SPFI;
    private _logSource: string;
    /**
     *
     */
    constructor(options?: {
        graph?: GraphFI,
        sp?: SPFI;
        logSource?: string;
    }) {
        this._graph = options?.graph || pnpjs.graph();
        this._sp = options?.sp;
        this._logSource = options?.logSource || 'NA';
    }

    public async getSiteIdsBatched(siteUrls: string[]): Promise<{ [siteUrl: string]: string }> {
        const [batch, execute] = this._graph.batched();
        const queries = siteUrls.map(s => ({
            hostname: new URL(s).hostname,
            siteServerRelativeUrl: new URL(s).pathname
        })).map(s =>
            batch.sites.getByUrl(s.hostname, s.siteServerRelativeUrl).then(site => site.select('id')().then(s => s.id))
        );
        await execute();
        return (await Promise.all(queries)).reduce((acc, curr, i) => {
            acc[siteUrls[i]] = curr;
            return acc;
        }, {} as { [siteUrl: string]: string });
    }
    public async getDriveIdsBatched(graphSitesMap: { [siteUrl: string]: string }, libSitePairs: { siteUrl: string, librarySiteRelativeUrl: string }[]): Promise<{ [libraryUrl: string]: string }> {
        const [batch, execute] = this._graph.batched();
        const queries = libSitePairs.map(l => {
            return batch.sites.getById(graphSitesMap[l.siteUrl])
                .drives.select('webUrl', 'id', 'name')()
                .then(drives => drives.find(d => d.webUrl == (`${l.siteUrl}${l.librarySiteRelativeUrl}`)?.replace(/\s/g, '%20'))?.id);
        });
        await execute();
        return (await Promise.all(queries)).reduce((acc, curr, i) => {
            acc[libSitePairs[i].siteUrl + libSitePairs[i].librarySiteRelativeUrl] = curr;
            return acc;
        }, {} as { [libraryUrl: string]: string });
    }

    public async getDownloadInfoBatched(files: { fileServerRelativeUrl: string; siteUrl: string }[]): Promise<{ [fileServerRelativeUrl: string]: { downloadUrl: string, size: number } }> {
        const siteIds = await this.getSiteIdsBatched(files.map(f => f.siteUrl).filter((v, i, a) => a.indexOf(v) === i));
        const driveIds = await this.getDriveIdsBatched(
            siteIds,
            files.map(f => {
                const { siteUrl, fileServerRelativeUrl } = f;
                const siteServerRelativeUrl = new URL(siteUrl).pathname;
                const libraryUrl = fileServerRelativeUrl.replace(siteServerRelativeUrl, '').split('/').slice(0, 2).join('/');
                return { siteUrl, librarySiteRelativeUrl: libraryUrl };
            })
        );
        const [batch, execute] = this._graph.batched();
        const queries = files.map(f => {
            const { siteUrl, fileServerRelativeUrl } = f;
            const siteServerRelativeUrl = new URL(siteUrl).pathname;
            const libraryUrl = fileServerRelativeUrl.replace(siteServerRelativeUrl, '').split('/').slice(0, 2).join('/');
            const libraryId = driveIds[siteUrl + libraryUrl];
            if (!libraryId) {
                throw new Error(`Library not found for url: ${f.siteUrl}${f.fileServerRelativeUrl}. Drive Ids: ${JSON.stringify(driveIds)}`);
            }
            const fileLibraryRelativeUrl = fileServerRelativeUrl.replace(siteServerRelativeUrl, '').replace(libraryUrl, '').substring(1);
            return batch.drives.getById(libraryId)
                .getItemByPath(fileLibraryRelativeUrl)
                .select("@microsoft.graph.downloadUrl")()
                .then(file => ({ [f.fileServerRelativeUrl]: { downloadUrl: file["@microsoft.graph.downloadUrl"], size: file.size } }));
        });
        await execute();
        return (await Promise.all(queries)).reduce((acc, curr) => ({ ...acc, ...curr }), {} as { [fileServerRelativeUrl: string]: { downloadUrl: string, size: number } });
    }

    public async getFileMapFromRequest(request: IBaseLineExportEvent, hostname: string): Promise<IBaselineItem[]> {
        const rootBaseline = { ...request.RootBaseLine, relatedBaseline: request.ExportDocuments };
        let map: { itemZipPath: string; itemServerRelativeUrl: string; siteUrl: string, id: string, type: 'docSet' | 'folder' | 'file' }[] = [
            {
                id: '' + rootBaseline.ItemId,
                itemZipPath: `/{blName:${rootBaseline.SitePath}:${rootBaseline.ItemId}}`,
                itemServerRelativeUrl: '',
                siteUrl: `${hostname}${rootBaseline.SitePath}`,
                type: 'folder'
            }
        ];
        this._buildFileMapRecursive(
            `/{blName:${rootBaseline.SitePath}:${rootBaseline.ItemId}}`,
            rootBaseline,
            map, hostname);
        await this._addMissingMetadata(map, hostname);
        //Sort rows so baseline comes first, then each docset followed by its files
        let res = map
            .filter(i => i.type == 'folder')
            .sort((a, b) => a.itemZipPath.localeCompare(b.itemZipPath))
            .map(f =>
                [f,
                    ...map.filter(i => i.type == 'docSet' && i.itemZipPath.startsWith(f.itemZipPath) && i.itemZipPath.replace(f.itemZipPath, '').split('/').length == 2)
                        .sort((a, b) => a.itemZipPath.localeCompare(b.itemZipPath))
                        .map(d =>
                            [d,
                                ...map
                                    .filter(i => i.type == 'file' && i.itemZipPath.startsWith(d.itemZipPath) && i.itemZipPath.replace(d.itemZipPath, '').split('/').length == 2)
                                    .sort((a, b) => a.itemZipPath.localeCompare(b.itemZipPath))
                            ]
                        )
                ]

            ).flat(2);
        return res;
    }

    private _buildFileMapRecursive(zipRoot: string, baseline: IBaselineEvent, map: IBaselineItem[], hostname: string): void {
        if (baseline.relatedBaseline?.technical) {
            baseline.relatedBaseline.technical.forEach(d => {
                map.push({
                    id: d.DocumentId,
                    itemZipPath: `${zipRoot}/${d.DocumentLink?.split('/').pop() || d.FileName}`,
                    itemServerRelativeUrl: d.DocumentLink,
                    siteUrl: `${hostname}${d.SitePath}`,
                    type: 'docSet'
                });
            });
        }
        if (baseline.relatedBaseline?.baseline) {
            baseline.relatedBaseline.baseline.forEach(b => {
                map.push({
                    id: '' + b.ItemId,
                    itemServerRelativeUrl: '',
                    siteUrl: `${hostname}${b.SitePath}`,
                    itemZipPath: `${zipRoot}/{blName:${b.SitePath}:${b.ItemId}}`,
                    type: 'folder'
                });
                this._buildFileMapRecursive(`${zipRoot}/{blName:${b.SitePath}:${b.ItemId}}`, b, map, hostname)
            });
        }
    }

    private async _addMissingMetadata(map: IBaselineItem[], hostname: string): Promise<void> {
        const allSites = map.map(m => m.siteUrl).filter((v, i, a) => a.indexOf(v) === i);
        const allResults: { spItem?: any, baselineItem: IBaselineItem, files?: (IFileInfo & { ListItemAllFields: { [field: string]: string } })[] }[] = [];
        for (const siteUrl of allSites) {
            const siteServerRelativeUrl = new URL(siteUrl).pathname;
            const sp = pnpjs.sp(siteServerRelativeUrl);
            const { baselineList } = await this._getLists(sp, siteServerRelativeUrl);
            const items = map.filter(m => m.siteUrl == siteUrl && m.type != 'file');
            //Process the items in batches of 100
            const batchSize = 100;
            const includeFieldsDocset = Object.values(AppSettings.BaselineExportFieldConfiguration)
                .map(f => f.file?.name).filter(n => n);
            const includeFieldsFile = Object.values(AppSettings.BaselineExportFieldConfiguration)
                .map(f => f.file?.name).filter(n => n).map(n => `ListItemAllFields/${n}`);
            for (let i = 0; i < items.length; i += batchSize) {
                const batchItems = items.slice(i, i + batchSize);
                const [batch, execute] = sp.batched();
                const queries = batchItems.map(m => {
                    const currentItem = { ...m };
                    switch (currentItem.type) {
                        case 'docSet':
                            return Promise.all(
                                [
                                    batch.web.getFolderByServerRelativePath(currentItem.itemServerRelativeUrl).listItemAllFields
                                        .select(...includeFieldsDocset)(),
                                    batch.web.getFolderByServerRelativePath(currentItem.itemServerRelativeUrl).files
                                        .expand('ListItemAllFields').select(...['ServerRelativeUrl', ...includeFieldsFile])()
                                ]
                            )
                                .then(([item, files]) => ({ spItem: item, baselineItem: currentItem, files }))
                                .catch(e => ({ baselineItem: { ...currentItem, error: JSON.stringify(e) + '\n' } }));
                        case 'folder':
                            return batch.web.lists.getById(baselineList.Id).items.getById(parseInt(currentItem.id))()
                                .then(item => ({ spItem: item, baselineItem: currentItem }))
                                .catch(e => ({ baselineItem: { ...currentItem, error: JSON.stringify(e) + '\n' } }));
                        default:
                            throw new Error(`Unknown type: ${currentItem.type}`);
                    }
                });
                await execute();
                allResults.push(...await Promise.all(queries));
            }
        }
        const folderNameMap: { [urlAndId: string]: string } = {};

        //TODO: Fix zip path everywhere (missing leafref)
        //Add the files to the map
        allResults.forEach(r => {
            if (r.files) {
                r.files.forEach(f => {
                    map.push({
                        id: f.ListItemAllFields.ID,
                        itemZipPath: r.baselineItem.itemZipPath + '/' + f.ServerRelativeUrl.split('/').pop(),
                        itemServerRelativeUrl: f.ServerRelativeUrl,
                        siteUrl: r.baselineItem.siteUrl,
                        type: 'file',
                        metadata: Object.keys(f.ListItemAllFields)
                            .reduce((acc, curr) => {
                                acc[curr] = f.ListItemAllFields[curr] || r.spItem[curr];
                                return acc;
                            }, {} as { [field: string]: string })
                    });
                });
            }
        });
        //Go through the baseline items and add match with sharepoint items
        //Replace the path params with the name of the item in sharepoint
        map.forEach(m => {
            //No need to get file metadata
            if (m.type !== 'file') {
                const result = allResults.find(r => r.baselineItem.id == m.id && r.baselineItem.type == m.type && r.baselineItem.siteUrl == m.siteUrl);
                if (result) {
                    m.metadata = result.spItem;
                    m.error = result.baselineItem.error;
                } else {
                    m.error = `Could not find a matching item in SP\n`;
                }
            }
            let pathParams = /{blName\:([^\:]+)\:([^\}]+)}/.exec(m.itemZipPath);
            while (pathParams) {
                if (pathParams.length != 3) {
                    m.error += (m.error || '') + `Wrong path params\n`;
                } else {
                    const toReplace = pathParams[0];
                    const sitePath = pathParams[1];
                    const itemId = pathParams[2];
                    const bl = allResults.find(r => {
                        return '' + r.baselineItem.id == '' + itemId && r.baselineItem.siteUrl == `${hostname}${sitePath}`;
                    }
                    );
                    if (bl) {
                        let folderName = folderNameMap[`${bl.baselineItem.siteUrl}${bl.baselineItem.id}`];
                        if (!folderName) {
                            folderName = AppSettings.FolderNameFields.map(f => bl.spItem[f]).join('-');
                            let counter = 0;
                            while (Object.values(folderNameMap).includes(folderName)) {
                                folderName = folderName.replace(/(\s\(\d+\))?$/, ` (${++counter})`);
                            }
                            folderNameMap[`${bl.baselineItem.siteUrl}${bl.baselineItem.id}`] = folderName;
                        }
                        m.itemZipPath = m.itemZipPath.replace(toReplace, `${folderName}`);
                    } else {
                        m.itemZipPath = m.itemZipPath.replace(toReplace, `unknown`);
                        m.error = (m.error || '') + `BL reference in path not found\n`;
                    }
                }
                pathParams = /{blName\:([^\:]+)\:([^\}]+)}/.exec(m.itemZipPath);
            }
        });
    }
    private async _getLists(sp: SPFI, siteServerRelativeUrl: string): Promise<{ baselineList: IListInfo, draftLibrary: IListInfo, applicableLibrary: IListInfo, oldRevLibrary: IListInfo }> {
        const lists = await sp.web.lists.select('Id', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
        const baselineList = lists.find(l => l.RootFolder.ServerRelativeUrl == `${siteServerRelativeUrl}/Lists/${AppSettings.BaselineListName}`);
        if (!baselineList) {
            throw new Error(`Baseline list not found in site: ${siteServerRelativeUrl}`);
        }
        const draftLibrary = lists.find(l => l.RootFolder.ServerRelativeUrl == `${siteServerRelativeUrl}/Draft`);
        if (!draftLibrary) {
            throw new Error(`Draft library not found in site: ${siteServerRelativeUrl}`);
        }
        const applicableLibrary = lists.find(l => l.RootFolder.ServerRelativeUrl == `${siteServerRelativeUrl}/ApplicableDocuments`);
        if (!applicableLibrary) {
            throw new Error(`Applicable library not found in site: ${siteServerRelativeUrl}`);
        }
        const oldRevLibrary = lists.find(l => l.RootFolder.ServerRelativeUrl == `${siteServerRelativeUrl}/PreviousVersions`);
        if (!oldRevLibrary) {
            throw new Error(`OldRev library not found in site: ${siteServerRelativeUrl}`);
        }
        return { baselineList, draftLibrary, applicableLibrary, oldRevLibrary };
    }

    public async getArchiveFromFileMap(map: IBaselineItem[]): Promise<archiver.Archiver> {
        const archive = archiver('zip', { zlib: { level: 1 } });
        const downloadInfos = await this.getDownloadInfoBatched(
            map.filter(m => m.type === 'file').map(m => ({ siteUrl: m.siteUrl, fileServerRelativeUrl: m.itemServerRelativeUrl }))
        );
        const filesWithDownloadInfos = map.filter(m => m.type == 'file').map(m =>
            ({ ...m, ...downloadInfos[m.itemServerRelativeUrl] })
        );
        map.filter(m => m.type !== 'file').forEach(f => {
            logger.trackTrace({
                message: `Appending to zip ${JSON.stringify(f)}`,
                properties: { source: this._logSource, request_type: 'NA' },
                severity: 'Verbose'
            });
            archive.append(null, { name: f.itemZipPath + '/' });
        });
        for (const file of filesWithDownloadInfos) {
            const downStream = await axios.get(file.downloadUrl, { responseType: 'stream' });
            logger.trackTrace({
                message: `Appending to zip ${JSON.stringify(file)}`,
                properties: { source: this._logSource, request_type: 'NA' },
                severity: 'Verbose'
            });
            archive.append(downStream.data, { name: file.itemZipPath });
        }
        return archive;
    }

    public async uploadZipToSharepoint(zipPath: string, fileName: string,): Promise<string> {
        const webInfo = await this._sp.web.select("Url", "ServerRelativeUrl")();
        const sharedDocumentsRelativeUrl = `${webInfo.ServerRelativeUrl}/Shared Documents`;
        const chunkFactor = 10;
        const chunkSize = 327680 * chunkFactor;

        const readStream = createReadStream(zipPath, { highWaterMark: chunkSize });
        var stats = statSync(zipPath);
        const fileUploadOptions = {
            item: {
                name: fileName,
                fileSize: stats.size
            },
        };
        const site = await this._graph.sites.getByUrl(new URL(webInfo.Url).hostname, webInfo.ServerRelativeUrl);
        const driveRoot = await site.drive.root();
        // Create the upload session, must get the drive root folder id to call createUploadSession
        const uploadSession = await site.drive.getItemById(driveRoot.id).createUploadSession(fileUploadOptions);

        let totalSent = 0;
        for await (const chunk of readStream) {
            const range = `bytes ${totalSent}-${totalSent + chunk.length - 1}/${stats.size}`;
            await uploadSession.resumableUpload.upload(chunk.length, chunk, range);
            totalSent += chunk.length;
            logger.trackTrace({
                message: `Sent ${totalSent / 1024 / 1024}/${stats.size / 1024 / 1024} MB. ${chunk.length / 1024 / 1024} MB sent.`,
                properties: { source: this._logSource, request_type: 'NA' },
                severity: 'Verbose'
            });
        }
        const file = await this._sp.web.getFileByServerRelativePath(`${sharedDocumentsRelativeUrl}/${fileName}`).select('ServerRelativeUrl')();
        return file.ServerRelativeUrl;
    }

    public async getExcelFileFromMap(fMap: IBaselineItem[], baselineId: string, baselineName: string): Promise<Buffer> {
        //Workaround: Remove file names from the path
        let fileMap = [...fMap].map(r => {
            if (r.type == 'file') {
                r.itemZipPath = r.itemZipPath.replace('/' + r.metadata['FileLeafRef'], '');
            }
            return r;
        });
        //Remove root baseline from path
        fileMap = fileMap.map(r => {
            r.itemZipPath = '/' + r.itemZipPath.split('/').slice(2).join('/');
            return r;
        });
        const table = fileMap.map(item => {
            return Object.keys(AppSettings.BaselineExportFieldConfiguration).reduce((acc, key) => {
                const field = AppSettings.BaselineExportFieldConfiguration[key][item.type];
                let valueSource: { [key: string]: unknown; } | IBaselineItem = item;
                if (field?.source === 'metadata') {
                    valueSource = item.metadata;
                }
                if (field) {
                    acc[key] = valueSource[field.name];
                }
                return acc;
            }, {} as { [prop: keyof typeof AppSettings.BaselineExportFieldConfiguration]: string });
        });
        logger.trackTrace({
            message: `Created Table. ${JSON.stringify(table)}`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Verbose'
        });
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet(baselineName);
        sheet.columns = Object.keys(AppSettings.BaselineExportFieldConfiguration).map(key => ({
            header: key,
            key: key
        }));
        table.forEach(obj => {
            sheet.addRow(obj);
        });
        return await workbook.xlsx.writeBuffer() as unknown as Buffer;
    }
}