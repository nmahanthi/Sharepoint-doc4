import { FileHelper } from './FileHelper';
import { beforeEach, describe, expect, it } from '@jest/globals';
import { logger, pnpjs } from "az-common/dist/index.mjs";
import localSettings from '../../local.settings.json';
import { LogLevel } from 'az-common/node_modules/@pnp/logging';
import fs from 'fs';
import AppSettings from '../AppSettings.js';
describe('FileHelper', () => {
    let fileHelper: FileHelper;
    let mockRequest: typeof import('../../test/mock-request.json');

    beforeEach(async () => {
        mockRequest = (await import('../../test/mock-request.json')).default;
        for (const key in localSettings.Values) {
            process.env[key] = localSettings.Values[key];
        }
        logger.initWithConsole(LogLevel.Verbose, console);
        await pnpjs.Init();
        const sp = pnpjs.sp(mockRequest.SiteUrl);
        fileHelper = new FileHelper({ sp });
    });


    it('should build a filemap from a request', async () => {
        const hostUrl = 'https://alstomgrouppp.sharepoint.com';
        expect(mockRequest).toBeDefined();
        expect(mockRequest.RootBaseLine).toBeDefined();
        const fileMap = await fileHelper.getFileMapFromRequest(mockRequest, hostUrl);
        const testSuffix = `.${new Date().getTime()}.ignore`;
        try {
            fs.writeFileSync(`D:\\dms\\dms-az-baseline-export\\test\\file-map${testSuffix}.json`, JSON.stringify(fileMap), {
                encoding: 'utf8'
            });
        } catch (err) {
            console.error(err);
        }
    });

    it('should get the site id', async () => {
        const siteUrl = 'https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT';
        const siteIds = await fileHelper.getSiteIdsBatched([siteUrl]);
        expect(siteIds).toBeDefined();
        expect(siteIds[siteUrl]).toBeDefined();
    });
    it('should get the drive id', async () => {
        const siteUrl = 'https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT';
        const libraryUrl = '/Shared%20Documents';
        const siteMap = await fileHelper.getSiteIdsBatched([siteUrl]);
        const driveIds = await fileHelper.getDriveIdsBatched(siteMap, [{ siteUrl, librarySiteRelativeUrl: libraryUrl }]);
        expect(driveIds).toBeDefined();
        console.log(JSON.stringify(driveIds));
        expect(driveIds[siteUrl + libraryUrl]).toBeDefined();
    });

    it('should get the file download infos', async () => {
        const siteUrl = 'https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT';
        const downloadInfos = await fileHelper.getDownloadInfoBatched([
            { siteUrl, fileServerRelativeUrl: '/sites/DOC4A-INT/Shared Documents/test.txt' }
        ]);
        expect(downloadInfos).toBeDefined();
        expect(downloadInfos['/sites/DOC4A-INT/Shared Documents/test.txt'].downloadUrl).toBeDefined();
    });

    it('should build a valid zip from the file map', async () => {
        const hostUrl = 'https://alstomgrouppp.sharepoint.com';
        const testSuffix = `.${new Date().getTime()}.ignore`;
        const zipFile = `D:\\dms\\dms-az-baseline-export\\test\\bl-export${testSuffix}.zip`;
        const output = fs.createWriteStream(zipFile);
        const fileMap = await fileHelper.getFileMapFromRequest(mockRequest, hostUrl);
        try {
            fs.writeFileSync(`D:\\dms\\dms-az-baseline-export\\test\\file-map${testSuffix}.json`, JSON.stringify(fileMap), {
                encoding: 'utf8'
            });
        } catch (err) {
            console.error(err);
        }
        const archive = await fileHelper.getArchiveFromFileMap(fileMap);
        archive.on("error", function (err) {
            throw err;
        });
        // pipe archive data to the file
        archive.pipe(output);
        output.on("close", function () {
            console.log(archive.pointer() + " total bytes");
            console.log(
                "archiver has been finalized and the output file descriptor has closed.",
            );
        });
        output.on("end", function () {
            console.log("Data has been drained");
        });
        archive.finalize();
    }, 1000 * 60 * 5);

    it('should stream the zip to sharepoint', async () => {
        const hostUrl = 'https://alstomgrouppp.sharepoint.com';
        const fileMap = await fileHelper.getFileMapFromRequest(mockRequest, hostUrl);
        const archive = await fileHelper.getArchiveFromFileMap(fileMap);
        const testSuffix = `.${new Date().getTime()}.ignore`;
        const zipFile = `D:\\dms\\dms-az-baseline-export\\test\\bl-export${testSuffix}.zip`;
        const output = fs.createWriteStream(zipFile, { highWaterMark: 1024 * 1024 * 100 });
        archive.pipe(output);
        archive.finalize();
        const wiatForStream = new Promise<void>((resolve, reject) => {
            output.on('drain', () => {
                console.log('drain');
            });
            output.on('finish', () => {
                resolve();
            });
            output.on('error', (err) => {
                reject(err);
            });
        });
        await wiatForStream;
        const uploadedFile = await fileHelper.uploadZipToSharepoint(zipFile, 'test.zip');
        console.log(uploadedFile);
        expect(uploadedFile).toBeDefined();
    }, 1000 * 60 * 120);

    it('should generate a valid excel file', async () => {
        const hostUrl = 'https://alstomgrouppp.sharepoint.com';
        const fileMap = await fileHelper.getFileMapFromRequest(mockRequest, hostUrl);
        const baselineName = AppSettings.FolderNameFields.map(f => fileMap[0].metadata[f]).join('-');
        const buffer = await fileHelper.getExcelFileFromMap(fileMap, 'baselineId', baselineName);
        expect(buffer).toBeDefined();
        const testSuffix = `.${new Date().getTime()}.ignore`;
        try {
            fs.writeFileSync(`D:\\dms\\dms-az-baseline-export\\test\\file-map${testSuffix}.json`, JSON.stringify(fileMap));
            fs.writeFileSync(`D:\\dms\\dms-az-baseline-export\\test\\excel${testSuffix}.xlsx`, new Uint8Array(buffer));
        } catch (err) {
            console.error(err);
        }
    });
});