import { CommonSettings, logger } from 'az-common/dist/index.mjs';
import os from "os";

const AppSettings = {
    get BaselineExportFieldConfiguration() {
        if (!process.env['BASELINE_EXPORT_FIELD_CONFIGURATION']) {
            throw new Error('BASELINE_EXPORT_FIELD_CONFIGURATION is not set');
        }
        try {
            return JSON.parse(process.env['BASELINE_EXPORT_FIELD_CONFIGURATION'] || '{}') as
                {
                    [excelColumn: string]: { [key in 'docSet' | 'folder' | 'file']: {
                        source: "metadata" | "object";
                        type: string;
                        name: string;
                    } }
                };
        } catch (e) {
            logger.trackException({
                exception: new Error(`BASELINE_EXPORT_FIELD_CONFIGURATION is not a valid JSON: ${process.env['BASELINE_EXPORT_FIELD_CONFIGURATION']}. Details: ${e}`),
                properties: { source: 'NA', method: 'AppSettings' }
            });
            return {};
        }
    },
    get BaselineListName() {
        if (!process.env['BASELINE_LIST_NAME']) {
            throw new Error('BASELINE_LIST_NAME is not set');
        }
        return process.env['BASELINE_LIST_NAME'];
    },
    get FolderNameFields() {
        if (!process.env['FOLDER_NAME_FIELDS']) {
            throw new Error('FOLDER_NAME_FIELDS is not set');
        }
        return process.env['FOLDER_NAME_FIELDS'].split(',').map(f => f.trim());
    },
    get TempZipFolder() {
        if (!process.env['TEMP_ZIP_FOLDER']) {
            logger.trackTrace({
                message: 'TEMP_ZIP_FOLDER is not set. Using default',
                properties: { source: 'NA', method: 'AppSettings' }, severity: 'Information'
            });
            return os.tmpdir();
        }
        return process.env['TEMP_ZIP_FOLDER'];
    }
};
for (const key in CommonSettings) {
    Object.defineProperty(AppSettings, key, { get: () => CommonSettings[key] });
}

export default AppSettings as (typeof AppSettings) & (typeof CommonSettings);