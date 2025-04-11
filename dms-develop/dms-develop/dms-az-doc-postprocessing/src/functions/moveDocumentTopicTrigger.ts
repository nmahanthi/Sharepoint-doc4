import { app, InvocationContext } from "@azure/functions";
import { SPFI } from "@pnp/sp";
import { IDmsEvent, logger, MassImportService, pnpjs, SPUtilityService } from "az-common/dist/index.mjs";
import "@pnp/sp/batching.js";
import "@pnp/sp/presets/all.js";

const LOG_SOURCE = 'moveDocumentTopicTrigger';
const absFieldName = process.env.ABS_FIELD_NAME || 'ABS';
const pbsFieldName = process.env.PBS_FIELD_NAME || 'PBS';
const obsFieldName = process.env.OBS_FIELD_NAME || 'IssuingEntity';
const statusFieldName = process.env.STATUS_FIELD_NAME || 'DocumentStatus';
const approvedValue = process.env.APPROVED_STATUS_VALUE || 'Applicable';
const refFieldName = process.env.REF_FIELD_NAME || 'ProjectReference';
const docIdFieldName = process.env.DOC_ID_FIELD_NAME || '_dlc_DocId';
const docIdFieldNameRest = process.env.DOC_ID_FIELD_NAME_REST || 'OData__dlc_DocId';
const docLibNames = process.env.DOC_LIB_NAMES?.split(',') || ['Draft', 'ApplicableDocuments', 'PreviousVersions'];
const createdApplicableDocSessionId = process.env.CREATED_APPLICABLE_DOC_SID || 'GenRef';

export async function moveDocumentTopicTrigger(message: IDmsEvent, context: InvocationContext): Promise<void> {
    logger.init(context);
    const massImportService = new MassImportService();
    try {
         logger.trackEvent({
            name: 'projectHttpTrigger',
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(message) }
        });
        await pnpjs.Init(true);
        if (
            !message.action ||
            !message.itemId ||
            !message.libraryId ||
            !message.siteUrl ||
            !message.fields ||
            !message.fields[absFieldName] ||
            !message.fields[pbsFieldName] ||
            !message.fields[obsFieldName]
        ) throw new Error(`Invalid message: ${JSON.stringify(message)}`);
        const sp = pnpjs.sp(message.siteUrl);
        const spUtilitySvc = new SPUtilityService(sp);

        //Get list Ids
        const lists = await Promise.all(
            docLibNames.map(
                libName => {
                    return spUtilitySvc.getListFromName(libName, 'Id', undefined, true);
                }
            )
        );

        if (message.fields[statusFieldName] != approvedValue) {
            logger.trackTrace({
                message: `Document not approved. Status: ${message.fields[statusFieldName]}`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
            return;
        }
        const documentUniqueId = message.fields[docIdFieldName];
        //Get previous versions         
        const prevItems = await getAllItemsFromLargeList(sp, lists[1].Id, message.fields[refFieldName], documentUniqueId);

        const prevItemsRootFolder = await sp.web.lists.getById(lists[2].Id).rootFolder.select('ServerRelativeUrl')();
        for (const prevItem of prevItems) {
            const prevItemTargetFolder = `${prevItemsRootFolder.ServerRelativeUrl}/${_getSafeFolderName(message.fields[obsFieldName])}/${_getSafeFolderName(message.fields[absFieldName])}/${_getSafeFolderName(message.fields[pbsFieldName])}`;
            // await _moveDocument(sp, lists[1].Id, prevItem.Id, prevItemTargetFolder, prevItemsRootFolder.ServerRelativeUrl);
            await moveItembyCountandFolderStructure(sp, prevItemTargetFolder, lists[1].Id, prevItem.Id, prevItemsRootFolder.ServerRelativeUrl, lists[2].Id);
        }

        const applicableRootFolder = await sp.web.lists.getById(lists[1].Id).rootFolder.select('ServerRelativeUrl')();
        const targetFolder = `${applicableRootFolder.ServerRelativeUrl}/${_getSafeFolderName(message.fields[obsFieldName])}/${_getSafeFolderName(message.fields[absFieldName])}/${_getSafeFolderName(message.fields[pbsFieldName])}`;

        const movedItem = await moveItembyCountandFolderStructure(sp, targetFolder, message.libraryId, message.itemId, applicableRootFolder.ServerRelativeUrl, lists[1].Id)

        if (message.fromBackend) {
            logger.trackTrace({
                message: `Initializing mass import service`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
            massImportService.setSP(sp);
            logger.trackTrace({
                message: `Publishing applicable doc creation event ${movedItem.libraryId}, ${movedItem.itemId}, ${message.user}`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
            await massImportService.publishDocumentChangeEvent(movedItem.libraryId, movedItem.itemId, message.user, 'create', createdApplicableDocSessionId);
        }
        logger.trackTrace({
            message: `Document moved to target folder`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(error);
    }
    finally {
        massImportService && massImportService.cleanup();
    }
}
function isNumber(str: string) {
    if (!str.trim()) {
        return false;
    }
    const num = Number(str);
    return !isNaN(num);
}
async function moveItembyCountandFolderStructure(sp: SPFI, targetFolder: string, sourceListId: any, sourceItemId: number, targetLibRelativeUrl: string, targetListId: any) {

    //Get the all the Folders to determine where the incoming document set should get moved
    const folder = await sp.web.getFolderByServerRelativePath(targetFolder);

    const folders = await folder.folders();

    //Get the folder name to move document set & Calculate the new Folder name based on existing folders
    const systemfolders = folders.filter(f => f.ProgID === null);
    const folderName: number[] = [];
    systemfolders.map(f => {
        const num = Number(f.Name);
        if (!isNaN(num) && isFinite(num)) {
            folderName.push(num);
        }
    });
    folderName.sort((a, b) => a - b);
    let newFolderName = 1;
    if (folderName.length > 0) {
        newFolderName = folderName[folderName.length - 1];
        //Check the item count on the existing folder
        const exisingFolder = await sp.web.getFolderByServerRelativePath(`${targetFolder}/${folderName[folderName.length - 1].toString()}`);
        const existingFolders = await exisingFolder.folders();
        //If existing folder count more than specified amount, then create new numbered folder and ove the incoming document set
        if (existingFolders.length >= 4000) // update back to 4000
            newFolderName = folderName[folderName.length - 1] + 1;
    }
    //Move the incoming document set to the above caluculated folder
    let movedItem = null;
    if (newFolderName > 1)
        movedItem = await _moveDocument(sp, sourceListId, sourceItemId, `${targetFolder}/${newFolderName.toString()}`, targetLibRelativeUrl);
    else
        movedItem = await _moveDocument(sp, sourceListId, sourceItemId, targetFolder, targetLibRelativeUrl);

    //In the initial folder stucture(OBS/ABS/PBS)  document sets outside of the folders , then create a new numbered folder and move all Doc sets to it
    if (folders.length >= 4000) { //update back to 4000

        //Rename the target Folder to numbered
        const item = await folder.getItem();
        const result = await item.update({ FileLeafRef: newFolderName.toString() });
        //Create new folder with 3rd level
        await sp.web.folders.addUsingPath(targetFolder);

        //Move numbered folder to 3rd level
        await _moveDocument(sp, targetListId, item["ID"], targetFolder, targetLibRelativeUrl);

        //Move all numbered folders back to 3rd Level folders
        const updatedFolder = await sp.web.getFolderByServerRelativePath(`${targetFolder}/${newFolderName}`);
        const updatedFolders = await updatedFolder.folders();
        const defaultFold = updatedFolders.filter(f => f.ProgID === null);
        for (let i = 0; i < defaultFold.length; i++) {
            const _folder = await sp.web.getFolderByServerRelativePath(defaultFold[i].ServerRelativeUrl);
            const _item = await _folder.getItem();
            await _moveDocument(sp, targetListId, _item["ID"], targetFolder, targetLibRelativeUrl);
        }
    }
    return movedItem;
}

async function getAllItemsFromLargeList(sp: SPFI, listID: string, refValue: string, docID: string) {
    let allItems = [];
    let nextPage = null; // Tracks the skipToken for pagination
    let itemsPerPage = 2000; // Fetch 500 items per request (can be adjusted)

    try {
        // Loop to handle pagination
        do {
            //   let query =  await sp.web.lists.getById(listID).items.filter(`(${refFieldName} eq '${refValue}') and (${docIdFieldNameRest} ne '${docID}')`).select('Id', 'File').top(itemsPerPage);
            // .expand("ContentType") // Expand for list item and content type
            //       .select("ServerRelativeUrl", "Id", "FileRef", "FileLeafRef", "ContentType/Name", "ContentType/Id")

            let query = await sp.web.lists.getById(listID).items.select('Id', refFieldName, docIdFieldNameRest).top(itemsPerPage);

            // If there is a nextPage (skipToken), append it to the query
            if (nextPage) {
                query = query.skip(nextPage);
            }

            // Fetch the current batch of items
            const currentItems = await query();

            // Append current batch to the allItems array
            allItems = allItems.concat(currentItems);

            // Check if there's more data (look for nextPage in response)
            nextPage = currentItems.length > 0 ? currentItems[currentItems.length - 1].Id : null;

            console.log(`Fetched ${currentItems.length} items, Total items so far: ${allItems.length}`);

        } while (nextPage);

        const resultItems = allItems.filter(itm => itm[refFieldName] === `${refValue}` && itm[docIdFieldNameRest] !== `${docID}`)

        return resultItems;
    } catch (error) {
        logger.trackTrace({
            message: `Issue in Reading items from Applicable library: ${error}`,
            properties: { source: LOG_SOURCE },
            severity: 'Warning'
        });
    }
    return [];
}


async function _moveDocument(sp: SPFI, libraryId: string, itemId: number, targetFolder: string, rootFolder: string): Promise<{ itemId: number; libraryId: string }> {
    logger.trackTrace({
        message: `Target Folder ${targetFolder}`,
        properties: { source: LOG_SOURCE },
        severity: 'Information'
    });
    const itemFolder = await sp.web.lists.getById(libraryId).items.getById(itemId).folder.select('ServerRelativeUrl')();
    const currentFolder = itemFolder.ServerRelativeUrl.substring(0, itemFolder.ServerRelativeUrl.lastIndexOf('/'));
    const currentName = itemFolder.ServerRelativeUrl.substring(itemFolder.ServerRelativeUrl.lastIndexOf('/') + 1);
    if (currentFolder === targetFolder) {
        logger.trackTrace({
            message: `Already in target folder`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
        return;
    }
    const spTargetFolder = await sp.web.getFolderByServerRelativePath(targetFolder).select('Exists')();
    logger.trackTrace({
        message: `Target Folder Exists: ${spTargetFolder.Exists}`,
        properties: { source: LOG_SOURCE },
        severity: 'Information'
    });
    if (!spTargetFolder.Exists) {
        await _ensureFolder(sp, rootFolder, targetFolder);
        logger.trackTrace({
            message: `Target Folder Created`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
    }
    //const movedFolder = await sp.web.lists.getById(libraryId).items.getById(itemId).folder.copyByPath(`${targetFolder}/${currentName}`, {
    const movedFolder = await sp.web.lists.getById(libraryId).items.getById(itemId).folder.moveByPath(`${targetFolder}/${currentName}`, {
        RetainEditorAndModifiedOnMove: true,
        ShouldBypassSharedLocks: true
    });
    const movedFolderMeta = (await (await movedFolder.getItem()).select('Id', 'ParentList/Id').expand('ParentList')());
    return {
        itemId: movedFolderMeta.Id,
        libraryId: movedFolderMeta.ParentList.Id
    };
}

function _getSafeFolderName(taxField: { Label: string } | string): string {
    // if ((taxField as string).charAt) return (taxField as string).replace(/[^a-z0-9\s_-]/gi, '');
     if ((taxField as string).charAt) return (taxField as string);
     if (!(taxField as { Label: string })?.Label) throw new Error('Invalid taxonomy field value');
    // return (taxField as { Label: string }).Label.replace(/[^a-z0-9\s_-]/gi, '');
     //return (taxField as { Label: string }).Label;
     return (taxField as { Label: string }).Label.replaceAll(":","").replaceAll("*","").replaceAll("/","");
 }

async function _ensureFolder(sp: SPFI, rootFolder: string, targetFolder: string): Promise<void> {
    const pathComponents = targetFolder.replace(rootFolder, '').split('/').filter(p => !!p);
    let path = rootFolder;
    while (pathComponents.length > 0) {
        path += `/${pathComponents.splice(0, 1)[0]}`;
        let spFolder = await sp.web.getFolderByServerRelativePath(path).select('Exists')();
        if (!spFolder.Exists) {
            try {
                await sp.web.folders.addUsingPath(path);
            } catch (error) {
                if (error.message.includes('-2130575257')) {
                    spFolder = await sp.web.getFolderByServerRelativePath(path).select('Exists')();
                    if (!spFolder.Exists) throw new Error(error)
                }
            }
            logger.trackTrace({
                message: `Folder Created: ${path}`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
        }
    }
}

app.serviceBusTopic('moveDocumentTopicTrigger', {
    connection: 'dmsdev1_SERVICEBUS',
    topicName: 'documentchanges',
    subscriptionName: 'Auto-Classify',
    isSessionsEnabled: true,
    handler: moveDocumentTopicTrigger
});


