import { SPFI } from "@pnp/sp";
import { SPUtilityService } from "./SPUtilityService.js";

const LOG_SOURCE = "ClassificationService";
export class ClassificationService {
    private _sputility: SPUtilityService;
    constructor(private _sp: SPFI) {
        this._sputility = new SPUtilityService(_sp);
    }
    // public async ClassifyDocument(applicableLibId:string,prevVersionLibId:string, revisionsFilter:string, classificationValues:string[]): Promise<void> {
    //     //Get previous versions
    //     const prevItems = await this._sp.web.lists.getById(applicableLibId).items.filter(revisionsFilter).select('Id', 'File')();
    //     const prevItemsRootFolder = await this._sp.web.lists.getById(prevVersionLibId).rootFolder.select('ServerRelativeUrl')();
    //     for (const prevItem of prevItems) {
    //         const prevItemTargetFolder = `${prevItemsRootFolder.ServerRelativeUrl}/${classificationValues.map(v=>this._getSafeFolderName(v)).join('/')}`;
    //         await this._sputility.moveDocument(applicableLibId, prevItem.Id, prevItemTargetFolder, prevItemsRootFolder.ServerRelativeUrl);
    //     }

    //     const rootFolder = await this._sp.web.lists.getById(applicableLibId).rootFolder.select('ServerRelativeUrl')();
    //     const targetFolder = `${rootFolder.ServerRelativeUrl}/${classificationValues.map(v=>this._getSafeFolderName(v)).join('/')}`;
    //     await this._sputility.moveDocument(message.libraryId, message.itemId, targetFolder, rootFolder.ServerRelativeUrl);
    // }
    private _getSafeFolderName(taxField: { Label: string } | string): string {
        if ((taxField as string).charAt) return (taxField as string).replace(/[^a-z0-9\s_-]/gi, '');
        if (!(taxField as { Label: string })?.Label) throw new Error('Invalid taxonomy field value');
        return (taxField as { Label: string }).Label.replace(/[^a-z0-9\s_-]/gi, '');
    }
}