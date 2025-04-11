import "@pnp/sp/files";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/content-types/list";
import "@pnp/sp/content-types/web";
import "@pnp/sp/content-types";
import "@pnp/sp/sharing";
import "@pnp/sp/batching";
import "@pnp/sp/security";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { FormCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import "@pnp/sp/search";
import { ICustomSearchResults, ISPColumn, ISPType } from "../interfaces/IGlobalInterfaces";
import CONSTANTS from "../_Constants";
import { AadHttpClient } from "@microsoft/sp-http";
import { DmsDocumentFormConstants } from "../constants";
import { Web } from "@pnp/sp/webs";
import { getSP } from "../pnpjs-config";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { RenderListDataOptions } from "@pnp/sp/lists";

export interface ISPSvcProps {
  context: FormCustomizerContext | WebPartContext;
}

class BaselineService {
  private sp: SPFI;
  // private graph: GraphFI;
  private context: FormCustomizerContext | WebPartContext | ListViewCommandSetContext;
  private httpTriggerEndPoint: string;
  private _dmsClient: AadHttpClient;

  public async init(context: FormCustomizerContext | WebPartContext | ListViewCommandSetContext, appId?: string) {
    if (!!context) {
      this.sp = getSP(context);
      this.context = context;
    }
    this._dmsClient = await this.context.aadHttpClientFactory.getClient(appId || DmsDocumentFormConstants.dmsUserApiAppId);
    return this.sp;
  }

  public setHttpTriggerEndPoint(endpoint: string) {
    this.httpTriggerEndPoint = endpoint;
  }


  public async getHubSiteID() {
    return await this.sp.site.select("HubSiteId")();
  }

  public async getSiteContentTypes(): Promise<ISPType[]> {

    const contentType = await this.sp.web.contentTypes.select("Id", "Name", "Description").filter("Name eq 'Technical Document' or Name eq 'Baseline'")();

    const ctType: ISPType[] = [];
    contentType.map(ctype => {
      ctType.push({
        id: ctype.Id.StringValue,
        title: ctype.Name
      });
    })
    return ctType;
  }

  public async getContentTypeFields(contentTypeId: string): Promise<ISPColumn[]> {
    const fields = await this.sp.web
      .contentTypes
      .getById(contentTypeId)
      .fields();
    const flds: ISPColumn[] = [];
    fields.map(fld => {
      //   if (!fld.Hidden && !fld.ReadOnlyField)
      flds.push({
        id: fld.Id,
        group: fld.Group,
        hidden: fld.Hidden,
        internalName: fld.InternalName,
        title: fld.Title,
        typeAsString: fld.TypeAsString
      });
    })
    return flds;
  }
  public async getFieldChoices(fieldName: string): Promise<any> {
    if ('list' in this.context) {
      const choiceField = await this.sp.web.lists.getByTitle(this.context.list.title).fields.getByTitle(fieldName)();
      return choiceField.Choices;
    }
  }
  public async getCurrentUser() {
    return await this.sp.web.currentUser();
  }

  public async getUserId(email: string) {
    const userInfo = await this.sp.web.ensureUser(email);
    return userInfo.Id
  }

  public async getUserById(id: number) {
    // if(id)
    // {
    return await this.sp.web.getUserById(id)();
    // }

  }

  public async getBaseLineContentTypeId() {
    if ('list' in this.context) {
      const cType = await this.sp.web.lists.getByTitle(this.context.list.title).contentTypes();

      return cType.filter(c => c.Name === CONSTANTS.ContentTypeNames.Baseline)[0]?.Id?.StringValue;
    }
  }

  public async createListItem(listTitle: string, object: any) {
    const i = await this.sp.web.lists.getByTitle(listTitle).items.add(object);
    return i;
  }
  public async updateListItem(listTitle: string, object: any, itemId: number) {
    const i = await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update(object);
    return i;
  }

  public async batchDelete(listTitle: string, itemId: number[]) {

    const [batchedSP, execute] = this.sp.batched();

    const res = [];
    for (let i = 0; i < itemId.length; i++)
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      batchedSP.web.lists.getByTitle(listTitle).items.getById(itemId[i]).delete().then(r => res.push(r));

    await execute();

  }

  public async getReferenceListItem(listName: string, baselineId: string, webUrl: string) {
    const web = Web([this.sp.web, webUrl]);
    const retResult = await web.lists.getByTitle(listName).items.filter(`${CONSTANTS.TechnicalRefernceListFieldNames.BaselineParentID} eq '${baselineId}'`)();
    if (CONSTANTS.ListNames.BaselineRefernceList === listName) {
      await Promise.all(retResult.map(async (b) => {
        const childweb = Web([this.sp.web, b.SiteURL]);
        const retResult1 = await childweb.lists.getByTitle("Baselines").items.filter(`ID eq '${b.BaselineChildID}'`)();
        b.BaselineStatus = retResult1[0] ? retResult1[0].BaselineStatus : "";
        b.Code = retResult1[0] ? retResult1[0].Code : "";
        b.BaselineVersion = retResult1[0] ? retResult1[0].BaselineVersion : "";
        b.BaselineChildID = retResult1[0] ? retResult1[0].Id.toString() : "0";
        return b;
      }));
      console.log(retResult);
      return retResult.filter(res => res.BaselineChildID !== "0");
    }
    return retResult;
  }

  // public async getReferenceListItem(listName: string, baselineId: string, webUrl: string) {
  //   const web = Web([this.sp.web, webUrl]);
  //   const retResult = await web.lists.getByTitle(listName).items.filter(`${CONSTANTS.TechnicalRefernceListFieldNames.BaselineParentID} eq '${baselineId}'`)();
  //   if (CONSTANTS.ListNames.BaselineRefernceList === listName) {
  //     await Promise.all(retResult.map(async (b) => {
  //       const childweb= Web([this.sp.web, b.SiteURL]);
  //       const retResult1 = await childweb.lists.getByTitle("Baselines").items.filter(`ID eq '${b.BaselineChildID}'`)();
  //       b.BaselineStatus = retResult1[0].BaselineStatus;
  //       b.Code=retResult1[0].Code;
  //       b.BaselineVersion=retResult1[0].BaselineVersion;
  //       return b;
  //     }));

  //     return retResult;
  //   }
  //   return retResult;
  // }

  // public getSearchResults = async (query: string): Promise<ICustomSearchResults[]> => {

  //   const results: ICustomSearchResults[] = [];
  //   const currentResults: any = await this.sp.search({
  //     Querytext: query,
  //     RowLimit: 500,
  //     TrimDuplicates: false, ClientType: 'ContentSearchRegular',
  //     SelectProperties: ["DlcDocId", "DocId", "ListItemID", "ContentType", "ProjectReferenceOWSTEXT", "ProjectRevisionOWSTEXT", "CRIDOWSTEXT", "ChangeRequestIDOWSTEXT", "CodeOWSTEXT", "Title", "SitePath", "LabelOWSTEXT", "BaselineStatusOWSTEXT", "DocumentStatusOWSCHCS","DocumentLink","Path"]
  //   });

  //   currentResults.PrimarySearchResults.map((r: any, index: number) => {
  //     results.push(
  //       {
  //         // eslint-disable-next-line dot-notation
  //         DocumentId: r['DlcDocId'],
  //         ItemId: r["ListItemID"],
  //         ProjectReference: r["ProjectReferenceOWSTEXT"],
  //         ProjectRevision: r["ProjectRevisionOWSTEXT"],
  //         CRID: r["CRIDOWSTEXT"],
  //         ChangeRequestT: r["ChangeRequestIDOWSTEXT"],
  //         Code: r["CodeOWSTEXT"],
  //         Title: r["Title"],
  //         SitePath: r["SitePath"],
  //         isSelected: false,
  //         DocId: r["DocId"],
  //         ContentType: r["ContentType"],
  //         Label: r["LabelOWSTEXT"],
  //         BaselineStatus: r["BaselineStatusOWSTEXT"],
  //         index: index,
  //         DocumentStatus: r["DocumentStatusOWSCHCS"],
  //         DocumentLink: r["Path"]
  //       }
  //     );
  //   });
  //   return results;
  // }

  public getSearchResults = async (query: string): Promise<ICustomSearchResults[]> => {
    const rowLimit = 500;
    const results: ICustomSearchResults[] = [];
    const currentResults: any = await this.sp.search({
      Querytext: query,
      RowLimit: rowLimit,
      TrimDuplicates: false, ClientType: 'ContentSearchRegular',
      SelectProperties: ["DlcDocId", "DocId", "ProjectNameOWSTEXT", "ProjectCodeOWSTEXT", "ListItemID", "ContentType", "ProjectReferenceOWSTEXT", "ProjectRevisionOWSTEXT", "CRIDOWSTEXT", "ChangeRequestIDOWSTEXT", "CodeOWSTEXT", "Title", "SitePath", "LabelOWSTEXT", "BaselineStatusOWSTEXT", "DocumentStatusOWSCHCS", "DocumentLink", "Path"]
    });

    currentResults.PrimarySearchResults.map((r: any, index: number) => {
      results.push(
        {
          // eslint-disable-next-line dot-notation
          DocumentId: r['DlcDocId'],
          ItemId: r["ListItemID"],
          ProjectReference: r["ProjectReferenceOWSTEXT"],
          ProjectRevision: r["ProjectRevisionOWSTEXT"],
          CRID: r["CRIDOWSTEXT"],
          ChangeRequestT: r["ChangeRequestIDOWSTEXT"],
          Code: r["CodeOWSTEXT"],
          Title: r["Title"],
          SitePath: r["SitePath"],
          isSelected: false,
          DocId: r["DocId"],
          ContentType: r["ContentType"],
          Label: r["LabelOWSTEXT"],
          BaselineStatus: r["BaselineStatusOWSTEXT"],
          index: index,
          DocumentStatus: r["DocumentStatusOWSCHCS"],
          DocumentLink: r["Path"],
          ProjectName: r["ProjectNameOWSTEXT"],
          ProjectCode: r["ProjectCodeOWSTEXT"]
        }
      );
    });

    for (let i = 2; i <= (currentResults.TotalRows / rowLimit) + 1; i++) {
      const result = await currentResults.getPage(i);
      console.log(result);
      if (result) {
        result.PrimarySearchResults.map((r: any, index: number) => {
          results.push(
            {
              // eslint-disable-next-line dot-notation
              DocumentId: r['DlcDocId'],
              ItemId: r["ListItemID"],
              ProjectReference: r["ProjectReferenceOWSTEXT"],
              ProjectRevision: r["ProjectRevisionOWSTEXT"],
              CRID: r["CRIDOWSTEXT"],
              ChangeRequestT: r["ChangeRequestIDOWSTEXT"],
              Code: r["CodeOWSTEXT"],
              Title: r["Title"],
              SitePath: r["SitePath"],
              isSelected: false,
              DocId: r["DocId"],
              ContentType: r["ContentType"],
              Label: r["LabelOWSTEXT"],
              BaselineStatus: r["BaselineStatusOWSTEXT"],
              index: index,
              DocumentStatus: r["DocumentStatusOWSCHCS"],
              DocumentLink: r["Path"],
              ProjectName: r["ProjectNameOWSTEXT"],
              ProjectCode: r["ProjectCodeOWSTEXT"]
            }
          );
        });
      }
    }

    console.log(results);
    return results;
  }

  public getSearchBatchResults = async (query: any[]): Promise<ICustomSearchResults[]> => {

    const [batchedSP, execute] = this.sp.batched();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const res: any[] = [];

    query.map(t => {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      batchedSP.search({
        Querytext: `DlcDocId=${t[CONSTANTS.TechnicalRefernceListFieldNames.DocumentID]}`,
        RowLimit: 500,
        TrimDuplicates: false, ClientType: 'ContentSearchRegular',
        SelectProperties: ["DlcDocId", "DocId", "ListItemID", "ContentType", "ProjectNameOWSTEXT", "ProjectCodeOWSTEXT", "ProjectReferenceOWSTEXT", "ProjectRevisionOWSTEXT", "CRIDOWSTEXT", "ChangeRequestIDOWSTEXT", "CodeOWSTEXT", "Title", "SitePath", "LabelOWSTEXT", "BaselineStatusOWSTEXT", "DocumentStatusOWSCHCS", "DocumentLink", "Path"]
      }).then(r => res.push(r));
    });
    await execute();
    const results: ICustomSearchResults[] = [];
    for (let i = 0; i < res.length; i++) {
      res[i].PrimarySearchResults.map((r: any, index: number) => {
        results.push(
          {
            // eslint-disable-next-line dot-notation
            DocumentId: r['DlcDocId'],
            ItemId: r["ListItemID"],
            ProjectReference: r["ProjectReferenceOWSTEXT"],
            ProjectRevision: r["ProjectRevisionOWSTEXT"],
            CRID: r["CRIDOWSTEXT"],
            ChangeRequestT: r["ChangeRequestIDOWSTEXT"],
            Code: r["CodeOWSTEXT"],
            Title: r["Title"],
            SitePath: r["SitePath"],
            isSelected: false,
            DocId: r["DocId"],
            ContentType: r["ContentType"],
            Label: r["LabelOWSTEXT"],
            BaselineStatus: r["BaselineStatusOWSTEXT"],
            index: index,
            DocumentStatus: r["DocumentStatusOWSCHCS"],
            DocumentLink: r["Path"],
            ProjectName: r["ProjectNameOWSTEXT"],
            ProjectCode: r["ProjectCodeOWSTEXT"]
          }
        );
      });
    }
    return results;
  }

  /*public getVersionDetailold = async (code: string, version: string, currentItemId?: number) => {
    const filterQuery = currentItemId
      ? `Code eq '${code}' and BaselineVersion eq '${version.trim()}' and ID ne ${currentItemId}`
      : `Code eq '${code}' and BaselineVersion eq '${version.trim()}'`;

    // Cast sp.web to any to bypass type error
    const hasInvalidVersionCheck = await (this.sp.web as any).lists
      .getByTitle('Baselines')
      .items.filter(filterQuery)();

    return hasInvalidVersionCheck;
  };*/

  public getVersionDetail = async (code: string, version: string, currentItemId?: number) => {
    const filterQuery = currentItemId
      ? `<View Scope='RecursiveAll'><Query><Where><And><And><Eq><FieldRef Name='Code' /><Value Type='Text'>${code}</Value></Eq><Eq><FieldRef Name='BaselineVersion' /><Value Type='Text'>${version.trim()}</Value></Eq></And><Neq><FieldRef Name='ID' /><Value Type='Counter'>${currentItemId}</Value></Neq></And></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields><RowLimit>1</RowLimit></View>`
      : `<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='Code' /><Value Type='Text'>${code}</Value></Eq><Eq><FieldRef Name='BaselineVersion' /><Value Type='Text'>${version.trim()}</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields><RowLimit>1</RowLimit></View>`;

    // Cast sp.web to any to bypass type error
    const hasInvalidVersionCheck = await this.sp.web.lists.getByTitle('Baselines').renderListDataAsStream({
            ViewXml: filterQuery,
            RenderOptions: RenderListDataOptions.ListData,
        });

    return hasInvalidVersionCheck;
  };

  public async sendmessage(payload: any) {

    //const httpTriggerEndpoint = "http://localhost:7071/api/baseline";
    const httpTriggerEndpoint = this.httpTriggerEndPoint;

    if (!httpTriggerEndpoint) {
      console.error('httpTriggerEndPoint is not set');
      return;
    }

    // const requestHeaders: Headers = new Headers();
    // requestHeaders.append("Content-type", "text/plain");
    // requestHeaders.append("Cache-Control", "no-cache");
    // const postOptions: IHttpClientOptions = {
    //   headers: requestHeaders,
    //   body: `${JSON.stringify(payload)}`
    // };

    //  return await this.context.httpClient.post(httpTriggerEndpoint, HttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
    //     if (response.ok) {
    //       console.log(response);
    //       return response.statusText;
    //     }
    //   }).catch((response: any) => {
    //     console.log(response);
    //     return response
    //   });
    return await this._dmsClient.post(httpTriggerEndpoint, AadHttpClient.configurations.v1, {
      body: JSON.stringify(payload)
    }).then(response => {
      if (!response.ok) {
        throw new Error(response.statusText);
      }
      return response.statusText;
    }).catch(error => {
      console.error("Error sending message:", error);
      return error;
    });


  }

  public async updateBaseliFrozeListItem(listTitle: string, itemId: number, fieldName: string, value: any): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update({
        [fieldName]: value

      })

    } catch (error) {
      console.error(error)

    }

  }


  public getSearchResultsbyProperties = async (query: string, selectProps: string[]): Promise<any[]> => {
    const rowLimit = 500;
    let results: any[] = [];
    const currentResults: any = await this.sp.search({
      Querytext: query,
      RowLimit: rowLimit,
      TrimDuplicates: false, ClientType: 'ContentSearchRegular',
      SelectProperties: selectProps//["DlcDocId", "DocId","ProjectNameOWSTEXT", "ProjectCodeOWSTEXT", "ListItemID", "ContentType", "ProjectReferenceOWSTEXT", "ProjectRevisionOWSTEXT", "CRIDOWSTEXT", "ChangeRequestIDOWSTEXT", "CodeOWSTEXT", "Title", "SitePath", "LabelOWSTEXT", "BaselineStatusOWSTEXT", "DocumentStatusOWSCHCS", "DocumentLink", "Path"]
    });
    results = [...results, ...currentResults.PrimarySearchResults];

    for (let i = 2; i <= (currentResults.TotalRows / rowLimit) + 1; i++) {
      const result = await currentResults.getPage(i);
      console.log(result);
      if (result) {
        results = [...results, ...result.PrimarySearchResults];
      }
    }
    return results;
  }


  // public async setItemPermissionsForSiteGroups(

  //   listName: string,
  //   itemId: number,

  // ): Promise<void> {

  //   // try {

  //   //   const list = this.sp.web.lists.getByTitle(listName);
  //   //   const item = list.items.getById(itemId);
  //   //   await item.breakRoleInheritance(false, true);
  //   //   console.log(`Permissions inheritance broken for item ${itemId}.`);

  //   //   const  siteUrl=this.context.pageContext.web.absoluteUrl;

  //   //   console.log("siteUrl",siteUrl);

  //   //   const currentUser = await this.sp.web.currentUser();
  //   //   const currentUserId = currentUser.Id;
  //   //   const groups = await this.sp.web.siteGroups();

  //   //   const readRoleDefinition = await this.sp.web.roleDefinitions.getByName("Read")();

  //   //   console.log("readRoleDefinition", readRoleDefinition);

  //   //   for (const group of groups) {

  //   //       console.log(`Skipping User8 :${group.Title}`)
  //   //     try {
  //   //       console.log(`Group ${group.Title} retrieved successfully.`);

  //   //       await item.roleAssignments.add(group.Id, readRoleDefinition.Id);
  //   //       console.log(`Read access granted to group ${group.Title} for item ${itemId}.`);
  //   //     } catch (error) {
  //   //       console.error(`Error setting permissions for group ${group.Title}:`, error);
  //   //     }
  //   //     const roleAssignments = await item.roleAssignments();

  //   //     console.log("currentUserId",currentUserId);
  //   //     console.log("roleAssignments",roleAssignments);

  //   //   }

  //   // } catch (error) {
  //   //   console.error(`Error setting permissions for item ${itemId}:`, error);
  //   // }
  // }


}

export default new BaselineService();