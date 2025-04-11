
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { AadHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IProjectPropertiesService } from "../../../service/IProjectPropertiesService";
import { IPermissionsService } from "../../../service/PermissionsService";


export interface IAdvancedSearchProps {
  context:WebPartContext;
  technicalFields:any;
  baseLineFields:any;
  theme: IReadonlyTheme;
  dmsClient: AadHttpClient;
  projectPropertiesService: IProjectPropertiesService;
  permissionsService: IPermissionsService;
  SearchRequestListName:string;
}
