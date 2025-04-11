import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { AadHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IProjectPropertiesService } from "../../../service/IProjectPropertiesService";

export interface IProjectAdminPageProps {
  context: WebPartContext;
  theme: IReadonlyTheme;
  projectEndpoint: string;
  dmsClient: AadHttpClient;
  projectConfigurationPageURL: string;
  projectPropertiesService: IProjectPropertiesService;
  projectTemplateEndpoint:string
}
