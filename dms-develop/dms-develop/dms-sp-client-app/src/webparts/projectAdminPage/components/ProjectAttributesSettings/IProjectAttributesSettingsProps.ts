import { AadHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectAttributesSettingsProps {
  context: WebPartContext;
  onError: (errorMessage: string) => void;
  libraryIds: string[];
  projectEndpoint: string;
  dmsClient: AadHttpClient;
}