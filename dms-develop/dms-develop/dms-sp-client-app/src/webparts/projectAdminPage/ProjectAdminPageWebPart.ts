import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProjectAdminPageWebPartStrings';
import ProjectAdminPage from './components/ProjectAdminPage';
import { IProjectAdminPageProps } from './components/IProjectAdminPageProps';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/content-types';
import "@pnp/sp/sites";
import { getSP } from '../../pnpjs-config';
import { DmsDocumentFormConstants, projectEndpoint, projectTemplateEndPoint } from '../../constants';
import { AadHttpClient } from '@microsoft/sp-http';
import { ProjectPropertiesService } from '../../service/ProjectPropertiesService';
import { IProjectPropertiesService } from '../../service/IProjectPropertiesService';



export interface IProjectAdminPageWebPartProps {
  dmsUserApiAppId: string;
  projectEndpoint: string;
  projectConfigurationPageURL: string;
  projectTemplateEndpoint:string
}

export default class ProjectAdminPageWebPart extends BaseClientSideWebPart<IProjectAdminPageWebPartProps> {
  private _theme: IReadonlyTheme;
  private _dmsClient: AadHttpClient;
  private _projectPropertiesService: IProjectPropertiesService;

  public render(): void {
    const element: React.ReactElement<IProjectAdminPageProps> = React.createElement(
      ProjectAdminPage,
      {
        context: this.context,
        theme: this._theme,
        projectEndpoint: this.properties.projectEndpoint || projectEndpoint,
        dmsClient: this._dmsClient,
        projectConfigurationPageURL: this.properties.projectConfigurationPageURL,
        projectPropertiesService: this._projectPropertiesService,
        projectTemplateEndpoint:this.properties.projectTemplateEndpoint||projectTemplateEndPoint
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    const serviceInitPromise = new Promise<void>((resolve) => {
      this.context.serviceScope.whenFinished(() => {
        this._projectPropertiesService = this.context.serviceScope.consume(ProjectPropertiesService.serviceKey);
        resolve();
      });
    });
    await serviceInitPromise;
    this._dmsClient = await this.context.aadHttpClientFactory
      .getClient(this.properties.dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);
    getSP(this.context);
  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;
    this._theme = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('projectEndpoint', {
                  label: strings.ProjectEndpointFieldLabel
                }),

                PropertyPaneTextField('projectConfigurationPageURL', {
                  label: strings.projectConfigurationPageURL
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
