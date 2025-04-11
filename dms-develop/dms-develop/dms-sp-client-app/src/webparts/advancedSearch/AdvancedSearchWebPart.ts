import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'AdvancedSearchWebPartStrings';
import AdvancedSearch from './components/AdvancedSearch';
import { IAdvancedSearchProps } from './components/IAdvancedSearchProps';
import BaselineService from '../../service/BaselineService';
import { AadHttpClient } from '@microsoft/sp-http';
import { IProjectPropertiesService } from '../../service/IProjectPropertiesService';
import { ProjectPropertiesService } from '../../service/ProjectPropertiesService';
import { IPermissionsService, PermissionsService } from '../../service/PermissionsService';
export interface IAdvancedSearchWebPartProps {
  TechnicalFields: string;
  BaseLineFields: string;
  dmsUserApiAppId: any;
  description: string;
  SearchRequestListName:string;
}

export default class AdvancedSearchWebPart extends BaseClientSideWebPart<IAdvancedSearchWebPartProps> {
  private _theme: IReadonlyTheme;
  private _dmsClient: AadHttpClient;
  private _projectPropertiesService: IProjectPropertiesService;
  private _permissionsService: IPermissionsService;
  public render(): void {
    const element: React.ReactElement<IAdvancedSearchProps> = React.createElement(
      AdvancedSearch,
      {
        context: this.context,
        theme: this._theme,
        dmsClient: this._dmsClient,
        technicalFields:this.properties.TechnicalFields ||  "Title,LinkFilename,ProjectRevision,SupplierReference,SupplierRevision,OtherReference,OtherRevision,Keywords,FolioNumber,AlstomPartNumbers,MoCStatus,AlstomAssetID,MaterialID,LocationID,JobPlanID,InspectionID,FailureID,MaintenancePlanID,HazardID,PermitToWorkID,MoCID,IncidentID,HandoverDocumentFilename,HandoverPBSLevelCode,HandoverDocumentStatus,HandoverDocumentRevision,CustomerResponseExpectedDate,CustomerResponseReference,CustomerResponseDate,CustomerResponseStatus,ReforecastSubmissionDate,ActualSubmissionDate,CustomerMilestone,TransmittalReference,TransmittalDate,ApprovalCycleLaunchDate,Verifiedby,VerificationDate,Validatedby,ValidationDate,Approvedby,ApprovalDate,CustomerComments,EffectiveDate,ExpiryDate,ReplacementFor,ReplacedBy,ChangeFromNumber,ChangeToNumber,SafetyCriticity,CustomernameFinal,ProductName,ProjectAlias1,ScopeCountry,WorkflowStatus,DocumentStatus,CustomernameDirect,ProjectName,ProjectCode,DocumentSetDescription,ExternalURL,LineNumber,PLOwnership,DocumentType,DocumentLanguage,ABS,PBS,Contractual,SupplierName,CustomerRevision,RequirementID,ChangeRequestID,LinkedtoRequirementsDB,SubmissionCategory,DBS,FBS,SBS,GRMilestone,Authors,_Comments,WBS,PlannedSubmissionDate,IssuingEntity,ConfidentialityLevel,_dlc_DocId,AlstomPartNoReference,AlstomReference,ProjectReference",
        baseLineFields:this.properties.BaseLineFields || "BaselineComments,Code,CRID,IsLatest,Title,BaselineVersion",
        projectPropertiesService: this._projectPropertiesService,
        permissionsService:this._permissionsService,
        SearchRequestListName:this.properties.SearchRequestListName||"SearchRequests"
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
    this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);
    await serviceInitPromise;
    await BaselineService.init(this.context,this.properties.dmsUserApiAppId);
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const {
      semanticColors
    } = currentTheme;

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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
