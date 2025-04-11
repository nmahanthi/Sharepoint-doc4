import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import BaselineForms, { IBaselineFormsProps } from './components/BaselineForms';
import BaselineService from '../../service/BaselineService';
import { DmsDocumentFormConstants } from '../../constants';
import { AadHttpClient } from '@microsoft/sp-http';
import { ILinkedItems } from '../../interfaces/IGlobalInterfaces';
import { INotificationService, NotificationService } from '../../service/NotificationService';
import { IPermissionsService, PermissionsService } from '../../service/PermissionsService';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IBaselineFormsFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
  httpTriggerEndPoint: string;
  freezeEndpoint: string;
  dmsUserApiAppId: string;
  TechnicalFields?:string;
  BaseLineFields?:string;

}

const LOG_SOURCE: string = 'BaselineFormsFormCustomizer';

export default class BaselineFormsFormCustomizer
  extends BaseFormCustomizer<IBaselineFormsFormCustomizerProperties> {
  private _dmsClient: AadHttpClient;
  private _notificationService: INotificationService;
  private _permissionsService: IPermissionsService;
  public async onInit(): Promise<void> {
    const { dmsUserApiAppId, httpTriggerEndPoint: propHttpTriggerEndPoint } = this.properties;
    Log.info(LOG_SOURCE, 'Activated BaselineFormsFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));

    const httpTriggerEndPoint = propHttpTriggerEndPoint || "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/baseline?code=g9KiiJXZFE9gcdGui_5Y6reGDlLxqiJ-nkvy0yjFFwMJAzFujkcXmA%3D%3D";
    await BaselineService.init(this.context,dmsUserApiAppId);
    BaselineService.setHttpTriggerEndPoint(httpTriggerEndPoint);
    this._dmsClient = await this.context.aadHttpClientFactory.getClient(dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);
    const serviceReady = new Promise<void>((resolve, reject) => {
      this.context.serviceScope.whenFinished(() => {
        this._permissionsService = this.context.serviceScope.consume(PermissionsService.serviceKey);
        this._notificationService = this.context.serviceScope.consume(NotificationService.serviceKey);
        resolve();
      });
    });
    await serviceReady;
    return Promise.resolve();
  }

  public render(): void {
    // Use this method to perform your custom rendering.
    const freezeEndpoint = this.properties.freezeEndpoint || "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/baseline/freeze?code=g9KiiJXZFE9gcdGui_5Y6reGDlLxqiJ-nkvy0yjFFwMJAzFujkcXmA%3D%3D";

    const baselineForms: React.ReactElement<{}> =
      React.createElement(BaselineForms, {
        context: this.context,
        displayMode: this.displayMode,
        itemID: this.context.itemId,
        item: this.context.item as unknown as ILinkedItems,
        onSave: this._onSave,
        onClose: this._onClose,
        dmsClient: this._dmsClient,
        freezeEndpoint: freezeEndpoint,
        notificationService: this._notificationService,
         technicalFields:this.properties.TechnicalFields ||  "Title,LinkFilename,ProjectRevision,SupplierReference,SupplierRevision,OtherReference,OtherRevision,Keywords,FolioNumber,AlstomPartNumbers,MoCStatus,AlstomAssetID,MaterialID,LocationID,JobPlanID,InspectionID,FailureID,MaintenancePlanID,HazardID,PermitToWorkID,MoCID,IncidentID,HandoverDocumentFilename,HandoverPBSLevelCode,HandoverDocumentStatus,HandoverDocumentRevision,CustomerResponseExpectedDate,CustomerResponseReference,CustomerResponseDate,CustomerResponseStatus,ReforecastSubmissionDate,ActualSubmissionDate,CustomerMilestone,TransmittalReference,TransmittalDate,ApprovalCycleLaunchDate,Verifiedby,VerificationDate,Validatedby,ValidationDate,Approvedby,ApprovalDate,CustomerComments,EffectiveDate,ExpiryDate,ReplacementFor,ReplacedBy,ChangeFromNumber,ChangeToNumber,SafetyCriticity,CustomernameFinal,ProductName,ProjectAlias1,ScopeCountry,WorkflowStatus,DocumentStatus,CustomernameDirect,ProjectName,ProjectCode,DocumentSetDescription,ExternalURL,LineNumber,PLOwnership,DocumentType,DocumentLanguage,ABS,PBS,Contractual,SupplierName,CustomerRevision,RequirementID,ChangeRequestID,LinkedtoRequirementsDB,SubmissionCategory,DBS,FBS,SBS,GRMilestone,Authors,_Comments,WBS,PlannedSubmissionDate,IssuingEntity,ConfidentialityLevel,_dlc_DocId,AlstomPartNoReference,AlstomReference,ProjectReference",
         baseLineFields:this.properties.BaseLineFields || "BaselineComments,Code,CRID,IsLatest,Title,BaselineVersion,Created,Author,Modified,Editor",
         permissionsService:this._permissionsService

      } as IBaselineFormsProps);

    ReactDOM.render(baselineForms, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose = (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
