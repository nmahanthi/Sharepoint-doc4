import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './DmsApplicationActions.module.scss';
import { SPFI } from '@pnp/sp';
import { FieldCustomizerContext, ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { DmsRole, IPermissionsService } from '../../../service/PermissionsService';
import {
  DefaultButton,
  IContextualMenuItem,
  ContextualMenuItemType,
  assign,
  Icon,
  ContextualMenu
} from '@fluentui/react';
import { PermissionKind } from '@pnp/sp/security';
import { CommonFields, ContentTypeSiteIds, DmsDocumentFormConstants, FieldNames, ListUrls } from '../../../constants';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items";
import "@pnp/sp/security";
import ConfirmationPopup from '../../../components/Applicable/ApplicableStatus';
import CreateRevisionDialogContainer from '../../../components/CreateRevision/createRevisionDialogContainer';
import { IRequestService } from '../../../service/IRequestService';
import WorkflowsPanel, { IWorkflowsPanelProps } from '../../../components/WorkflowsPanel/WorkflowsPanel';
import { IItem } from '@pnp/sp/items';
import WorkflowDialogContent from '../../../components/Workflow/WorkflowDialogContent';
import { INotificationService } from '../../../service/NotificationService';
import ProofreadingDialogContent from '../../../components/ProofReading/ProofreadingDialogContent';
import { AadHttpClient } from '@microsoft/sp-http';


export interface IDmsApplicationActionsProps {
  sp: SPFI,
  listId: string,
  listUrl: string,
  webRelativeUrl: string,
  currentItem: ListItemAccessor,
  permissionService: IPermissionsService,
  strings: IDmsApplicationActionsFieldCustomizerStrings,
  currentUser: string,
  context: FieldCustomizerContext;
  workflowUrl: string;
  requestService: IRequestService;
  AVVAWorkflowUrl: string;
  notificationService: INotificationService;
  proofreadingWorkflowUrl: string;
  logHistoryHttpTriggerEndPoint: string;
  dmsUserApiAppId: string;
}

interface IDmsApplicationActionsState {
  hasEditAccess: number,
  currentRoles: DmsRole[],
  isContextMenuVisible: boolean,
  menuItems: IContextualMenuItem[],
  target:any
}

const LOG_SOURCE: string = 'DmsApplicationActions';

export default class DmsApplicationActions extends React.Component<IDmsApplicationActionsProps, IDmsApplicationActionsState> {
  validateContainer: any;
  private fileLeafRefs: { projectReference: string; projectRevision: string }[] | undefined;
  private _requestService: IRequestService;
  private panelPlaceHolder: HTMLDivElement | null = null;
  private errorMessage: string | null = null;
  private itemData: any;
  private AVVAcontainer: HTMLDivElement | null = null;
  private proofreadingcontainer: HTMLDivElement | null = null;
  private _dmsClient: AadHttpClient | null = null;

  constructor(props: IDmsApplicationActionsProps) {
    super(props);
    this.state = { hasEditAccess: -1, currentRoles: [], isContextMenuVisible: false, menuItems: [], target: null };
    this._getContextMenuItems = this._getContextMenuItems.bind(this);
    this.dismissMenu = this.dismissMenu.bind(this);
    this._onItemClick = this._onItemClick.bind(this);
  }

  public componentDidMount() {
    Log.info(LOG_SOURCE, 'React Element: DmsApplicationActions mounted');
    console.log("DmsApplicationActions ComponentDidMount");
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: DmsApplicationActions unmounted');
  }

  public render(): React.ReactElement<{}> {
    const { strings } = this.props;
    const { isContextMenuVisible, menuItems, target } = this.state
    if(!this._isDocumentSet()){
      return <></>
    }
    return (
      <div className={styles.dmsApplicationActions}>
        <DefaultButton className={styles.dmsChevButton} onClick={ this._getContextMenuItems }><Icon iconName={strings.MainMenuIcon} /></DefaultButton>
        <ContextualMenu items={menuItems} hidden={!isContextMenuVisible} onDismiss={this.dismissMenu} onItemClick={ this._onItemClick } target={target}/>
      </div>
    );
  }

  //#region private methods
  //get contextual menu item


  private dismissMenu() {
    this.setState({ isContextMenuVisible: false });
  }

  private _getContextMenuItems(event:any) {
    const { strings } = this.props;
    this.setState({ menuItems: [{ key: "Loading", text: strings.LoadingText, title: strings.LoadingText,
                                  iconProps: { iconName: strings.LoadingIcon, style: { color: strings.LoadingIconColor } }}], isContextMenuVisible: true, target: event.currentTarget || event.target });

    Promise.all([
      this._getCurrentUserRoles(),
      this._doesUserHaveEditAccess(),
      this._getDmsClient()
    ]).then(() => {
      const menus = [];
      const { hasEditAccess, currentRoles } = this.state;
      const { currentItem, listUrl } = this.props;

      if (this._isDocumentSet()) {
        //Edit menu
        if (hasEditAccess == 1) {
          menus.push({
            key: "Edit",
            text: strings.EditText,
            title: strings.EditText,
            iconProps: { iconName: strings.EditIcon, style: { color: strings.EditIconColor } }
          });
        }

        if (currentRoles.includes(DmsRole.Contributor) && currentItem.getValueByName(FieldNames.DocumentStatus) === "Applicable" && listUrl == ListUrls.ApplicableDocuments) {
          menus.push({
            key: "CreateRevision",
            text: strings.CreateRevisionText,
            title: strings.CreateRevisionText,
            iconProps: { iconName: strings.CreateRevisionIcon, style: { color: strings.CreateRevisionIconColor } }
          });
        }

        if (currentRoles.includes(DmsRole.DocumentController) && currentItem.getValueByName(FieldNames.DocumentStatus) === "Draft" && listUrl == ListUrls.Draft && currentItem.getValueByName("ItemChildCount") > 0) {
          menus.push({
            key: "ValidateDocument",
            text: strings.ValidateText,
            title: strings.ValidateText,
            iconProps: { iconName: strings.ValidateIcon, style: { color: strings.ValidateIconColor } }
          });
        }

        if (menus.length > 0) {
          //divider
          menus.push({ key: "EditDivider", itemType: ContextualMenuItemType.Divider });
        }

        //workflow detail panel
        menus.push({
          key: "WorkflowDetail",
          text: strings.WFDetailText,
          title: strings.WFDetailText,
          iconProps: { iconName: strings.WFDetailIcon, style: { color: strings.WFDetailIconColor } }
        });

        //360 panel
        menus.push({
          key: "RevisionPanel",
          text: strings.RevisionDetailText,
          title: strings.RevisionDetailText,
          iconProps: { iconName: strings.RevisionDetailIcon, style: { color: strings.RevisionDetailIconColor } }
        });

        //divider
        menus.push({ key: "DetailDivider", itemType: ContextualMenuItemType.Divider });

        //Proofreading and AVVA
        if (currentRoles.includes(DmsRole.Contributor) && currentItem.getValueByName(FieldNames.DocumentStatus) === "Draft" && listUrl == ListUrls.Draft && currentItem.getValueByName("ItemChildCount") > 0) {
          menus.push({
            key: "Proofreading",
            text: strings.ProofreadingText,
            title: strings.ProofreadingText,
            iconProps: { iconName: strings.ProofreadingIcon, style: { color: strings.ProofreadingIconColor } }
          });

          menus.push({
            key: "AVVA",
            text: strings.AVVAText,
            title: strings.AVVAText,
            iconProps: { iconName: strings.AVVAIcon, style: { color: strings.AVVAIconColor } }
          });
        }
      }

      this.setState({ menuItems: menus, isContextMenuVisible: true, target: event.currentTarget || event.target });
    }).catch((err) => {
      console.log(err);
      this.setState({ menuItems: [{ key: "Error", text: strings.ErrorText, title: strings.ErrorText,
        iconProps: { iconName: strings.ErrorIcon, style: { color: strings.ErrorIconColor } }}], isContextMenuVisible: true, target: event.currentTarget || event.target });
    });
  }

  //check if user has edit access to the current item
  private async _doesUserHaveEditAccess(): Promise<void> {
    const { sp, listId, currentItem } = this.props;
    const { hasEditAccess } = this.state;
    if (hasEditAccess == -1) {
      return sp.web.lists.getById(listId).items.getById(currentItem.getValueByName('ID')).currentUserHasPermissions(PermissionKind.EditListItems).then(edt => {
        this.setState({ hasEditAccess: edt ? 1 : 0 });
      });
    }
  }

  private async _getCurrentUserRoles(): Promise<void> {
    const { permissionService } = this.props;
    const { currentRoles } = this.state;
    if (!(currentRoles && currentRoles.length > 0)) {
      return permissionService.getCurrentUserRoles().then((roles) => {
        this.setState({ currentRoles: roles });
      });
    }
  }

  private async _getDmsClient(): Promise<void> {
    if (!this._dmsClient) {
      this._dmsClient = await this.props.context.aadHttpClientFactory.getClient(this.props.dmsUserApiAppId || DmsDocumentFormConstants.dmsUserApiAppId);
    }
  }

  private _onItemClick(_?: unknown | undefined, item?: IContextualMenuItem | undefined): boolean | void | undefined {
    const { listId, webRelativeUrl, currentItem } = this.props;
    if (!item) {
      return;
    }

    switch (item.key) {
      case "Edit":
        const arCurrentItemUrl = [...currentItem.getValueByName('FileRef').split("/")];
        arCurrentItemUrl.pop();
        const editUrl = `${webRelativeUrl}${ListUrls.EditForm.replace("[ListId]", listId).replace("[ItemId]", currentItem.getValueByName('ID')).replace("[Src]", encodeURIComponent(arCurrentItemUrl.join("/")))}`;
        window.location.href = editUrl;
        break;
      case "CreateRevision":
        this._showCreateRevisionDialog();
        break;
      case "ValidateDocument":
        this._validateDocumentShowDialog();
        break;
      case "WorkflowDetail":
        this._renderWfPanel();
        break;
      case "RevisionPanel":
        this.showRevisionDetails();
        break;
      case "AVVA":
        this._startAVVAWFDialog();
        break;
      case "Proofreading":
        this._startProofreadingDialog();
        break;
    }
  }

  private _isDocumentSet(): boolean {
    const { currentItem } = this.props;
    return currentItem.getValueByName('ContentTypeId')?.startsWith(ContentTypeSiteIds.TechnicalDocuments);
  }
  //#endregion

  //#region Validate Document
  private _validateDocumentShowDialog(): void {
    const { currentItem, webRelativeUrl } = this.props
    if (this.validateContainer) {
      ReactDOM.unmountComponentAtNode(this.validateContainer);
      document.body.removeChild(this.validateContainer);
      this.validateContainer = null;
    }
    this.validateContainer = document.createElement('div');
    document.body.appendChild(this.validateContainer);

    const onClose = () => {
      if (this.validateContainer) {
        ReactDOM.unmountComponentAtNode(this.validateContainer);
        document.body.removeChild(this.validateContainer);
        this.validateContainer = null;
        window.location.href = window.location.href;
        //this.raiseOnChange();
      }
    };

    const onConfirm = () => {
      //this.raiseOnChange();
      onClose();
    };
    this.fileLeafRefs = [];
    this.fileLeafRefs.push({
      projectReference: currentItem.getValueByName("ProjectReference") || "",
      projectRevision: currentItem.getValueByName("ProjectRevision") || ""
    });

    const element = React.createElement(ConfirmationPopup, {
      webUrl: webRelativeUrl,
      onClose,
      onConfirm,
      getSelection: this._getValidateSelection.bind(this),
      getselectedId: this._getValidateSelectedId.bind(this),
      fileLeafRefs: this.fileLeafRefs
    });
    ReactDOM.render(element, this.validateContainer);
  }

  private _getValidateSelection(): readonly ListItemAccessor[] | undefined {
    const { currentItem } = this.props;
    const listItems: ListItemAccessor[] = [];
    listItems.push(currentItem);
    return listItems;
  }

  private _getValidateSelectedId(): string | undefined {
    const { listId } = this.props;
    return listId ?? undefined;
  }

  //#endregion

  //#region Create Revision
  private _showCreateRevisionDialog() {
    const { currentItem, sp } = this.props;
    const projReference = currentItem.getValueByName(FieldNames.ProjectReference);
    const docSetServerRelativeUrl = currentItem.getValueByName(CommonFields.FileRef);
    const dialog = new CreateRevisionDialogContainer(projReference, sp, docSetServerRelativeUrl);
    dialog.show();
  }
  //#endregion

  //#region Workflow Panel
  private _dismissPanel() {
    this._renderWFPanelComponent({ showPanel: false });
  }

  private _showWFPanel(itemId: number) {
    if (!this.props.context.pageContext) {
      return;
    }
    //const httpTriggerEndPoint = this.props.logHistoryHttpTriggerEndPoint || "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/documentChangeHttpTrigger";
    const workflowUrl = this.props.workflowUrl || 'https://prod-150.westeurope.logic.azure.com:443/workflows/e2e86240560746d0a5d8191a09ed10ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=t-PsdqesaVR5HFrJKke4oqYiJo-G6a1vMOGZIvD9ceE';

    this._renderWFPanelComponent({
      showPanel: true,
      setShowPanel: this._dismissPanel.bind(this),
      documentId: itemId,
      requestService: this.props.requestService,
      cultureName: this.props.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
      currentUserLogin: this.props.context.pageContext.user.loginName,
      workflowUrl: workflowUrl
    });
  }

  private _renderWFPanelComponent(props: any) {
    if (!props.showPanel) {
      ReactDOM.unmountComponentAtNode(this.panelPlaceHolder!);
      return;
    }
    const element: React.ReactElement<IWorkflowsPanelProps> = React.createElement(WorkflowsPanel, assign({
      showPanel: false,
      setShowPanel: null,
      documentId: null,
      requestService: this._requestService,
      cultureName: this.props.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
      currentUserLogin: this.props.context.pageContext.user.loginName,
      isthreeSixtyDegree: false
    }, props));
    ReactDOM.render(element, this.panelPlaceHolder!);
  }

  private _renderWfPanel() {
    const { currentItem, requestService, listId, sp, context } = this.props;
    requestService.configure(
      listId,
      sp,
      { listId: listId, sp: sp, httpClient: context.httpClient, siteAbsoluteUrl: context.pageContext.web.absoluteUrl, currentItem: this.props.currentItem }
    )
    //console.log(event);
    const listItemId = currentItem.getValueByName('ID') as number;
    this._showWFPanel(listItemId);
  }
  //#endregion

  //#region  360 panel
  private async getDetails(projectReference: string): Promise<{ draftItems: IItem[], applicableItems: IItem[], previousVersionsItems: IItem[] }> {
    const draftItems = await this.getItems('Draft', projectReference);
    const applicableItems = await this.getItems('Applicable Documents', projectReference);
    const previousVersionsItems = await this.getItems('Previous Versions', projectReference);
    return { draftItems, applicableItems, previousVersionsItems };
  }

  private async getItems(listTitle: string, projectReference: string): Promise<IItem[]> {
    try {
      const items: IItem[] = await this.props.sp.web.lists.getByTitle(listTitle).items
        .select('Title', 'Modified', 'Author/Title', 'ProjectReference', 'ProjectRevision', 'FileRef')
        .expand('Author')
        .filter(`ProjectReference eq '${projectReference}'`)();
      return items;
    } catch (error) {
      console.error(`Error fetching items from ${listTitle}:`, error);
      this.errorMessage = `Error fetching items from ${listTitle}. Please try again later.`;
      return [];
    }
  }

  private async showRevisionDetails(): Promise<void> {

    this.props.requestService.configure(
      this.props.listId,
      this.props.sp,
      { listId: this.props.listId, sp: this.props.sp, httpClient: this.props.context.httpClient, siteAbsoluteUrl: this.props.context.pageContext.web.absoluteUrl, currentItem: this.props.currentItem }
    )
    const selectedRows = this.props.currentItem;
    if (!selectedRows) {
      return;
    }

    const selectedItem = this.props.currentItem;
    const projectReference = selectedItem.getValueByName('ProjectReference');

    if (!projectReference) {
      this.errorMessage = 'No Project Reference found for the selected item.';
      this._render360PanelComponent({ showPanel: true });
      return;
    }

    try {
      const details = await this.getDetails(projectReference);

      this.itemData = [
        ...details.draftItems,
        ...details.applicableItems,
        ...details.previousVersionsItems
      ].map(item => ({
        title: (item as any).Title,
        modified: new Date((item as any).Modified).toLocaleDateString('en-GB'),
        createdBy: (item as any).Author ? (item as any).Author.Title : 'Unknown',
        projectRevision: (item as any).ProjectRevision,
        link: (item as any).FileRef
      }));

      if (this.itemData && this.itemData.length > 0) {
        const listItemId = selectedItem.getValueByName('ID') as number;
        this._show360Panel(listItemId);
      } else {
        this.errorMessage = 'No document data available for the selected project.';
        this._render360PanelComponent({ showPanel: true });
      }
    } catch (error) {
      console.error('Error fetching document details:', error);
      this.errorMessage = 'An error occurred while retrieving document details. Please try again later.';
      this._render360PanelComponent({ showPanel: true });
    }
  }

  private _show360Panel(itemId: number) {
    try {
      if (!this.props.context.pageContext || !this.props.currentItem || !this.itemData) {
        this.errorMessage = 'An error occurred while preparing the panel. Please try again.';
        this._render360PanelComponent({ showPanel: true });
        return;
      }

      const httpTriggerEndPoint = this.props.logHistoryHttpTriggerEndPoint || "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/documentChangeHttpTrigger";
      const workflowUrl = this.props.workflowUrl || 'https://prod-150.westeurope.logic.azure.com:443/workflows/e2e86240560746d0a5d8191a09ed10ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=t-PsdqesaVR5HFrJKke4oqYiJo-G6a1vMOGZIvD9ceE';

      this._render360PanelComponent({
        showPanel: true,
        setShowPanel: this._dismissPanel.bind(this),
        documentId: itemId,
        requestService: this.props.requestService,
        cultureName: this.props.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
        currentUserLogin: this.props.context.pageContext.user.loginName,
        workflowUrl: workflowUrl,
        itemData: this.itemData,
        context: this.props.context,
        dmsClient: this._dmsClient,
        httpTriggerEndPoint: `${httpTriggerEndPoint}?documentId=${encodeURIComponent(this.props.currentItem.getValueByName("_dlc_DocIdUrl.desc"))}`
      });
    } catch (error) {
      console.error('Error showing panel:', error);
      this.errorMessage = 'An error occurred while displaying the panel. Please try again.';
      this._render360PanelComponent({ showPanel: true });
    }
  }

  private _render360PanelComponent(props: any) {
    try {
      if (!props.showPanel) {
        ReactDOM.unmountComponentAtNode(this.panelPlaceHolder!);
        return;
      }
      const element: React.ReactElement<IWorkflowsPanelProps> = React.createElement(WorkflowsPanel, assign({
        showPanel: false,
        setShowPanel: null,
        documentId: null,
        requestService: this.props.requestService,
        cultureName: this.props.context.pageContext.cultureInfo.currentUICultureName.toLowerCase(),
        currentUserLogin: this.props.context.pageContext.user.loginName,
        isthreeSixtyDegree: true,
        itemData: this.itemData,
        errorMessage: this.errorMessage // Pass the error message to the panel component
      }, props));
      ReactDOM.render(element, this.panelPlaceHolder!);
    } catch (error) {
      console.error('Error rendering panel component:', error);
      this.errorMessage = 'An error occurred while rendering the panel. Please try again.';
      this._render360PanelComponent({ showPanel: true });
    }
  }
  //#endregion

  //#region Start AVVA workflow
  private _startAVVAWFDialog(): void {
    if (
      !this.props.currentItem ||
      !this.props.listId) {
      console.error('No selected rows');
      return;
    }
    if (this.AVVAcontainer) {
      ReactDOM.unmountComponentAtNode(this.AVVAcontainer);
      document.body.removeChild(this.AVVAcontainer);
      this.AVVAcontainer = null;
    }

    this.AVVAcontainer = document.createElement('div');

    document.body.appendChild(this.AVVAcontainer);

    const onClose = () => {
      if (this.AVVAcontainer) {
        ReactDOM.unmountComponentAtNode(this.AVVAcontainer);
        document.body.removeChild(this.AVVAcontainer);
        this.AVVAcontainer = null;
      }
    };
    const defaultWorkflowUrl = 'https://prod-44.westeurope.logic.azure.com:443/workflows/b5d5e9e2f34f476fbd47a205c836d9d0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=z_4_ooxtm1cBxgazpBZ72HQ3mIeYLiQb-EtsQFj8CX0';
    const itemId = this.props.currentItem.getValueByName('ID') as number;
    const author = this.props.currentItem.getValueByName('Author');
    const createdBy: { EMail: string; Title: string; UserName: string; ID: number; } = {
      EMail: author[0].email,
      Title: author[0].title,
      UserName: author[0].sip,
      ID: author[0].id
    }
    const serverUrl = this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, '');
    const element = React.createElement(WorkflowDialogContent, {
      onClose,
      context: this.props.context,
      itemID: itemId,
      libraryID: this.props.listId,
      siteURL: this.props.context.pageContext.web.absoluteUrl,
      currentUser: this.props.context.pageContext.user.displayName,
      currentUserEmail: this.props.context.pageContext.user.email,
      projReference: this.props.currentItem.getValueByName(FieldNames.ProjectReference),
      projRevision: this.props.currentItem.getValueByName(FieldNames.Revision),
      docTitle: this.props.currentItem.getValueByName('Title'),
      itemLink: `${serverUrl}/${this.props.currentItem.getValueByName('FileRef')}`,
      docGUId: this.props.currentItem.getValueByName('_dlc_DocIdUrl.desc'),// Passing the URL as a property
      createdBy: createdBy,
      workflowUrl: this.props.AVVAWorkflowUrl || defaultWorkflowUrl,
      notificationService: this.props.notificationService,
    });
    ReactDOM.render(element, this.AVVAcontainer);

  }
  //#endregion

  //#region Start Proofreading
  private _startProofreadingDialog(): void {
    if (this.proofreadingcontainer) {
      ReactDOM.unmountComponentAtNode(this.proofreadingcontainer);
      document.body.removeChild(this.proofreadingcontainer);
      this.proofreadingcontainer = null;
    }
    this.proofreadingcontainer = document.createElement('div');
    document.body.appendChild(this.proofreadingcontainer);

    const onClose = () => {
      if (this.proofreadingcontainer) {
        ReactDOM.unmountComponentAtNode(this.proofreadingcontainer);
        document.body.removeChild(this.proofreadingcontainer);
        this.proofreadingcontainer = null;
        console.log("after raise change");
      }
    };

    const defaultWorkflowUrl = "https://prod-104.westeurope.logic.azure.com:443/workflows/d30aa75d1dfc40f18d4e51544189d2bf/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oJi2TWhsh1Syjh69gTmyImGuVZFxNVkJgaEKKoH3xVc";
    if (
      !this.props.currentItem ||
      !this.props.listId) {
      console.error('No selected rows');
      return;
    }
    const serverUrl = this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, '');
    const itemId = this.props.currentItem.getValueByName('ID') as number;
    const author = this.props.currentItem.getValueByName('Author');
    const createdBy: { EMail: string; Title: string; UserName: string; ID: number; } = {
      EMail: author[0].email,
      Title: author[0].title,
      UserName: author[0].sip,
      ID: author[0].id
    }

    const element = React.createElement(ProofreadingDialogContent, {
      onClose,
      context: this.props.context,
      itemID: itemId,
      libraryID: this.props.listId,
      siteURL: this.props.context.pageContext.web.absoluteUrl,
      currentUser: this.props.context.pageContext.user.displayName,
      currentUserEmail: this.props.context.pageContext.user.email,
      docGUId: this.props.currentItem.getValueByName('_dlc_DocIdUrl.desc'),
      createdBy: createdBy,
      projReference: this.props.currentItem.getValueByName(FieldNames.ProjectReference),
      projRevision: this.props.currentItem.getValueByName(FieldNames.Revision),
      docTitle: this.props.currentItem.getValueByName('Title'),
      itemLink: `${serverUrl}/${this.props.currentItem.getValueByName('FileRef')}`,
      workflowUrl: this.props.proofreadingWorkflowUrl || defaultWorkflowUrl,
      notificationService: this.props.notificationService
    });
    ReactDOM.render(element, this.proofreadingcontainer);

  }
  //#endregion
}
