import { Guid, Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './MassImportStatus.module.scss';
import {
  Callout, DetailsList, DetailsListLayoutMode, Icon, IconButton, IContextualMenuItem,
  IContextualMenuItemProps, IContextualMenuProps, TextField, TooltipHost, Text,
  SelectionMode,
  IColumn
} from '@fluentui/react';
import strings from 'MassImportStatusFieldCustomizerStrings';
import { BatchImportConstants } from '../../../constants';
import { AadHttpClient } from '@microsoft/sp-http';
import { INotificationService } from '../../../service/NotificationService';
import { IMassImportRequest } from '../../../interfaces/IMassImportRequest';
import { IImportJob } from '../../../interfaces/mass-import';
import { SPFI } from '@pnp/sp';

export interface IMassImportStatusProps {
  massImportEndpointUrl: string;
  fileName: string;
  status: string;
  dmlFilePrefix: string;
  midFilePrefix: string;
  miaFilePrefix: string;
  dmsClient: AadHttpClient;
  notificationService: INotificationService,
  fileServerRelativeUrl: string;
  siteUrl: string;
  spId: number;
  spListId: string;
  sp: SPFI;
  currentUser: string;
  maxQueueTime: number;
}

const LOG_SOURCE: string = 'MassImportStatus';

const TOOLTIP_COLS: IColumn[] = [
  { key: 'id', name: 'ID', fieldName: 'id', minWidth: 100, maxWidth: 200, isResizable: true },
  { key: 'status', name: 'Status', fieldName: 'status', minWidth: 50, maxWidth: 200, isResizable: true },
  { key: 'error', name: 'Error', fieldName: 'friendlyError', minWidth: 100, maxWidth: 300, isResizable: true, styles: { root: { whiteSpace: 'normal' } } },
];

export default class MassImportStatus extends React.Component<IMassImportStatusProps, {
  importJob?: IImportJob;
  statusError?: boolean;
  refreshInterval?: NodeJS.Timeout;
  id?: string;
  filter?: string;
  showDetails?: boolean;
  errorNumber: number;
  isLoading?: boolean;
  lastModifiedBy?: string;
  initialStatus?: string;
  queueExpired?: boolean;
}> {
  private _reset(): void {
    const { sp, spId, spListId, notificationService } = this.props;
    sp.web.lists.getById(spListId).items.getById(spId).update({
      [BatchImportConstants.BatchImportListFields.JobId]: null,
      [BatchImportConstants.BatchImportListFields.Status]: null,
      [BatchImportConstants.BatchImportListFields.Error]: null,
      [BatchImportConstants.BatchImportListFields.ErrorDetails]: null
    }).then(() => {
      window.location.reload();
    }).catch((error) => {
      Log.error(LOG_SOURCE, error);
      console.error(error);
      notificationService.notify(strings.UpdateJobError, strings.UpdateJobErrorMessage, 'error', true);
    });
  }

  private _onFilter(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    this.setState({ filter: newValue });
  }

  constructor(props: IMassImportStatusProps) {
    super(props);
    this.state = { errorNumber: 0 };
  }

  private _refreshStatus(initialLoad?: boolean): void {
    const { dmsClient, massImportEndpointUrl, siteUrl, sp, currentUser } = this.props;
    const { id } = this.state;
    const { importJob: currentImportJob, errorNumber, refreshInterval, initialStatus, lastModifiedBy } = this.state;
    if (id) {
      this.setState({ isLoading: initialLoad });
      dmsClient.get(`${massImportEndpointUrl}&siteUrl=${encodeURIComponent(siteUrl)}&id=${id}`, AadHttpClient.configurations.v1).then(res => {
        if (res.ok) {
          res.json().then((importJob: IImportJob) => {
            if (initialLoad && currentImportJob && currentImportJob?.status !== importJob.status && currentUser === lastModifiedBy) {
              sp.web.lists.getById(this.props.spListId).items.getById(this.props.spId).update({
                [BatchImportConstants.BatchImportListFields.Status]: importJob.status
              }).catch(err => {
                console.error(err);
                Log.error(LOG_SOURCE, err);
              });
            }
            if (currentImportJob?.status !== importJob.status ||
              currentImportJob.documents?.some((doc, index) => doc.status !== importJob.documents[index].status)
            ) {
              this.setState({ importJob, statusError: false, errorNumber: 0, isLoading: false });
            }
            if ((importJob.status === BatchImportConstants.BatchImportStatus.Completed || importJob.status === BatchImportConstants.BatchImportStatus.Failed) &&
              refreshInterval) {
              clearInterval(refreshInterval);
            }
          }).catch(err => {
            console.error(err);
            Log.error(LOG_SOURCE, err);
            this.setState({ statusError: true, errorNumber: errorNumber + 1, isLoading: false });
            if (errorNumber > 3 && refreshInterval) {
              clearInterval(refreshInterval);
            }
          });
        } else {
          if (res.status === 404 && initialStatus === BatchImportConstants.BatchImportStatus.Queued) {
            this.setState({ isLoading: false });
            return;
          }
          console.error(res.statusText);
          Log.error(LOG_SOURCE, new Error(res.statusText));
          this.setState({ statusError: true, errorNumber: errorNumber + 1 });
          if (errorNumber > 3 && refreshInterval) {
            clearInterval(refreshInterval);
          }
        }
      }).catch(err => {
        console.error(err);
        Log.error(LOG_SOURCE, err);
        this.setState({ statusError: true, errorNumber: errorNumber + 1 });
        if (errorNumber > 3 && refreshInterval) {
          clearInterval(refreshInterval);
        }
      });
    }
  }
  public componentDidMount(): void {
    const { sp, spListId, spId, maxQueueTime } = this.props;
    Log.info(LOG_SOURCE, 'React Element: MassImportStatus mounted');
    sp.web.lists.getById(spListId).items.getById(spId).select(
      BatchImportConstants.BatchImportListFields.JobId,
      BatchImportConstants.BatchImportListFields.Status,
      'Editor/EMail',
      'Modified'
    ).expand('Editor')().then((item) => {
      const minutesSinceLastModified = (new Date().getTime() - new Date(item.Modified).getTime()) / 60000;
      console.log('minutesSinceLastModified', minutesSinceLastModified);
      this.setState({
        id: item[BatchImportConstants.BatchImportListFields.JobId],
        initialStatus: item[BatchImportConstants.BatchImportListFields.Status],
        lastModifiedBy: item.Editor.EMail,
        queueExpired: minutesSinceLastModified > maxQueueTime
      }, () => this._refreshStatus(true));
    }).catch((error) => {
      Log.error(LOG_SOURCE, error);
      console.error(error);
      this.setState({ statusError: true });
    });
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: MassImportStatus unmounted');
    const { refreshInterval } = this.state;
    if (refreshInterval) {
      clearInterval(refreshInterval);
    }
  }

  public render(): React.ReactElement<{}> {
    const { refreshInterval, statusError, isLoading, initialStatus } = this.state;
    const { fileName } = this.props;
    const { Completed, Failed } = BatchImportConstants.BatchImportStatus;
    if (!fileName?.endsWith('.xlsx')) {
      return <></>;
    }

    if (initialStatus && !refreshInterval && (initialStatus !== Completed && initialStatus !== Failed)) {
      this.setState({ refreshInterval: setInterval(this._refreshStatus.bind(this), 10000) });
    }

    return (
      <div className={styles.massImportStatus}>
        {this._renderStatus()}
        {statusError && <TooltipHost
          content={strings.ErrorGettingStatus}
          id={`${fileName}-tooltip-error`}
        >
          <Icon iconName='error' className={styles.errorIcon} />
        </TooltipHost>}
        {isLoading && <Icon iconName='Sync' className={styles.loadingIcon} />}
      </div>
    );
  }

  private _renderStatus(): React.ReactElement<{}> {
    const { importJob, filter, statusError, showDetails, queueExpired } = this.state;
    const { status: spStatus, fileName, } = this.props;
    const { NotStarted, Queued, Processing, Completed, Failed, Requested } = BatchImportConstants.BatchImportStatus;
    const status = importJob?.status || spStatus;
    const labelId = `${fileName}-callout-label`;
    const buttonId = `B${Guid.newGuid().toString().replace(/-/g, '')}`;
    const menuProps: IContextualMenuProps = {
      items: this._getMenuItems(),
      contextualMenuItemAs: (props: IContextualMenuItemProps) => <div className={styles.commandBarMenuButton}>
        <div>{props.item.text}</div>
        {props.item.disabled &&
          <TooltipHost
            content={strings.DisabledTooltip}
            id={`${props.item.key}-tooltip`}
          >
            <div className={styles.tooltipIcon}>
              <Icon iconName='Unknown' />
            </div>
          </TooltipHost>
        }
      </div>,
      onItemClick: this._onItemClick?.bind(this)
    };

    const items = importJob?.documents?.map(d =>
      ({ ...d, id: d.id.replace(`${importJob.id}_`, '') })
    )?.filter(doc => !filter || doc.id.toLowerCase().includes(filter.toLowerCase()));
    switch (status) {
      case Requested:
        return <>{status}</>;
      case Queued:
      case Processing:
      case Completed:
      case Failed:
        return <>
          <div>{status}</div>
          {(status === Completed || status === Failed || (status === Queued && queueExpired)) && <IconButton iconProps={{ iconName: 'Refresh' }} onClick={this._reset.bind(this)} />}
          {(items && items?.length > 0 && !statusError) && <>
            <IconButton iconProps={{ iconName: 'BulletedListMirrored' }} id={buttonId}
              onClick={() => this.setState({ showDetails: true })} />
            {showDetails && <Callout
              className={styles.callout}
              ariaLabelledBy={labelId}
              role="dialog"
              gapSpace={0}
              target={`#${buttonId}`}
              onDismiss={() => this.setState({ showDetails: false })}
              setInitialFocus
            >
              <div className={styles.calloutContent}>
                <div className={styles.calloutHeader}>
                  <Text as="h1" block variant="mediumPlus" className={styles.calloutTitle} id={labelId}>
                    {strings.JobDetails}
                  </Text>
                  <TextField
                    className={styles.filterTextField}
                    label={strings.Filter}
                    underlined
                    onChange={this._onFilter.bind(this)}
                    value={filter}
                  />
                </div>
                <div className={styles.tableContainer}>
                  <DetailsList
                    selectionMode={SelectionMode.none}
                    items={items}
                    columns={TOOLTIP_COLS}
                    layoutMode={DetailsListLayoutMode.justified}
                  />
                </div>
              </div>
            </Callout>}
          </>
          }
        </>;
      case NotStarted:
      default:
        return <>
          <div>{strings.StartImport}</div>
          <IconButton
            iconProps={{ iconName: 'BulkUpload' }}
            menuProps={menuProps}
          />
        </>;
    }
  }
  private _getMenuItems(): IContextualMenuItem[] {
    const { dmlFilePrefix, midFilePrefix, miaFilePrefix, fileName } = this.props;
    return [
      { key: 'DML', text: strings.DML, disabled: !fileName?.startsWith(dmlFilePrefix) },
      { key: 'MIDraft', text: strings.MIDraft, disabled: !fileName?.startsWith(midFilePrefix) },
      { key: 'MIApplicable', text: strings.MIApplicable, disabled: !fileName?.startsWith(miaFilePrefix) },
    ];
  }
  private _onItemClick(_?: unknown | undefined, item?: IContextualMenuItem | undefined): boolean | void | undefined {
    const {
      massImportEndpointUrl, dmsClient,
      notificationService,
      fileServerRelativeUrl, siteUrl, spId, spListId, sp } = this.props;
    if (!item) {
      return;
    }
    let operationName = "";
    switch (item.key) {
      case 'DML':
        operationName = strings.DMLOperation;
        break;
      case 'MIDraft':
        operationName = strings.MIDraftOperation;
        break;
      case 'MIApplicable':
        operationName = strings.MIApplicableOperation;
        break;
    }
    const id = Guid.newGuid().toString();
    const notificationId = notificationService.addOngoingOperation(operationName);
    sp.web.lists.getById(spListId).items.getById(spId).update({
      [BatchImportConstants.BatchImportListFields.JobId]: id
    }).then(() => {
      dmsClient.post(massImportEndpointUrl, AadHttpClient.configurations.v1, {
        body: JSON.stringify({
          fileServerRelativeUrl,
          id,
          siteUrl,
          spId,
          spListId,
          type: item.key
        } as IMassImportRequest)
      })
        .then((result) => {
          if (result.ok) {
            sp.web.lists.getById(spListId).items.getById(spId).update({
              [BatchImportConstants.BatchImportListFields.Status]: BatchImportConstants.BatchImportStatus.Queued
            }).then(() => {
              notificationService.setOperationCompleted(notificationId, true);
              setTimeout(() => { window.location.reload() }, 1000);
            }).catch((err) => {
              console.error(err);
              Log.error(LOG_SOURCE, err);
              notificationService.setOperationCompleted(notificationId, false, strings.UpdateStatusFailed);
            });
          } else {
            sp.web.lists.getById(spListId).items.getById(spId).update({
              [BatchImportConstants.BatchImportListFields.JobId]: null
            }).catch((err) => {
              console.error(err);
              Log.error(LOG_SOURCE, err);
              notificationService.notify(strings.UpdateJobError, strings.ResetIdError, 'error', true);
            });
            notificationService.notify(operationName, result.statusText, 'error', true);
            notificationService.setOperationCompleted(notificationId, false, result.statusText);
            Log.error(LOG_SOURCE, new Error(result.statusText));
            console.error(result.statusText);
          }
        })
        .catch((error) => {
          notificationService.setOperationCompleted(notificationId, false, strings.ErrorStartingImport);
          Log.error(LOG_SOURCE, error);
          console.error(error);
        });
    })
      .catch((error) => {
        notificationService.setOperationCompleted(notificationId, false, strings.ErrorStartingImport);
        Log.error(LOG_SOURCE, error);
        console.error(error);
      });
  }

}
