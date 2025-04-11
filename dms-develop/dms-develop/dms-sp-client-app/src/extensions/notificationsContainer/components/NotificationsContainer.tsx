import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './NotificationsContainer.module.scss';
import strings from 'NotificationsContainerApplicationCustomizerStrings';
import { IAddOperationEventArgs, INotificationService, INotifyEventArgs, IOperationCompletedEventArgs } from '../../../service/NotificationService';
import { ActionButton, Dialog, DialogFooter, DialogType, Icon, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';

export interface INotificationsContainerProps {
    notificationService: INotificationService
}

const LOG_SOURCE: string = 'NotificationsContainer';
const dialogStyles = { main: { maxWidth: 450 } };

export default class NotificationsContainer extends React.Component<INotificationsContainerProps, {
    notification?: INotifyEventArgs,
    operations: IAddOperationEventArgs[],
    showDetails?: boolean
}> {

    constructor(props: INotificationsContainerProps) {
        super(props);
        this.state = {
            notification: undefined,
            operations: []
        };
    }
    public componentDidMount(): void {
        const { notificationService } = this.props;
        Log.info(LOG_SOURCE, 'React Element: NotificationsContainer mounted');
        document.addEventListener(notificationService.notifyEventType, this._handleNotify.bind(this));
        document.addEventListener(notificationService.operationCompletedEventType, this._handleOperationCompleted.bind(this));
        document.addEventListener(notificationService.addOperationEventType, this._handleAddOperation.bind(this));
    }

    public componentWillUnmount(): void {
        Log.info(LOG_SOURCE, 'React Element: NotificationsContainer unmounted');
    }

    public render(): React.ReactElement<{}> {
        const { notification, operations, showDetails } = this.state;
        if (notification?.autoClose) {
            setTimeout(() => this.setState({ notification: undefined }), 3000);
        }
        const lastOperation = operations[operations.length - 1];
        return (
            <div className={styles.notificationsContainer}>
                {notification &&
                    <Dialog
                        hidden={false}
                        onDismiss={() => this.setState({ notification: undefined })}
                        dialogContentProps={{
                            type: DialogType.largeHeader,
                            title: notification.title
                        }}
                        modalProps={{
                            className: styles.notificationDialog,
                        }}
                    >
                        <div className={styles.notificationsContainer}>
                            <div className={styles.notificationMessage}>
                                <span >{notification.message}</span>
                                <Icon iconName={notification.notificationType === 'success' ? 'Accept' : 'ErrorBadge'}
                                    className={notification.notificationType === 'success' ? styles.success : styles.error} />
                            </div>
                            {notification.details && notification.details.length > 0 &&
                                <div className={styles.notificationDetails}>
                                    <ActionButton iconProps={{ iconName: showDetails ? 'ChevronDown' : 'ChevronRight' }}
                                        onClick={() => this.setState({ showDetails: !showDetails })} >
                                        {strings.ShowDetails}
                                    </ActionButton>
                                    {showDetails &&
                                        <div className={styles.detailsContainer}>
                                            {notification.details.length === 1 && <span>{notification.details[0]}</span>}
                                            {notification.details.length > 1 &&
                                                <ul>
                                                    {notification.details.map((detail, index) => <li key={index}>{detail}</li>)}
                                                </ul>
                                            }
                                        </div>
                                    }
                                </div>
                            }
                        </div>
                        {!notification?.autoClose &&
                            <DialogFooter>
                                <PrimaryButton onClick={() => this.setState({ notification: undefined })} text={strings.CloseNotification} />
                            </DialogFooter>
                        }
                    </Dialog>
                }
                {lastOperation &&
                    <Dialog
                        hidden={false}
                        dialogContentProps={{
                            type: DialogType.largeHeader,
                            title: lastOperation.operation,
                        }}
                        modalProps={{
                            styles: dialogStyles,
                            isBlocking: true
                        }}
                    >
                        <div className={styles.operationMessage}>
                            <Spinner label={strings.OperationInProgress} labelPosition="top" size={SpinnerSize.large} />
                        </div>
                    </Dialog>
                }
            </div>
        );
    }
    private _handleNotify(event: CustomEvent<INotifyEventArgs>): void {
        this.setState({
            notification: event.detail
        });
    }
    private _handleOperationCompleted(event: CustomEvent<IOperationCompletedEventArgs>): void {
        const operation = this.state.operations.find(operation => operation.operationId === event.detail.operationId);
        if (operation) {
            this.setState({
                operations: this.state.operations.filter(operation => operation.operationId !== event.detail.operationId),
                notification: event.detail.isSuccess !== undefined ? {
                    message: `${event.detail.isSuccess ? strings.OperationCompletedSuccessfully : strings.OperationFailed}`,
                    notificationType: event.detail.isSuccess ? 'success' : 'error',
                    autoClose: event.detail.isSuccess ? true : false,
                    title: operation.operation,
                    details: event.detail.details
                } : undefined
            });
        }
    }
    private _handleAddOperation(event: CustomEvent<IAddOperationEventArgs>): void {
        this.setState({
            operations: [...this.state.operations, event.detail]
        });
    }
}