import { Guid, ServiceKey } from "@microsoft/sp-core-library";

export interface INotificationService {
    addOngoingOperation(operation: string): Guid;
    setOperationCompletedSilent(operationId: Guid): void;
    setOperationCompleted(operationId: Guid, isSuccess: boolean, error?: string): void;
    setOperationCompletedWithDetails(operationId: Guid, isSuccess: boolean, details: string[], error?: string): void;
    notify(title: string, message: string, notificationType: 'success' | 'error', autoClose?: boolean): void;
    notifyWithDetails(title: string, message: string, details: string[], notificationType: 'success' | 'error', autoClose?: boolean): void;
    get addOperationEventType(): string;
    get operationCompletedEventType(): string;
    get notifyEventType(): string;
}
export interface IAddOperationEventArgs {
    operationId: Guid;
    operation: string;
}
export interface IOperationCompletedEventArgs {
    operationId: Guid;
    isSuccess?: boolean;
    details?: string[];
    error?: string;
}
export interface INotifyEventArgs {
    title: string;
    message: string;
    details?: string[];
    notificationType: 'success' | 'error';
    autoClose?: boolean;
}
export class NotificationService implements INotificationService {
    get addOperationEventType(): string {
        return 'dms:operationStarted';
    }
    get operationCompletedEventType(): string {
        return 'dms:operationCompleted';
    }
    get notifyEventType(): string {
        return 'dms:notify';
    }
    public static readonly serviceKey: ServiceKey<INotificationService> =
        ServiceKey.create<INotificationService>('dms:INotificationService', NotificationService);

    public addOngoingOperation(operation: string): Guid {
        const operationId = Guid.newGuid();
        document.dispatchEvent(new CustomEvent('dms:operationStarted', {
            detail: {
                operationId,
                operation
            }
        }));
        return operationId;
    }

    public setOperationCompleted(operationId: Guid, isSuccess: boolean, error?: string): void {
        document.dispatchEvent(new CustomEvent('dms:operationCompleted', {
            detail: {
                operationId,
                isSuccess,
                error
            }
        }));
    }
    public setOperationCompletedSilent(operationId: Guid): void {
        document.dispatchEvent(new CustomEvent('dms:operationCompleted', {
            detail: { operationId }
        }));
    }
    public setOperationCompletedWithDetails(operationId: Guid, isSuccess: boolean, details: string[], error?: string): void {
        document.dispatchEvent(new CustomEvent('dms:operationCompleted', {
            detail: {
                operationId,
                isSuccess,
                details,
                error
            }
        }));
    }


    public notify(title: string, message: string, notificationType: "success" | "error", autoClose?: boolean): void {
        document.dispatchEvent(new CustomEvent('dms:notify', {
            detail: {
                title,
                message,
                notificationType,
                autoClose
            }
        }));
    }
    public notifyWithDetails(title: string, message: string, details: string[], notificationType: "success" | "error", autoClose?: boolean): void {
        document.dispatchEvent(new CustomEvent('dms:notify', {
            detail: {
                title,
                message,
                details,
                notificationType,
                autoClose
            }
        }));
    }
}