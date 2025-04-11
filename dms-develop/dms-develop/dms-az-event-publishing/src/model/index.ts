export interface ISpEvent {
    //The unique identifier for the subscription resource.
    subscriptionId: string;
    //An optional string value that is passed back in the notification. message for this subscription.
    clientState: string;
    //The date and time when the subscription expires if not updated or renewed.
    expirationDateTime: string;
    //Unique identifier of the list where the subscription is registered.
    resource: string;
    tenantId: string;
    siteUrl: string;
    webId: string;
}
export interface ISpNotification {
    value: ISpEvent[]
}

export interface IDmsEvent {
    siteServerRelativeUrl: string;
    siteUrl: string;
    libraryId: string;
    itemId: number;
    user: string;
    spContentType: string;
    action: 'create' | 'update' | 'delete';
    modified: string;
    fields: { [key: string]: any };
}

export enum SPChangeType {
    Activity = 22,
    Add = 1,
    AssignmentAdd = 11,
    AssignmentDelete = 12,
    DeleteObject = 3,
    Dirty = 21,
    ListContentTypeAdd = 19,
    ListContentTypeDelete = 20,
    MemberAdd = 13,
    MemberDelete = 14,
    MoveAway = 5,
    MoveInto = 6,
    Navigation = 16,
    NoChange = 0,
    Rename = 4,
    Restore = 7,
    RoleAdd = 8,
    RoleDelete = 9,
    RoleUpdate = 10,
    ScopeAdd = 17,
    ScopeDelete = 18,
    SystemUpdate = 15,
    Update = 2
}

export interface ISPChange {
    ChangeToken: {
        StringValue: string;
    };
    ChangeType: SPChangeType;
    SiteId: string;
    Time: string;
    Editor?: string;
    EditorEmailHint?: string;
    ItemId: number
    ListId: string;
    ServerRelativeUrl?: string;
    UniqueId: string;
    WebId: string
}