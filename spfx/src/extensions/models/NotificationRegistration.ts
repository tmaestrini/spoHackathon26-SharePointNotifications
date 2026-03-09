export type UUID = string;

export type NotificationRegistration = {
    id?: UUID;
    userId: UUID; //ENTRA ID
    changeType: ChangeType;
    siteId: UUID;
    siteUrl: string;
    webId: UUID;
    listId: UUID;
    itemId?: number;
    notificationChannel: NotificationChannel[];
    description?: string;
};

export enum ChangeType {
    CREATED = 'CREATED',
    UPDATED = 'UPDATED',
    DELETED = 'DELETED',
    ALL = 'ALL'
}

export enum NotificationChannel {
    Teams = 'TEAMS',
    TeamsChannel = 'TEAMS_CHANNEL',
    Email = 'EMAIL',
}