import * as React from 'react';
import { NotificationChannel, ChangeType, NotificationRegistration } from '../models/NotificationRegistration';
import { useApplicationContext } from '@spteck/react-controls-v2';
import BackendAPIService, { IBackendAPIService } from '../services/BackendAPIService';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { IConfiguration } from '../models/Configuration';
import MockBackendAPIService from '../services/MockBackendAPIService';


export type NotificationSettings = {
    title?: string | number;
    recipientId?: string;
    deliveryMethod?: NotificationChannel[];
    changeType?: ChangeType;
}

export interface INotificationSettingsContext {
    changeSetting: (setting: Partial<NotificationSettings>) => void;
    notificationSettings: NotificationSettings;
    registration: NotificationRegistration;
    backendService: IBackendAPIService;
}

const NotificationSettingsContext = React.createContext<INotificationSettingsContext | undefined>(undefined);

export const NotificationSettingsProvider: React.FC<{
    spoContext: ListViewCommandSetContext,
    configuration: IConfiguration,
    children: React.ReactNode
}> = ({ spoContext, configuration, children }) => {
    const application = useApplicationContext();
    const [notificationSettings, setNotificationSettings] = React.useState<NotificationSettings>({});

    //TODO Remove Mock
    const backendService: IBackendAPIService = MockBackendAPIService.init(
        spoContext,
        configuration
    );

    const changeSetting = (setting: Partial<NotificationSettings>): void => {
        setNotificationSettings(prev => ({ ...prev, ...setting }));
    };

    const registration: NotificationRegistration = {
        changeType: notificationSettings.changeType || ChangeType.ALL,
        userId: notificationSettings.recipientId || application?.pageContext?.aadInfo?.userId || "",
        siteId: application?.pageContext?.site?.id || "",
        siteUrl: application?.pageContext?.site?.absoluteUrl || "",
        webId: application?.pageContext?.web?.id || "",
        listId: application?.pageContext?.list?.id || "",
        notificationChannel: notificationSettings.deliveryMethod || [],
        description: notificationSettings.title?.toString()
    }

    return (
        <NotificationSettingsContext.Provider value={{ changeSetting, notificationSettings, registration, backendService }}>
            {children}
        </NotificationSettingsContext.Provider>
    );
};

export const useNotificationContext = (): INotificationSettingsContext => {
    const context = React.useContext(NotificationSettingsContext);
    if (!context) {
        throw new Error('useNotificationContext must be used within a NotificationProvider');
    }
    return context;
};