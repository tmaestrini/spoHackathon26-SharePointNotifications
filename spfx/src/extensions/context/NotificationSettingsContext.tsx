import * as React from 'react';
import { NotificationChannel, ChangeType } from '../models/NotificationRegistration';


export type NotificationSettings = {
    title?: string | number;
    recipientAddress?: string;
    deliveryMethod?: NotificationChannel[];
    changeType?: ChangeType;
}

export interface INotificationSettingsContext {
    changeSetting: (setting: Partial<NotificationSettings>) => void;
    notificationSettings: NotificationSettings;
}

const NotificationSettingsContext = React.createContext<INotificationSettingsContext | undefined>(undefined);

export const NotificationSettingsProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [notificationSettings, setNotificationSettings] = React.useState<NotificationSettings>({});

    const changeSetting = (setting: Partial<NotificationSettings>): void => {
        setNotificationSettings(prev => ({ ...prev, ...setting }));
    };

    return (
        <NotificationSettingsContext.Provider value={{ changeSetting, notificationSettings }}>
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