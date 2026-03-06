import * as React from 'react';
import { NotificationChannel, ChangeType } from '../models/NotificationRegistration';


export type NotificationSettings = {
    title?: string | number;
    recipientAddress?: string;
    deliveryMethod?: NotificationChannel[];
    changeType?: ChangeType;
}

export interface INotificationContext {
    changeSetting: (setting: Partial<NotificationSettings>) => void;
    getSettings: () => NotificationSettings;
}

const NotificationContext = React.createContext<INotificationContext | undefined>(undefined);

export const NotificationProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [notificationSettings, setNotificationSettings] = React.useState<NotificationSettings>({});

    const changeSetting = (setting: Partial<NotificationSettings>) => {
        setNotificationSettings(prev => ({ ...prev, ...setting }));
    };

    const getSettings = (): NotificationSettings => {
        return notificationSettings;
    };

    return (
        <NotificationContext.Provider value={{ changeSetting, getSettings }}>
            {children}
        </NotificationContext.Provider>
    );
};

export const useNotificationContext = (): INotificationContext => {
    const context = React.useContext(NotificationContext);
    if (!context) {
        throw new Error('useNotificationContext must be used within a NotificationProvider');
    }
    return context;
};