import { RenderDialog, StackV2, TypographyControl, useApplicationContext } from '@spteck/react-controls-v2';
import * as React from 'react'; import {
    Warning24Regular, Info24Regular,
    DismissCircle24Regular,
    Save24Regular,
} from '@fluentui/react-icons';
import { Button, makeStyles, SelectTabData, Tab, TabList, TabValue, tokens } from '@fluentui/react-components';
import NotificationSettings from './NotificationSettings';
import { NotificationSettingsProvider, useNotificationContext } from '../context/NotificationSettingsContext';
import BackendAPIService from '../services/BackendAPIService';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { NotificationRegistration } from '../models/NotificationRegistration';
import { IConfiguration } from '../models/Configuration';


const useStyles = makeStyles({
    panels: {
        padding: "0 10px",
    },
});

export interface ISPONotificationProps {
    spoContext: ListViewCommandSetContext
    onClose: () => void;
    configuration: IConfiguration;
}

const SPONotification: React.FC<ISPONotificationProps> = ({ spoContext, onClose, configuration }) => {
    const context = useApplicationContext();
    const { notificationObject } = useNotificationContext();

    const [dialogOpen, setDialogOpen] = React.useState(true);
    const [selectedTab, setSelectedTab] = React.useState<TabValue>('settings');
    const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);
    
    const onSave = async (): Promise<void> => {
        // TODO: call backend API to save the settings (get the service URL from admin context)
        const backendService = BackendAPIService.init(
            spoContext,
            configuration
        );
        try {
            await backendService.createRegistration(notificationObject)
            // setDialogOpen(false);
            // onClose();
        } catch (error) {
            setErrorMessage(error instanceof Error ? error.message : 'Failed to save notification settings');
            console.error('Failed to save notification settings:', error);
        }
    }

    const styles = useStyles();
    return (
            <RenderDialog
                isOpen={dialogOpen}

                dialogTitle={
                    <StackV2 direction="horizontal" gap="s" alignItems="center">
                        <Warning24Regular style={{ color: tokens.colorPaletteYellowForeground2 }} />
                        <TypographyControl fontWeight="semibold" fontSize="l">
                            Notification settings for {context?.pageContext?.user?.displayName || 'current user'}
                        </TypographyControl>
                    </StackV2>
                }
                minWidth={'1170px'}
                maxWidth={'1170px'}
                minHeight={'700px'}
                maxHeight={'700px'}

                dialogActions={
                    <StackV2 paddingTop="m" direction="horizontal" gap="s"
                        justifyContent="flex-end" style={{ width: '100%' }}>
                        <Button appearance="secondary" icon={<DismissCircle24Regular />}
                            onClick={() => { setDialogOpen(false); onClose(); }}>Cancel</Button>
                        <Button appearance="primary" icon={<Save24Regular />}
                            onClick={onSave}>Save</Button>
                    </StackV2>
                }
                onDismiss={() => { setDialogOpen(false); onClose(); }}
            >
                <StackV2 direction="vertical" gap="l">
                    <TypographyControl>
                        Alert me when items change
                    </TypographyControl>

                    {errorMessage &&
                        <StackV2 direction="horizontal" gap="s" alignItems="center" padding="m"
                            style={{
                                borderRadius: tokens.borderRadiusMedium,
                                backgroundColor: tokens.colorNeutralBackground3,
                                border: `1px solid ${tokens.colorNeutralStroke2}`
                            }}>
                            <Info24Regular style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
                            <TypographyControl fontSize="xs" color={tokens.colorNeutralForeground3}>
                                {errorMessage}
                            </TypographyControl>
                        </StackV2>
                    }

                    <TabList defaultSelectedValue={selectedTab} onTabSelect={(_, data: SelectTabData) => {
                        setSelectedTab(data.value);
                    }}>
                        <Tab value="settings">Notification Settings</Tab>
                        <Tab value="tab2">Alerts on this List</Tab>
                    </TabList>

                    <div className={styles.panels}>
                        {selectedTab === 'settings' && <NotificationSettings />}
                    </div>

                </StackV2>
            </RenderDialog >
    );
};

export default SPONotification;
