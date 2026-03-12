import { RenderDialog, StackV2, TypographyControl, useApplicationContext } from '@spteck/react-controls-v2';
import * as React from 'react'; import {
    Warning24Regular, Info24Regular,
    DismissCircle24Regular,
    Save24Regular,
} from '@fluentui/react-icons';
import { Button, makeStyles, SelectTabData, Tab, TabList, TabValue, tokens } from '@fluentui/react-components';
import NotificationSettings from './NotificationSettings';
import { useNotificationContext } from '../context/NotificationSettingsContext';
import NotificationRegistrations from './NotificationRegistrations';


const useStyles = makeStyles({
    panels: {
        padding: "0 10px",
    },
});

enum Tabs {
    Settings = 'settings',
    Alerts = 'alerts'
}

export interface ISPONotificationProps {
    onClose: () => void;
}

const SPONotification: React.FC<ISPONotificationProps> = ({ onClose }) => {
    const context = useApplicationContext();
    const { registration, backendService } = useNotificationContext();

    const [dialogOpen, setDialogOpen] = React.useState(true);
    const [selectedTab, setSelectedTab] = React.useState<TabValue>(Tabs.Settings);
    const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

    const onSave = async (): Promise<void> => {
        try {
            await backendService.createRegistration(registration)
            setDialogOpen(false);
            onClose();
        } catch (error) {
            setErrorMessage(error instanceof Error ? error.message : 'Failed to save notification settings');
            console.error('Failed to save notification settings:', error);
        }
    }

    const onCloseDialog = (): void => {
        setDialogOpen(false);
        onClose();
    }

    const DialogActions: React.FC = () => {
        return (
            <StackV2 paddingTop="m" direction="horizontal" gap="s"
                justifyContent="flex-end" 
                style={{ width: '100%', position: 'absolute', bottom: '24px', right: '24px'}}>
                {selectedTab === Tabs.Settings &&
                    <Button appearance="primary" icon={<Save24Regular />}
                        onClick={onSave}>Save</Button>
                }
                <Button appearance="secondary" icon={<DismissCircle24Regular />}
                    onClick={onCloseDialog}>Close</Button>
            </StackV2>
        )
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
                <DialogActions />
            }
            onDismiss={() => { setDialogOpen(false); onClose(); }}
        >
            <StackV2 direction="vertical" gap="l" style={{ maxHeight: 'calc(700px - 160px)', overflowY: 'auto' }}>
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
                    <Tab value={Tabs.Settings}>Notification Settings</Tab>
                    <Tab value={Tabs.Alerts}>Alerts on this List</Tab>
                </TabList>

                <div className={styles.panels}>
                    {selectedTab === Tabs.Settings && <NotificationSettings />}
                    {selectedTab === Tabs.Alerts && <NotificationRegistrations />}
                </div>

            </StackV2>
        </RenderDialog >
    );
};

export default SPONotification;
