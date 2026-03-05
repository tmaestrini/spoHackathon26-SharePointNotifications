import { InputField, DropdownField, RenderDialog, StackV2, TypographyControl, useApplicationContext } from '@spteck/react-controls-v2';
import * as React from 'react'; import {
    Warning24Regular, Info24Regular,
    DismissCircle24Regular,
    Save24Regular,
} from '@fluentui/react-icons';
import { Button, Radio, RadioGroup, RadioGroupOnChangeData, tokens } from '@fluentui/react-components';
import ConfigItem from './ConfigItem';

export interface ISPONotificationProps {
    onClose: () => void;
}

const SPONotification: React.FC<ISPONotificationProps> = ({ onClose }) => {
    const [dialogOpen, setDialogOpen] = React.useState(true);
    const context = useApplicationContext();

    // TODO: Adjust the state and handlers according to our needs (config items)
    const [deliveryMethod, setDeliveryMethod] = React.useState<string | undefined>(undefined);
    const [changeType, setChangeType] = React.useState<string | undefined>(undefined);

    console.log(changeType, deliveryMethod);

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
                        onClick={() => { setDialogOpen(false); onClose(); }}>Save</Button>
                </StackV2>
            }
            onDismiss={() => { setDialogOpen(false); onClose(); }}
        >
            <StackV2 direction="vertical" gap="l">
                <TypographyControl>
                    Alert me when items change
                </TypographyControl>
                <StackV2 direction="horizontal" gap="s" alignItems="center" padding="m"
                    style={{
                        borderRadius: tokens.borderRadiusMedium,
                        backgroundColor: tokens.colorNeutralBackground3,
                        border: `1px solid ${tokens.colorNeutralStroke2}`
                    }}>
                    <Info24Regular style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
                    <TypographyControl fontSize="xs" color={tokens.colorNeutralForeground3}>
                        This dialog is rendered using RenderDialog from @spteck/react-controls-v2.
                    </TypographyControl>
                </StackV2>

                {/* TODO: Adjust config items according to our needs */}
                <ConfigItem title="Alert Title"
                    label="Enter the title for this alert. This is included in the subject of the notification sent for this alert.">
                    <InputField label="" placeholder="Set the title of the notification" />
                </ConfigItem>
                <ConfigItem title="Send Alerts To"
                    label="You can enter user names or email addresses. Separate them with semicolons">
                    <InputField label="" placeholder="Set the email addresses of the recipients" />
                </ConfigItem>
                <ConfigItem title="Delivery Method"
                    label="Specify how you want the alerts delivered.">
                    <DropdownField
                        label="Send me alerts by:"
                        placeholder="Select the channel"
                        required
                        options={[
                            { value: 'EMAIL', text: 'Email' },
                            { value: 'TEAMS', text: 'Microsoft Teams (Chat)' },
                            { value: 'TEAMS_CHANNEL', text: 'Microsoft Teams Channel' },
                        ]}
                        renderItem={(option: { value: string, text?: string }) => (
                            <StackV2 direction="horizontal" gap="s" alignItems="center">
                                <TypographyControl>{option.text}</TypographyControl>
                            </StackV2>
                        )}
                        onChange={(value: string) => setDeliveryMethod(value)}
                        hint="Only one delivery method can be selected at the moment."
                    />

                </ConfigItem>
                <ConfigItem title="Change Type"
                    label="Specify the type of changes that you want to be alerted to.">
                    <RadioGroup
                        onChange={(ev: React.FormEvent<HTMLDivElement>, data: RadioGroupOnChangeData) => setChangeType(data.value)}>
                        <Radio value="ALL" label="All changes" />
                        <Radio value="CREATED" label="New items are added" />
                        <Radio value="UPDATED" label="Existing items are modified" />
                        <Radio value="DELETED" label="Existing items are deleted" />
                    </RadioGroup>
                </ConfigItem>
            </StackV2>
        </RenderDialog >
    );
};

export default SPONotification;
