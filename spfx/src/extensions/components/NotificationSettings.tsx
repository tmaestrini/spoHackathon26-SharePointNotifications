import { InputField, StackV2, TypographyControl } from '@spteck/react-controls-v2';
import * as React from 'react'; import { Dropdown, Option, OptionOnSelectData, Radio, RadioGroup, RadioGroupOnChangeData } from '@fluentui/react-components';
import ConfigItem from './ConfigItem';
import { ChangeType, NotificationChannel } from '../models/NotificationRegistration';

const NotificationSettings: React.FC = () => {
    // TODO: Adjust the state and handlers according to our needs (config items)
    const [title, setTitle] = React.useState<string | number | undefined>(undefined);
    const [recipientAddress, setRecipientAddress] = React.useState<string | number | undefined>(undefined);
    const [deliveryMethod, setDeliveryMethod] = React.useState<string[] | undefined>(undefined);
    const [changeType, setChangeType] = React.useState<string | undefined>(undefined);

    console.log(deliveryMethod)
    return (
        <>
            {/* TODO: Adjust config items according to our needs */}
            < ConfigItem title="Alert Title"
                label="Enter the title for this alert. This is included in the subject of the notification sent for this alert." >
                <InputField label="" placeholder="Set the title of the notification"
                    onChange={(value: string | number) => setTitle(value)} />
            </ConfigItem >

            <ConfigItem title="Send Alerts To"
                label="You can enter user names or email addresses. Separate them with semicolons">
                <InputField label="" placeholder="Set the email addresses of the recipients"
                    onChange={(value: string | number) => setRecipientAddress(value)} />
            </ConfigItem>

            <ConfigItem title="Delivery Method"
                label="Specify how you want the alerts delivered.">
                <Dropdown
                    multiselect
                    placeholder="Select the channel"
                    onOptionSelect={(_: any, data: OptionOnSelectData) => setDeliveryMethod(data.selectedOptions)}>
                    <Option key={NotificationChannel.Email}>Email</Option>
                    <Option key={NotificationChannel.Teams}>Microsoft Teams (Chat)</Option>
                    {/* <Option key={NotificationChannel.TeamsChannel}>Microsoft Teams Channel</Option> */}
                </Dropdown>
            </ConfigItem>

            <ConfigItem title="Change Type"
                label="Specify the type of changes that you want to be alerted to.">
                <RadioGroup
                    onChange={(_: any, data: RadioGroupOnChangeData) => setChangeType(data.value)}>
                    <Radio value={ChangeType.ALL} label="All changes" />
                    <Radio value={ChangeType.CREATED} label="New items are added" />
                    <Radio value={ChangeType.UPDATED} label="Existing items are modified" />
                    <Radio value={ChangeType.DELETED} label="Existing items are deleted" />
                </RadioGroup>
            </ConfigItem>
        </>
    )
}

export default NotificationSettings;