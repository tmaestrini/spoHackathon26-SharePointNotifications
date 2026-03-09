import { InputField } from '@spteck/react-controls-v2';
import * as React from 'react'; import { Dropdown, Option, OptionOnSelectData, Radio, RadioGroup, RadioGroupOnChangeData, SelectionEvents } from '@fluentui/react-components';
import ConfigItem from './ConfigItem';
import { ChangeType, NotificationChannel } from '../models/NotificationRegistration';
import { useNotificationContext } from '../context/NotificationSettingsContext';

const NotificationSettings: React.FC = () => {
    const { changeSetting } = useNotificationContext();

    return (
        <>
            < ConfigItem title="Alert Title"
                label="Enter the title for this alert. This is included in the subject of the notification sent for this alert." >
                <InputField label="" placeholder="Set the title of the notification"
                    // onChange={(value: string | number) => setTitle(value)} />
                    onChange={(value: string | number) => changeSetting({ title: value })} />
            </ConfigItem >

            <ConfigItem title="Send Alerts To"
                label="You can enter user names or email addresses. Separate them with semicolons">
                <InputField label="" placeholder="Set the email addresses of the recipients"
                    onChange={(value: string | number) => changeSetting({ recipientAddress: value.toString() })} />
            </ConfigItem>

            <ConfigItem title="Delivery Method"
                label="Specify how you want the alerts delivered.">
                <Dropdown
                    multiselect
                    placeholder="Select the channel"
                    onOptionSelect={(_: SelectionEvents, data: OptionOnSelectData) => changeSetting({ deliveryMethod: [data.optionValue as NotificationChannel] })}>
                    <Option key={NotificationChannel.Email}>Email</Option>
                    <Option key={NotificationChannel.Teams}>Microsoft Teams (Chat)</Option>
                    {/* <Option key={NotificationChannel.TeamsChannel}>Microsoft Teams Channel</Option> */}
                </Dropdown>
            </ConfigItem>

            <ConfigItem title="Change Type"
                label="Specify the type of changes that you want to be alerted to.">
                <RadioGroup
                    onChange={(_: React.FormEvent, data: RadioGroupOnChangeData) => changeSetting({ changeType: data.value as ChangeType })}>
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