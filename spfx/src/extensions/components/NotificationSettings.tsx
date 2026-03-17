import { InputField, useApplicationContext } from '@spteck/react-controls-v2';
import * as React from 'react'; import { Dropdown, Option, OptionOnSelectData, Radio, RadioGroup, RadioGroupOnChangeData, SelectionEvents } from '@fluentui/react-components';
import ConfigItem from './ConfigItem';
import { ChangeType, NotificationChannel } from '../models/NotificationRegistration';
import { useNotificationContext } from '../context/NotificationSettingsContext';
import { PeoplePicker } from './PeoplePicker';

const NotificationSettings: React.FC = (props) => {
    const { changeSetting } = useNotificationContext();
    const context = useApplicationContext();
    return (
        <>
            < ConfigItem title="Alert Title"
                label="Enter the title for this alert. This is included in the subject of the notification sent for this alert." >
                <InputField placeholder="Set the title of the notification"
                    // onChange={(value: string | number) => setTitle(value)} />
                    required
                    onChange={(value: string | number) => changeSetting({ title: value })} />
            </ConfigItem >

            <ConfigItem title="Send Alerts To"
                label="You can enter user names or email addresses. Separate them with semicolons">
                <PeoplePicker
                    placeholder="Select recipient"
                    maxSelectedOptions={1}
                    defaultSelectedIds={context?.pageContext?.user?.userId ? [context.pageContext.user.userId] : undefined}
                    onPeopleChange={(val) => {
                        changeSetting({ recipientId: val[0] ?? null })
                    }}
                />
            </ConfigItem>

            <ConfigItem title="Delivery Method"
                label="Specify how you want the alerts delivered.">
                <Dropdown
                    multiselect
                    placeholder="Select the channel"
                    onOptionSelect={(_: SelectionEvents, data: OptionOnSelectData) => changeSetting({ deliveryMethod: data.selectedOptions as NotificationChannel[] })}>
                    <Option key={NotificationChannel.Email} value={NotificationChannel.Email}>Email</Option>
                    <Option key={NotificationChannel.Teams} value={NotificationChannel.Teams}>Microsoft Teams (Chat)</Option>
                    {/* <Option key={NotificationChannel.TeamsChannel}>Microsoft Teams Channel</Option> */}
                </Dropdown>
            </ConfigItem>

            <ConfigItem title="Change Type"
                label="Specify the type of changes that you want to be alerted to.">
                <RadioGroup
                    onChange={(_: React.FormEvent, data: RadioGroupOnChangeData) => changeSetting({ changeType: data.value as ChangeType })}>
                    <Radio defaultChecked={true} value={ChangeType.ALL} label="All changes" />
                    <Radio value={ChangeType.CREATED} label="New items are added" />
                    <Radio value={ChangeType.UPDATED} label="Existing items are modified" />
                    <Radio value={ChangeType.DELETED} label="Existing items are deleted" />
                </RadioGroup>
            </ConfigItem>
        </>
    )
}

export default NotificationSettings;