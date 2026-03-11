import React, { CSSProperties } from "react";
import {
    DeleteRegular,
} from "@fluentui/react-icons";
import {
    Badge,
    Button,
    Table,
    TableBody,
    TableHeader,
    TableHeaderCell,
    TableRow
} from "@fluentui/react-components";
import { ChangeType, NotificationChannel, NotificationRegistration } from "../models/NotificationRegistration";

type NotificationItem = Required<Pick<NotificationRegistration, 'id' | 'description' | 'notificationChannel' | 'changeType'>>;

const NotificationRegistrations: React.FC = () => {

    const items: NotificationItem[] = [
        {
            id: '1',
            description: "Notification for item created",
            notificationChannel: [NotificationChannel.Email],
            changeType: ChangeType.CREATED
        },
        {
            'id': '2',
            description: "Notification for item updated",
            notificationChannel: [NotificationChannel.Teams],
            changeType: ChangeType.UPDATED
        },
        {
            id: '3',
            description: "Notification for item deleted",
            notificationChannel: [NotificationChannel.Teams],
            changeType: ChangeType.DELETED
        },
        {
            id: '4',
            description: "Notification for all changes",
            notificationChannel: [NotificationChannel.Email, NotificationChannel.Teams],
            changeType: ChangeType.ALL
        }
    ];

    const columns: { columnId: string, label: string, style?: CSSProperties }[] = [
        { columnId: 'description', label: 'Title', style: { fontWeight: 600 } },
        { columnId: 'changeType', label: 'Change Type', style: { fontWeight: 600, width: '180px' } },
        { columnId: 'notificationChannel', label: 'Delivery Method', style: { fontWeight: 600, width: '180px' } },
        { columnId: 'actions', label: 'Actions', style: { fontWeight: 600, width: '80px' } },
    ];

    return (
        <Table size="small">
            <TableHeader>
                <TableRow>
                    {columns.map(column => (
                        <TableHeaderCell key={column.columnId} style={column.style}>{column.label}</TableHeaderCell>
                    ))}
                </TableRow>
            </TableHeader>
            <TableBody>
                {items.map((item, index) => (
                    <TableRow key={index}>
                        <TableHeaderCell>{item.description}</TableHeaderCell>
                        <TableHeaderCell><Badge appearance="outline" color="informative" shape="rounded">{item.changeType}</Badge></TableHeaderCell>
                        <TableHeaderCell>{item.notificationChannel.join(', ')}</TableHeaderCell>
                        <TableHeaderCell>
                            <Button appearance="subtle" size="small" icon={<DeleteRegular />}
                            onClick={() => {}}
                            />
                        </TableHeaderCell>
                    </TableRow>
                ))}
            </TableBody>
        </Table>
    );
}

export default NotificationRegistrations;