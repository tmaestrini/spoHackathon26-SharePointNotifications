import React, { CSSProperties, useEffect } from "react";
import {
    DeleteRegular,
} from "@fluentui/react-icons";
import {
    Badge,
    Button,
    Spinner,
    Table,
    TableBody,
    TableHeader,
    TableHeaderCell,
    TableRow
} from "@fluentui/react-components";
import { ChangeType, NotificationChannel, NotificationRegistration } from "../models/NotificationRegistration";
import { useNotificationContext } from "../context/NotificationSettingsContext";


type NotificationItem = Required<Pick<NotificationRegistration, 'id' | 'description' | 'notificationChannel' | 'changeType'>>;

const NotificationRegistrations: React.FC = () => {
    const { backendService } = useNotificationContext();
    
    const [isLoading, setIsLoading] = React.useState<boolean>(true);
    const [registrations, setRegistrations] = React.useState<NotificationItem[]>([]);

    useEffect(() => {
        loadRegistrations();
    }, []);
    
    const loadRegistrations = async () => {
        try {
            // TODO: call backend API to get the list of notification registrations and convert them to NotificationItem type (get the service URL from admin context)
            const registrationData: NotificationRegistration[] = await backendService.loadRegistrations();
            console.log('Loaded registrations from backend API:', registrationData);
        } catch (error) {
            console.error('Failed to load notification registrations:', error);
        }

        // SIMPLE MOCK DATA FOR TESTING - REPLACE WITH API DATA
        console.log('Loading mock notification registrations...');
        const mockData: NotificationItem[] = [
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
        setIsLoading(true);
        // TODO: replace mock data with actual data from backend API
        setRegistrations(mockData);
        // TODO: remove lazy loading simulation and set loading to false after data is loaded
        setTimeout(() => setIsLoading(false), Math.random() * 2600 + Math.random() * 400);
    };

    const deleteRegistration = async (id: string): Promise<void> => {
        try {
            setRegistrations(prev => prev.filter(reg => reg.id !== id));
        } catch (error) {
            console.error(`Failed to delete registration with id ${id}:`, error);
        }
    }

    const columns: { columnId: string, label: string, style?: CSSProperties }[] = [
        { columnId: 'description', label: 'Title', style: { fontWeight: 600 } },
        { columnId: 'changeType', label: 'Change Type', style: { fontWeight: 600, width: '180px' } },
        { columnId: 'notificationChannel', label: 'Delivery Method', style: { fontWeight: 600, width: '180px' } },
        { columnId: 'actions', label: 'Actions', style: { fontWeight: 600, width: '80px' } },
    ];

    return (
        <>
            {isLoading && <Spinner size="tiny" labelPosition="below" label="Loading registered notifications..."></Spinner>}
            {!isLoading &&
                <Table size="small">
                    <TableHeader>
                        <TableRow>
                            {columns.map(column => (
                                <TableHeaderCell key={column.columnId} style={column.style}>{column.label}</TableHeaderCell>
                            ))}
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {registrations && registrations.map((item: NotificationItem, index) => (
                            <TableRow key={index}>
                                <TableHeaderCell>{item.description}</TableHeaderCell>
                                <TableHeaderCell><Badge appearance="outline" color="informative" shape="rounded">{item.changeType}</Badge></TableHeaderCell>
                                <TableHeaderCell>{item.notificationChannel.join(', ')}</TableHeaderCell>
                                <TableHeaderCell>
                                    <Button appearance="subtle" size="small" icon={<DeleteRegular />}
                                        onClick={() => deleteRegistration(item.id)}
                                    />
                                </TableHeaderCell>
                            </TableRow>
                        ))}
                    </TableBody>
                </Table>
            }
        </>
    );
}

export default NotificationRegistrations;