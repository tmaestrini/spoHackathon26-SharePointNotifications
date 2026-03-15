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
import { NotificationRegistration } from "../models/NotificationRegistration";
import { useNotificationContext } from "../context/NotificationSettingsContext";


type NotificationItem = Required<Pick<NotificationRegistration, 'id' | 'description' | 'notificationChannels' | 'changeType'>>;

const NotificationRegistrations: React.FC = () => {
    const { backendService } = useNotificationContext();

    const [isLoading, setIsLoading] = React.useState<boolean>(true);
    const [registrations, setRegistrations] = React.useState<NotificationItem[]>([]);

    useEffect(() => {
        loadRegistrations().catch(error => {
            console.error('Error loading registrations:', error);
            setIsLoading(false);
        });
    }, []);

    async function loadRegistrations(): Promise<void> {
        try {
            const registrationData: NotificationRegistration[] = await backendService.loadRegistrations();
            console.log('Loaded registrations from backend API:', registrationData);
            setRegistrations(registrationData.map(item => ({
                changeType: item.changeType,
                description: item.description ?? "",
                id: item.id ?? "",
                notificationChannels: item.notificationChannels,
            })));
            setIsLoading(false);
        } catch (error) {
            console.error('Failed to load notification registrations:', error);
            setIsLoading(false);
        }
    }

    async function deleteRegistration(id: string): Promise<void> {
        try {
            await backendService.deleteRegistration(id);
            console.log(`Registration with id ${id} deleted successfully.`)
            // locally update the list of registrations after deletion due to performance reasons
            setRegistrations(prev => prev.filter(reg => reg.id !== id));
            // Refresh the list of registrations from the backend - to ensure consistency
            loadRegistrations().catch(error => {
                console.error('Error reloading registrations after deletion:', error);
            });

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
                                <TableHeaderCell>{item.notificationChannels?.join(', ')}</TableHeaderCell>
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