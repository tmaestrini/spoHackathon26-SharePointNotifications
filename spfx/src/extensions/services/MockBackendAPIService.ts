
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { NotificationRegistration } from '../models/NotificationRegistration';
import { IConfiguration } from '../models/Configuration';
import { IBackendAPIService } from './BackendAPIService';
import { Guid } from '@microsoft/sp-core-library';

type BackendAPIServiceResponse = {
    status: number;
    result: 'success' | 'error';
    message: string;
}

export default class MockBackendAPIService implements IBackendAPIService {
    private static instance: MockBackendAPIService;
    private registrations: NotificationRegistration[] = [];

    private constructor() {  }

    public static init(context: ListViewCommandSetContext, configuration: IConfiguration): MockBackendAPIService {
        const instance = this.getInstance();

        return instance;
    }

    private static getInstance(): MockBackendAPIService {
        if (!MockBackendAPIService.instance) {
            MockBackendAPIService.instance = new MockBackendAPIService();
        }
        return MockBackendAPIService.instance;
    }

    private randomDelay(): Promise<void> {
        return new Promise(resolve => setTimeout(resolve, Math.random() * 1000));
    }

    public async deleteRegistration(id: string): Promise<BackendAPIServiceResponse> {
        await this.randomDelay();

        if (this.registrations.find(reg => reg.id === id)) {
            this.registrations = this.registrations.filter(reg => reg.id !== id);
            return Promise.resolve({
                status: 200,
                result: 'success',
                message: 'Notification registration deleted successfully.'
            });
        } else {
            console.error('Error deleting notification registration: Registration not found');
            return Promise.reject({
                status: 500,
                result: 'error',
                message: `An error occurred while deleting the notification registration: Registration not found`
            });
        }
    }

    public async loadRegistrations(): Promise<NotificationRegistration[]> {
        await this.randomDelay();
        console.log('Loading notification registrations...');
        return Promise.resolve(this.registrations);
    }

    public async createRegistration(registration: NotificationRegistration): Promise<BackendAPIServiceResponse> {
        await this.randomDelay();
        console.log('Creating notification registration with data:', registration);
        this.registrations.push({...registration, id: Guid.newGuid().toString()});
        return Promise.resolve({
            status: 200,
            result: 'success',
            message: 'Notification registration created successfully.'
        });
    }
}