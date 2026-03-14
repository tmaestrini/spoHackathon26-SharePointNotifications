
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { NotificationRegistration } from '../models/NotificationRegistration';
import { IConfiguration } from '../models/Configuration';
import { IUserInfo } from '@spteck/react-controls-v2';

type BackendAPIServiceResponse = {
    status: number;
    result: 'success' | 'error';
    message: string;
}

export interface IBackendAPIService {
    deleteRegistration(id: string): Promise<BackendAPIServiceResponse>;
    loadRegistrations(): Promise<NotificationRegistration[]>;
    createRegistration(registration: NotificationRegistration): Promise<BackendAPIServiceResponse>;
}


export default class BackendAPIService implements IBackendAPIService {
    private static instance: BackendAPIService;
    private context: ListViewCommandSetContext | undefined = undefined;
    private userContext: IUserInfo | undefined = undefined;
    private AZURE_FUNCTION_CLIENT_ID: string | undefined = undefined;
    private AZURE_FUNCTION_BASE_URL: string | undefined = undefined;
    private AZURE_FUNCTION_KEY: string | undefined = undefined;

    private constructor() { }

    public static init(context: ListViewCommandSetContext, configuration: IConfiguration, userContext?: IUserInfo): BackendAPIService {
        const instance = this.getInstance();
        instance.context = context;
        instance.userContext = userContext;
        instance.AZURE_FUNCTION_CLIENT_ID = configuration.AZURE_FUNCTION_CLIENT_ID;
        instance.AZURE_FUNCTION_BASE_URL = configuration.AZURE_FUNCTION_BASE_URL;
        instance.AZURE_FUNCTION_KEY = configuration.AZURE_FUNCTION_KEY;

        return instance;
    }

    private static getInstance(): BackendAPIService {
        if (!BackendAPIService.instance) {
            BackendAPIService.instance = new BackendAPIService();
        }
        return BackendAPIService.instance;
    }

    private async query(endpoint: string, method: 'POST' | 'GET' | 'PUT' | 'DELETE', body?: any): Promise<Response> {

        const response = await fetch(this.AZURE_FUNCTION_BASE_URL + endpoint, {
            method,
            headers: {
                'Content-Type': 'application/json',
                'x-functions-key': this.AZURE_FUNCTION_KEY ?? '',
            },
            body: body && JSON.stringify(body)
        });

        if (!response || !response.ok) {
            console.error('API request failed:', response);
            return Promise.reject('API request failed');
        }

        return response;
    }

    public async deleteRegistration(id: string): Promise<BackendAPIServiceResponse> {
        try {
            await this.query(`registrations/${this.userContext?.userId}/${encodeURIComponent(id)}`, 'DELETE');
            return Promise.resolve({
                status: 200,
                result: 'success',
                message: 'Notification registration deleted successfully.'
            });
        } catch (error) {
            console.error('Error deleting notification registration:', error);
            return Promise.reject({
                status: 500,
                result: 'error',
                message: `An error occurred while deleting the notification registration: (${error})`
            });
        }
    }

    public async loadRegistrations(): Promise<NotificationRegistration[]> {
        try {
            const response = await this.query(`registrations/${this.userContext?.userId}`, 'GET');
            const data: [NotificationRegistration] = await response.json();
            const registrationsForThisSite = data.filter(item => item.siteUrl === this.context?.pageContext.site.absoluteUrl);
console.log(registrationsForThisSite);
            console.log('API response for loading registrations:', registrationsForThisSite);
            return Promise.resolve(registrationsForThisSite);
        } catch (error) {
            console.error('Error loading notification registrations:', error);
            return Promise.reject({
                status: 500,
                result: 'error',
                message: `An error occurred while loading notification registrations: (${error})`
            });
        }
    }

    public async createRegistration(registration: NotificationRegistration): Promise<BackendAPIServiceResponse> {
        try {
            console.log('Creating notification registration with data:', registration);
            const response = await this.query('registrations', 'POST', registration);

            if (!response) {
                throw new Error('Failed to get response from API.');
            }

            console.log('API response:', await response.json());

            return Promise.resolve({
                status: response.status,
                result: 'success',
                message: 'Notification settings saved successfully.'
            });
        } catch (error) {
            console.error('Error saving notification settings:', error);
            return Promise.reject({
                status: 500,
                result: 'error',
                message: `An error occurred while saving notification settings: (${error})`
            });
        }
    }
}