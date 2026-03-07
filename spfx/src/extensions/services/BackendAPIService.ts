
import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { NotificationRegistration } from '../models/NotificationRegistration';

type BackendAPIServiceResponse = {
    status: number;
    result: 'success' | 'error';
    message: string;
}

export default class BackendAPIService {
    private static instance: BackendAPIService;
    private aadHttpClientFactory: AadHttpClientFactory | undefined = undefined;
    private client: AadHttpClient | undefined = undefined;

    private AZURE_FUNCTION_BASE_URL = "https://spo-notifications-api.azurewebsites.net/api/";
    private AZURE_FUNCTION_ClIENT_ID = "d1b0cbb7-8c9e-4a3b-9cbd-1c3fddbfaa5b"; // client ID of the Azure Function app registration Entra ID

    private constructor() { }

    public static init(baseUrl: string, context: ListViewCommandSetContext): BackendAPIService {
        const instance = this.getInstance();
        instance.AZURE_FUNCTION_BASE_URL = baseUrl;
        instance.aadHttpClientFactory = context.aadHttpClientFactory;
        // TODO: Set the client ID from admin configuration

        return instance;
    }

    private static getInstance(): BackendAPIService {
        if (!BackendAPIService.instance) {
            BackendAPIService.instance = new BackendAPIService();
        }
        return BackendAPIService.instance;
    }

    private async query(endpoint: string, method: 'POST' | 'GET' | 'PUT' | 'DELETE', body?: any): Promise<HttpClientResponse> {
        if (!this.client) {
            this.client = await this.aadHttpClientFactory?.getClient(this.AZURE_FUNCTION_ClIENT_ID);
        }

        const response = await this.client?.fetch(this.AZURE_FUNCTION_BASE_URL + endpoint, AadHttpClient.configurations.v1, {
            method,
            headers: {
                'Content-Type': 'application/json'
            },
            body: body && JSON.stringify(body)
        });

        if(!response || !response.ok) {
            console.error('API request failed:', response);
            return Promise.reject('API request failed');
        }

        return response;
    }

    public async createRegistration(registration: NotificationRegistration): Promise<BackendAPIServiceResponse> {
        try {
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
            return {
                status: 500,
                result: 'error',
                message: `An error occurred while saving notification settings: (${error})`
            };
        }
    }
}