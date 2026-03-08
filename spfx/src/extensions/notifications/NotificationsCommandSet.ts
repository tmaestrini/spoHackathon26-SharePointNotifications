import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import SPONotificationDialog from '../components/SPONotificationDialog';
import { getThemeColor } from '../util/themeHelper';
import { Dialog } from '@microsoft/sp-dialog';
import { IConfiguration } from '../models/Configuration';


/**
 * The properties for the NotificationsCommandSet command set.
 * @remarks
 * The properties defined in this interface are expected to be passed in the manifest of the extension. 
 * They can be configured by the tenant admin in the app catalog.
 * For more information, see https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions
 */
export interface INotificationsCommandSetProperties {
  AZURE_FUNCTION_BASE_URL: string;
  AZURE_FUNCTION_CLIENT_ID: string;
}

const LOG_SOURCE: string = 'NotificationsCommandSet';
const COMMAND_NAME: string = 'COMMAND_Notifications';

export default class NotificationsCommandSet extends BaseListViewCommandSet<INotificationsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized NotificationsCommandSet');
    this.handleCommands();

    const command = this.tryGetCommand(COMMAND_NAME);
    if (command) {
      const fillColor = getThemeColor('themePrimary').replace('#', '%23');
      const svg = `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 2048 2048'%3E%3Cpath d='M1792 1536v128h-512q0 53-20 99t-55 82-81 55-100 20q-53 0-99-20t-82-55-55-81-20-100H256v-128h128V768q0-88 23-170t64-153 100-129 130-100 153-65 170-23q88 0 170 23t153 64 129 100 100 130 65 153 23 170v768h128zm-256 0V768q0-106-40-199t-110-162-163-110-199-41q-106 0-199 40T663 406 553 569t-41 199v768h1024zm-512 256q27 0 50-10t40-27 28-41 10-50H896q0 27 10 50t27 40 41 28 50 10z' fill='${fillColor}'%3E%3C/path%3E%3C/svg%3E`
      command.iconImageUrl = svg;
    }

    if(this.properties.AZURE_FUNCTION_BASE_URL == null || this.properties.AZURE_FUNCTION_CLIENT_ID == null) {
      console.error('Azure Function configuration is missing.');
      Dialog.alert('NotificationsCommandSet - Azure Function configuration is missing. Please check the configuration and try again.');
    }

    return Promise.resolve();
  }

  private handleCommands(): void {
    const notificationsCommand: Command = this.tryGetCommand(COMMAND_NAME);
    if (notificationsCommand) {
      notificationsCommand.visible = true;
    }
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case COMMAND_NAME: {
        const configuration: IConfiguration = {
          AZURE_FUNCTION_BASE_URL: this.properties.AZURE_FUNCTION_BASE_URL,
          AZURE_FUNCTION_CLIENT_ID: this.properties.AZURE_FUNCTION_CLIENT_ID
        }
        const dialog = new SPONotificationDialog(this.context, configuration);
        dialog.show().catch(error => console.error(error));
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    this.raiseOnChange();
  }
}
