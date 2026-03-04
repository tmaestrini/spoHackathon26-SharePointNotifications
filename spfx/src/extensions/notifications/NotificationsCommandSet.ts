import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';
import {
  BaseListViewCommandSet,
  Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import SPONotificationDialog from '../components/SPONotificationDialog';
import SPONotification from '../components/SPONotification';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INotificationsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'NotificationsCommandSet';

export default class NotificationsCommandSet extends BaseListViewCommandSet<INotificationsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized NotificationsCommandSet');
    this.handleCommands();

    return Promise.resolve();
  }

  private handleCommands(): void {
    const notificationsCommand: Command = this.tryGetCommand('COMMAND_Notifications');
    if (notificationsCommand) {
      notificationsCommand.visible = true;
    }
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_Notifications': {
        const dialog = new SPONotificationDialog(this.context);
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
