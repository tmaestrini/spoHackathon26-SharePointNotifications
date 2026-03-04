import * as React from 'react';
import { SPFxContextAdapter } from '@spteck/react-controls-v2-spfx-adapter';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { UniversalProvider } from '@spteck/react-controls-v2';
import { FluentProvider, webLightTheme, webDarkTheme, IdPrefixProvider } from '@fluentui/react-components';
import SPONotification from './SPONotification';
import ReactDOM from 'react-dom';

export default class SPONotificationDialog {
    private container: HTMLDivElement;
    private context: ListViewCommandSetContext;

    constructor(context: ListViewCommandSetContext) {
        this.context = context;
        this.container = document.createElement('div');
    }

    public show(): Promise<void> {
        document.body.appendChild(this.container);
        const appContext = SPFxContextAdapter.adapt(this.context, 'SPONotificationDialog');
        const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        
        ReactDOM.render(
            <IdPrefixProvider value="spo-notification-">
                <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
                    <UniversalProvider context={appContext as any}>
                        <SPONotification onClose={() => this.close()} context={this.context} />
                    </UniversalProvider>
                </FluentProvider>
            </IdPrefixProvider>,
            this.container
        );
        return Promise.resolve();
    }

    public close(): void {
        ReactDOM.unmountComponentAtNode(this.container);
        this.container.remove();
    }
}

