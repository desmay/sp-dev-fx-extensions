import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'AppInsightsAdvancedApplicationCustomizerStrings';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";

const LOG_SOURCE: string = 'AppInsightsAdvancedApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppInsightsAdvancedApplicationCustomizerProperties {
    // This is an example; replace with your own property
    appInsightsKey: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppInsightsAdvancedApplicationCustomizer
    extends BaseApplicationCustomizer<IAppInsightsAdvancedApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        const browserHistory = createBrowserHistory({ basename: '' });
        var reactPlugin = new ReactPlugin();
        var appInsights = new ApplicationInsights({
            config: {
                maxBatchInterval: 0,
                connectionString:'InstrumentationKey=74134282-0275-f4b5-9d85-80e665c7ca1d;EndpointSuffix=applicationinsights.us;IngestionEndpoint=https://usgovvirginia-1.in.applicationinsights.azure.us/;AADAudience=https://monitor.azure.us/;ApplicationId=275ff6ae-ecb1-4bad-8948-32623ba15e82',
                disableFetchTracking: true,
                disableAjaxTracking: true,
                //enableAutoRouteTracking: true,
                //instrumentationKey: `${this.properties.appInsightsKey}`,
                extensions: [reactPlugin],
                extensionConfig: {
                    [reactPlugin.identifier]: { history: browserHistory }
                }
            }
        });
        appInsights.loadAppInsights();
        appInsights.trackPageView({
            name: document.title, uri: window.location.href,
            properties: {
                ["CustomProps"]: {
                    WebAbsUrl: this.context.pageContext.web.absoluteUrl,
                    WebSerUrl: this.context.pageContext.web.serverRelativeUrl,
                    WebId: this.context.pageContext.web.id,
                    UserTitle: this.context.pageContext.user.displayName,
                    UserEmail: this.context.pageContext.user.email,
                    UserLoginName: this.context.pageContext.user.loginName
                }
            }
        });        
        return Promise.resolve();
    }
}
