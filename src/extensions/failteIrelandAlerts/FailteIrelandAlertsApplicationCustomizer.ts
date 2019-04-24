import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'FailteIrelandAlertsApplicationCustomizerStrings';
import { IAlert, IAlertItem, AlertType } from './IAlert';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { ISiteData, ISiteDataResponse } from './ISiteData';
import { IAlertNotificationsProps, AlertNotifications } from './components/index';
import * as React from 'react';
import * as ReactDom from 'react-dom';

const LOG_SOURCE: string = 'FailteIrelandAlertsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFailteIrelandAlertsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class FailteIrelandAlertsApplicationCustomizer
  extends BaseApplicationCustomizer<IFailteIrelandAlertsApplicationCustomizerProperties> {

    private static _topPlaceholder?: PlaceholderContent;


  @override
  public onInit(): Promise<void> {

    // add handler to he page navigated event which occurs both when
    // the user opens and leaves the page
    this.context.application.navigatedEvent.add(this, this._render);

    return Promise.resolve();
  }//end onInit()

  private _render():void{ 

    if(!FailteIrelandAlertsApplicationCustomizer._topPlaceholder){
      FailteIrelandAlertsApplicationCustomizer._topPlaceholder = this.context.placeholderProvider
        .tryCreateContent(PlaceholderName.Top, { onDispose: this._handleDispose });
    }

    if (!FailteIrelandAlertsApplicationCustomizer._topPlaceholder) {
      return;
    }

    this
      //._getConnectedSiteData("string")
      //.then((connectedHubSiteData: ISiteData): Promise<IAlert[]> => {
      //  return this._loadUpcomingAlerts(connectedHubSiteData.url);
      //})
      .getAlerts()
      
      
      .then((upcomingAlerts: IAlert[]): void => {
        if (upcomingAlerts.length === 0) {
          console.info('No upcoming alerts found');
          return;
        }

        const element: React.ReactElement<IAlertNotificationsProps> = React.createElement(
          AlertNotifications,
          {
            alerts: upcomingAlerts
          }
        );

        // render the UI using a React component
        ReactDom.render(element, FailteIrelandAlertsApplicationCustomizer._topPlaceholder.domElement);
      });

  }//end _render


  private _handleDispose(): void {
  }//end _handleDispose

  private getAlerts(): Promise<IAlert[]> {
    return this._loadUpcomingAlerts(this.context.pageContext.web.absoluteUrl);
  }

  private _loadUpcomingAlerts(SiteUrl: string): Promise<IAlert[]> {
    return new Promise<IAlert[]>((resolve: (upcomingAlerts: IAlert[]) => void, reject: (error: any) => void): void => {
      // suppress loading metadata to minimize the amount of data sent over the network
      const headers: Headers = new Headers();
      headers.append('accept', 'application/json;odata.metadata=none');

      // current date in the ISO format used to retrieve active alerts
      const nowString: string = new Date().toISOString();

      // from the Alerts list in the hub site, load the list of upcoming alerts sorted
      // ascending by their end time so that alerts that expire first, are shown on top
      this.context.spHttpClient
        .get(`${SiteUrl}/_api/web/lists/getByTitle('Alerts')/items?$filter=PnPAlertStartDateTime le datetime'${nowString}' and PnPAlertEndDateTime ge datetime'${nowString}'&$select=PnPAlertType,PnPAlertMessage,PnPAlertMoreInformation&$orderby=PnPAlertEndDateTime`, SPHttpClient.configurations.v1, {
          headers: headers
        })
        .then((res: SPHttpClientResponse): Promise<{ value: IAlertItem[] }> => {
          return res.json();
        })
        .then((res: { value: IAlertItem[] }): void => {
          // change the alert list item to alert object
          const upcomingAlerts: IAlert[] = res.value.map(alert => {
            return {
              type: AlertType[alert.PnPAlertType],
              message: alert.PnPAlertMessage,
              moreInformationUrl: alert.PnPAlertMoreInformation ? alert.PnPAlertMoreInformation.Url : null,
              moreInformationUrlText : alert.PnPAlertMoreInformation ? alert.PnPAlertMoreInformation.Description : null

            };
          });
          resolve(upcomingAlerts);
        })
        .catch((error: any): void => {
          reject(error);
        });
    });
  }
}
