import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ExternalApiConsumerWebPart.module.scss';
import * as strings from 'ExternalApiConsumerWebPartStrings';

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface IExternalApiConsumerWebPartProps {
  description: string;
}

export default class ExternalApiConsumerWebPart extends BaseClientSideWebPart<IExternalApiConsumerWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.externalApiConsumer}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Current user claims from Azure Function:</p>
            </div>
          </div>
        </div>
      </div>
      <div class="${styles.azFuncTablecontainer}">
            <table class='azFuncClaimsTable'>
            </table>
      </div>`;

    this.context.aadHttpClientFactory
      .getClient('https://<your-tenant>.onmicrosoft.com/6b347c27-f360-47ac-b4d4-af78d0da4223')
      .then((client: AadHttpClient): void => {
        client
          .get('https://<function-app-name>.azurewebsites.net/api/', AadHttpClient.configurations.v1)
          //Use /api/CurrentInfoFromSharePoint to call back to SharePoint
          //Use /api/CurrentUserFromGraph to call the Mirosoft Graph
          .then((response: HttpClientResponse): Promise<JSON> => {
            return response.json();
          })
          .then((responseJSON: JSON): void => {
            //Display the JSON in a table
            var claimsTable = this.domElement.getElementsByClassName("azFuncClaimsTable")[0];
            for (var key in responseJSON) {
              var trElement = document.createElement("tr");
              trElement.innerHTML = `<td class="${styles.azFuncCell}">${key}</td><td class="${styles.azFuncCell}">${responseJSON[key]}</td>`;
              claimsTable.appendChild(trElement);
            }
          });
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
