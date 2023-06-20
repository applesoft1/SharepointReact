import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FapWebPartStrings';
import Fap from './components/Fap';
import { IFapProps } from './components/IFapProps';

export interface IFapWebPartProps {
  description: string;
  test: boolean;
  test1: string;
}

export default class FapWebPart extends BaseClientSideWebPart<IFapWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFapProps> = React.createElement(
      Fap,
      {
        description: this.properties.description,
        test: this.properties.test,
        test1: this.context.pageContext.web.absoluteUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                }),
                PropertyPaneToggle('test', {
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('test1', {
                  label: 'Multi-line Text Field',
                  multiline: true
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  // private _getListData(): Promise<ISPLists> {
  //   return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
  //       return response.json();
  //     })
  //     .catch(() => {});
  // }
  // private _renderList(items: ISPList[]): void {
  //   let html: string = '';
  //   items.forEach((item: ISPList) => {
  //     html += `
  //   <ul class="${styles.list}">
  //     <li class="${styles.listItem}">
  //       <span class="ms-font-l">${item.Title}</span>
  //     </li>
  //   </ul>`;
  //   });
  
  //   const listContainer: Element = this.domElement.querySelector('#spListContainer');
  //   listContainer.innerHTML = html;
  // }
  // private _renderListAsync(): void {
  //   this._getListData()
  //     .then((response) => {
  //       this._renderList(response.value);
  //     })
  //     .catch(() => {});
  // }

}
