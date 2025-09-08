import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DocumentLibraryWebPartStrings';
import DocumentLibrary from './components/DocumentLibrary';
import { IDocumentLibraryProps } from './components/IDocumentLibraryProps';

export interface IDocumentLibraryWebPartProps {
  title: string;
  description: string;
  listName: string;
  itemsPerPage: number;
}

export default class DocumentLibraryWebPart extends BaseClientSideWebPart<IDocumentLibraryWebPartProps> {


  public render(): void {
  const element: React.ReactElement<IDocumentLibraryProps> = React.createElement(
    DocumentLibrary,
    {
      title: this.properties.title || 'Document Library',
      listName: this.properties.listName || 'Documents/Compliance',
      context: this.context
    }
  );

  ReactDom.render(element, this.domElement);
}

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

    const { semanticColors } = currentTheme;

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
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: 'Compliance'
                }),
                PropertyPaneTextField('description', {
                  label: 'Description',
                  value: 'Below are documents related to Compliance.'
                }),
                PropertyPaneTextField('listName', {
                  label: 'Document Library Name',
                  value: 'Documents',
                  description: 'Enter the name of the SharePoint document library'
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: 'Items per page',
                  min: 5,
                  max: 50,
                  value: 10,
                  showValue: true,
                  step: 5
                })
              ]
            }
          ]
        }
      ]
    };
  }
}