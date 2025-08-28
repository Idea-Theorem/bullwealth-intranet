import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import DocumentLibrary from './components/DocumentLibrary';
import { IDocumentLibraryProps } from './components/IDocumentLibraryProps';

export interface IDocumentLibraryWebPartProps {
  title: string;
  description: string;
  libraryName: string;
  showUploadButton: boolean;
  showActions: boolean;
  pageSize: number;
}

export default class DocumentLibraryWebPart extends BaseClientSideWebPart<IDocumentLibraryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDocumentLibraryProps> = React.createElement(
      DocumentLibrary,
      {
        title: this.properties.title || 'Compliance',
        description: this.properties.description || 'Below are documents related to Compliance.',
        libraryName: this.properties.libraryName || 'Policies',
        showUploadButton: this.properties.showUploadButton !== false,
        showActions: this.properties.showActions !== false,
        pageSize: this.properties.pageSize || 10,
        context: this.context,
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
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = 'Office';
              break;
            case 'Outlook':
              environmentMessage = 'Outlook';
              break;
            case 'Teams':
              environmentMessage = 'Teams';
              break;
            default:
              environmentMessage = 'SharePoint';
          }
          return environmentMessage;
        });
    }

    return Promise.resolve('SharePoint');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
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
            description: 'Configure Document Library Settings'
          },
          groups: [
            {
              groupName: 'Basic Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title'
                }),
                PropertyPaneTextField('description', {
                  label: 'Description',
                  multiline: true,
                  rows: 2
                }),
                PropertyPaneTextField('libraryName', {
                  label: 'Library Name',
                  description: 'Name of the document library to display'
                }),
                PropertyPaneSlider('pageSize', {
                  label: 'Items per page',
                  min: 5,
                  max: 50,
                  value: 10,
                  showValue: true,
                  step: 5
                }),
                PropertyPaneToggle('showUploadButton', {
                  label: 'Show Upload Button',
                  checked: true,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('showActions', {
                  label: 'Show Actions Menu',
                  checked: true,
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}