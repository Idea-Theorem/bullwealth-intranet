import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import HRDocuments from './components/HRDocuments';
import { IHRDocumentsProps } from './components/IHRDocumentsProps';

export interface IHRDocument {
  id: string;
  title: string;
  date: string;
  iconData: string; // Base64 encoded icon or blob URL
  iconType: 'word' | 'pdf' | 'custom';
  documentUrl: string;
}

export interface IHRDocumentsWebPartProps {
  title: string;
  documents: string; // JSON string to store documents
  columnsPerRow: number;
  showDate: boolean;
  allowUpload: boolean;
}

export default class HRDocumentsWebPart extends BaseClientSideWebPart<IHRDocumentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    let documents: IHRDocument[] = [];
    
    try {
      documents = this.properties.documents ? JSON.parse(this.properties.documents) : this.getDefaultDocuments();
    } catch {
      documents = this.getDefaultDocuments();
    }

    const element: React.ReactElement<IHRDocumentsProps> = React.createElement(
      HRDocuments,
      {
        title: this.properties.title || 'Common HR Documents',
        documents: documents,
        columnsPerRow: this.properties.columnsPerRow || 4,
        showDate: this.properties.showDate !== false,
        allowUpload: this.properties.allowUpload !== false,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        onDocumentsUpdate: (updatedDocs: IHRDocument[]) => {
          this.properties.documents = JSON.stringify(updatedDocs);
          this.context.propertyPane.refresh();
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getDefaultDocuments(): IHRDocument[] {
    return [
      {
        id: '1',
        title: 'Employee Handbook',
        date: 'June 23, 2025',
        iconData: '',
        iconType: 'word',
        documentUrl: '#'
      },
      {
        id: '2',
        title: 'Employee Handbook',
        date: 'June 23, 2025',
        iconData: '',
        iconType: 'word',
        documentUrl: '#'
      },
      {
        id: '3',
        title: 'Training and Onboarding',
        date: 'June 23, 2025',
        iconData: '',
        iconType: 'pdf',
        documentUrl: '#'
      },
      {
        id: '4',
        title: 'Training and Onboarding',
        date: 'June 23, 2025',
        iconData: '',
        iconType: 'pdf',
        documentUrl: '#'
      }
    ];
  }

  protected onInit(): Promise<void> {
    if (!this.properties.title) {
      this.properties.title = 'Common HR Documents';
    }
    if (!this.properties.columnsPerRow) {
      this.properties.columnsPerRow = 4;
    }
    if (!this.properties.documents) {
      this.properties.documents = JSON.stringify(this.getDefaultDocuments());
    }

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
            description: 'Configure HR Documents Display'
          },
          groups: [
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Section Title'
                }),
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per Row',
                  min: 2,
                  max: 6,
                  value: this.properties.columnsPerRow || 4,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneToggle('showDate', {
                  label: 'Show Dates',
                  checked: this.properties.showDate !== false,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('allowUpload', {
                  label: 'Allow Icon Upload',
                  checked: this.properties.allowUpload !== false,
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