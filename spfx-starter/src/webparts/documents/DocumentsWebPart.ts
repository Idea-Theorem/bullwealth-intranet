import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import Documents from './components/Documents';
import { IDocumentsProps } from './components/IDocumentsProps';

export interface IDocumentCategory {
  id: string;
  title: string;
  imageUrl: string;
  documentUrl: string;
  viewAllUrl: string;
}

export interface IDocumentsWebPartProps {
  title: string;
  columnsPerRow: number;
  
  // Support up to 8 document categories
  category1Title: string;
  category1ImageUrl: string;
  category1DocumentUrl: string;
  category1ViewAllUrl: string;
  
  category2Title: string;
  category2ImageUrl: string;
  category2DocumentUrl: string;
  category2ViewAllUrl: string;
  
  category3Title: string;
  category3ImageUrl: string;
  category3DocumentUrl: string;
  category3ViewAllUrl: string;
  
  category4Title: string;
  category4ImageUrl: string;
  category4DocumentUrl: string;
  category4ViewAllUrl: string;
  
  category5Title: string;
  category5ImageUrl: string;
  category5DocumentUrl: string;
  category5ViewAllUrl: string;
  
  category6Title: string;
  category6ImageUrl: string;
  category6DocumentUrl: string;
  category6ViewAllUrl: string;
  
  category7Title: string;
  category7ImageUrl: string;
  category7DocumentUrl: string;
  category7ViewAllUrl: string;
  
  category8Title: string;
  category8ImageUrl: string;
  category8DocumentUrl: string;
  category8ViewAllUrl: string;
}

export default class DocumentsWebPart extends BaseClientSideWebPart<IDocumentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const categories: IDocumentCategory[] = [];
    
    const defaultTitles = [
      'Compliance',
      'Research & Investments',
      'Advisory Group',
      'Operations',
      'Business Development',
      'Tax & Accounting',
      'Employee Directory',
      'Human Resources'
    ];

    // Build categories array from properties
    for (let i = 1; i <= 8; i++) {
      const props = this.properties as any;
      if (props[`category${i}Title`]) {
        categories.push({
          id: i.toString(),
          title: props[`category${i}Title`],
          imageUrl: props[`category${i}ImageUrl`] || '',
          documentUrl: props[`category${i}DocumentUrl`] || '#',
          viewAllUrl: props[`category${i}ViewAllUrl`] || '#'
        });
      }
    }

    // If no categories configured, use defaults
    if (categories.length === 0) {
      for (let i = 0; i < 6; i++) {
        categories.push({
          id: (i + 1).toString(),
          title: defaultTitles[i],
          imageUrl: '',
          documentUrl: '#',
          viewAllUrl: '#'
        });
      }
    }

    const element: React.ReactElement<IDocumentsProps> = React.createElement(
      Documents,
      {
        title: this.properties.title || 'Documents',
        categories: categories,
        columnsPerRow: this.properties.columnsPerRow || 4,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Set default values
    if (!this.properties.title) {
      this.properties.title = 'Documents';
    }
    if (!this.properties.columnsPerRow) {
      this.properties.columnsPerRow = 4;
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
    const categoryGroups = [];
    
    for (let i = 1; i <= 8; i++) {
      categoryGroups.push({
        groupName: `Category ${i}`,
        groupFields: [
          PropertyPaneTextField(`category${i}Title`, {
            label: 'Title',
            placeholder: 'Enter category title'
          }),
          PropertyPaneTextField(`category${i}ImageUrl`, {
            label: 'Image URL',
            placeholder: 'https://your-site/image.jpg'
          }),
          PropertyPaneTextField(`category${i}DocumentUrl`, {
            label: 'Document Library URL',
            placeholder: 'https://your-site/Shared%20Documents'
          }),
          PropertyPaneTextField(`category${i}ViewAllUrl`, {
            label: 'View All Documents URL',
            placeholder: 'https://your-site/Shared%20Documents/Forms/AllItems.aspx'
          })
        ]
      });
    }

    return {
      pages: [
        {
          header: {
            description: 'Configure Document Categories'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Section Title'
                }),
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per row',
                  min: 2,
                  max: 6,
                  value: this.properties.columnsPerRow || 4,
                  showValue: true,
                  step: 1
                })
              ]
            },
            ...categoryGroups
          ]
        }
      ]
    };
  }
}