import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import DocumentLibrary from './components/DocumentLibrary';

export interface IDocumentLibraryWebPartProps {
  title: string;
  description: string;
  listName: string;
}

export default class DocumentLibraryWebPart extends BaseClientSideWebPart<IDocumentLibraryWebPartProps> {
  public render(): void {
    // ✅ FIXED: Read library parameter from URL
    const urlParams = new URLSearchParams(window.location.search);
    const libraryPath = urlParams.get('library');

    // Use URL parameter OR default
    let listName = libraryPath || this.properties.listName || 'Documents';

    if (libraryPath) {
      listName = decodeURIComponent(libraryPath);
      console.log('✅ DocumentLibrary WebPart - Using library from URL:', listName);
    } else {
      console.log('📋 DocumentLibrary WebPart - Using default listName:', listName);
    }

    const element: React.ReactElement = React.createElement(DocumentLibrary, {
      context: this.context,
      title: this.properties.title || 'Document Library',
      listName: listName
    });

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure Document Library WebPart'
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: this.properties.title || 'Document Library'
                }),
                PropertyPaneTextField('listName', {
                  label: 'Default Library Path',
                  value: this.properties.listName || 'Documents',
                  description: 'Used if no ?library= URL parameter is provided'
                }),
                PropertyPaneTextField('description', {
                  label: 'Description',
                  value: this.properties.description || '',
                  description: 'Optional description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
