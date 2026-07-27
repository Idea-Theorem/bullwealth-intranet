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
  // ✅ FIXED: Added error handling wrapper
  public render(): void {
    const urlParams = new URLSearchParams(window.location.search);
    const libraryPath = urlParams.get('library');

    let listName = libraryPath || this.properties.listName || 'Documents';

    if (libraryPath) {
      listName = decodeURIComponent(libraryPath);
      console.log('✅ DocumentLibrary WebPart - Using library from URL:', listName);
    } else {
      console.log('📋 DocumentLibrary WebPart - Using default listName:', listName);
    }

    try {
      const element: React.ReactElement = React.createElement(DocumentLibrary, {
        context: this.context,
        title: this.properties.title || 'Document Library',
        listName: listName
      });

      ReactDom.render(element, this.domElement);
    } catch (error) {
      console.error('Error rendering DocumentLibrary:', error);
      this.domElement.innerHTML = `
        <div style="padding: 20px; color: #d13438; text-align: center; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
          <h3 style="margin: 0 0 12px 0; font-size: 20px;">Something went wrong</h3>
          <p style="margin: 8px 0; font-size: 14px; color: #605e5c;">Please refresh the page. If the problem persists, contact the site administrator.</p>
          <button 
            onclick="window.location.reload()" 
            style="padding: 8px 16px; margin-top: 16px; cursor: pointer; background: #0078d4; color: white; border: none; border-radius: 2px; font-size: 14px;"
          >
            Refresh Page
          </button>
          <div style="margin-top: 20px; padding: 12px; background: #f3f2f1; border-radius: 2px; font-size: 12px; color: #605e5c; text-align: left;">
            <strong>Technical Details:</strong><br/>
            ${error instanceof Error ? error.message : String(error)}
          </div>
        </div>
      `;
    }
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
    return Version.parse('1.0');
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
