// DynamicDocumentPageWebPart.ts
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import DocumentPageRouter from './components/DocumentPageRouter';

export interface IDynamicDocumentPageWebPartProps {
  documentLibraryName: string;
  baseLibraryPath: string;
}

export default class DynamicDocumentPageWebPart extends BaseClientSideWebPart<IDynamicDocumentPageWebPartProps> {
  public render(): void {
    // Extract folder name from URL query parameter
    const urlParams = new URLSearchParams(window.location.search);
    const folderName = urlParams.get('folder') || 'Unknown';

    const element: React.ReactElement = React.createElement(
      DocumentPageRouter,
      {
        context: this.context,
        folderName: folderName,
        documentLibraryName: this.properties.documentLibraryName || 'Documents',
        baseLibraryPath: this.properties.baseLibraryPath || 'BullWealth Documents'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Dynamic Document Page Configuration' },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('documentLibraryName', {
                  label: 'Document Library Name',
                  value: this.properties.documentLibraryName || 'Documents'
                }),
                PropertyPaneTextField('baseLibraryPath', {
                  label: 'Base Library Path',
                  value: this.properties.baseLibraryPath || 'BullWealth Documents'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
