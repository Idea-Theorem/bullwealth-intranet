import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'QuickAccessWebPartStrings';
import QuickAccess from './components/QuickAccess';
import { IQuickAccessProps } from './components/IQuickAccessProps';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IQuickAccessWebPartProps {
  title: string;
  quickLinks: any[];
  supports: any[];
}

export default class QuickAccessWebPart extends BaseClientSideWebPart<IQuickAccessWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuickAccessProps> = React.createElement(
      QuickAccess,
      {
        title: this.properties.title,
        quickLinks: this.properties.quickLinks,
        supports: this.properties.supports
      }
    );

    ReactDom.render(element, this.domElement);
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
                  label: 'Section Title'
                }),
                PropertyFieldCollectionData('quickLinks', {
                  key: 'quickLinks',
                  label: 'Quick Links',
                  panelHeader: 'Edit Quick Links',
                  manageBtnLabel: 'Manage Quick Links',
                  value: this.properties.quickLinks,
                  fields: [
                    {
                      id: 'title',
                      title: 'Title',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'iconUrl',
                      title: 'Icon URL (upload image to SharePoint library and paste absolute URL)',
                      type: CustomCollectionFieldType.url,
                      required: true
                    },
                    {
                      id: 'url',
                      title: 'URL (leave blank if no link)',
                      type: CustomCollectionFieldType.url
                    }
                  ],
                  disabled: false
                }),
                PropertyFieldCollectionData('supports', {
                  key: 'supports',
                  label: 'Tech Supports',
                  panelHeader: 'Edit Tech Supports',
                  manageBtnLabel: 'Manage Tech Supports',
                  value: this.properties.supports,
                  fields: [
                    {
                      id: 'name',
                      title: 'Name (e.g., Clover)',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'email',
                      title: 'Email',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'phone',
                      title: 'Phone',
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}