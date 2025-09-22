import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MessagesWebPartStrings';
import Messages from './components/Messages';
import { IMessagesProps } from './components/IMessagesProps';

export interface IMessagesWebPartProps {
  description: string;
  listName: string;
  title: string;
  columnsPerRow: number;
}

export default class MessagesWebPart extends BaseClientSideWebPart<IMessagesWebPartProps> {

  // Update the render method to use correct list name
public render(): void {
  const element: React.ReactElement<IMessagesProps> = React.createElement(
    Messages,
    {
      description: this.properties.description,
      context: this.context,
      listName: this.properties.listName || 'Archived-Messages',
      title: this.properties.title || 'Newsletters and Messages',
      columnsPerRow: this.properties.columnsPerRow || 4
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
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name'
                }),
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per Row',
                  min: 1,
                  max: 5,
                  value: 4,
                  showValue: true,
                  step: 1
                }),
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
