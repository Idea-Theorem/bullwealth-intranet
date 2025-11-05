import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import VideoBanner from './components/VideoBanner';
import { IVideoBannerProps } from './components/IVideoBannerProps';

export interface IVideoBannerWebPartProps {
  description: string;
}

export default class VideoBannerWebPart extends BaseClientSideWebPart<IVideoBannerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IVideoBannerProps> = React.createElement(
      VideoBanner,
      {
        context: this.context
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
            description: 'Video Banner - Content managed from "Archived-Messages" list'
          },
          groups: [
            {
              groupName: 'Information',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'All content is managed in the "Archived-Messages" SharePoint list.',
                  value: 'Title, Content, PublishedDate, FeaturedImage, and NewsletterVideo columns.',
                  multiline: true,
                  disabled: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
