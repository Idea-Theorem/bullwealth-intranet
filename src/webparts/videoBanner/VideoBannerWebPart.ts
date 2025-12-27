import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import VideoBanner from './components/VideoBanner';
import { IVideoBannerProps } from './components/IVideoBannerProps';

export interface IVideoBannerWebPartProps {
  description: string;
  backgroundImages: string;
  autoSlide: boolean;
  slideInterval: number;
}

export default class VideoBannerWebPart extends BaseClientSideWebPart<IVideoBannerWebPartProps> {
  
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    const bgImages = this.properties.backgroundImages 
      ? this.properties.backgroundImages.split(',').map(url => url.trim()).filter(url => url.length > 0)
      : [];

    const element: React.ReactElement<IVideoBannerProps> = React.createElement(
      VideoBanner,
      {
        context: this.context,
        backgroundImages: bgImages,
        autoSlide: this.properties.autoSlide !== undefined ? this.properties.autoSlide : true,
        slideInterval: this.properties.slideInterval || 5
      }
    );
    
    // Enable full-width rendering by removing container constraints
    if (this.domElement) {
      this.domElement.style.maxWidth = 'none';
      this.domElement.style.width = '100%';
    }
    
    if (this.domElement.parentElement) {
      this.domElement.parentElement.style.maxWidth = 'none';
      this.domElement.parentElement.style.padding = '0';
      this.domElement.parentElement.style.margin = '0';
    }
    
    // Also try to remove constraints from grandparent
    if (this.domElement.parentElement?.parentElement) {
      this.domElement.parentElement.parentElement.style.maxWidth = 'none';
    }
    
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _openSiteAssets = (): void => {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    window.open(`${siteUrl}/SiteAssets/Forms/AllItems.aspx`, '_blank');
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const imageCount = this.properties.backgroundImages 
      ? this.properties.backgroundImages.split(',').filter(url => url.trim().length > 0).length 
      : 0;

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
            },
            {
              groupName: 'Background Carousel Settings',
              groupFields: [
                PropertyPaneLabel('imageCount', {
                  text: `Currently selected: ${imageCount} image(s)`
                }),
                PropertyPaneButton('openLibrary', {
                  text: '📁 Open SharePoint Image Library',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._openSiteAssets
                }),
                PropertyPaneLabel('instructions', {
                  text: 'Steps: 1) Click button above 2) Upload/select images 3) Right-click → Copy link 4) Paste URLs below'
                }),
                PropertyPaneTextField('backgroundImages', {
                  label: 'Background Image URLs (comma-separated)',
                  description: 'Paste image URLs from SharePoint, separated by commas',
                  value: this.properties.backgroundImages || '',
                  multiline: true,
                  rows: 6,
                  placeholder: 'https://tenant.sharepoint.com/sites/site/SiteAssets/bg1.jpg,\nhttps://tenant.sharepoint.com/sites/site/SiteAssets/bg2.jpg,\nhttps://tenant.sharepoint.com/sites/site/SiteAssets/bg3.jpg'
                }),
                PropertyPaneToggle('autoSlide', {
                  label: 'Auto-advance slides',
                  checked: this.properties.autoSlide !== undefined ? this.properties.autoSlide : true,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneSlider('slideInterval', {
                  label: 'Slide interval (seconds)',
                  min: 3,
                  max: 15,
                  step: 1,
                  value: this.properties.slideInterval || 5,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
