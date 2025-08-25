import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WelcomeBannerPartWebPartStrings';
import WelcomeBannerPart from './components/WelcomeBannerPart';
import { IWelcomeBannerPartProps } from './components/IWelcomeBannerPartProps';

export default class WelcomeBannerPartWebPart extends BaseClientSideWebPart<IWelcomeBannerPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWelcomeBannerPartProps> = React.createElement(
      WelcomeBannerPart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        
        // Pass all configurable properties
        messageTitle: this.properties.messageTitle,
        ceoName: this.properties.ceoName,
        ceoTitle: this.properties.ceoTitle,
        ceoMessage: this.properties.ceoMessage,
        ceoExpandedMessage: this.properties.ceoExpandedMessage,
        backgroundImageUrl: this.properties.backgroundImageUrl,
        ceoImageUrl: this.properties.ceoImageUrl,
        videoUrl: this.properties.videoUrl,
        showVideo: this.properties.showVideo,
        readMoreText: this.properties.readMoreText,
        readLessText: this.properties.readLessText
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Set default values if not already set
    if (!this.properties.messageTitle) {
      this.properties.messageTitle = 'Message from CEO';
    }
    if (!this.properties.ceoName) {
      this.properties.ceoName = 'Sarah Johnson';
    }
    if (!this.properties.ceoTitle) {
      this.properties.ceoTitle = 'Chief Executive Officer';
    }
    if (!this.properties.ceoMessage) {
      this.properties.ceoMessage = "We wouldn't be where we are today without each and every one of you. Thank you for making us successful!";
    }
    if (!this.properties.ceoExpandedMessage) {
      this.properties.ceoExpandedMessage = "Your dedication, hard work, and commitment to excellence have been instrumental in our journey. Together, we've overcome challenges, celebrated victories, and built something truly remarkable.";
    }
    if (!this.properties.readMoreText) {
      this.properties.readMoreText = 'Read More';
    }
    if (!this.properties.readLessText) {
      this.properties.readLessText = 'Read Less';
    }
    if (this.properties.showVideo === undefined) {
      this.properties.showVideo = true;
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

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
            description: 'Configure the CEO Welcome Banner'
          },
          groups: [
            {
              groupName: 'Content Settings',
              groupFields: [
                PropertyPaneTextField('messageTitle', {
                  label: 'Message Title',
                  description: 'The main title displayed (e.g., "Message from CEO")',
                  value: this.properties.messageTitle
                }),
                PropertyPaneTextField('ceoName', {
                  label: 'CEO Name',
                  description: 'Name of the CEO',
                  value: this.properties.ceoName
                }),
                PropertyPaneTextField('ceoTitle', {
                  label: 'CEO Title',
                  description: 'Title/Position of the CEO',
                  value: this.properties.ceoTitle
                }),
                PropertyPaneTextField('ceoMessage', {
                  label: 'Main Message',
                  description: 'The main message that is always visible',
                  multiline: true,
                  rows: 3,
                  value: this.properties.ceoMessage
                }),
                PropertyPaneTextField('ceoExpandedMessage', {
                  label: 'Extended Message',
                  description: 'Additional message shown when "Read More" is clicked',
                  multiline: true,
                  rows: 4,
                  value: this.properties.ceoExpandedMessage
                })
              ]
            },
            {
              groupName: 'Media Settings',
              groupFields: [
                PropertyPaneTextField('backgroundImageUrl', {
                  label: 'Background Image URL',
                  description: 'URL for the background image (city skyline)',
                  value: this.properties.backgroundImageUrl,
                  placeholder: 'https://your-site/SiteAssets/city-skyline.jpg'
                }),
                PropertyPaneTextField('ceoImageUrl', {
                  label: 'CEO Image URL',
                  description: 'URL for the CEO photo/video thumbnail',
                  value: this.properties.ceoImageUrl,
                  placeholder: 'https://your-site/SiteAssets/ceo-photo.jpg'
                }),
                PropertyPaneToggle('showVideo', {
                  label: 'Show Video Section',
                  checked: this.properties.showVideo,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneTextField('videoUrl', {
                  label: 'Video URL',
                  description: 'URL for the CEO video (if video section is enabled)',
                  value: this.properties.videoUrl,
                  placeholder: 'https://your-site/video-url',
                  disabled: !this.properties.showVideo
                })
              ]
            },
            {
              groupName: 'Button Labels',
              groupFields: [
                PropertyPaneTextField('readMoreText', {
                  label: 'Read More Button Text',
                  value: this.properties.readMoreText
                }),
                PropertyPaneTextField('readLessText', {
                  label: 'Read Less Button Text',
                  value: this.properties.readLessText
                })
              ]
            }
          ]
        }
      ]
    };
  }
}