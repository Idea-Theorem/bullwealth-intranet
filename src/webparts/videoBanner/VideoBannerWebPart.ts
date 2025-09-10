import * as React from 'react';
import * as ReactDom from 'react-dom';
// Import main styles (this loads the font faces)
import '../../styles/main.scss';

import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

// Import PnP property controls
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import VideoBanner from './components/VideoBanner';
import { IVideoBannerProps } from './components/IVideoBannerProps';

export interface IVideoBannerWebPartProps {
  title: string;
  message: string;
  buttonText: string;
  videoUrl: string;
  videoFile?: IFilePickerResult;
  thumbnailUrl: string;
  thumbnailFile?: IFilePickerResult;
  backgroundImageUrl: string;
  backgroundImageFile?: IFilePickerResult;
  autoPlay: boolean;
  showInModal: boolean;
}

export default class VideoBannerWebPart extends BaseClientSideWebPart<IVideoBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // Safe property access with optional chaining
    const thumbnailUrl = this.properties.thumbnailFile?.fileAbsoluteUrl || this.properties.thumbnailUrl;
    const backgroundImageUrl = this.properties.backgroundImageFile?.fileAbsoluteUrl || this.properties.backgroundImageUrl;
    const videoUrl = this.properties.videoFile?.fileAbsoluteUrl || this.properties.videoUrl;

    const element: React.ReactElement<IVideoBannerProps> = React.createElement(
      VideoBanner,
      {
        title: this.properties.title || 'Message from CEO',
        message: this.properties.message || '"We wouldn\'t be where we are today without each and every one of you. Thank you for making us successful!"',
        buttonText: this.properties.buttonText || 'Read More',
        videoUrl: videoUrl || '',
        thumbnailUrl: thumbnailUrl || '',
        backgroundImageUrl: backgroundImageUrl || '',
        autoPlay: this.properties.autoPlay || false,
        showInModal: this.properties.showInModal !== false,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
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
            description: 'Configure Video Banner'
          },
          groups: [
            {
              groupName: 'Content Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title',
                  placeholder: 'Message from CEO'
                }),
                PropertyPaneTextField('message', {
                  label: 'Message',
                  multiline: true,
                  rows: 4,
                  placeholder: 'Enter your message here'
                }),
                PropertyPaneTextField('buttonText', {
                  label: 'Button Text',
                  placeholder: 'Read More'
                })
              ]
            },
            {
              groupName: 'Media Settings',
              groupFields: [
                // Video File Picker
                PropertyFieldFilePicker('videoFile', {
                  context: this.context as any,
                  key: 'videoFilePicker',
                  buttonLabel: 'Select Video',
                  label: 'Video',
                  accepts: ['.mp4', '.avi', '.mov', '.wmv', '.flv', '.webm'],
                  buttonIcon: 'Video',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.videoFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    this.properties.videoFile = filePickerResult;
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.videoFile = filePickerResult;
                  },
                  hideStockImages: true,
                  hideWebSearchTab: true
                }),
                PropertyPaneTextField('videoUrl', {
                  label: 'Or Video URL',
                  placeholder: 'https://your-video-url.mp4 or YouTube/Stream URL'
                }),
                
                // Thumbnail File Picker
                PropertyFieldFilePicker('thumbnailFile', {
                  context: this.context as any,
                  key: 'thumbnailFilePicker',
                  buttonLabel: 'Select Thumbnail',
                  label: 'Video Thumbnail',
                  accepts: ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'],
                  buttonIcon: 'FileImage',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.thumbnailFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    this.properties.thumbnailFile = filePickerResult;
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.thumbnailFile = filePickerResult;
                  }
                }),
                PropertyPaneTextField('thumbnailUrl', {
                  label: 'Or Thumbnail URL',
                  placeholder: 'https://your-site/thumbnail.jpg'
                }),

                // Background Image File Picker
                PropertyFieldFilePicker('backgroundImageFile', {
                  context: this.context as any,
                  key: 'backgroundImageFilePicker',
                  buttonLabel: 'Select Background Image',
                  label: 'Background Image',
                  accepts: ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'],
                  buttonIcon: 'Photo2',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.backgroundImageFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    this.properties.backgroundImageFile = filePickerResult;
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.backgroundImageFile = filePickerResult;
                  }
                }),
                PropertyPaneTextField('backgroundImageUrl', {
                  label: 'Or Background Image URL',
                  placeholder: 'https://your-site/background.jpg'
                })
              ]
            },
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneToggle('showInModal', {
                  label: 'Play video in modal',
                  onText: 'Modal',
                  offText: 'Inline'
                }),
                PropertyPaneToggle('autoPlay', {
                  label: 'Auto-play video',
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
