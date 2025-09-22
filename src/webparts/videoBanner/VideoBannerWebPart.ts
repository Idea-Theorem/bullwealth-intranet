/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import '../../styles/main.scss';

import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

// 🔧 FIXED: Correct DateTimePicker import
import {
  PropertyFieldDateTimePicker,
  DateConvention,
  TimeConvention,
  IDateTimeFieldValue
} from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

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
  lastUpdate: number;
  // 🔧 FIXED: Use correct type for date picker
  publishedDate?: IDateTimeFieldValue;
  readMoreUrl?: string;
}

export default class VideoBannerWebPart extends BaseClientSideWebPart<IVideoBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const thumbnailUrl = this.properties.thumbnailFile?.fileAbsoluteUrl || this.properties.thumbnailUrl;
    const backgroundImageUrl = this.properties.backgroundImageFile?.fileAbsoluteUrl || this.properties.backgroundImageUrl;
    const videoUrl = this.properties.videoFile?.fileAbsoluteUrl || this.properties.videoUrl;

    // 🔧 FIXED: Properly format date for display
    let formattedDate: string | undefined;
    if (this.properties.publishedDate && this.properties.publishedDate.value) {
      const date = new Date(this.properties.publishedDate.value);
      formattedDate = `Published ${date.toLocaleDateString('en-GB', { 
        day: 'numeric',
        month: 'short', 
        year: 'numeric' 
      })}`;
    }

    console.log('🎥 Video Banner Render:', {
      videoUrl,
      thumbnailUrl,
      backgroundImageUrl,
      publishedDate: formattedDate,
      readMoreUrl: this.properties.readMoreUrl
    });

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
        publishedDate: formattedDate, // 🔧 FIXED: Pass formatted date string
        readMoreUrl: this.properties.readMoreUrl, // 🔧 ENSURE: Read more URL is passed
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private forceRefresh(): void {
    this.properties.lastUpdate = Date.now();
    this.context.propertyPane.refresh();
    this.render();
  }

  private clearVideoFile(): void {
    this.properties.videoFile = undefined;
    this.properties.videoUrl = '';
    this.forceRefresh();
  }

  private clearThumbnailFile(): void {
    this.properties.thumbnailFile = undefined;
    this.properties.thumbnailUrl = '';
    this.forceRefresh();
  }

  private clearBackgroundFile(): void {
    this.properties.backgroundImageFile = undefined;
    this.properties.backgroundImageUrl = '';
    this.forceRefresh();
  }

  private async uploadFileToSharePoint(file: File, type: 'video' | 'thumbnail' | 'background'): Promise<string> {
    try {
      console.log(`🚀 Starting SharePoint upload for ${type}:`, file.name);

      const maxSize = 100 * 1024 * 1024; // 100MB
      if (file.size > maxSize) {
        throw new Error(`File size (${(file.size / 1024 / 1024).toFixed(2)}MB) exceeds 100MB limit`);
      }

      const digestResponse = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        }
      });
      
      if (!digestResponse.ok) {
        throw new Error(`Failed to get request digest: ${digestResponse.status}`);
      }
      
      const digestData = await digestResponse.json();
      const requestDigest = digestData.d.GetContextWebInformation.FormDigestValue;

      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      const folderName = 'SiteAssets/VideoBanner';
      const timestamp = Date.now();
      const cleanFileName = file.name.replace(/[^a-zA-Z0-9.-]/g, '_');
      const fileName = `${type}_${timestamp}_${cleanFileName}`;

      try {
        await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/folders`, {
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.Folder' },
            'ServerRelativeUrl': `/${folderName}`
          })
        });
        
        console.log('✅ Folder created successfully');
      } catch (folderError) {
        console.log('📁 Folder might already exist:', folderError);
      }

      const uploadUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${folderName}')/Files/Add(url='${fileName}', overwrite=true)`;
      
      const uploadResponse = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        },
        body: arrayBuffer
      });

      if (!uploadResponse.ok) {
        const errorText = await uploadResponse.text();
        console.error('❌ Upload response error:', errorText);
        throw new Error(`Upload failed: ${uploadResponse.status} - ${errorText}`);
      }

      const fileUrl = `${this.context.pageContext.web.absoluteUrl}/${folderName}/${fileName}`;
      
      console.log('✅ SharePoint upload successful:', fileUrl);
      return fileUrl;

    } catch (error) {
      console.error(`❌ SharePoint upload failed for ${type}:`, error);
      throw error;
    }
  }

  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  private handleNativeVideoUpload(): void {
    this.createFileInput('.mp4,.avi,.mov,.wmv,.flv,.webm,.mkv', (file) => {
      this.processFileUpload(file, 'video');
    });
  }

  private handleNativeThumbnailUpload(): void {
    this.createFileInput('.jpg,.jpeg,.png,.gif,.bmp,.svg,.webp', (file) => {
      this.processFileUpload(file, 'thumbnail');
    });
  }

  private handleNativeBackgroundUpload(): void {
    this.createFileInput('.jpg,.jpeg,.png,.gif,.bmp,.svg,.webp', (file) => {
      this.processFileUpload(file, 'background');
    });
  }

  private createFileInput(accept: string, callback: (file: File) => void): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = (event: Event) => {
      const target = event.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        callback(file);
      }
    };
    input.click();
  }

  private async processFileUpload(file: File, type: 'video' | 'thumbnail' | 'background'): Promise<void> {
    try {
      console.log(`Uploading ${type}... Please wait.`);

      const fileUrl = await this.uploadFileToSharePoint(file, type);
      
      const filePickerResult: IFilePickerResult = {
        fileAbsoluteUrl: fileUrl,
        fileName: file.name,
        fileNameWithoutExtension: file.name.replace(/\.[^/.]+$/, ''),
        downloadFileContent: () => {
          return fetch(fileUrl)
            .then(response => response.blob())
            .then(blob => {
              return new File([blob], file.name, {
                type: blob.type,
                lastModified: Date.now()
              });
            });
        }
      };

      switch (type) {
        case 'video':
          this.properties.videoFile = filePickerResult;
          this.properties.videoUrl = fileUrl;
          break;
        case 'thumbnail':
          this.properties.thumbnailFile = filePickerResult;
          this.properties.thumbnailUrl = fileUrl;
          break;
        case 'background':
          this.properties.backgroundImageFile = filePickerResult;
          this.properties.backgroundImageUrl = fileUrl;
          break;
      }

      this.forceRefresh();
      alert(`✅ ${type.charAt(0).toUpperCase() + type.slice(1)} uploaded successfully!`);

    } catch (error) {
      console.error(`❌ Upload failed for ${type}:`, error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      alert(`❌ Failed to upload ${type}: ${errorMessage}\n\nPlease try again or use a different file.`);
    }
  }

  private handleVideoUrlInput(): void {
    const url = prompt(
      'Enter Video URL:\n\n• Direct video file URL (MP4, AVI, etc.)\n• YouTube URL\n• Microsoft Stream URL\n• Any public video URL',
      'https://'
    );
    
    if (url && url !== 'https://') {
      this.properties.videoUrl = url;
      this.properties.videoFile = undefined;
      this.forceRefresh();
    }
  }

  private handleThumbnailUrlInput(): void {
    const url = prompt(
      'Enter Thumbnail Image URL:\n\n• Direct image file URL\n• Any public image URL',
      'https://'
    );
    
    if (url && url !== 'https://') {
      this.properties.thumbnailUrl = url;
      this.properties.thumbnailFile = undefined;
      this.forceRefresh();
    }
  }

  private handleBackgroundUrlInput(): void {
    const url = prompt(
      'Enter Background Image URL:\n\n• Direct image file URL\n• Any public image URL',
      'https://'
    );
    
    if (url && url !== 'https://') {
      this.properties.backgroundImageUrl = url;
      this.properties.backgroundImageFile = undefined;
      this.forceRefresh();
    }
  }

  protected onInit(): Promise<void> {
    this.properties.lastUpdate = this.properties.lastUpdate || Date.now();
    
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
            default:
              environmentMessage = 'SharePoint';
          }
          return environmentMessage;
        });
    }
    return Promise.resolve('SharePoint');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
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
            description: 'Configure Video Banner - Multiple Upload Options Available'
          },
          groups: [
            {
              groupName: 'Content Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title',
                  placeholder: 'Message from CEO',
                  value: this.properties.title
                }),
                // 🔧 FIXED: Correct DateTimePicker configuration
                PropertyFieldDateTimePicker('publishedDate', {
                  label: 'Published Date',
                  initialDate: this.properties.publishedDate || { value: new Date(), displayValue: new Date().toLocaleDateString() },
                  dateConvention: DateConvention.Date,
                  timeConvention: TimeConvention.Hours24,
                  key: `publishedDatePicker_${this.properties.lastUpdate}`,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: () => '',
                  deferredValidationTime: 0,
                  showLabels: false
                }),
                PropertyPaneTextField('message', {
                  label: 'Message',
                  multiline: true,
                  rows: 4,
                  placeholder: 'Enter your message here',
                  value: this.properties.message
                }),
                PropertyPaneTextField('buttonText', {
                  label: 'Button Text',
                  placeholder: 'Read More',
                  value: this.properties.buttonText
                }),
                PropertyPaneTextField('readMoreUrl', {
                  label: 'Read More URL',
                  placeholder: 'https://your-site.com/article',
                  value: this.properties.readMoreUrl,
                  description: 'URL to open when Read More button is clicked'
                })
              ]
            },
            {
              groupName: '🎥 Video Upload Options',
              groupFields: [
                PropertyFieldFilePicker('videoFile', {
                  context: this.context as any,
                  key: `videoFilePicker_${this.properties.lastUpdate}`,
                  buttonLabel: '📁 Browse Files (Recent/Sites/OneDrive)',
                  label: 'Option 1: Browse SharePoint Files',
                  accepts: ['.mp4', '.avi', '.mov', '.wmv', '.flv', '.webm', '.mkv'],
                  buttonIcon: 'Video',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.videoFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log('📁 Video file selected:', filePickerResult);
                    this.properties.videoFile = filePickerResult;
                    this.properties.videoUrl = filePickerResult.fileAbsoluteUrl || '';
                    this.forceRefresh();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.videoFile = filePickerResult;
                  },
                  hideStockImages: true,
                  hideWebSearchTab: true,
                  hideLinkUploadTab: false,
                  hideOrganisationalAssetTab: false,
                  hideRecentTab: false,
                  hideOneDriveTab: false,
                  hideSiteFilesTab: false,
                  storeLastActiveTab: true,
                  required: false
                }),

                PropertyPaneButton('uploadVideoNative', {
                  text: '💻 Option 2: Upload from Computer',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleNativeVideoUpload.bind(this)
                }),

                PropertyPaneButton('videoUrlInput', {
                  text: '🔗 Option 3: Enter Video URL',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleVideoUrlInput.bind(this)
                }),

                PropertyPaneTextField('videoUrl', {
                  label: '🔗 Or paste Video URL directly',
                  placeholder: 'https://your-video-url.mp4 or YouTube/Stream URL',
                  value: this.properties.videoUrl,
                  multiline: true
                }),
                
                ...(this.properties.videoFile?.fileName || this.properties.videoUrl ? [
                  PropertyPaneTextField('currentVideo', {
                    label: '📹 Current Selection',
                    value: this.properties.videoFile?.fileName || this.properties.videoUrl || 'None',
                    disabled: true
                  }),
                  PropertyPaneButton('clearVideoFile', {
                    text: '🗑️ Clear Video Selection',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.clearVideoFile.bind(this)
                  })
                ] : [])
              ]
            },
            {
              groupName: '🖼️ Thumbnail Upload Options',
              groupFields: [
                PropertyFieldFilePicker('thumbnailFile', {
                  context: this.context as any,
                  key: `thumbnailFilePicker_${this.properties.lastUpdate}`,
                  buttonLabel: '📁 Browse Images (Recent/Sites/OneDrive)',
                  label: 'Option 1: Browse SharePoint Images',
                  accepts: ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'],
                  buttonIcon: 'FileImage',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.thumbnailFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log('🖼️ Thumbnail selected:', filePickerResult);
                    this.properties.thumbnailFile = filePickerResult;
                    this.properties.thumbnailUrl = filePickerResult.fileAbsoluteUrl || '';
                    this.forceRefresh();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.thumbnailFile = filePickerResult;
                  },
                  hideStockImages: false,
                  hideWebSearchTab: false,
                  hideLinkUploadTab: false,
                  hideOrganisationalAssetTab: false,
                  hideRecentTab: false,
                  hideOneDriveTab: false,
                  hideSiteFilesTab: false,
                  storeLastActiveTab: true,
                  required: false
                }),

                PropertyPaneButton('uploadThumbnailNative', {
                  text: '💻 Option 2: Upload from Computer',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleNativeThumbnailUpload.bind(this)
                }),

                PropertyPaneButton('thumbnailUrlInput', {
                  text: '🔗 Option 3: Enter Image URL',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleThumbnailUrlInput.bind(this)
                }),

                PropertyPaneTextField('thumbnailUrl', {
                  label: '🔗 Or paste Thumbnail URL directly',
                  placeholder: 'https://your-site/thumbnail.jpg',
                  value: this.properties.thumbnailUrl,
                  multiline: true
                }),

                ...(this.properties.thumbnailFile?.fileName || this.properties.thumbnailUrl ? [
                  PropertyPaneTextField('currentThumbnail', {
                    label: '🖼️ Current Selection',
                    value: this.properties.thumbnailFile?.fileName || this.properties.thumbnailUrl || 'None',
                    disabled: true
                  }),
                  PropertyPaneButton('clearThumbnailFile', {
                    text: '🗑️ Clear Thumbnail Selection',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.clearThumbnailFile.bind(this)
                  })
                ] : [])
              ]
            },
            {
              groupName: '🎨 Background Upload Options',
              groupFields: [
                PropertyFieldFilePicker('backgroundImageFile', {
                  context: this.context as any,
                  key: `backgroundImageFilePicker_${this.properties.lastUpdate}`,
                  buttonLabel: '📁 Browse Images (Recent/Sites/OneDrive)',
                  label: 'Option 1: Browse SharePoint Images',
                  accepts: ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'],
                  buttonIcon: 'Photo2',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  filePickerResult: this.properties.backgroundImageFile || ({} as IFilePickerResult),
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log('🎨 Background selected:', filePickerResult);
                    this.properties.backgroundImageFile = filePickerResult;
                    this.properties.backgroundImageUrl = filePickerResult.fileAbsoluteUrl || '';
                    this.forceRefresh();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    this.properties.backgroundImageFile = filePickerResult;
                  },
                  hideStockImages: false,
                  hideWebSearchTab: false,
                  hideLinkUploadTab: false,
                  hideOrganisationalAssetTab: false,
                  hideRecentTab: false,
                  hideOneDriveTab: false,
                  hideSiteFilesTab: false,
                  storeLastActiveTab: true,
                  required: false
                }),

                PropertyPaneButton('uploadBackgroundNative', {
                  text: '💻 Option 2: Upload from Computer',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleNativeBackgroundUpload.bind(this)
                }),

                PropertyPaneButton('backgroundUrlInput', {
                  text: '🔗 Option 3: Enter Image URL',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleBackgroundUrlInput.bind(this)
                }),

                PropertyPaneTextField('backgroundImageUrl', {
                  label: '🔗 Or paste Background URL directly',
                  placeholder: 'https://your-site/background.jpg',
                  value: this.properties.backgroundImageUrl,
                  multiline: true
                }),

                ...(this.properties.backgroundImageFile?.fileName || this.properties.backgroundImageUrl ? [
                  PropertyPaneTextField('currentBackground', {
                    label: '🎨 Current Selection',
                    value: this.properties.backgroundImageFile?.fileName || this.properties.backgroundImageUrl || 'None',
                    disabled: true
                  }),
                  PropertyPaneButton('clearBackgroundFile', {
                    text: '🗑️ Clear Background Selection',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.clearBackgroundFile.bind(this)
                  })
                ] : [])
              ]
            },
            {
              groupName: '⚙️ Display Settings',
              groupFields: [
                PropertyPaneToggle('showInModal', {
                  label: 'Play video in modal',
                  onText: 'Modal',
                  offText: 'Inline',
                  checked: this.properties.showInModal !== false
                }),
                PropertyPaneToggle('autoPlay', {
                  label: 'Auto-play video',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.autoPlay || false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log(`🔧 Property changed: ${propertyPath} = ${newValue}`);
    
    if (propertyPath.includes('File') || propertyPath === 'publishedDate') {
      this.forceRefresh();
    }
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }
}
