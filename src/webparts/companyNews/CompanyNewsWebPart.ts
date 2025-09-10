import * as React from 'react';
import * as ReactDom from 'react-dom';
// Import main styles (this loads the font faces)
import '../../styles/main.scss';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// Import PnP Property Controls
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { DatePicker } from '@fluentui/react/lib/DatePicker';

import CompanyNews from './components/CompanyNews';
import { ICompanyNewsProps, INewsItem } from './components/ICompanyNewsProps';

// Import PnP JS - FIXED
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/folders';
import '@pnp/sp/files';

import '@pnp/polyfill-ie11';

export interface ICompanyNewsWebPartProps {
  title: string;
  newsItems: INewsItem[];
  autoScroll: boolean;
  autoScrollInterval: number;
  itemsToShow: number;
  showDots: boolean;
  showArrows: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

export default class CompanyNewsWebPart extends BaseClientSideWebPart<ICompanyNewsWebPartProps> {

  private _sp: ReturnType<typeof spfi>;

  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // FIXED: Initialize PnP correctly
    this._sp = spfi().using(SPFx(this.context));

    // Initialize with sample data if empty
    if (!this.properties.newsItems || this.properties.newsItems.length === 0) {
      this.properties.newsItems = [
        {
          id: '1',
          title: 'Quarterly Company Meeting',
          author: 'John Doe',
          date: new Date().toISOString(),
          imageUrl: '',
          readMoreUrl: '#',
          shareUrl: ''
        },
        {
          id: '2',
          title: 'Announcements',
          author: 'John Doe',
          date: new Date().toISOString(),
          imageUrl: '',
          readMoreUrl: '#',
          shareUrl: ''
        },
        {
          id: '3',
          title: 'New HR Policy Update',
          author: 'John Doe',
          date: new Date().toISOString(),
          imageUrl: '',
          readMoreUrl: '#',
          shareUrl: ''
        },
        {
          id: '4',
          title: 'Wealth Podcasts',
          author: 'John Doe',
          date: new Date().toISOString(),
          imageUrl: '',
          readMoreUrl: '#',
          shareUrl: ''
        }
      ];
    }

    // Set default values - Remove arrows, keep dots only
    if (!this.properties.itemsToShow) this.properties.itemsToShow = 4;
    if (this.properties.showDots === undefined) this.properties.showDots = true;
    this.properties.showArrows = false; // ✅ Always disable arrows
  }

  public render(): void {
    const element: React.ReactElement<ICompanyNewsProps> = React.createElement(
      CompanyNews,
      {
        title: this.properties.title || 'Company News & Updates',
        newsItems: this.properties.newsItems || [],
        autoScroll: this.properties.autoScroll || false,
        autoScrollInterval: this.properties.autoScrollInterval || 5000,
        itemsToShow: this.properties.itemsToShow || 4,
        showDots: this.properties.showDots,
        showArrows: false, // ✅ Always false - no arrows
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // Method to upload image to SharePoint - FIXED
  private async uploadImageToSharePoint(file: File): Promise<string> {
    try {
      const timestamp = new Date().getTime();
      const fileName = `news_${timestamp}_${file.name}`;
      const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/SiteAssets/NewsImages`;
      
      try {
        await this._sp.web.getFolderByServerRelativePath(folderUrl)();
      } catch {
        await this._sp.web.folders.addUsingPath('SiteAssets/NewsImages');
      }
      
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
      await folder.files.addUsingPath(fileName, file, { Overwrite: true });
      
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      return `${siteUrl}/SiteAssets/NewsImages/${fileName}`;
      
    } catch (error) {
      console.error('Error uploading file:', error);
      throw new Error('Failed to upload image to SharePoint');
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, _oldValue: any, newValue: any): void {
    (this.properties as any)[propertyPath] = newValue;
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the Company News display'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneSlider('itemsToShow', {
                  label: 'Cards to Show at Once',
                  min: 1,
                  max: 8,
                  value: this.properties.itemsToShow || 4,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle('autoScroll', {
                  label: 'Auto Scroll Carousel',
                  checked: this.properties.autoScroll || false
                }),
                PropertyPaneSlider('autoScrollInterval', {
                  label: 'Auto Scroll Interval (seconds)',
                  min: 2,
                  max: 10,
                  value: (this.properties.autoScrollInterval || 5000) / 1000,
                  step: 1,
                  showValue: true,
                  disabled: !this.properties.autoScroll
                }),
                PropertyPaneToggle('showDots', {
                  label: 'Show Navigation Dots',
                  checked: this.properties.showDots
                })
              ]
            },
            {
              groupName: 'News Items Management',
              groupFields: [
                PropertyFieldCollectionData("newsItems", {
                  key: "newsItems",
                  label: "News Items",
                  panelHeader: "Manage News Items",
                  manageBtnLabel: "Manage News Items",
                  value: this.properties.newsItems,
                  fields: [
                    {
                      id: "title",
                      title: "Title",
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: "Enter news title"
                    },
                    {
                      id: "author",
                      title: "Author",
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: "Enter author name"
                    },
                    {
                      id: "date",
                      title: "Date",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return React.createElement("div", { style: { margin: "10px 0" } },
                          React.createElement("label", { 
                            style: { 
                              display: "block", 
                              fontWeight: 600, 
                              marginBottom: "5px", 
                              fontSize: "14px" 
                            } 
                          }, "Select Date"),
                          React.createElement(DatePicker, {
                            value: value ? new Date(value) : undefined,
                            placeholder: "Select a date",
                            onSelectDate: (date: Date | null | undefined) => {
                              if (date) {
                                onUpdate(field.id, date.toISOString());
                              }
                            },
                            formatDate: (date?: Date) => {
                              return date ? date.toLocaleDateString('en-US', { 
                                year: 'numeric', 
                                month: 'long', 
                                day: 'numeric' 
                              }) : '';
                            },
                            style: { 
                              width: "100%", 
                              padding: "8px", 
                              border: "1px solid #ccc", 
                              borderRadius: "4px" 
                            }
                          })
                        );
                      }
                    },
                    {
                      id: "imageUrl",
                      title: "Image",
                      type: CustomCollectionFieldType.custom,
                      required: false,
                      onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return React.createElement("div", { style: { margin: "10px 0" } },
                          value && React.createElement("img", {
                            src: value,
                            alt: "Preview",
                            style: { 
                              width: "120px", 
                              height: "80px", 
                              objectFit: "cover", 
                              marginBottom: "10px", 
                              display: "block",
                              border: "1px solid #ccc",
                              borderRadius: "4px"
                            }
                          }),
                          React.createElement("div", { style: { marginBottom: "10px" } },
                            React.createElement("input", {
                              type: "file",
                              accept: "image/*",
                              style: { 
                                padding: "8px",
                                border: "1px solid #ccc",
                                borderRadius: "4px",
                                width: "100%"
                              },
                              onChange: async (e: any) => {
                                const file = e.target.files[0];
                                if (file) {
                                  try {
                                    const uploadedUrl = await this.uploadImageToSharePoint(file);
                                    onUpdate(field.id, uploadedUrl);
                                  } catch (error) {
                                    alert('Upload failed: ' + (error as Error).message);
                                  }
                                }
                              }
                            })
                          ),
                          React.createElement("input", {
                            type: "text",
                            placeholder: "Or enter image URL",
                            value: value || "",
                            onChange: (e: any) => onUpdate(field.id, e.target.value),
                            style: { 
                              width: "100%", 
                              padding: "8px",
                              border: "1px solid #ccc",
                              borderRadius: "4px"
                            }
                          })
                        );
                      }
                    },
                    {
                      id: "readMoreUrl",
                      title: "Read More URL",
                      type: CustomCollectionFieldType.url,
                      required: false,
                      placeholder: "https://example.com/article"
                    },
                    {
                      id: "shareUrl",
                      title: "Share URL",
                      type: CustomCollectionFieldType.url,
                      required: false,
                      placeholder: "https://example.com/share"
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
