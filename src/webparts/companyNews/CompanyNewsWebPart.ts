import * as React from 'react';
import * as ReactDom from 'react-dom';
<<<<<<< HEAD
// Import main styles (this loads the font faces)
import '../../styles/main.scss';
=======
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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

<<<<<<< HEAD
// Import PnP JS - FIXED
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/folders';
import '@pnp/sp/files';

=======
// Import PnP JS
import { sp } from '@pnp/sp/presets/all';
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
import '@pnp/polyfill-ie11';

export interface ICompanyNewsWebPartProps {
  title: string;
  newsItems: INewsItem[];
  autoScroll: boolean;
  autoScrollInterval: number;
  itemsToShow: number;
  showDots: boolean;
  showArrows: boolean;
<<<<<<< HEAD
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
=======
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
  [key: string]: any;
}

export default class CompanyNewsWebPart extends BaseClientSideWebPart<ICompanyNewsWebPartProps> {

<<<<<<< HEAD
  private _sp: ReturnType<typeof spfi>;

  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // FIXED: Initialize PnP correctly
    this._sp = spfi().using(SPFx(this.context));
=======
  protected async onInit(): Promise<void> {
    await super.onInit();
    
    sp.setup({
      spfxContext: this.context as any
    });
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec

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

<<<<<<< HEAD
  // Method to upload image to SharePoint - FIXED
=======
  // Method to upload image to SharePoint
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
  private async uploadImageToSharePoint(file: File): Promise<string> {
    try {
      const timestamp = new Date().getTime();
      const fileName = `news_${timestamp}_${file.name}`;
      const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/SiteAssets/NewsImages`;
      
      try {
<<<<<<< HEAD
        await this._sp.web.getFolderByServerRelativePath(folderUrl)();
      } catch {
        await this._sp.web.folders.addUsingPath('SiteAssets/NewsImages');
      }
      
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
=======
        await sp.web.getFolderByServerRelativeUrl(folderUrl).get();
      } catch {
        await sp.web.folders.addUsingPath(folderUrl);
      }
      
      const folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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

<<<<<<< HEAD
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
=======
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
<<<<<<< HEAD
=======
                // ✅ Removed showArrows toggle - arrows always hidden
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
<<<<<<< HEAD
=======
                    // ✅ Date Picker Field
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
<<<<<<< HEAD
=======
                    // ✅ Image Upload Field
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
                    {
                      id: "imageUrl",
                      title: "Image",
                      type: CustomCollectionFieldType.custom,
                      required: false,
                      onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return React.createElement("div", { style: { margin: "10px 0" } },
<<<<<<< HEAD
=======
                          // Image preview
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
<<<<<<< HEAD
=======
                          // File upload button
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
<<<<<<< HEAD
=======
                          // URL input as fallback
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
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
