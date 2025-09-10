import * as React from 'react';
import * as ReactDom from 'react-dom';
<<<<<<< HEAD
// Import global styles - ADD THIS LINE
import '../../styles/main.scss';

=======
>>>>>>> 9c3c809eaa69f41d431d5185d9da9217288dffec
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import Documents from './components/Documents';

export interface IDocumentCategory {
  id: string;
  title: string;
  imageData: string; // Base64 encoded image
  libraryUrl: string; // Main URL field (changed from libraryName)
  viewAllUrl?: string; // Optional: Alternative view URL
}

export interface IDocumentsWebPartProps {
  title: string;
  columnsPerRow: number;
  categories: string; // JSON string to store categories
}

export default class DocumentsWebPart extends BaseClientSideWebPart<IDocumentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    let categories: IDocumentCategory[] = [];
    
    try {
      categories = this.properties.categories ? JSON.parse(this.properties.categories) : this.getDefaultCategories();
    } catch {
      categories = this.getDefaultCategories();
    }

    const element = React.createElement(
      Documents,
      {
        title: this.properties.title || 'Documents', // ✅ Fixed: Now matches interface
        categories: categories,
        columnsPerRow: this.properties.columnsPerRow || 4,
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        onCategoriesUpdate: (updatedCategories: IDocumentCategory[]) => {
          this.properties.categories = JSON.stringify(updatedCategories);
          this.context.propertyPane.refresh();
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getDefaultCategories(): IDocumentCategory[] {
    const baseUrl = this.context.pageContext.web.absoluteUrl;
    return [
      {
        id: '1',
        title: 'Compliance',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Compliance`,
        viewAllUrl: ''
      },
      {
        id: '2',
        title: 'Research & Investments',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Research`,
        viewAllUrl: ''
      },
      {
        id: '3',
        title: 'Advisory Group',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Advisory`,
        viewAllUrl: ''
      },
      {
        id: '4',
        title: 'Operations',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Operations`,
        viewAllUrl: ''
      },
      {
        id: '5',
        title: 'Business Development',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Business`,
        viewAllUrl: ''
      },
      {
        id: '6',
        title: 'Tax & Accounting',
        imageData: '',
        libraryUrl: `${baseUrl}/Shared Documents/Tax`,
        viewAllUrl: ''
      }
    ];
  }

  private _uploadImage = (categoryIndex: number): void => {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
          const categories: IDocumentCategory[] = JSON.parse(this.properties.categories || '[]');
          if (categories[categoryIndex]) {
            categories[categoryIndex].imageData = event.target?.result as string;
            this.properties.categories = JSON.stringify(categories);
            this.context.propertyPane.refresh();
            this.render();
          }
        };
        reader.readAsDataURL(file);
      }
    };
    fileInput.click();
  }

  // Helper function to validate URL
  private _isValidUrl = (url: string): boolean => {
    try {
      if (!url) return false;
      // Allow relative URLs or full URLs
      if (url.startsWith('/') || url.startsWith('http://') || url.startsWith('https://')) {
        return true;
      }
      new URL(url);
      return true;
    } catch {
      return false;
    }
  }

  // ✅ FIXED: Removed unused _formatUrl function

  protected onInit(): Promise<void> {
    // Set default values
    if (!this.properties.title) {
      this.properties.title = 'Documents';
    }
    if (!this.properties.columnsPerRow) {
      this.properties.columnsPerRow = 4;
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
    const categories: IDocumentCategory[] = JSON.parse(this.properties.categories || '[]');
    
    const categoryGroups = categories.map((category, index) => ({
      groupName: `Category: ${category.title}`,
      groupFields: [
        PropertyPaneTextField(`tempTitle_${index}`, {
          label: 'Title',
          value: category.title,
          onGetErrorMessage: (value: string) => {
            if (!value) return 'Title is required';
            const cats = JSON.parse(this.properties.categories || '[]');
            cats[index].title = value;
            this.properties.categories = JSON.stringify(cats);
            this.render();
            return '';
          }
        }),
        // ✅ UPDATED: URL field instead of Library Name
        PropertyPaneTextField(`tempLibraryUrl_${index}`, {
          label: 'Library URL',
          value: category.libraryUrl,
          placeholder: 'https://yourtenant.sharepoint.com/sites/yoursite/DocumentLibrary',
          description: 'Enter the complete URL to the SharePoint document library',
          multiline: false,
          onGetErrorMessage: (value: string) => {
            const cats = JSON.parse(this.properties.categories || '[]');
            cats[index].libraryUrl = value;
            this.properties.categories = JSON.stringify(cats);
            this.render();
            
            // Validate URL format
            if (value && !this._isValidUrl(value)) {
              return 'Please enter a valid URL (e.g., https://yourtenant.sharepoint.com/sites/yoursite/DocumentLibrary)';
            }
            return '';
          }
        }),
        // ✅ ADDED: Optional View All URL field
        PropertyPaneTextField(`tempViewAllUrl_${index}`, {
          label: 'View All URL (Optional)',
          value: category.viewAllUrl || '',
          placeholder: 'https://yourtenant.sharepoint.com/sites/yoursite/DocumentLibrary/Forms/AllItems.aspx',
          description: 'Optional: Custom URL for "View Documents" link',
          multiline: false,
          onGetErrorMessage: (value: string) => {
            const cats = JSON.parse(this.properties.categories || '[]');
            cats[index].viewAllUrl = value;
            this.properties.categories = JSON.stringify(cats);
            this.render();
            
            // Validate URL format if provided
            if (value && !this._isValidUrl(value)) {
              return 'Please enter a valid URL';
            }
            return '';
          }
        }),
        PropertyPaneButton(`uploadImage_${index}`, {
          text: category.imageData ? 'Change Image' : 'Upload Image',
          buttonType: PropertyPaneButtonType.Normal,
          onClick: () => this._uploadImage(index)
        }),
        PropertyPaneButton(`deleteCategory_${index}`, {
          text: 'Delete Category',
          buttonType: PropertyPaneButtonType.Normal,
          onClick: () => {
            if (confirm('Are you sure you want to delete this category?')) {
              const cats = JSON.parse(this.properties.categories || '[]');
              cats.splice(index, 1);
              this.properties.categories = JSON.stringify(cats);
              this.context.propertyPane.refresh();
              this.render();
            }
          }
        })
      ]
    }));

    return {
      pages: [
        {
          header: {
            description: 'Configure Document Categories'
          },
          groups: [
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per row',
                  min: 2,
                  max: 6,
                  value: this.properties.columnsPerRow || 4,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneButton('addCategory', {
                  text: 'Add New Category',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    const cats = JSON.parse(this.properties.categories || '[]');
                    const baseUrl = this.context.pageContext.web.absoluteUrl;
                    cats.push({
                      id: Date.now().toString(),
                      title: 'New Category',
                      imageData: '',
                      libraryUrl: `${baseUrl}/Shared Documents`,
                      viewAllUrl: ''
                    });
                    this.properties.categories = JSON.stringify(cats);
                    this.context.propertyPane.refresh();
                    this.render();
                  }
                })
              ]
            },
            ...categoryGroups
          ]
        }
      ]
    };
  }
}
