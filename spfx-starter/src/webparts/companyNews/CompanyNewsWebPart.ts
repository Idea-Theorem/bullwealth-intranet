import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import CompanyNews from './components/CompanyNews';
import { ICompanyNewsProps } from './components/ICompanyNewsProps';

export interface INewsItem {
  id: string;
  title: string;
  author: string;
  date: string;
  imageUrl: string;
  description: string;
  shareUrl?: string;
  readMoreUrl?: string;
}

export interface ICompanyNewsWebPartProps {
  title: string;
  itemsToShow: number;
  autoScroll: boolean;
  autoScrollInterval: number;
  totalNewsItems: number;
  
  // Support up to 8 news items
  newsItem1Title: string;
  newsItem1Author: string;
  newsItem1Date: string;
  newsItem1ImageUrl: string;
  newsItem1ReadMoreUrl: string;
  
  newsItem2Title: string;
  newsItem2Author: string;
  newsItem2Date: string;
  newsItem2ImageUrl: string;
  newsItem2ReadMoreUrl: string;
  
  newsItem3Title: string;
  newsItem3Author: string;
  newsItem3Date: string;
  newsItem3ImageUrl: string;
  newsItem3ReadMoreUrl: string;
  
  newsItem4Title: string;
  newsItem4Author: string;
  newsItem4Date: string;
  newsItem4ImageUrl: string;
  newsItem4ReadMoreUrl: string;
  
  newsItem5Title: string;
  newsItem5Author: string;
  newsItem5Date: string;
  newsItem5ImageUrl: string;
  newsItem5ReadMoreUrl: string;
  
  newsItem6Title: string;
  newsItem6Author: string;
  newsItem6Date: string;
  newsItem6ImageUrl: string;
  newsItem6ReadMoreUrl: string;
  
  newsItem7Title: string;
  newsItem7Author: string;
  newsItem7Date: string;
  newsItem7ImageUrl: string;
  newsItem7ReadMoreUrl: string;
  
  newsItem8Title: string;
  newsItem8Author: string;
  newsItem8Date: string;
  newsItem8ImageUrl: string;
  newsItem8ReadMoreUrl: string;
}

export default class CompanyNewsWebPart extends BaseClientSideWebPart<ICompanyNewsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // Build news items array from properties
    const allNewsItems: INewsItem[] = [];
    const defaultTitles = [
      'Quarterly Company Meeting',
      'Announcements', 
      'New HR Policy Update',
      'Wealth Podcasts',
      'Team Building Event',
      'Product Launch',
      'Training Workshop',
      'Company Awards'
    ];
    
    const defaultAuthors = [
      'John Doe',
      'Jane Smith',
      'HR Team',
      'Finance Team',
      'Events Team',
      'Product Team',
      'L&D Team',
      'Leadership Team'
    ];

    // Generate news items based on totalNewsItems setting
    const itemCount = this.properties.totalNewsItems || 4;
    
    for (let i = 1; i <= Math.min(itemCount, 8); i++) {
      // Use type assertion to access dynamic properties
      const props = this.properties as any;
      const item: INewsItem = {
        id: i.toString(),
        title: props[`newsItem${i}Title`] || defaultTitles[i - 1],
        author: props[`newsItem${i}Author`] || defaultAuthors[i - 1],
        date: props[`newsItem${i}Date`] || new Date().toLocaleDateString(),
        imageUrl: props[`newsItem${i}ImageUrl`] || '',
        description: `Description for ${defaultTitles[i - 1]}`,
        readMoreUrl: props[`newsItem${i}ReadMoreUrl`] || '#',
        shareUrl: ''
      };
      allNewsItems.push(item);
    }

    const element: React.ReactElement<ICompanyNewsProps> = React.createElement(
      CompanyNews,
      {
        title: this.properties.title || 'Company News & Updates',
        newsItems: allNewsItems,
        itemsToShow: this.properties.itemsToShow || 4,
        autoScroll: this.properties.autoScroll !== false,
        autoScrollInterval: this.properties.autoScrollInterval || 5000,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Set default values
    if (!this.properties.title) {
      this.properties.title = 'Company News & Updates';
    }
    if (this.properties.itemsToShow === undefined) {
      this.properties.itemsToShow = 4;
    }
    if (this.properties.totalNewsItems === undefined) {
      this.properties.totalNewsItems = 4;
    }
    if (this.properties.autoScroll === undefined) {
      this.properties.autoScroll = true;
    }
    if (!this.properties.autoScrollInterval) {
      this.properties.autoScrollInterval = 5000;
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

  private _getNewsItemGroups(): any[] {
    const groups = [];
    const itemCount = this.properties.totalNewsItems || 4;
    // Use type assertion to access dynamic properties
    const props = this.properties as any;
    
    for (let i = 1; i <= Math.min(itemCount, 8); i++) {
      groups.push({
        groupName: `News Item ${i}`,
        groupFields: [
          PropertyPaneTextField(`newsItem${i}Title`, {
            label: 'Title',
            placeholder: 'Enter news title',
            value: props[`newsItem${i}Title`]
          }),
          PropertyPaneTextField(`newsItem${i}Author`, {
            label: 'Author',
            placeholder: 'Enter author name',
            value: props[`newsItem${i}Author`]
          }),
          PropertyPaneTextField(`newsItem${i}Date`, {
            label: 'Date',
            placeholder: 'MM/DD/YYYY',
            value: props[`newsItem${i}Date`]
          }),
          PropertyPaneTextField(`newsItem${i}ImageUrl`, {
            label: 'Image URL',
            placeholder: 'https://your-site/image.jpg',
            value: props[`newsItem${i}ImageUrl`]
          }),
          PropertyPaneTextField(`newsItem${i}ReadMoreUrl`, {
            label: 'Read More URL',
            placeholder: 'https://your-site/article',
            value: props[`newsItem${i}ReadMoreUrl`]
          })
        ]
      });
    }
    
    return groups;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure Company News & Updates'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Section Title'
                }),
                PropertyPaneSlider('totalNewsItems', {
                  label: 'Total number of news items',
                  min: 1,
                  max: 8,
                  value: this.properties.totalNewsItems || 4,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneSlider('itemsToShow', {
                  label: 'Items visible at once',
                  min: 1,
                  max: 8,
                  value: this.properties.itemsToShow || 4,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneToggle('autoScroll', {
                  label: 'Auto-scroll',
                  checked: this.properties.autoScroll !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('autoScrollInterval', {
                  label: 'Auto-scroll interval (seconds)',
                  min: 2,
                  max: 10,
                  value: (this.properties.autoScrollInterval || 5000) / 1000,
                  showValue: true,
                  step: 1,
                  disabled: !this.properties.autoScroll
                })
              ]
            },
            ...this._getNewsItemGroups()
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'autoScrollInterval' && !isNaN(newValue)) {
      this.properties.autoScrollInterval = newValue * 1000;
    }
    
    if (propertyPath === 'totalNewsItems') {
      // Refresh property pane to show/hide news item groups
      this.context.propertyPane.refresh();
    }
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}