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
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import CompliancePolicies from './components/CompliancePolicies';
import { ICompliancePoliciesProps } from './components/ICompliancePoliciesProps';

export interface IPolicyDocument {
  id: string;
  title: string;
  dateAdded?: string;
  dateModified?: string;
  description: string;
  documentUrl: string;
  category?: string;
}

export interface ICompliancePoliciesWebPartProps {
  sectionTitle: string;
  categoryTitle: string;
  viewAllText: string;
  viewAllUrl: string;
  columnsPerRow: number;
  showExportButton: boolean;
  showShareButton: boolean;
  numberOfPolicies: number;
  
  // Support up to 12 policy documents
  policy1Title: string;
  policy1Description: string;
  policy1Date: string;
  policy1Url: string;
  
  policy2Title: string;
  policy2Description: string;
  policy2Date: string;
  policy2Url: string;
  
  policy3Title: string;
  policy3Description: string;
  policy3Date: string;
  policy3Url: string;
  
  policy4Title: string;
  policy4Description: string;
  policy4Date: string;
  policy4Url: string;
  
  policy5Title: string;
  policy5Description: string;
  policy5Date: string;
  policy5Url: string;
  
  policy6Title: string;
  policy6Description: string;
  policy6Date: string;
  policy6Url: string;
  
  policy7Title: string;
  policy7Description: string;
  policy7Date: string;
  policy7Url: string;
  
  policy8Title: string;
  policy8Description: string;
  policy8Date: string;
  policy8Url: string;
  
  policy9Title: string;
  policy9Description: string;
  policy9Date: string;
  policy9Url: string;
  
  policy10Title: string;
  policy10Description: string;
  policy10Date: string;
  policy10Url: string;
  
  policy11Title: string;
  policy11Description: string;
  policy11Date: string;
  policy11Url: string;
  
  policy12Title: string;
  policy12Description: string;
  policy12Date: string;
  policy12Url: string;
}

export default class CompliancePoliciesWebPart extends BaseClientSideWebPart<ICompliancePoliciesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const policies: IPolicyDocument[] = [];
    const props = this.properties as any;
    
    // Build policies array from properties
    for (let i = 1; i <= Math.min(this.properties.numberOfPolicies || 3, 12); i++) {
      if (props[`policy${i}Title`]) {
        policies.push({
          id: i.toString(),
          title: props[`policy${i}Title`],
          description: props[`policy${i}Description`] || 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
          dateAdded: props[`policy${i}Date`] ? `Added ${props[`policy${i}Date`]}` : undefined,
          dateModified: props[`policy${i}Date`] ? `Modified ${props[`policy${i}Date`]}` : undefined,
          documentUrl: props[`policy${i}Url`] || '#',
          category: this.properties.categoryTitle
        });
      }
    }

    // Add default policies if none configured
    if (policies.length === 0) {
      policies.push(
        {
          id: '1',
          title: 'Record Keeping Policy',
          dateAdded: 'Added Dec 22, 2023',
          description: 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
          documentUrl: '#'
        },
        {
          id: '2',
          title: 'Cross Trade Policy',
          dateModified: 'Modified Dec 22, 2023',
          description: 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
          documentUrl: '#'
        },
        {
          id: '3',
          title: 'Compliance Manual Deck',
          dateModified: 'Modified Dec 22, 2023',
          description: 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
          documentUrl: '#'
        }
      );
    }

    const element: React.ReactElement<ICompliancePoliciesProps> = React.createElement(
      CompliancePolicies,
      {
        sectionTitle: this.properties.sectionTitle || 'Compliance',
        categoryTitle: this.properties.categoryTitle || 'Policies',
        viewAllText: this.properties.viewAllText || 'View all',
        viewAllUrl: this.properties.viewAllUrl || '#',
        policies: policies,
        columnsPerRow: this.properties.columnsPerRow || 3,
        showExportButton: this.properties.showExportButton !== false,
        showShareButton: this.properties.showShareButton !== false,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Set default values
    if (!this.properties.sectionTitle) {
      this.properties.sectionTitle = 'Compliance';
    }
    if (!this.properties.categoryTitle) {
      this.properties.categoryTitle = 'Policies';
    }
    if (!this.properties.viewAllText) {
      this.properties.viewAllText = 'View all';
    }
    if (!this.properties.columnsPerRow) {
      this.properties.columnsPerRow = 3;
    }
    if (!this.properties.numberOfPolicies) {
      this.properties.numberOfPolicies = 3;
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

  private _getPolicyGroups(): any[] {
    const groups = [];
    const numberOfPolicies = this.properties.numberOfPolicies || 3;
    
    for (let i = 1; i <= Math.min(numberOfPolicies, 12); i++) {
      groups.push({
        groupName: `Policy ${i}`,
        groupFields: [
          PropertyPaneTextField(`policy${i}Title`, {
            label: 'Title',
            placeholder: 'Enter policy title'
          }),
          PropertyPaneTextField(`policy${i}Description`, {
            label: 'Description',
            multiline: true,
            rows: 3,
            placeholder: 'Enter policy description'
          }),
          PropertyPaneTextField(`policy${i}Date`, {
            label: 'Date',
            placeholder: 'Dec 22, 2023'
          }),
          PropertyPaneTextField(`policy${i}Url`, {
            label: 'Document URL',
            placeholder: 'https://your-site/document.pdf'
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
            description: 'Configure Compliance Policies Display'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('sectionTitle', {
                  label: 'Section Title'
                }),
                PropertyPaneTextField('categoryTitle', {
                  label: 'Category Title'
                }),
                PropertyPaneTextField('viewAllText', {
                  label: 'View All Button Text'
                }),
                PropertyPaneTextField('viewAllUrl', {
                  label: 'View All URL',
                  placeholder: '/sites/compliance/documents'
                }),
                PropertyPaneSlider('numberOfPolicies', {
                  label: 'Number of Policies',
                  min: 1,
                  max: 12,
                  value: this.properties.numberOfPolicies || 3,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per Row',
                  min: 1,
                  max: 4,
                  value: this.properties.columnsPerRow || 3,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneToggle('showExportButton', {
                  label: 'Show Export Button',
                  checked: this.properties.showExportButton !== false,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('showShareButton', {
                  label: 'Show Share Button',
                  checked: this.properties.showShareButton !== false,
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            },
            ...this._getPolicyGroups()
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'numberOfPolicies') {
      // Refresh property pane to show/hide policy groups
      this.context.propertyPane.refresh();
    }
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}