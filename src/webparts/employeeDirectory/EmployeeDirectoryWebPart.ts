import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';

export interface IEmployeeDirectoryWebPartProps {
  title: string;
  maxEmployeesToShow: number;
  orgChartLink: string;
  listName: string;
  selectedCompany: string;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {

  private companyOptions: IPropertyPaneDropdownOption[] = [];
  private companiesFetched: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // ✅ Set default list name if not set
    if (!this.properties.listName) {
      this.properties.listName = 'Employees';
    }
    
    // ✅ Set default to show All if not set
    if (!this.properties.selectedCompany) {
      this.properties.selectedCompany = 'All';
    }
    
    // ✅ Auto-fetch companies on initialization
    await this.fetchCompanyOptions();
  }

  public render(): void {
    const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(
      EmployeeDirectory,
      {
        title: this.properties.title || 'Employee Directory',
        maxEmployeesToShow: this.properties.maxEmployeesToShow || 5,
        orgChartLink: this.properties.orgChartLink || '',
        listName: this.properties.listName || 'Employees', // ✅ Default
        selectedCompany: this.properties.selectedCompany || 'All', // ✅ Default shows all
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async fetchCompanyOptions(): Promise<void> {
    // ✅ Always use 'Employees' as default list name
    const listToFetch = this.properties.listName || 'Employees';
    
    if (!listToFetch || this.companiesFetched) {
      return;
    }

    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listToFetch)}')/items?$select=CompanyName&$top=1000`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        
        const companies = Array.from(
          new Set(
            data.value
              .map((item: any) => item.CompanyName)
              .filter((name: string) => name && name.trim() !== '')
          )
        ).sort();

        this.companyOptions = [
          { key: 'All', text: 'All Employees' }, // ✅ Default option
          ...companies.map((company: any) => ({
            key: company,
            text: company
          }))
        ];

        this.companiesFetched = true;
        console.log('✅ Fetched company options:', this.companyOptions);
      }
    } catch (error) {
      console.error('❌ Error fetching companies:', error);
      this.companyOptions = [{ key: 'All', text: 'All Employees' }];
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'listName' && newValue !== oldValue) {
      // ✅ Reset company options when list name changes
      this.companiesFetched = false;
      this.properties.selectedCompany = 'All'; // ✅ Reset to All
      await this.fetchCompanyOptions();
      this.context.propertyPane.refresh();
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (!this.companiesFetched) {
      await this.fetchCompanyOptions();
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure Employee Directory"
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: this.properties.title || 'Employee Directory'
                }),
                PropertyPaneTextField('orgChartLink', {
                  label: 'Organization Chart Link (URL)',
                  placeholder: 'https://your-org-chart-link'
                }),
                PropertyPaneSlider('maxEmployeesToShow', {
                  label: 'Max Employees Per Page',
                  min: 1,
                  max: 20,
                  value: this.properties.maxEmployeesToShow || 5,
                  showValue: true
                })
              ]
            },
            {
              groupName: "Data Source",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name',
                  description: 'Default: Employees',
                  placeholder: 'Employees',
                  value: this.properties.listName || 'Employees' // ✅ Show default
                }),
                PropertyPaneDropdown('selectedCompany', {
                  label: 'Filter by Company Name',
                  selectedKey: this.properties.selectedCompany || 'All',
                  options: this.companyOptions.length > 0 
                    ? this.companyOptions 
                    : [{ key: 'All', text: 'Loading companies...' }],
                  disabled: !this.companiesFetched // ✅ Disable until loaded
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
