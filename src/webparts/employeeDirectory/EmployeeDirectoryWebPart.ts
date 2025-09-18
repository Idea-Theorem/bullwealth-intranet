import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps, ISection } from './components/IEmployeeDirectoryProps';

export interface IEmployeeDirectoryWebPartProps {
  title: string;
  maxEmployeesToShow: number;
  orgChartLink: string;
  sections: ISection[];
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(
      EmployeeDirectory,
      {
        title: this.properties.title || 'Employee Directory',
        maxEmployeesToShow: this.properties.maxEmployeesToShow || 5,
        orgChartLink: this.properties.orgChartLink || '',
        sections: this.properties.sections || [
          { title: 'Compliance', listName: 'Employee - Compliance' },
          { title: 'Research & Investment', listName: 'Employee - Research & Investment' }
        ],
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
            description: "Configure Employee Directory"
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('orgChartLink', {
                  label: 'Org Chart Link (URL)'
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
              groupName: "Department Sections",
              groupFields: [
                PropertyFieldCollectionData('sections', {
                  key: 'sections',
                  label: 'Employee Sections',
                  panelHeader: 'Configure directory sections',
                  manageBtnLabel: 'Manage Sections',
                  value: this.properties.sections,
                  fields: [
                    {
                      id: 'title',
                      title: 'Section Title',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'listName',
                      title: 'SharePoint List Name',
                      type: CustomCollectionFieldType.string,
                      required: true
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
