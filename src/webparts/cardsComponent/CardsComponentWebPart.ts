import * as React from 'react';
import * as ReactDom from 'react-dom';
// Import global styles - ADD THIS LINE
import '../../styles/main.scss';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CardsComponentWebPartStrings';
import CardsComponent from './components/CardsComponent';
import { ICardsComponentProps } from './components/ICardsComponentProps';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export default class CardsComponentWebPart extends BaseClientSideWebPart<ICardsComponentProps> {

  public render(): void {
    const element: React.ReactElement<ICardsComponentProps> = React.createElement(
      CardsComponent,
      {
        title: this.properties.title,
        intro: this.properties.intro,
        cards: this.properties.cards
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Main Title'
                }),
                PropertyPaneTextField('intro', {
                  label: 'Introductory Paragraph',
                  multiline: true
                }),
                PropertyFieldCollectionData('cards', {
                  key: 'cards',
                  label: 'Core Value Cards',
                  panelHeader: 'Edit Cards',
                  manageBtnLabel: 'Manage Cards',
                  value: this.properties.cards,
                  fields: [
                    {
                      id: 'cardTitle',
                      title: 'Card Title',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'cardDescription',
                      title: 'Card Description',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'iconUrl',
                      title: 'Icon',
                      type: CustomCollectionFieldType.custom,
                      required: true,
                     onCustomRender: (field, value, onUpdate, item, itemId, fieldId) => {
                                            return React.createElement('input', {
                                              type: 'file',
                                              accept: '.png,.jpg,.jpeg,.svg,.gif',
                                              onChange: (event: React.ChangeEvent<HTMLInputElement>) => {
                                                const file = event.target.files && event.target.files[0];
                                                if (file) {
                                                  const reader = new FileReader();
                                                  reader.onload = (e: ProgressEvent<FileReader>) => {
                                                    onUpdate(field.id, e.target?.result as string);
                                                  };
                                                  reader.readAsDataURL(file);
                                                }
                                              },
                                              style: { width: '100%' }
                                            });
                                          }
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