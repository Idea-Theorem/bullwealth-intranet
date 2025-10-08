import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneToggle,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import CarouselCards from './components/CarouselCards';
import { ICarouselCardsProps, ICarouselCard } from './components/ICarouselCardsProps';

export interface ICarouselCardsWebPartProps {
  title: string;
  subtitle: string;
  cards: string;
}

export default class CarouselCardsWebPart extends BaseClientSideWebPart<ICarouselCardsWebPartProps> {

  private _cardsArray: ICarouselCard[] = [];

  protected onInit(): Promise<void> {
    if (this.properties.cards) {
      try {
        this._cardsArray = JSON.parse(this.properties.cards);
      } catch (error) {
        console.error('Error parsing cards:', error);
        this._cardsArray = this._getDefaultCards();
      }
    } else {
      this._cardsArray = this._getDefaultCards();
    }

    if (!this.properties.title) {
      this.properties.title = 'Corporate Values - Definition and Behaviours';
    }
    if (!this.properties.subtitle) {
      this.properties.subtitle = 'At Bullwealth, our core values guide everything we do. They define who we are as a company and how we serve our clients.';
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ICarouselCardsProps> = React.createElement(
      CarouselCards,
      {
        title: this.properties.title,
        subtitle: this.properties.subtitle,
        cards: this._cardsArray
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    const cardTitleMatch = propertyPath.match(/^cardTitle(\d+)$/);
    const cardIconMatch = propertyPath.match(/^cardIcon(\d+)$/);
    const cardIconColorMatch = propertyPath.match(/^cardIconColor(\d+)$/);
    const cardDescriptionMatch = propertyPath.match(/^cardDescription(\d+)$/);
    const cardBulletMatch = propertyPath.match(/^cardBullet(\d+)_(\d+)$/);
    const cardVisibilityMatch = propertyPath.match(/^cardVisible(\d+)$/);

    if (cardTitleMatch) {
      const cardIndex = parseInt(cardTitleMatch[1]);
      this._updateCard(cardIndex, 'title', newValue);
    } else if (cardIconMatch) {
      const cardIndex = parseInt(cardIconMatch[1]);
      this._updateCard(cardIndex, 'icon', newValue);
    } else if (cardIconColorMatch) {
      const cardIndex = parseInt(cardIconColorMatch[1]);
      this._updateCard(cardIndex, 'iconColor', newValue);
    } else if (cardDescriptionMatch) {
      const cardIndex = parseInt(cardDescriptionMatch[1]);
      this._updateCard(cardIndex, 'description', newValue);
    } else if (cardBulletMatch) {
      const cardIndex = parseInt(cardBulletMatch[1]);
      const pointIndex = parseInt(cardBulletMatch[2]);
      this._updateBulletPoint(cardIndex, pointIndex, newValue);
    } else if (cardVisibilityMatch) {
      const cardIndex = parseInt(cardVisibilityMatch[1]);
      this._updateCard(cardIndex, 'isVisible', newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _getDefaultCards(): ICarouselCard[] {
    return [
      {
        id: '1',
        icon: 'Lightbulb',
        iconType: 'fluent',
        iconColor: '#90EE90',
        title: 'Insight',
        description: 'We invest our time to develop and share our knowledge and insights, so together, we can make the most informed decisions.',
        bulletPoints: [
          'Our knowledge base remains current and aligned with evolving client needs',
          'We take time to assess situations through reflection',
          'We ask, encourage and consider different perspectives',
          'Our experiences are consistently shared with colleagues',
          'We use exploratory questions to deepen understanding',
          'We encourage knowledge expansion through best practices'
        ],
        isVisible: true
      },
      {
        id: '2',
        icon: 'People',
        iconType: 'fluent',
        iconColor: '#FFB6C1',
        title: 'Collaboration',
        description: 'Working together towards common goals, sharing knowledge and expertise to achieve better outcomes.',
        bulletPoints: [
          'Foster open communication across all levels',
          'Encourage team participation and input',
          'Share resources and information freely',
          'Build strong working relationships',
          'Support collective decision-making',
          'Celebrate team achievements together'
        ],
        isVisible: true
      },
      {
        id: '3',
        icon: 'Trophy',
        iconType: 'fluent',
        iconColor: '#FFD700',
        title: 'Excellence',
        description: 'Striving for the highest standards in everything we do, continuously improving and innovating.',
        bulletPoints: [
          'Maintain high quality standards',
          'Pursue continuous improvement',
          'Embrace innovation and creativity',
          'Take ownership of outcomes',
          'Learn from successes and failures',
          'Deliver exceptional results consistently'
        ],
        isVisible: true
      }
    ];
  }

  private _addNewCard(): void {
    const newCard: ICarouselCard = {
      id: Date.now().toString(),
      icon: 'Lightbulb',
      iconType: 'fluent',
      iconColor: '#90EE90',
      title: 'New Card',
      description: 'Enter card description here',
      bulletPoints: ['Point 1', 'Point 2', 'Point 3'],
      isVisible: true
    };

    this._cardsArray.push(newCard);
    this.properties.cards = JSON.stringify(this._cardsArray);
    this.render();
    this.context.propertyPane.refresh();
  }

  private _updateCard(index: number, field: string, value: any): void {
    if (this._cardsArray[index]) {
      (this._cardsArray[index] as any)[field] = value;
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
    }
  }

  private _updateBulletPoint(cardIndex: number, pointIndex: number, value: string): void {
    if (this._cardsArray[cardIndex] && this._cardsArray[cardIndex].bulletPoints[pointIndex] !== undefined) {
      this._cardsArray[cardIndex].bulletPoints[pointIndex] = value;
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
    }
  }

  private _addBulletPoint(cardIndex: number): void {
    if (this._cardsArray[cardIndex]) {
      this._cardsArray[cardIndex].bulletPoints.push('New point');
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
      this.context.propertyPane.refresh();
    }
  }

  private _removeBulletPoint(cardIndex: number, pointIndex: number): void {
    if (this._cardsArray[cardIndex]) {
      this._cardsArray[cardIndex].bulletPoints.splice(pointIndex, 1);
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
      this.context.propertyPane.refresh();
    }
  }

  private _deleteCard(index: number): void {
    this._cardsArray.splice(index, 1);
    this.properties.cards = JSON.stringify(this._cardsArray);
    this.render();
    this.context.propertyPane.refresh();
  }

  private _moveCardUp(index: number): void {
    if (index > 0) {
      const temp = this._cardsArray[index];
      this._cardsArray[index] = this._cardsArray[index - 1];
      this._cardsArray[index - 1] = temp;
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
      this.context.propertyPane.refresh();
    }
  }

  private _moveCardDown(index: number): void {
    if (index < this._cardsArray.length - 1) {
      const temp = this._cardsArray[index];
      this._cardsArray[index] = this._cardsArray[index + 1];
      this._cardsArray[index + 1] = temp;
      this.properties.cards = JSON.stringify(this._cardsArray);
      this.render();
      this.context.propertyPane.refresh();
    }
  }

  private _switchToFluentIcon(cardIndex: number): void {
    this._updateCard(cardIndex, 'iconType', 'fluent');
    this._updateCard(cardIndex, 'icon', 'Lightbulb');
    this.context.propertyPane.refresh();
  }

  private _switchToUploadIcon(cardIndex: number): void {
    this._updateCard(cardIndex, 'iconType', 'upload');
    this._handleImageUpload(cardIndex);
  }

  private _handleImageUpload(cardIndex: number): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    
    input.onchange = (e: Event) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event: ProgressEvent<FileReader>) => {
          const base64 = event.target?.result as string;
          this._updateCard(cardIndex, 'icon', base64);
          this._updateCard(cardIndex, 'iconType', 'upload');
          this.context.propertyPane.refresh();
        };
        reader.readAsDataURL(file);
      }
    };
    
    input.click();
  }

  private _getCardPropertyFields(): IPropertyPaneField<any>[] {
    const fields: IPropertyPaneField<any>[] = [];

    this._cardsArray.forEach((card, cardIndex) => {
      fields.push(
        PropertyPaneButton(`card${cardIndex}Header`, {
          text: `Card ${cardIndex + 1}: ${card.title}`,
          buttonType: PropertyPaneButtonType.Hero,
          onClick: (): void => {
            // Header button
          }
        })
      );

      fields.push(
        PropertyPaneToggle(`cardVisible${cardIndex}`, {
          label: 'Show Card',
          checked: card.isVisible !== false,
          onText: 'Visible',
          offText: 'Hidden'
        })
      );

      fields.push(
        PropertyPaneTextField(`cardTitle${cardIndex}`, {
          label: 'Card Title',
          value: card.title,
          onGetErrorMessage: (value: string): string => {
            if (!value || value.trim() === '') {
              return 'Title is required';
            }
            return '';
          },
          deferredValidationTime: 500
        })
      );

      if (card.iconType === 'upload') {
        fields.push(
          PropertyPaneButton(`switchToFluent${cardIndex}`, {
            text: 'Switch to Fluent UI Icon',
            buttonType: PropertyPaneButtonType.Normal,
            onClick: (): void => {
              this._switchToFluentIcon(cardIndex);
            }
          })
        );

        fields.push(
          PropertyPaneButton(`uploadIcon${cardIndex}`, {
            text: card.icon && card.icon.startsWith('data:') ? 'Image Uploaded - Click to Change' : 'Upload Icon Image',
            buttonType: PropertyPaneButtonType.Normal,
            onClick: (): void => {
              this._handleImageUpload(cardIndex);
            }
          })
        );
      } else {
        fields.push(
          PropertyPaneTextField(`cardIcon${cardIndex}`, {
            label: 'Icon Name (Fluent UI)',
            value: card.icon,
            description: 'E.g., Lightbulb, People, Trophy, Heart, Rocket, Bullseye'
          })
        );

        fields.push(
          PropertyPaneButton(`switchToUpload${cardIndex}`, {
            text: 'Switch to Upload Image',
            buttonType: PropertyPaneButtonType.Normal,
            onClick: (): void => {
              this._switchToUploadIcon(cardIndex);
            }
          })
        );
      }

      fields.push(
        PropertyPaneTextField(`cardIconColor${cardIndex}`, {
          label: 'Icon Background Color',
          value: card.iconColor,
          description: 'Hex color code (e.g., #90EE90)'
        })
      );

      fields.push(
        PropertyPaneTextField(`cardDescription${cardIndex}`, {
          label: 'Description',
          value: card.description,
          multiline: true,
          rows: 3
        })
      );

      card.bulletPoints.forEach((point, pointIndex) => {
        fields.push(
          PropertyPaneTextField(`cardBullet${cardIndex}_${pointIndex}`, {
            label: `Bullet Point ${pointIndex + 1}`,
            value: point,
            multiline: true,
            rows: 2
          })
        );

        fields.push(
          PropertyPaneButton(`removeBullet${cardIndex}_${pointIndex}`, {
            text: 'Remove Point',
            buttonType: PropertyPaneButtonType.Normal,
            onClick: (): void => {
              this._removeBulletPoint(cardIndex, pointIndex);
            }
          })
        );
      });

      fields.push(
        PropertyPaneButton(`addBullet${cardIndex}`, {
          text: 'Add Bullet Point',
          buttonType: PropertyPaneButtonType.Normal,
          onClick: (): void => {
            this._addBulletPoint(cardIndex);
          }
        })
      );

      fields.push(
        PropertyPaneButton(`moveUp${cardIndex}`, {
          text: 'Move Up',
          buttonType: PropertyPaneButtonType.Normal,
          disabled: cardIndex === 0,
          onClick: (): void => {
            this._moveCardUp(cardIndex);
          }
        })
      );

      fields.push(
        PropertyPaneButton(`moveDown${cardIndex}`, {
          text: 'Move Down',
          buttonType: PropertyPaneButtonType.Normal,
          disabled: cardIndex === this._cardsArray.length - 1,
          onClick: (): void => {
            this._moveCardDown(cardIndex);
          }
        })
      );

      fields.push(
        PropertyPaneButton(`deleteCard${cardIndex}`, {
          text: 'Delete Card',
          buttonType: PropertyPaneButtonType.Normal,
          onClick: (): void => {
            const confirmed = confirm(`Are you sure you want to delete "${card.title}"?`);
            if (confirmed) {
              this._deleteCard(cardIndex);
            }
          }
        })
      );
    });

    return fields;
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
            description: 'Configure your carousel cards'
          },
          groups: [
            {
              groupName: 'Header Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Main Title',
                  value: this.properties.title,
                  placeholder: 'Corporate Values - Definition and Behaviours'
                }),
                PropertyPaneTextField('subtitle', {
                  label: 'Subtitle',
                  value: this.properties.subtitle,
                  multiline: true,
                  rows: 3,
                  placeholder: 'At Bullwealth, our core values guide everything we do...'
                })
              ]
            },
            {
              groupName: 'Carousel Settings',
              groupFields: [
                PropertyPaneButton('addCard', {
                  text: 'Add New Card',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: (): void => {
                    this._addNewCard();
                  }
                })
              ]
            },
            {
              groupName: 'Cards Configuration',
              groupFields: this._getCardPropertyFields()
            }
          ]
        }
      ]
    };
  }
}
