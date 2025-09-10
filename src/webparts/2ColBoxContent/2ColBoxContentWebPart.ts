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
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TwoColBoxContentWebPartStrings';
// Make sure the file exists at the correct path and is named 'TwoColBoxContent.tsx' or 'TwoColBoxContent.ts'
// Update the import path or file name if needed, for example:
import TwoColBoxContent from './components/2ColBoxContent';

// Or, if the file is named differently, update accordingly:
// import TwoColBoxContent from './components/TwoColBoxContentComponent';
// Update the import path below if your file is named differently, e.g., 'TwoColBoxContentProps.ts'
import { I2ColBoxContentProps, IContactCard } from './components/I2ColBoxContentProps';

export interface ITwoColBoxContentWebPartProps {
  // Layout properties
  columnLayout: string;
  containerBackgroundColor: string;
  cardSpacing: number;
  
  // Left card properties
  leftCardTitle: string;
  leftCardSubtitle: string;
  leftCardName: string;
  leftCardEmail: string;
  leftCardPhone: string;
  leftCardEmailButtonText: string;
  leftCardPhoneButtonText: string;
  leftCardShowEmailButton: boolean;
  leftCardShowPhoneButton: boolean;
  leftCardBackgroundColor: string;
  leftCardTitleColor: string;
  leftCardSubtitleColor: string;
  leftCardNameColor: string;
  leftCardContactColor: string;
  leftCardEmailButtonColor: string;
  leftCardPhoneButtonColor: string;
  
  // Right card properties
  rightCardTitle: string;
  rightCardSubtitle: string;
  rightCardName: string;
  rightCardEmail: string;
  rightCardPhone: string;
  rightCardEmailButtonText: string;
  rightCardPhoneButtonText: string;
  rightCardShowEmailButton: boolean;
  rightCardShowPhoneButton: boolean;
  rightCardBackgroundColor: string;
  rightCardTitleColor: string;
  rightCardSubtitleColor: string;
  rightCardNameColor: string;
  rightCardContactColor: string;
  rightCardEmailButtonColor: string;
  rightCardPhoneButtonColor: string;
}

export default class TwoColBoxContentWebPart extends BaseClientSideWebPart<ITwoColBoxContentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const leftCard: IContactCard = {
      title: this.properties.leftCardTitle,
      subtitle: this.properties.leftCardSubtitle,
      name: this.properties.leftCardName,
      email: this.properties.leftCardEmail,
      phone: this.properties.leftCardPhone,
      emailButtonText: this.properties.leftCardEmailButtonText,
      phoneButtonText: this.properties.leftCardPhoneButtonText,
      showEmailButton: this.properties.leftCardShowEmailButton,
      showPhoneButton: this.properties.leftCardShowPhoneButton,
      cardBackgroundColor: this.properties.leftCardBackgroundColor,
      titleColor: this.properties.leftCardTitleColor,
      subtitleColor: this.properties.leftCardSubtitleColor,
      nameColor: this.properties.leftCardNameColor,
      contactColor: this.properties.leftCardContactColor,
      emailButtonColor: this.properties.leftCardEmailButtonColor,
      phoneButtonColor: this.properties.leftCardPhoneButtonColor
    };

    const rightCard: IContactCard = {
      title: this.properties.rightCardTitle,
      subtitle: this.properties.rightCardSubtitle,
      name: this.properties.rightCardName,
      email: this.properties.rightCardEmail,
      phone: this.properties.rightCardPhone,
      emailButtonText: this.properties.rightCardEmailButtonText,
      phoneButtonText: this.properties.rightCardPhoneButtonText,
      showEmailButton: this.properties.rightCardShowEmailButton,
      showPhoneButton: this.properties.rightCardShowPhoneButton,
      cardBackgroundColor: this.properties.rightCardBackgroundColor,
      titleColor: this.properties.rightCardTitleColor,
      subtitleColor: this.properties.rightCardSubtitleColor,
      nameColor: this.properties.rightCardNameColor,
      contactColor: this.properties.rightCardContactColor,
      emailButtonColor: this.properties.rightCardEmailButtonColor,
      phoneButtonColor: this.properties.rightCardPhoneButtonColor
    };

    const element: React.ReactElement<I2ColBoxContentProps> = React.createElement(
      TwoColBoxContent,
      {
        leftCard: leftCard,
        rightCard: rightCard,
        columnLayout: this.properties.columnLayout,
        containerBackgroundColor: this.properties.containerBackgroundColor,
        cardSpacing: this.properties.cardSpacing,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneDropdown('columnLayout', {
                  label: strings.ColumnLayoutFieldLabel,
                  options: [
                    { key: 'left-right', text: 'Left - Right' },
                    { key: 'right-left', text: 'Right - Left' }
                  ],
                  selectedKey: this.properties.columnLayout || 'left-right'
                }),
                PropertyPaneTextField('containerBackgroundColor', {
                  label: strings.ContainerBackgroundColorFieldLabel,
                  value: this.properties.containerBackgroundColor || 'transparent',
                  description: 'Enter a valid CSS color value (e.g., #ffffff or transparent)'
                }),
                PropertyPaneSlider('cardSpacing', {
                  label: strings.CardSpacingFieldLabel,
                  min: 10,
                  max: 50,
                  value: this.properties.cardSpacing || 20,
                  showValue: true
                })
              ]
            },
            {
              groupName: strings.LeftCardGroupName,
              groupFields: [
                PropertyPaneTextField('leftCardTitle', {
                  label: strings.CardTitleFieldLabel,
                  value: this.properties.leftCardTitle || 'Technical Support (BullWealth)'
                }),
                PropertyPaneTextField('leftCardSubtitle', {
                  label: strings.CardSubtitleFieldLabel,
                  multiline: true,
                  rows: 2,
                  value: this.properties.leftCardSubtitle || 'Contact information for technical assistance'
                }),
                PropertyPaneTextField('leftCardName', {
                  label: strings.NameFieldLabel,
                  value: this.properties.leftCardName || 'Jolari HD'
                }),
                PropertyPaneTextField('leftCardEmail', {
                  label: strings.EmailFieldLabel,
                  value: this.properties.leftCardEmail || 'joralad@company.com'
                }),
                PropertyPaneTextField('leftCardPhone', {
                  label: strings.PhoneFieldLabel,
                  value: this.properties.leftCardPhone || '555-123-4567'
                }),
                PropertyPaneTextField('leftCardEmailButtonText', {
                  label: strings.EmailButtonTextFieldLabel,
                  value: this.properties.leftCardEmailButtonText || 'Email'
                }),
                PropertyPaneTextField('leftCardPhoneButtonText', {
                  label: strings.PhoneButtonTextFieldLabel,
                  value: this.properties.leftCardPhoneButtonText || 'Phone'
                }),
                PropertyPaneToggle('leftCardShowEmailButton', {
                  label: strings.ShowEmailButtonFieldLabel,
                  checked: this.properties.leftCardShowEmailButton !== false
                }),
                PropertyPaneToggle('leftCardShowPhoneButton', {
                  label: strings.ShowPhoneButtonFieldLabel,
                  checked: this.properties.leftCardShowPhoneButton !== false
                })
              ]
            },
            {
              groupName: strings.RightCardGroupName,
              groupFields: [
                PropertyPaneTextField('rightCardTitle', {
                  label: strings.CardTitleFieldLabel,
                  value: this.properties.rightCardTitle || 'Technical Support (Clover)'
                }),
                PropertyPaneTextField('rightCardSubtitle', {
                  label: strings.CardSubtitleFieldLabel,
                  multiline: true,
                  rows: 2,
                  value: this.properties.rightCardSubtitle || 'Contact information for technical assistance'
                }),
                PropertyPaneTextField('rightCardName', {
                  label: strings.NameFieldLabel,
                  value: this.properties.rightCardName || 'Blue Triangle'
                }),
                PropertyPaneTextField('rightCardEmail', {
                  label: strings.EmailFieldLabel,
                  value: this.properties.rightCardEmail || 'john.doe@company.com'
                }),
                PropertyPaneTextField('rightCardPhone', {
                  label: strings.PhoneFieldLabel,
                  value: this.properties.rightCardPhone || '555-123-4567'
                }),
                PropertyPaneTextField('rightCardEmailButtonText', {
                  label: strings.EmailButtonTextFieldLabel,
                  value: this.properties.rightCardEmailButtonText || 'Email'
                }),
                PropertyPaneTextField('rightCardPhoneButtonText', {
                  label: strings.PhoneButtonTextFieldLabel,
                  value: this.properties.rightCardPhoneButtonText || 'Phone'
                }),
                PropertyPaneToggle('rightCardShowEmailButton', {
                  label: strings.ShowEmailButtonFieldLabel,
                  checked: this.properties.rightCardShowEmailButton !== false
                }),
                PropertyPaneToggle('rightCardShowPhoneButton', {
                  label: strings.ShowPhoneButtonFieldLabel,
                  checked: this.properties.rightCardShowPhoneButton !== false
                })
              ]
            },
            {
              groupName: strings.DesignGroupName,
              groupFields: [
                PropertyPaneTextField('leftCardBackgroundColor', {
                  label: `Left ${strings.CardBackgroundColorFieldLabel}`,
                  value: this.properties.leftCardBackgroundColor || '#ffffff',
                  description: 'Enter a valid CSS color value (e.g., #ffffff, #000000, or transparent)'
                }),
                PropertyPaneTextField('rightCardBackgroundColor', {
                  label: `Right ${strings.CardBackgroundColorFieldLabel}`,
                  value: this.properties.rightCardBackgroundColor || '#ffffff',
                  description: 'Enter a valid CSS color value (e.g., #ffffff, #000000, or transparent)'
                }),
                PropertyPaneTextField('leftCardEmailButtonColor', {
                  label: `Left ${strings.EmailButtonColorFieldLabel}`,
                  value: this.properties.leftCardEmailButtonColor || '#5cb85c',
                  description: 'Enter a valid CSS color value (e.g., #5cb85c)'
                }),
                PropertyPaneTextField('rightCardEmailButtonColor', {
                  label: `Right ${strings.EmailButtonColorFieldLabel}`,
                  value: this.properties.rightCardEmailButtonColor || '#5cb85c',
                  description: 'Enter a valid CSS color value (e.g., #5cb85c)'
                }),
                PropertyPaneTextField('leftCardPhoneButtonColor', {
                  label: `Left ${strings.PhoneButtonColorFieldLabel}`,
                  value: this.properties.leftCardPhoneButtonColor || '#5bc0de',
                  description: 'Enter a valid CSS color value (e.g., #5bc0de)'
                }),
                PropertyPaneTextField('rightCardPhoneButtonColor', {
                  label: `Right ${strings.PhoneButtonColorFieldLabel}`,
                  value: this.properties.rightCardPhoneButtonColor || '#5bc0de',
                  description: 'Enter a valid CSS color value (e.g., #5bc0de)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
