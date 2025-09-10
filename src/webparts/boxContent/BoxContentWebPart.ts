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
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BoxContentWebPartStrings';
// Ensure the file exists at the correct path and with the correct name (BoxContent.tsx or BoxContent.ts)
// Ensure the file exists as BoxContent.tsx or BoxContent.ts in the components folder
import BoxContent from './components/BoxContent';
// If your file is named BoxContent.tsx, you can also use:
// import BoxContent from './components/BoxContent.tsx';
import { IBoxContentProps } from './components/IBoxContentProps';

export interface IBoxContentWebPartProps {
  title: string;
  description: string;
  duration: string;
  buttonText: string;
  buttonUrl: string;
  buttonIcon: string;
  backgroundColor: string;
  titleColor: string;
  descriptionColor: string;
  buttonColor: string;
  showDuration: boolean;
}

export default class BoxContentWebPart extends BaseClientSideWebPart<IBoxContentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBoxContentProps> = React.createElement(
      BoxContent,
      {
        title: this.properties.title,
        description: this.properties.description,
        duration: this.properties.duration,
        buttonText: this.properties.buttonText,
        buttonUrl: this.properties.buttonUrl,
        buttonIcon: this.properties.buttonIcon,
        backgroundColor: this.properties.backgroundColor,
        titleColor: this.properties.titleColor,
        descriptionColor: this.properties.descriptionColor,
        buttonColor: this.properties.buttonColor,
        showDuration: this.properties.showDuration,
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
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: this.properties.title || 'HR Platform Introduction'
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  rows: 3,
                  value: this.properties.description || 'Get started with our HR platform. HR Platform Introductory Call Recording'
                }),
                PropertyPaneTextField('duration', {
                  label: strings.DurationFieldLabel,
                  value: this.properties.duration || '25 minutes'
                }),
                PropertyPaneToggle('showDuration', {
                  label: strings.ShowDurationFieldLabel,
                  checked: this.properties.showDuration !== false
                }),
                PropertyPaneTextField('buttonText', {
                  label: strings.ButtonTextFieldLabel,
                  value: this.properties.buttonText || 'Watch'
                }),
                PropertyPaneTextField('buttonUrl', {
                  label: strings.ButtonUrlFieldLabel,
                  placeholder: 'https://example.com or /sites/sitename/pages/pagename.aspx',
                  value: this.properties.buttonUrl || ''
                }),
                PropertyPaneTextField('buttonIcon', {
                  label: strings.ButtonIconFieldLabel,
                  placeholder: 'Play, VideoSolid, etc.',
                  value: this.properties.buttonIcon || 'Play'
                })
              ]
            },
            {
              groupName: strings.DesignGroupName,
              groupFields: [
                PropertyPaneSlider('backgroundColor', {
                  label: strings.BackgroundColorFieldLabel,
                  min: 0,
                  max: 100,
                  value: 0
                }),
                PropertyPaneTextField('titleColor', {
                  label: strings.TitleColorFieldLabel,
                  placeholder: '#323130',
                  value: this.properties.titleColor || '#323130'
                }),
                PropertyPaneTextField('descriptionColor', {
                  label: strings.DescriptionColorFieldLabel,
                  placeholder: '#605e5c',
                  value: this.properties.descriptionColor || '#605e5c'
                }),
                PropertyPaneTextField('buttonColor', {
                  label: strings.ButtonColorFieldLabel,
                  placeholder: '#5cb85c',
                  value: this.properties.buttonColor || '#5cb85c'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
