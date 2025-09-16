import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import NavigationMenu from './components/NavigationMenu';
import { INavigationMenuProps } from './components/INavigationProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { NavigationService } from './services/NavigationService';
import { initializeIcons } from '@fluentui/react/lib/Icons';

const LOG_SOURCE: string = 'BullWealthNavigationApplicationCustomizer';

export interface IBullWealthNavigationApplicationCustomizerProperties {
  homeUrl?: string;
  bullWealthUrl?: string;
  cloverUrl?: string;
  hrUrl?: string;
  itPolicyUrl?: string;
  helpUrl?: string;
}

export default class BullWealthNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IBullWealthNavigationApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _navigationService: NavigationService;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${LOG_SOURCE}`);

    // Load Fabric icons first
    this.loadFabricIconsImmediately();

    // Initialize navigation service with correct URL
    const correctSiteUrl = 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet';
    this._navigationService = new NavigationService(
      this.context.spHttpClient,
      correctSiteUrl
    );

    // Wait for placeholders
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private loadFabricIconsImmediately(): void {
  // ONLY load Fabric CSS - don't override fonts
  SPComponentLoader.loadCss('https://res.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css');
  
  // Initialize icons but DON'T override font-family
  initializeIcons();
  
  console.log('✅ Fabric icons loaded - letting SharePoint handle fonts');
}


  private _renderPlaceHolders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        console.error('❌ Top placeholder not found');
        return;
      }

      if (this._topPlaceholder.domElement) {
        this._navigationService.getNavigationItems()
          .then(navigationItems => {
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
              items: navigationItems,
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet'
            });

            if (this._topPlaceholder && this._topPlaceholder.domElement) {
              ReactDom.render(element, this._topPlaceholder.domElement);
              console.log('✅ Navigation rendered');
            }
          })
          .catch(error => {
            console.error('❌ Navigation error:', error);
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
              items: [],
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet'
            });

            if (this._topPlaceholder && this._topPlaceholder.domElement) {
              ReactDom.render(element, this._topPlaceholder.domElement);
            }
          });
      }
    }
  }

  private _onDispose(): void {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
  }
}
