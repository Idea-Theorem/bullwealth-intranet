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

    // 🚨 THE 3-LINE FIX: ONLY SHOW ON MRKEDCAPITALINTRANET
    const currentSiteUrl = this.context.pageContext.web.absoluteUrl.toLowerCase();
    if (!currentSiteUrl.includes('/sites/mrkedcapitalintranet')) {
      return Promise.resolve(); // EXIT - Navigation won't show
    }

    // Load Fabric icons first
    this.loadFabricIconsImmediately();

    // Initialize navigation service with correct URL
    const correctSiteUrl = 'https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/';
    this._navigationService = new NavigationService(
      this.context.spHttpClient,
      correctSiteUrl
    );

    // Wait for placeholders to be available
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private loadFabricIconsImmediately(): void {
    // ONLY load Fabric CSS - don't override fonts
    SPComponentLoader.loadCss('https://res.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css');
    
    // Initialize icons but DON'T override font-family
    initializeIcons();

    const globalOverrideCSS = document.createElement('style');
    globalOverrideCSS.id = 'sharepoint-layout-override';
    globalOverrideCSS.innerHTML = `
      @media screen and (min-width: 1024px) {
        .r_cJ1Dm_y298L:not(.f_XsZ2U_y298L) .s_ywkch_y298L {
          max-width: 1440px !important;
        }
      }
    `;
    
    document.head.appendChild(globalOverrideCSS);
    
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
            console.log('📋 Navigation items fetched:', navigationItems);
            
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
              items: navigationItems,
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/'
            });

            if (this._topPlaceholder && this._topPlaceholder.domElement) {
              ReactDom.render(element, this._topPlaceholder.domElement);
              console.log('✅ Custom navigation rendered successfully with', navigationItems.length, 'items');
              
              // Force show and position your navigation after render
              setTimeout(() => {
                const customNav = document.querySelector('[data-sp-placeholder-name="Top"]');
                if (customNav) {
                  (customNav as HTMLElement).style.display = 'block';
                  (customNav as HTMLElement).style.visibility = 'visible';
                  (customNav as HTMLElement).style.position = 'sticky';
                  (customNav as HTMLElement).style.top = '48px';
                  (customNav as HTMLElement).style.zIndex = '1000';
                  (customNav as HTMLElement).style.background = 'white';
                  (customNav as HTMLElement).style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
                  console.log('✅ Custom navigation forced visible and positioned');
                }
                
                // Hide SharePoint elements after custom nav is rendered
                this.hideSpecificElements();
              }, 100);
            }
          })
          .catch(error => {
            console.error('❌ Error loading navigation items:', error);
            
            // Render with fallback navigation
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
              items: [],
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/'
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
