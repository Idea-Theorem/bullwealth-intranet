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

    // Load Fabric icons and apply targeted CSS
    this.loadFabricIconsImmediately();

    // Initialize navigation service with correct URL
    const correctSiteUrl = 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet';
    console.log('🏠 Using site URL:', correctSiteUrl);

    this._navigationService = new NavigationService(
      this.context.spHttpClient,
      correctSiteUrl
    );

    // Wait for placeholders to be available
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    
    // Call render in case placeholders are already available
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private loadFabricIconsImmediately(): void {
    // Load Fabric CSS for icons
    SPComponentLoader.loadCss('https://res.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css');
    
    // Initialize icons (let SharePoint handle font-family)
    initializeIcons();
    
    // TARGETED CSS - Hide only center SharePoint header
    const hideHeaderCSS = document.createElement('style');
    hideHeaderCSS.id = 'hide-center-header';
    hideHeaderCSS.innerHTML = `
      /* KEEP M365 Suite Bar visible */
      div[class*="od-SuiteHeader"],
      #O365_NavHeader,
      #SuiteNavPlaceHolder,
      div[class*="suiteBar"],
      div[id*="appBar"] {
        display: block !important;
        visibility: visible !important;
        position: sticky !important;
        top: 0 !important;
        z-index: 2000 !important;
      }
      
      /* KEEP your custom navigation visible and positioned */
      div[data-sp-placeholder-name="Top"] {
        display: block !important;
        visibility: visible !important;
        position: sticky !important;
        top: 48px !important;
        z-index: 1000 !important;
        background: white !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
      }
      
      div[data-sp-placeholder-name="Top"] *,
      div[data-sp-placeholder-name="Top"] nav,
      div[data-sp-placeholder-name="Top"] ul,
      div[data-sp-placeholder-name="Top"] li,
      div[data-sp-placeholder-name="Top"] a {
        display: block !important;
        visibility: visible !important;
      }
      
      /* Hide the center SharePoint header section */
      div[data-sp-feature-tag="Site Header"],
      div[data-automation-id="SiteHeader"],
      header[data-sp-feature-instance-id*="SiteHeader"],
      div[class*="siteHeader"]:not([class*="custom"]):not([data-sp-placeholder-name="Top"]),
      div[class*="HeaderWrapper"]:not([class*="custom"]):not([data-sp-placeholder-name="Top"]),
      div[class*="pageHeader"]:not([class*="custom"]):not([data-sp-placeholder-name="Top"]),
      header[role="banner"]:not([class*="custom"]):not([data-sp-placeholder-name="Top"]) {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      
      /* KEEP page commands visible (New, Edit, etc.) */
      div[data-sp-feature-tag*="Command"],
      div[class*="commandBar"]:has(button),
      div[class*="pageCommand"],
      div[class*="toolbar"],
      div:has(button[title*="New"]),
      div:has(button[title*="Edit"]) {
        display: block !important;
        visibility: visible !important;
      }
      
      /* Remove spacing from hidden header */
      #contentRow,
      #contentBox,
      div[role="main"],
      .CanvasZone {
        margin-top: 0 !important;
        padding-top: 0 !important;
      }
    `;
    
    // Remove existing and add new
    const existing = document.getElementById('hide-center-header');
    if (existing) existing.remove();
    document.head.appendChild(hideHeaderCSS);
    
    // Use JavaScript for more precise hiding
    setTimeout(() => this.hideSpecificElements(), 1000);
    setTimeout(() => this.hideSpecificElements(), 3000);
    setTimeout(() => this.hideSpecificElements(), 5000);
    
    console.log('✅ Center SharePoint header hiding CSS applied');
  }

  private hideSpecificElements(): void {
    console.log('🔍 Hiding center SharePoint header elements...');
    
    // Method 1: Hide the entire site header section
    const siteHeaderSelectors = [
      'div[data-sp-feature-tag="Site Header"]',
      'div[data-automation-id="SiteHeader"]',
      'header[data-sp-feature-instance-id*="SiteHeader"]',
      'div[class*="siteHeader"]',
      'div[class*="HeaderWrapper"]',
      'div[class*="pageHeader"]'
    ];
    
    siteHeaderSelectors.forEach(selector => {
      document.querySelectorAll(selector).forEach((element: HTMLElement) => {
        // Only hide if it's not your custom navigation
        if (!element.closest('[data-sp-placeholder-name="Top"]') && 
            !element.querySelector('[data-sp-placeholder-name="Top"]')) {
          element.style.display = 'none';
          element.style.height = '0px';
          console.log('🗑️ Hidden site header:', selector);
        }
      });
    });
    
    // Method 2: Hide elements containing both logo and SharePoint navigation
    document.querySelectorAll('div').forEach((div: HTMLElement) => {
      const divText = div.textContent || '';
      
      // Look for the section with BI logo and site navigation
      if (divText.includes('BullWealth Intranet') && 
          divText.includes('Conversations') && 
          divText.includes('Notebook') &&
          divText.includes('Documents') &&
          !div.closest('[data-sp-placeholder-name="Top"]') &&
          !div.querySelector('[data-sp-placeholder-name="Top"]')) {
        
        // Check if this div doesn't contain page commands
        if (!divText.includes('New') || 
            (divText.includes('New') && divText.includes('Conversations'))) {
          div.style.display = 'none';
          console.log('🗑️ Hidden SharePoint header section');
        }
      }
    });
    
    // Method 3: Hide navigation with SharePoint default items
    document.querySelectorAll('nav').forEach((nav: HTMLElement) => {
      const navText = nav.textContent || '';
      
      if ((navText.includes('Home') && 
           navText.includes('Conversations') && 
           navText.includes('Notebook') &&
           navText.includes('Documents')) &&
          !nav.closest('[data-sp-placeholder-name="Top"]')) {
        
        // Hide the entire navigation container
        let containerToHide = nav.parentElement;
        while (containerToHide && containerToHide.tagName !== 'BODY') {
          const containerText = containerToHide.textContent || '';
          if (containerText.includes('BullWealth Intranet') && 
              containerText.includes('Conversations') &&
              !containerToHide.closest('[data-sp-placeholder-name="Top"]') &&
              !containerToHide.querySelector('[data-sp-placeholder-name="Top"]')) {
            
            // Don't hide if it contains page commands
            if (!containerText.includes('Page details') || 
                containerText.includes('Notebook')) {
              containerToHide.style.display = 'none';
              console.log('🗑️ Hidden SharePoint header container');
              break;
            }
          }
          containerToHide = containerToHide.parentElement;
        }
      }
    });
    
    // Method 4: Target by specific content (BI logo)
    document.querySelectorAll('*').forEach((element: HTMLElement) => {
      // Look for elements containing the BI logo area
      if (element.textContent?.trim() === 'BI' && 
          element.tagName !== 'TITLE' &&
          !element.closest('[data-sp-placeholder-name="Top"]')) {
        
        // Hide the logo container
        let logoContainer = element.parentElement;
        while (logoContainer && logoContainer.tagName !== 'BODY') {
          const containerText = logoContainer.textContent || '';
          if (containerText.includes('BullWealth Intranet') && 
              containerText.includes('Home') &&
              containerText.includes('Conversations') &&
              !logoContainer.closest('[data-sp-placeholder-name="Top"]') &&
              !logoContainer.querySelector('[data-sp-placeholder-name="Top"]')) {
            
            logoContainer.style.display = 'none';
            console.log('🗑️ Hidden BI logo container');
            break;
          }
          logoContainer = logoContainer.parentElement;
        }
      }
    });
    
    // Method 5: Hide headers but preserve page toolbar
    document.querySelectorAll('header, section').forEach((element: HTMLElement) => {
      const elementText = element.textContent || '';
      
      if (elementText.includes('BullWealth Intranet') && 
          elementText.includes('Conversations') &&
          !elementText.includes('Page details') &&
          !element.closest('[data-sp-placeholder-name="Top"]') &&
          !element.querySelector('[data-sp-placeholder-name="Top"]')) {
        
        element.style.display = 'none';
        console.log('🗑️ Hidden header/section element');
      }
    });
    
    // Force show your custom navigation
    const customNav = document.querySelector('[data-sp-placeholder-name="Top"]');
    if (customNav) {
      (customNav as HTMLElement).style.display = 'block';
      (customNav as HTMLElement).style.visibility = 'visible';
      (customNav as HTMLElement).style.position = 'sticky';
      (customNav as HTMLElement).style.top = '48px';
      (customNav as HTMLElement).style.zIndex = '1000';
      (customNav as HTMLElement).style.background = 'white';
      console.log('✅ Custom navigation kept visible and positioned');
    }
    
    // Force show page commands/toolbar
    const pageCommands = document.querySelectorAll('div[class*="commandBar"], div[class*="toolbar"], button[title*="New"], button[title*="Edit"]');
    pageCommands.forEach((cmd: HTMLElement) => {
      if (!cmd.textContent?.includes('Conversations')) {
        cmd.style.display = 'block';
        cmd.style.visibility = 'visible';
      }
    });
    
    console.log('✅ Specific element hiding completed');
  }

  private _renderPlaceHolders(): void {
    console.log('🚀 BullWealth Navigation: Attempting to render...');

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        console.error('❌ The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topPlaceholder.domElement) {
        console.log('✅ Top placeholder found, fetching navigation items...');
        console.log('📍 Placeholder DOM element:', this._topPlaceholder.domElement);
        
        // Fetch navigation items and render
        this._navigationService.getNavigationItems()
          .then(navigationItems => {
            console.log('📋 Navigation items fetched:', navigationItems);
            
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
              items: navigationItems,
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet'
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
              siteUrl: 'https://bullwealthmanagementgro.sharepoint.com/sites/BullWealthIntranet'
            });

            if (this._topPlaceholder && this._topPlaceholder.domElement) {
              ReactDom.render(element, this._topPlaceholder.domElement);
              console.log('✅ Navigation rendered with fallback items');
              
              // Force show your navigation
              setTimeout(() => {
                const customNav = document.querySelector('[data-sp-placeholder-name="Top"]');
                if (customNav) {
                  (customNav as HTMLElement).style.display = 'block';
                  (customNav as HTMLElement).style.visibility = 'visible';
                  (customNav as HTMLElement).style.position = 'sticky';
                  (customNav as HTMLElement).style.top = '48px';
                  (customNav as HTMLElement).style.zIndex = '1000';
                  (customNav as HTMLElement).style.background = 'white';
                  console.log('✅ Fallback navigation forced visible');
                }
                this.hideSpecificElements();
              }, 100);
            }
          });
      }
    }
  }

  private _onDispose(): void {
    console.log('[BullWealth Navigation] Disposed custom top placeholder.');
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
  }
}
