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

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${LOG_SOURCE}`);

    // Wait for placeholders to be available
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render in case placeholders are already available
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log('BullWealth Navigation: Attempting to render...');

    // Handle the top placeholder (header area)
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topPlaceholder.domElement) {
        console.log('✅ Top placeholder found, rendering navigation...');
        
        // Create React element
        const element: React.ReactElement<INavigationMenuProps> = React.createElement(NavigationMenu, {
          items: [], // Component defines its own items
          siteUrl: this.context.pageContext.web.absoluteUrl
        });

        ReactDom.render(element, this._topPlaceholder.domElement);
        console.log('✅ Navigation rendered successfully!');
      }
    }
  }

  private _onDispose(): void {
    console.log('[BullWealth Navigation] Disposed custom top placeholder.');
  }
}
