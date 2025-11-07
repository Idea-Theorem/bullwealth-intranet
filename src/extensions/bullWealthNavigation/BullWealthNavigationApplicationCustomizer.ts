/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDom from "react-dom";

// ✅ Fixed: Use default import (no TS1192 error)
import NavigationMenu from "./components/NavigationMenu";
import { INavigationMenuProps } from "./components/INavigationProps";

import { SPComponentLoader } from "@microsoft/sp-loader";
import { NavigationService } from "./services/NavigationService";
import { initializeIcons } from "@fluentui/react/lib/Icons";

const LOG_SOURCE: string = "BullWealthNavigationApplicationCustomizer";

export interface IBullWealthNavigationApplicationCustomizerProperties {
  homeUrl?: string;
  bullWealthUrl?: string;
  cloverUrl?: string;
  hrUrl?: string;
  itPolicyUrl?: string;
  helpUrl?: string;
}

export default class BullWealthNavigationApplicationCustomizer extends BaseApplicationCustomizer<IBullWealthNavigationApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _navigationService: NavigationService;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `✅ Initialized ${LOG_SOURCE}`);

    // Restrict to specific site
    const currentSiteUrl = this.context.pageContext.web.absoluteUrl.toLowerCase();
    if (!currentSiteUrl.includes("/sites/mrkedcapitalintranet")) {
      console.log("ℹ️ Navigation disabled on non-target site:", currentSiteUrl);
      return Promise.resolve();
    }

    // Load Fabric icons & CSS
    this._loadFabricIconsImmediately();

    // Initialize navigation service
    const correctSiteUrl =
      "https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/";
    this._navigationService = new NavigationService(
      this.context.spHttpClient,
      correctSiteUrl
    );

    // Listen to placeholder changes
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  /**
   * ✅ Load Fluent UI icons + Fabric CSS safely.
   */
  private _loadFabricIconsImmediately(): void {
    SPComponentLoader.loadCss(
      "https://res.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css"
    );

    initializeIcons();

    // Layout override for SharePoint modern width container
    const globalOverrideCSS = document.createElement("style");
    globalOverrideCSS.id = "sharepoint-layout-override";
    globalOverrideCSS.innerHTML = `
      @media screen and (min-width: 1024px) {
        .r_NLtZH_y298L:not(.f_bHim3_y298L) .s_wDEw-_y298L {
          max-width: 1440px !important;
        }
      }
    `;
    document.head.appendChild(globalOverrideCSS);

    console.log("✅ Fabric icons & layout styles applied");
  }

  /**
   * ✅ Render navigation into Top placeholder
   */
  private _renderPlaceHolders(): void {
    try {
      // Create Top placeholder if missing
      if (!this._topPlaceholder) {
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose }
        );

        if (!this._topPlaceholder) {
          console.error("❌ Top placeholder not found");
          return;
        }
      }

      // Render navigation inside the placeholder
      if (this._topPlaceholder?.domElement) {
        console.log("📦 Rendering navigation placeholder...");

        this._navigationService
          .getNavigationItems()
          .then((navigationItems) => {
            const element: React.ReactElement<INavigationMenuProps> = React.createElement(
              NavigationMenu,
              {
                items: navigationItems,
                siteUrl:
                  "https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/",
              }
            );

            // ✅ Use non-null assertion to silence TS2532
            ReactDom.render(element, this._topPlaceholder!.domElement);
            console.log("✅ Navigation rendered successfully");
          })
          .catch((error) => {
            console.error("❌ Error rendering navigation:", error);

            // Fallback: render an empty menu safely
            const fallbackElement: React.ReactElement<INavigationMenuProps> =
              React.createElement(NavigationMenu, {
                items: [],
                siteUrl:
                  "https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/",
              });

            if (this._topPlaceholder?.domElement) {
              ReactDom.render(fallbackElement, this._topPlaceholder!.domElement);
            }
          });
      }
    } catch (err) {
      console.error("💥 Render placeholder error:", err);
    }
  }

  /**
   * ✅ Clean up navigation when placeholder is removed
   */
  private _onDispose(): void {
    if (this._topPlaceholder?.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder!.domElement);
    }
  }
}
