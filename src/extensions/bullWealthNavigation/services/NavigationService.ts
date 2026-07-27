/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient } from "@microsoft/sp-http";
import { INavigationItem } from "../components/INavigationProps";

export class NavigationService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
  }

  // ✅ Main navigation fetcher (hybrid: static + dynamic)
  public async getNavigationItems(): Promise<INavigationItem[]> {
    try {
      const navListUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Navigation%20Items')/items?$select=Id,Title,URL,Icon,Parent,Order0,Order,IsActive,Library&$orderby=Id asc`;

      const response = await this.spHttpClient.get(navListUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        return this.getFallbackNavigation();
      }

      const data = await response.json();
      const items = data.value || [];
      const activeItems = items.filter(
        (item: any) => item.IsActive === true || item.IsActive === "Yes" || item.IsActive === 1
      );


      const baseNavigation = this.buildNavigationTree(activeItems);

      // ✅ Data-driven: any top-level nav item whose "Library" field (set on the
      // Navigation Items list) is populated gets its subfolders attached automatically.
      // No code change is needed to enable this for a new section going forward —
      // just set the Library column on that item in SharePoint.
      for (const parentNode of baseNavigation as (INavigationItem & { library?: string })[]) {
        const libraryName = parentNode.library;
        if (!libraryName) {
          continue;
        }

        const dynamicFolders = await this.getDynamicFoldersAndFiles(libraryName);
        if (dynamicFolders.length > 0) {
          // Show dynamic folders first, then static ones (no separator)
          parentNode.children = [
            ...dynamicFolders.sort((a, b) => (a.order ?? 999) - (b.order ?? 999)),
            ...(parentNode.children || []),
          ];
        }
      }

      return baseNavigation;
    } catch (err) {
      return this.getFallbackNavigation();
    }
  }

  // ✅ Reusable: fetch folders & files from any library
  private async getDynamicFoldersAndFiles(libraryName: string): Promise<INavigationItem[]> {
    try {
      const siteUrl = this.siteUrl;
      const webServerRelativeUrl = siteUrl.replace(/^https?:\/\/[^/]+/, ""); // e.g. /sites/MrkedCapitalIntranet
      const fullPath = `${webServerRelativeUrl}/Shared Documents/${libraryName}`.replace(/\/+/g, "/");


      const foldersApi = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
        fullPath
      )}')/Folders?$expand=Files&$select=Name,ServerRelativeUrl,Files/Name,Files/ServerRelativeUrl`;

      const res = await this.spHttpClient.get(foldersApi, SPHttpClient.configurations.v1);
      if (!res.ok) {
        console.error(`❌ Failed to fetch folders for ${libraryName}:`, res.status);
        return [];
      }

      const data = await res.json();
      const folders = data.value || [];
      console.log(`📁 Found ${folders.length} folders in ${libraryName}`);

      const folderItems: INavigationItem[] = [];

      for (const folder of folders) {
        const folderName = folder.Name;
        const folderServerUrl = folder.ServerRelativeUrl;

        // 🔹 Try to get "OrderBy" field for sorting
        let orderBy = 999;
        try {
          const orderUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
            folderServerUrl
          )}')/ListItemAllFields?$select=OrderBy`;
          const orderRes = await this.spHttpClient.get(orderUrl, SPHttpClient.configurations.v1);
          if (orderRes.ok) {
            const orderData = await orderRes.json();
            const parsed = parseInt(orderData.OrderBy, 10);
            if (!isNaN(parsed)) orderBy = parsed;
          }
        } catch {
          console.warn(`⚠️ No OrderBy found for folder: ${folderName}`);
        }

        // 🔹 Build view URL for DocumentLibrary.aspx
        const encodedPath = encodeURIComponent(`Shared Documents/${libraryName}/${folderName}`);
        const viewUrl = `${this.siteUrl}SitePages/DocumentLibrary.aspx?library=${encodedPath}`;

        folderItems.push({
          name: folderName,
          url: viewUrl,
          icon: "Folder",
          order: orderBy,
        });
      }

      folderItems.sort((a, b) => (a.order ?? 999) - (b.order ?? 999));
      return folderItems;
    } catch (err) {
      return [];
    }
  }

  // ✅ Builds base static parent-child structure
  private buildNavigationTree(items: any[]): INavigationItem[] {
    const convertedItems = items.map((item, index) => {
      let urlValue = "#";
      if (item.URL) {
        if (typeof item.URL === "string") urlValue = item.URL;
        else if (item.URL.Url) urlValue = item.URL.Url;
      }

      return {
        id: item.Id,
        name: item.Title || "Untitled",
        url: urlValue,
        icon: item.Icon || "Home",
        parent: item.Parent || null,
        order: item.Order0 || item.Order || index,
        library: item.Library || undefined,
      };
    });

    const parents = convertedItems.filter((i) => !i.parent);
    const children = convertedItems.filter((i) => i.parent);

    const navigationTree = parents
      .map((parent) => {
        const subItems = children
          .filter((child) => child.parent === parent.name)
          .sort((a, b) => (a.order ?? 999) - (b.order ?? 999))
          .map((child) => ({
            name: child.name,
            url: child.url,
            order: child.order,
          }));

        const navItem: INavigationItem & { library?: string } = {
          name: parent.name,
          url: parent.url,
          icon: parent.icon,
          order: parent.order ?? 999,
          library: parent.library,
        };

        if (subItems.length > 0) navItem.children = subItems;
        return navItem;
      })
      .sort((a, b) => (a.order ?? 999) - (b.order ?? 999));

    return navigationTree;
  }

  // ✅ Fallback static menu (used if list fails)
  private getFallbackNavigation(): INavigationItem[] {
    return [
      { name: "Home", url: "/", icon: "Home", order: 0 },
      {
        name: "BullWealth",
        url: "#",
        icon: "Building",
        children: [],
        order: 0,
      },
      {
        name: "Clover",
        url: "#",
        icon: "Leaf",
        children: [],
        order: 1,
      },
      {
        name: "HR & Finance",
        url: "#",
        icon: "People",
        children: [],
        order: 2,
      },
      {
        name: "Mrked",
        url: "#",
        icon: "Globe",
        children: [],
        order: 3,
      },
      { name: "Help Centre", url: "/sites/help", icon: "Help", order: 4 },
    ];
  }
}
