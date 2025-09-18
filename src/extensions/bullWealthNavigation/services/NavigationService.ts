/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { INavigationItem } from '../components/INavigationProps';

export interface INavigationListItem {
  Title: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  URL: any; // SharePoint URL field can have different structures
  // eslint-disable-next-line @rushstack/no-new-null
  Icon: string | null;
  // eslint-disable-next-line @rushstack/no-new-null
  Parent: string | null;
  Order0: number;
  IsActive: boolean;
  Id: number;
}

export class NavigationService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
    console.log('🏠 Navigation Service initialized with site URL:', this.siteUrl);
  }

  public async getNavigationItems(): Promise<INavigationItem[]> {
  try {
    // FIXED: Try different possible field names and remove problematic filter
    const itemsUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Navigation%20Items')/items?$select=Id,Title,URL,Icon,Parent,Order0,Order,IsActive&$orderby=Id asc`;
    console.log('🔍 Fetching navigation items (no filter):', itemsUrl);
    
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      itemsUrl,
      SPHttpClient.configurations.v1
    );

    console.log('📡 Response status:', response.status);

    if (response.ok) {
      const data = await response.json();
      console.log('📋 Raw SharePoint data:', data);
      
      const items = data.d?.results || data.value || [];
      console.log('📋 All items (before filter):', items);
      
      // FIXED: Apply filter in code instead of OData
      const activeItems = items.filter((item: any) => {
        console.log(`🔍 Checking item: ${item.Title}, IsActive: ${item.IsActive} (${typeof item.IsActive})`);
        return item.IsActive === true || item.IsActive === 'Yes' || item.IsActive === 1;
      });
      
      console.log('📋 Active items (after filter):', activeItems);
      
      if (activeItems && activeItems.length > 0) {
        const navigationTree = this.buildNavigationTree(activeItems);
        console.log('🌳 Built navigation tree:', navigationTree);
        return navigationTree;
      } else {
        console.warn('⚠️ No active items found, using fallback');
        return this.getFallbackNavigation();
      }
    } else {
      const errorText = await response.text();
      console.error('❌ HTTP Error:', response.status, response.statusText, errorText);
      return this.getFallbackNavigation();
    }
  } catch (error) {
    console.error('💥 Error fetching navigation items:', error);
    return this.getFallbackNavigation();
  }
}


    private buildNavigationTree(items: any[]): INavigationItem[] {
  console.log('🔨 Building navigation tree from', items.length, 'items');
  
  // Convert all items to proper format first
  const convertedItems = items.map((item, index) => {
    let urlValue = '#';
    if (item.URL) {
      if (typeof item.URL === 'string') {
        urlValue = item.URL;
      } else if (item.URL.Url) {
        urlValue = item.URL.Url;
      }
    }
    
    const iconValue = item.Icon || 'Home';
    const orderValue = item.Order0 || item.Order || index;
    
    return {
      id: item.Id,
      name: item.Title || 'Untitled',
      url: urlValue,
      icon: iconValue,
      parent: item.Parent || null,
      order: orderValue
    };
  });
  
  console.log('🔄 All converted items:', convertedItems);
  
  // FIXED: Separate parents and children
  const parentItems = convertedItems.filter(item => !item.parent || item.parent === '');
  const childItems = convertedItems.filter(item => item.parent && item.parent !== '');
  
  console.log('👨‍👦 Parent items:', parentItems);
  console.log('👶 Child items:', childItems);
  
  // Build navigation tree with proper parent-child relationships
  const navigationTree = parentItems.map(parent => {
    // Find all children for this parent
    const children = childItems
      .filter(child => child.parent === parent.name)
      .sort((a, b) => a.order - b.order)
      .map(child => ({
        name: child.name,
        url: child.url
      }));
    
    const navItem: INavigationItem = {
      name: parent.name,
      url: parent.url,
      icon: parent.icon
    };
    
    // Only add children if they exist
    if (children.length > 0) {
      navItem.children = children;
    }
    
    console.log(`📄 Built nav item: ${parent.name}`, navItem);
    return navItem;
  }).sort((a, b) => {
    const aOrder = convertedItems.find(item => item.name === a.name)?.order || 0;
    const bOrder = convertedItems.find(item => item.name === b.name)?.order || 0;
    return aOrder - bOrder;
  });
  
  console.log('🌳 Final navigation tree:', navigationTree);
  return navigationTree;
}



  private getFallbackNavigation(): INavigationItem[] {
    console.log('🔄 Using fallback navigation');
    return [
      { name: 'Home', url: '/', icon: 'Home' },
      { 
        name: 'BullWealth', 
        url: '#', 
        icon: 'Building', 
        children: [
          { name: 'Compliance', url: '/sites/bullwealth/compliance' },
          { name: 'Research & Investment', url: '/sites/bullwealth/research' }
        ]
      },
      { name: 'Human Resource', url: '/sites/hr', icon: 'People' },
      { name: 'IT Policy', url: '/sites/it-policy', icon: 'Shield' },
      { name: 'Help Centre', url: '/sites/help', icon: 'Help' }
    ];
  }
}