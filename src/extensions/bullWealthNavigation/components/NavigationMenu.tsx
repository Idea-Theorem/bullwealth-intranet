import * as React from 'react';
import { useState, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { INavigationMenuProps, INavigationItem } from './INavigationProps';
import styles from './NavigationMenu.module.scss';


const NavigationMenu: React.FC<INavigationMenuProps> = ({ items, siteUrl }) => {
  const [activeDropdown, setActiveDropdown] = useState<string | null>(null);
  const [currentUrl, setCurrentUrl] = useState<string>('');

  useEffect(() => {
    setCurrentUrl(window.location.href);
  }, []);

  const isActive = (item: INavigationItem): boolean => {
    const itemUrl = item.url.startsWith('http') ? item.url : `${siteUrl}${item.url}`;
    const normalizedItemUrl = itemUrl.toLowerCase().replace(/\/$/, '');
    const normalizedCurrentUrl = currentUrl.toLowerCase().replace(/\/$/, '');
    
    // Check if current URL matches this item
    if (normalizedCurrentUrl === normalizedItemUrl) {
      return true;
    }
    
    // Check if any child matches
    if (item.children) {
      return item.children.some(child => {
        const childUrl = child.url.startsWith('http') ? child.url : `${siteUrl}${child.url}`;
        const normalizedChildUrl = childUrl.toLowerCase().replace(/\/$/, '');
        return normalizedCurrentUrl === normalizedChildUrl || normalizedCurrentUrl.includes(normalizedChildUrl);
      });
    }
    
    return false;
  };

  const handleLinkClick = (url: string, external?: boolean): void => {
    if (external) {
      window.open(url, '_blank');
    } else {
      window.location.href = url.startsWith('http') ? url : `${siteUrl}${url}`;
    }
  };


  const getFallbackNavigation = (): INavigationItem[] => {
    return [
      { name: 'Home', url: '/', icon: 'Home' },
      { 
        name: 'BullWealth', 
        url: '/sites/BullWealthIntranet/bullwealth',
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
  };


  const navigationItems = (items && items.length > 0) ? items : getFallbackNavigation();


  return (
    <div className={styles.navigationWrapper}>
      <nav className={styles.navigationMenu}>
        <div className={styles.brand}>
          <a href="https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/SitePages/Home.aspx"><h1 className={styles.brandTitle}>Mrked Capital Intranet</h1></a>
        </div>


        <ul className={styles.navList}>
          {navigationItems.map((item, index) => (
            <li 
              key={index} 
              className={`${styles.navItem} ${item.children ? styles.dropdown : ''} ${isActive(item) ? styles.activeItem : ''}`}
              onMouseEnter={() => item.children && setActiveDropdown(item.name)}
              onMouseLeave={() => item.children && setActiveDropdown(null)}
            >
              <a 
                href="#"
                className={styles.navLink}
                onClick={(e) => {
                  e.preventDefault();
                  
                  if (item.children) {
                    if (item.url && item.url !== '#') {
                      console.log(`Navigating to parent: ${item.name} -> ${item.url}`);
                      handleLinkClick(item.url, item.external);
                    } else {
                      setActiveDropdown(activeDropdown === item.name ? null : item.name);
                    }
                  } else {
                    handleLinkClick(item.url, item.external);
                  }
                }}
                title={item.children && item.url && item.url !== '#' ? `Go to ${item.name} page` : undefined}
              >
                {item.icon && (
                  <Icon iconName={item.icon} className={styles.navIcon} />
                )}
                <span className={styles.navText}>{item.name}</span>
                {item.children && (
                  <Icon iconName="ChevronDown" className={styles.dropdownArrow} />
                )}
              </a>


              {item.children && (
                <div className={`${styles.dropdown} ${activeDropdown === item.name ? styles.show : ''}`}>
                  <div className={styles.dropdownHeader}>
                    {item.name} Dropdown
                  </div>
                  <ul className={styles.dropdownList}>
                    {item.children.map((child, childIndex) => (
                      <li key={childIndex} className={styles.dropdownItem}>
                        <a 
                          href="#"
                          className={styles.dropdownLink}
                          onClick={(e) => {
                            e.preventDefault();
                            handleLinkClick(child.url, child.external);
                            setActiveDropdown(null);
                          }}
                        >
                          {child.name}
                        </a>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </li>
          ))}
        </ul>
      </nav>
    </div>
  );
};


export default NavigationMenu;
