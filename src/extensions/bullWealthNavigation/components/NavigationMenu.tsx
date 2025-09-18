import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { INavigationMenuProps, INavigationItem } from './INavigationProps';
import styles from './NavigationMenu.module.scss';

const NavigationMenu: React.FC<INavigationMenuProps> = ({ items, siteUrl }) => {
  const [activeDropdown, setActiveDropdown] = useState<string | null>(null);

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
        url: '/sites/BullWealthIntranet/bullwealth', // Make parent clickable
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
              className={`${styles.navItem} ${item.children ? styles.dropdown : ''} ${item.name === 'Home' ? styles.activeItem : ''}`}
              onMouseEnter={() => item.children && setActiveDropdown(item.name)}
              onMouseLeave={() => item.children && setActiveDropdown(null)}
            >
              <a 
                href="#"
                className={styles.navLink}
                onClick={(e) => {
                  e.preventDefault();
                  
                  // UPDATED: Handle parent item clicks
                  if (item.children) {
                    // Parent has children - check if it should be clickable
                    if (item.url && item.url !== '#') {
                      // Parent is clickable - navigate to its URL
                      console.log(`Navigating to parent: ${item.name} -> ${item.url}`);
                      handleLinkClick(item.url, item.external);
                    } else {
                      // Parent is just a dropdown container - toggle dropdown
                      setActiveDropdown(activeDropdown === item.name ? null : item.name);
                    }
                  } else {
                    // No children, just navigate
                    handleLinkClick(item.url, item.external);
                  }
                }}
                // Add title to show it's clickable
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
