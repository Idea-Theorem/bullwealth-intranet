export interface INavigationItem {
  name: string;
  url: string;
  icon?: string;
  external?: boolean;
  children?: INavigationItem[];
}

export interface INavigationMenuProps {
  items: INavigationItem[];
  siteUrl: string;
}

// Additional interfaces for CSS modules
export interface INavigationStyles {
  navigationWrapper: string;
  navigationMenu: string;
  brand: string;
  brandTitle: string;
  navList: string;
  navItem: string;
  hasDropdown: string;
  activeItem: string;
  navLink: string;
  navIcon: string;
  navText: string;
  dropdownArrow: string;
  dropdown: string;
  show: string;
  dropdownHeader: string;
  dropdownList: string;
  dropdownItem: string;
  dropdownLink: string;
}
