export interface INavigationMenuProps {
  homeUrl: string;
  bullWealthUrl: string;
  cloverUrl: string;
  hrUrl: string;
  itPolicyUrl: string;
  helpUrl: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  context: any;
}

export interface INavigationItem {
  key: string;
  text: string;
  url: string;
  iconName: string;
  isActive?: boolean;
  hasDropdown?: boolean;
  subItems?: ISubItem[];
}

export interface ISubItem {
  key: string;
  text: string;
  url: string;
  iconName: string;
  description?: string;
}
