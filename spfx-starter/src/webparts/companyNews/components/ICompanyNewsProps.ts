import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface INewsItem {
  id: string;
  title: string;
  author: string;
  date: string;
  imageUrl: string;
  description: string;
  shareUrl?: string;
  readMoreUrl?: string;
}

export interface ICompanyNewsProps {
  title: string;
  newsItems: INewsItem[];
  itemsToShow: number;
  autoScroll: boolean;
  autoScrollInterval: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}