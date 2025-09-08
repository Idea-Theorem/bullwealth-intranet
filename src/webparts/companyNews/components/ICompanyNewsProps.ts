import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICompanyNewsProps {
  title: string;
  newsItems: INewsItem[];
  autoScroll?: boolean;
  autoScrollInterval?: number;
  itemsToShow?: number;
  showDots?: boolean;
  showArrows?: boolean;
  context: WebPartContext;
}

export interface INewsItem {
  id: string;
  title: string;
  author: string;
  date: string;
  imageUrl: string;
  readMoreUrl: string;
  shareUrl: string;
}
