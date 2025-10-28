export interface ICompanyNewsProps {
  title: string;
  newsItems: INewsItem[];
  itemsToShow: number;
  autoScroll: boolean;
  autoScrollInterval: number;
  showDots: boolean;
  context: any;
}

export interface INewsItem {
  id: string;
  title: string;
  author: string;
  date: string;
  imageUrl: string;
  readMoreUrl: string;
  shareUrl?: string; // ✅ Made optional with ?
}
