import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMessagesProps {
  description: string;
  context: WebPartContext;
  listName: string;
  title: string;
  columnsPerRow: number;
}

export interface IMessageItem {
  Id: number;
  Title: string;
  Content: string;
  PublishedDate: string;
  Category: string;
  Year: number;
  FeaturedImage?: {
    Url: string;
    Description?: string; // Make Description optional
  };
  Author: {
    Title: string;
  };
}

export interface IGroupedMessages {
  [year: string]: IMessageItem[];
}

export enum ViewMode {
  CardView = 'card',
  DetailView = 'detail'
}

export enum GridLayout {
  OneColumn = 1,
  TwoColumns = 2,
  ThreeColumns = 3,
  FourColumns = 4,
  FiveColumns = 5
}
