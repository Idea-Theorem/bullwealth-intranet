import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentCategory {
  imageUrl: string;
  id: string;
  title: string;
  imageData: string; // Base64 encoded image
  libraryUrl: string;
  libraryName: string;
  viewAllUrl: string;
}

export interface IDocumentsProps {
  title: string;
  categories: IDocumentCategory[];
  columnsPerRow: number;
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  onCategoriesUpdate: (categories: IDocumentCategory[]) => void;
}