import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentCategory {
  id: string;
  title: string;
  imageData: string; // Base64 encoded image
  libraryUrl: string; // Main URL field
  viewAllUrl?: string; // Optional: Alternative view URL
  viewDocumentsText?: string; // NEW: Customizable link text
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
