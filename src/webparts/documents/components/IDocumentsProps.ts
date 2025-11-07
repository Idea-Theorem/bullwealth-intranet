import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentCategory {
  id: string;
  title: string;
  imageData: string;
  libraryUrl: string;
  viewAllUrl?: string;
  viewDocumentsText?: string;
  folderName?: string; // NEW: Actual folder name
  pageUrl?: string; // NEW: Link to site page
  orderBy?: number; // ✅ NEW
  documentOrderMap?: { [key: string]: number }; // ✅ NEW: Document order map
  libraryParam?: string; // ✅ added field for folder navigation

   /** 👇 Add these two lines */
  modified?: string; // TimeLastModified (for Latest/Oldest sorting)
  created?: string;  // TimeCreated (optional, if needed later)
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
  // NEW: Dynamic mode settings
  isDynamicMode: boolean;
  documentLibraryName: string;
  folderPath: string;
  sitePageBasePath: string;
}
