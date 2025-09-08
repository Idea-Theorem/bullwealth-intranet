export interface IDocumentCategory {
  id: string;
  title: string;
  imageData: string; // Base64 encoded image
  libraryUrl: string; // Main URL field (changed from libraryName)
  viewAllUrl?: string; // Optional: Alternative view URL
}

export interface IDocumentsProps {
  title: string; // ✅ Added missing title property
  categories: IDocumentCategory[];
  columnsPerRow: number;
  context: any;
  isDarkTheme?: boolean; // ✅ Added optional properties
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  onCategoriesUpdate: (categories: IDocumentCategory[]) => void;
}
