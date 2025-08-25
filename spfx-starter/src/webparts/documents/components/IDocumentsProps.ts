export interface IDocumentCategory {
  id: string;
  title: string;
  imageUrl: string;
  documentUrl: string;
  viewAllUrl: string;
}

export interface IDocumentsProps {
  title: string;
  categories: IDocumentCategory[];
  columnsPerRow: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}