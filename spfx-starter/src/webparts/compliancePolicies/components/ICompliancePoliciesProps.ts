export interface IPolicyDocument {
  id: string;
  title: string;
  dateAdded?: string;
  dateModified?: string;
  description: string;
  documentUrl: string;
  category?: string;
}

export interface ICompliancePoliciesProps {
  sectionTitle: string;
  categoryTitle: string;
  viewAllText: string;
  viewAllUrl: string;
  policies: IPolicyDocument[];
  columnsPerRow: number;
  showExportButton: boolean;
  showShareButton: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}