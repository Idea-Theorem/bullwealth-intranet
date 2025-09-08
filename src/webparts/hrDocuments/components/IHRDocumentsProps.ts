export interface IHRDocument {
  id: string;
  title: string;
  date: string;
  iconData: string; // Base64 encoded icon or blob URL
  iconType: 'word' | 'pdf' | 'custom';
  documentUrl: string;
}

export interface IHRDocumentsProps {
  title: string;
  documents: IHRDocument[];
  columnsPerRow: number;
  showDate: boolean;
  allowUpload: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  onDocumentsUpdate: (documents: IHRDocument[]) => void;
}