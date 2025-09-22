export interface IHRDocument {
  id: string;
  title: string;
  date: string;
  author: string;
  iconData: string;
  iconType: 'word' | 'pdf' | 'video' | 'custom';
  documentUrl: string;
}

export interface IHRDocumentsProps {
  title: string;
  documents: IHRDocument[];
  columnsPerRow: number;
  showDate: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any;
  onDocumentsUpdate: (documents: IHRDocument[]) => void;
}
