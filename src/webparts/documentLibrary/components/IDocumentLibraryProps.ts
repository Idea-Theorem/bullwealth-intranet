export interface IDocument {
  id: number;
  name: string;
  fileType: string;
  modified: string;
  modifiedBy: string;
  serverRelativeUrl: string;
  downloadUrl: string;
  iconName: string;
  description: string;
  createdDate: string;
  // ✅ FIX 5: Add timestamp properties for accurate sorting
  modifiedTimestamp?: number;
  createdTimestamp?: number;
  orderBy?: number;  // ✅ ADD THIS LINE
}

export interface IDocumentLibraryProps {
  title: string;
  listName: string;
  context: any;
}
