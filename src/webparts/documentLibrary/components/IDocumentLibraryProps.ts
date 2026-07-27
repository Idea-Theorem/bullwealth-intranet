export interface IDocumentLibraryProps {
  context: any;
  title: string;
  listName: string;
}

export interface IDocument {
  title: string;
  id: string | number;
  name: string;
  fileType: string;
  modified: string;
  modifiedBy: string;
  serverRelativeUrl: string;
  downloadUrl: string;
  iconName: string;
  description: string;
  createdDate: string;
  modifiedTimestamp: number;
  createdTimestamp: number;
}
