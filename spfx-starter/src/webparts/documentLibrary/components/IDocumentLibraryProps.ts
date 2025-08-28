import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocument {
  execCommand(arg0: string): unknown;
  body: any;
  createElement(arg0: string): unknown;
  id: string;
  name: string;
  modified: string;
  modifiedBy: string;
  fileType: string;
  fileUrl: string;
  isSelected: boolean;
}

export interface IDocumentLibraryProps {
  title: string;
  description: string;
  libraryName: string;
  showUploadButton: boolean;
  showActions: boolean;
  pageSize: number;
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}