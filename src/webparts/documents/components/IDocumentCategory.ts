export interface IDocumentCategory {
  id: string;
  title: string;
  folderName: string; // ✅ REQUIRED - not optional
  imageData: string;
  libraryUrl: string;
  viewAllUrl?: string;
  viewDocumentsText?: string;
  pageUrl?: string;
  folderServerRelativeUrl?: string;
  orderBy?: number;
}
