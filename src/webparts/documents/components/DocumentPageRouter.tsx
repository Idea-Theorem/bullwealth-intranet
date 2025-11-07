/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';

interface IDynamicPageProps {
  context: any;
  folderName: string;
  documentLibraryName: string;
  baseLibraryPath: string;
}

export interface IDocumentItem {
  id: string;
  name: string;
  serverRelativeUrl: string;
  modified: string;
  modifiedBy: string;
  fileSize: number;
  contentTypeId: string;
}

const DocumentPageRouter: React.FC<IDynamicPageProps> = ({
  context,
  folderName,
  documentLibraryName,
  baseLibraryPath
}) => {
  const [documents, setDocuments] = useState<IDocumentItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchDocumentsForFolder();
  }, [folderName]);

  const fetchDocumentsForFolder = async () => {
    setLoading(true);
    setError(null);
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const sitePath = context.pageContext.web.serverRelativeUrl;

      const folderPath = `${sitePath}/${documentLibraryName}/${baseLibraryPath}/${folderName}`.replace(/\/+/g, '/');

      const filesApiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@path)/Files?@path='${encodeURIComponent(folderPath)}'&$select=Name,ServerRelativeUrl,TimeLastModified,ModifiedBy/Title,Length&$expand=ModifiedBy&$top=5000`;

      const response = await context.spHttpClient.get(
        filesApiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch documents: ${response.status}`);
      }

      const data = await response.json();
      const docs: IDocumentItem[] = (data.value || []).map((file: any) => ({
        id: file.Name,
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        modified: new Date(file.TimeLastModified).toLocaleDateString(),
        modifiedBy: file.ModifiedBy?.Title || 'Unknown',
        fileSize: file.Length,
        contentTypeId: file.ContentTypeId
      }));

      setDocuments(docs);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
      console.error('Document fetch error:', err);
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return <div>Loading documents...</div>;
  }

  if (error) {
    return <div style={{ color: 'red' }}>Error: {error}</div>;
  }

  return (
    <div style={{ padding: '20px' }}>
      <h2>{folderName}</h2>
      <p>Total documents: {documents.length}</p>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr style={{ borderBottom: '2px solid #ddd' }}>
            <th style={{ textAlign: 'left', padding: '10px' }}>Name</th>
            <th style={{ textAlign: 'left', padding: '10px' }}>Modified</th>
            <th style={{ textAlign: 'left', padding: '10px' }}>Modified By</th>
            <th style={{ textAlign: 'left', padding: '10px' }}>Size</th>
          </tr>
        </thead>
        <tbody>
          {documents.map((doc) => (
            <tr key={doc.id} style={{ borderBottom: '1px solid #eee' }}>
              <td style={{ padding: '10px' }}>
                <a href={`${context.pageContext.web.absoluteUrl}${doc.serverRelativeUrl}`} target="_blank" rel="noopener noreferrer">
                  {doc.name}
                </a>
              </td>
              <td style={{ padding: '10px' }}>{doc.modified}</td>
              <td style={{ padding: '10px' }}>{doc.modifiedBy}</td>
              <td style={{ padding: '10px' }}>{(doc.fileSize / 1024).toFixed(2)} KB</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default DocumentPageRouter;
