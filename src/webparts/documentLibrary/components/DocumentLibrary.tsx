/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-unused-vars */

import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './DocumentLibrary.module.scss';
import { IDocumentLibraryProps, IDocument } from './IDocumentLibraryProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { ContextualMenu, IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

interface IFolderWithDocuments {
  name: string;
  documents: IDocument[];
  allDocuments: IDocument[];
  folderPath: string;
  orderBy?: number;
}

const DocumentLibrary: React.FC<IDocumentLibraryProps> = (props) => {
  const [foldersWithDocuments, setFoldersWithDocuments] = useState<IFolderWithDocuments[]>([]);
  const [currentFolder, setCurrentFolder] = useState<string>('');
  const [currentDocuments, setCurrentDocuments] = useState<IDocument[]>([]);
  const [selectedDocument, setSelectedDocument] = useState<IDocument | null>(null);
  const [contextMenuTarget, setContextMenuTarget] = useState<HTMLElement | null>(null);
  const [checkedItems, setCheckedItems] = useState<Set<string>>(new Set());
  const [selectAllChecked, setSelectAllChecked] = useState<boolean>(false);
  const [message, setMessage] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [pageTitle, setPageTitle] = useState<string>('Documents');
  const [shareDoc, setShareDoc] = useState<IDocument | null>(null);
  const [showSharePanel, setShowSharePanel] = useState<boolean>(false);

  useEffect(() => {
    if (!currentFolder) {
      loadFolderStructure().catch((error) => {
        console.error('Failed to load folder structure:', error);
        setMessage(`❌ Error: ${error.message || 'Failed to load documents'}`);
        setIsLoading(false);
      });

      const urlParams = new URLSearchParams(window.location.search);
      const libraryPath = urlParams.get('library');
      if (libraryPath) {
        const pathParts = decodeURIComponent(libraryPath).split('/');
        const lastFolderName = pathParts[pathParts.length - 1];
        setPageTitle(lastFolderName || 'Documents');
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.listName, currentFolder]);

  useEffect(() => {
    if (currentFolder) {
      setCheckedItems(new Set());
      setSelectAllChecked(false);
    }
  }, [currentFolder]);

  const formatDate = (date: Date): string => {
    return new Intl.DateTimeFormat('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    }).format(date);
  };

  const sortDocumentsByOrderOnly = (documents: IDocument[]): IDocument[] => {
    return [...documents].sort((a, b) => {
      const aOrder = (a as any).orderBy !== undefined ? (a as any).orderBy : 999;
      const bOrder = (b as any).orderBy !== undefined ? (b as any).orderBy : 999;
      return aOrder - bOrder;
    });
  };

  const loadFolderStructure = async (): Promise<void> => {
    setIsLoading(true);
    try {
      const baseUrl = props.context.pageContext.web.absoluteUrl;
      const pathParts = props.listName.split('/');
      const mainLibrary = pathParts[0];
      const targetFolder = pathParts.slice(1).join('/');

      await tryFolderPaths(baseUrl, mainLibrary, targetFolder);
    } catch (error: any) {
      console.error('Error in loadFolderStructure:', error);
      setMessage(`❌ Error loading folder "${props.listName}": ${error.message || 'Unknown error'}`);
      setTimeout(() => setMessage(''), 8000);
      setFoldersWithDocuments([]);
    } finally {
      setIsLoading(false);
    }
  };

  const tryFolderPaths = async (baseUrl: string, mainLibrary: string, targetFolder: string): Promise<void> => {
    const siteName = baseUrl.split('/').pop();

    const folderPaths = [
      `/sites/${siteName}/${mainLibrary}/${targetFolder}`,
      `/${mainLibrary}/${targetFolder}`,
      `/sites/${siteName}/Shared Documents/${targetFolder}`,
      `/Shared Documents/${targetFolder}`,
      `/sites/${siteName}/Documents/${targetFolder}`,
      `/Documents/${targetFolder}`
    ];

    for (const folderPath of folderPaths) {
      try {
        const folderUrl =
          `${baseUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')` +
          `?$expand=Folders,Files,ListItemAllFields`;
        const response = await props.context.spHttpClient.get(
          folderUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          if (data && typeof data === 'object') {
            await processFolderData(data, folderPath);
            return;
          }
        }
      } catch (error) {
        console.error(`Failed to load folder path: ${folderPath}`, error);
        continue;
      }
    }

    setMessage(`⚠️ Could not find folder "${props.listName}"`);
    setTimeout(() => setMessage(''), 10000);
    setFoldersWithDocuments([]);
  };

  const getAllFolderMetadata = async (folderPath: string): Promise<Record<string, any>> => {
    const baseUrl = props.context.pageContext.web.absoluteUrl;

    const libraryNames = [
      'Documents',
      'Shared Documents',
      'Site Assets',
      'Style Library'
    ];

    for (const libraryTitle of libraryNames) {
      try {
        const itemsUrl =
          `${baseUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitle)}')/items?` +
          `$select=ID,Title,FileLeafRef,FileRef,FSObjType,OrderBy&` +
          `$top=5000`;

        const res = await props.context.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);

        if (res.ok) {
          const data = await res.json();

          if (data.value.length === 0) {
            continue;
          }

          const lookup: Record<string, any> = {};
          const folderNameLookup: Record<string, any> = {};

          let folderCount = 0;

          for (const item of data.value) {
            const isFolderItem = item.FSObjType === 1;

            if (isFolderItem) {
              folderCount++;
            }

            let orderByValue = 999;
            if (item.OrderBy !== null && item.OrderBy !== undefined && item.OrderBy !== '') {
              const parsed = parseInt(String(item.OrderBy), 10);
              if (!isNaN(parsed)) {
                orderByValue = parsed;
              }
            }

            const metadata = {
              OrderBy: orderByValue,
              FileRef: item.FileRef,
              FileLeafRef: item.FileLeafRef,
              FSObjType: item.FSObjType,
              isFolder: isFolderItem,
              Title: item.Title
            };

            if (item.FileRef) {
              const pathKey = item.FileRef.toLowerCase();
              lookup[pathKey] = metadata;
            }

            if (isFolderItem && item.FileLeafRef) {
              const names = [
                item.FileLeafRef,
                item.Title,
                item.FileLeafRef.replace(/-/g, ' '),
                item.FileLeafRef.replace(/\s+/g, '-'),
                item.FileLeafRef.replace(/[\s-]+/g, '').toLowerCase()
              ].filter(Boolean);

              for (const name of names) {
                if (name) {
                  const key = String(name).toLowerCase().trim();
                  folderNameLookup[key] = metadata;
                }
              }
            }
          }

          return {
            ...lookup,
            __folderNameLookup: folderNameLookup
          };
        }
      } catch (error: any) {
        continue;
      }
    }

    return {};
  };

  const extractUrlFromFile = async (serverRelativeUrl: string): Promise<string | null> => {
    const baseUrl = props.context.pageContext.web.absoluteUrl;

    try {
      const fileContentUrl = `${baseUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`;

      const response = await props.context.spHttpClient.get(
        fileContentUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const fileContent = await response.text();

        const urlPatterns = [
          /URL=(.+?)(?:\r|\n|$)/i,
          /URL\s*=\s*(.+?)(?:\r|\n|$)/i,
          /\[InternetShortcut\][\s\S]*?URL=(.+?)(?:\r|\n|$)/i,
          /(https?:\/\/[^\s\r\n]+)/i
        ];

        for (const pattern of urlPatterns) {
          const match = fileContent.match(pattern);
          if (match && match[1]) {
            const extractedUrl = match[1].trim();
            return extractedUrl;
          }
        }
      }
    } catch (error) {
      console.error('❌ Error extracting URL:', error);
    }

    return null;
  };

  const getDocumentsWithRealUsers = async (
    files: any[],
    metadataLookup: Record<string, any>
  ): Promise<IDocument[]> => {
    const documentPromises = files.map(async (file) => {
      const fileName = file.Name || 'Unknown';
      const serverRelativeUrl = file.ServerRelativeUrl || '';

      const lookupKey = serverRelativeUrl.toLowerCase();
      const itemMeta = metadataLookup[lookupKey];

      let modifiedBy = 'Unknown';
      let orderBy: number = 999;

      if (itemMeta) {
        modifiedBy = itemMeta.EditorTitle || 'Unknown';
        orderBy = itemMeta.OrderBy ?? 999;
      }

      const isUrlFile = fileName.toLowerCase().endsWith('.url');

      const isLink =
        file.ListItemAllFields &&
        file.ListItemAllFields.File_x0020_Type === 'url';

      let openUrl = `${window.location.protocol}//${window.location.host}${serverRelativeUrl}`;
      let actualUrl = openUrl;

      if (isUrlFile) {
        const extractedUrl = await extractUrlFromFile(serverRelativeUrl);
        if (extractedUrl) {
          actualUrl = extractedUrl;
        }
      } else if (isLink && file.ListItemAllFields.URL) {
        actualUrl = file.ListItemAllFields.URL;
      }

      const fileType = isUrlFile || isLink ? 'url' : (fileName.split('.').pop() || 'file');

      const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
      const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

      return {
        id: file.UniqueId || Math.random().toString(),
        name: fileName.replace(/\.[^/.]+$/, ''),
        fileType,
        modified: formatDate(modifiedDate),
        modifiedBy,
        serverRelativeUrl: actualUrl,
        originalServerRelativeUrl: serverRelativeUrl,
        downloadUrl: (isLink || isUrlFile) ? undefined : serverRelativeUrl,
        iconName: (isLink || isUrlFile) ? 'Globe' : getFileIcon(fileType),
        description: fileName.replace(/\.[^/.]+$/, ''),
        createdDate: formatDate(createdDate),
        modifiedTimestamp: modifiedDate.getTime(),
        createdTimestamp: createdDate.getTime(),
        orderBy
      } as any;
    });

    const resolvedDocuments = await Promise.all(documentPromises);
    return resolvedDocuments;
  };

  const processFolderData = async (data: any, folderPath: string): Promise<void> => {
    try {
      const subfolders = data.Folders || [];
      const files = data.Files || [];

      if (subfolders.length === 0 && files.length === 0) {
        setMessage(`⚠️ Folder "${props.listName}" is empty`);
        setTimeout(() => setMessage(''), 8000);
        setFoldersWithDocuments([]);
        return;
      }

      const allMetadata = await getAllFolderMetadata(folderPath);
      const folderNameLookup = allMetadata.__folderNameLookup || {};

      const folderGroups: { [key: string]: { documents: IDocument[], folderPath: string, orderBy?: number } } = {};

      if (files.length > 0) {
        const mappedDocuments = await getDocumentsWithRealUsers(files, allMetadata);
        const sortedDocuments = sortDocumentsByOrderOnly(mappedDocuments);

        const pathSegments = folderPath.split('/').filter(segment => segment.length > 0);
        let displayName = 'Documents';

        if (pathSegments.length > 0) {
          const lastSegment = pathSegments[pathSegments.length - 1];
          if (lastSegment && !['sites', 'Shared Documents', 'Documents'].includes(lastSegment)) {
            displayName = lastSegment;
          } else if (pathSegments.length > 1) {
            const secondLast = pathSegments[pathSegments.length - 2];
            if (secondLast && !['sites', 'Shared Documents', 'Documents'].includes(secondLast)) {
              displayName = secondLast;
            }
          }
        }

        if (displayName === 'Documents' || displayName === 'sites') {
          const propsParts = props.listName.split('/').filter(part => part.length > 0);
          if (propsParts.length > 1) {
            displayName = propsParts[propsParts.length - 1];
          }
        }

        folderGroups[displayName] = {
          documents: sortedDocuments,
          folderPath: folderPath
        };
      }

      const subfolderPromises = subfolders.map(async (subfolder: any) => {
        const subfolderName = subfolder.Name;
        const subfolderPath = subfolder.ServerRelativeUrl;

        try {
          const subfolderUrl =
            `${props.context.pageContext.web.absoluteUrl}` +
            `/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')` +
            `/Files?$expand=ListItemAllFields`;

          const subfolderResponse = await props.context.spHttpClient.get(
            subfolderUrl,
            SPHttpClient.configurations.v1
          );

          if (subfolderResponse.ok) {
            const subfolderData = await subfolderResponse.json();
            const subfolderFiles = subfolderData.value || [];

            if (subfolderFiles.length > 0) {
              const mappedDocuments = await getDocumentsWithRealUsers(subfolderFiles, allMetadata);
              const sortedDocuments = sortDocumentsByOrderOnly(mappedDocuments);

              let folderOrderBy: number = 999;

              const searchKeys = [
                subfolderName.toLowerCase().trim(),
                subfolderName.toLowerCase().replace(/\s+/g, '-'),
                subfolderName.toLowerCase().replace(/-/g, ' '),
                subfolderName.toLowerCase().replace(/[\s-]+/g, '')
              ];

              for (const searchKey of searchKeys) {
                if (folderNameLookup[searchKey]) {
                  folderOrderBy = folderNameLookup[searchKey].OrderBy ?? 999;
                  break;
                }
              }

              if (folderOrderBy === 999) {
                const pathKey = subfolderPath.toLowerCase();
                if (allMetadata[pathKey]) {
                  folderOrderBy = allMetadata[pathKey].OrderBy ?? 999;
                }
              }

              return {
                name: subfolderName,
                documents: sortedDocuments,
                folderPath: subfolderPath,
                orderBy: folderOrderBy
              };
            }
          }
        } catch (error) {
          console.error(`❌ Error processing subfolder ${subfolderName}:`, error);
        }
        return null;
      });

      const resolvedSubfolders = await Promise.all(subfolderPromises);

      resolvedSubfolders.forEach(subfolder => {
        if (subfolder && subfolder.documents.length > 0) {
          folderGroups[subfolder.name] = {
            documents: subfolder.documents,
            folderPath: subfolder.folderPath,
            orderBy: subfolder.orderBy
          };
        }
      });

      const foldersWithDocs: IFolderWithDocuments[] = [];

      for (const folderName of Object.keys(folderGroups)) {
        const sortedDocs = sortDocumentsByOrderOnly(folderGroups[folderName].documents);
        const folderOrderBy = folderGroups[folderName].orderBy ?? 999;

        foldersWithDocs.push({
          name: folderName,
          documents: sortedDocs.slice(0, 4),
          allDocuments: sortedDocs,
          folderPath: folderGroups[folderName].folderPath,
          orderBy: folderOrderBy
        });
      }

      foldersWithDocs.sort((a, b) => {
        const aOrder = a.orderBy !== undefined ? a.orderBy : 999;
        const bOrder = b.orderBy !== undefined ? b.orderBy : 999;
        return aOrder - bOrder;
      });

      if (foldersWithDocs.length === 0) {
        setFoldersWithDocuments([]);
        setMessage(`⚠️ No documents found`);
        setTimeout(() => setMessage(''), 8000);
      } else {
        setFoldersWithDocuments(foldersWithDocs);
        const totalDocs = foldersWithDocs.reduce((sum, folder) => sum + folder.allDocuments.length, 0);
        setMessage(`✅ Loaded ${totalDocs} documents from ${foldersWithDocs.length} folders`);
        setTimeout(() => setMessage(''), 5000);
      }
    } catch (error: any) {
      console.error('❌ Error in processFolderData:', error);
      setMessage(`❌ Error processing folder data: ${error.message || 'Unknown error'}`);
      setTimeout(() => setMessage(''), 8000);
      setFoldersWithDocuments([]);
    }
  };

  const handleDocumentClick = async (doc: IDocument): Promise<void> => {
    if (!doc.serverRelativeUrl || doc.serverRelativeUrl === '#') {
      setMessage('❌ Document URL not available');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    try {
      if (doc.fileType === 'url' || doc.iconName === 'Globe') {
        let urlToOpen = doc.serverRelativeUrl;

        if (urlToOpen.includes('.sharepoint.com') && urlToOpen.includes('.url')) {
          const originalUrl = (doc as any).originalServerRelativeUrl || doc.serverRelativeUrl;
          const extractedUrl = await extractUrlFromFile(originalUrl);

          if (extractedUrl) {
            urlToOpen = extractedUrl;
          }
        }

        window.open(urlToOpen, '_blank', 'noopener,noreferrer');
      } else {
        window.open(doc.serverRelativeUrl, '_blank');
      }
    } catch (error) {
      console.error('Error opening document:', error);
      setMessage('❌ Unable to open document. Please check your permissions.');
      setTimeout(() => setMessage(''), 5000);
    }
  };

  const getFileIcon = (fileType: string): string => {
    switch (fileType.toLowerCase()) {
      case 'url': return 'Globe';
      case 'pdf': return 'PDF';
      case 'doc':
      case 'docx': return 'WordDocument';
      case 'xls':
      case 'xlsx': return 'ExcelDocument';
      case 'ppt':
      case 'pptx': return 'PowerPointDocument';
      default: return 'Page';
    }
  };

  const handleDownloadDocument = async (doc: IDocument): Promise<void> => {
    if (doc.fileType === 'url') {
      await handleDocumentClick(doc);
      return;
    }

    if (doc.downloadUrl) {
      const downloadUrl = doc.downloadUrl.startsWith('http')
        ? doc.downloadUrl
        : `${window.location.protocol}//${window.location.host}${doc.downloadUrl}`;

      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `${doc.name}.${doc.fileType}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  // ─── SHARE FUNCTIONS ───────────────────────────────────────────────────────

  const handleShareDocument = (doc: IDocument): void => {
    setShareDoc(doc);
    setShowSharePanel(true);
  };

  const getEncodedUrl = (doc: IDocument): string => {
    const shareUrl = doc.serverRelativeUrl || window.location.href;
    const absoluteUrl = shareUrl.startsWith('http')
      ? shareUrl
      : `${window.location.protocol}//${window.location.host}${shareUrl}`;
    return absoluteUrl.replace(/ /g, '%20');
  };

  const copyToClipboard = (url: string, documentName: string): void => {
    if (navigator.clipboard) {
      navigator.clipboard.writeText(url).then(() => {
        setMessage(`Link for "${documentName}" copied to clipboard!`);
        setTimeout(() => setMessage(''), 3000);
      }).catch(() => {
        setMessage('Failed to copy link to clipboard');
        setTimeout(() => setMessage(''), 3000);
      });
    } else {
      const textArea = window.document.createElement('textarea');
      textArea.value = url;
      window.document.body.appendChild(textArea);
      textArea.select();
      try {
        window.document.execCommand('copy');
        setMessage(`Link for "${documentName}" copied to clipboard!`);
        setTimeout(() => setMessage(''), 3000);
      } catch {
        setMessage('Failed to copy link to clipboard');
        setTimeout(() => setMessage(''), 3000);
      }
      window.document.body.removeChild(textArea);
    }
  };

  // Opens Outlook desktop app via mailto with URL on its own line
  const openOutlookAppShare = async (doc: IDocument, encodedUrl: string): Promise<void> => {
  const displayName = doc.name || (doc as any).title;
  const subject = encodeURIComponent(`Sharing: ${displayName}`);

  let shareUrl = encodedUrl;

  try {
    const res = await fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(encodedUrl)}`);
    if (res.ok) {
      const short = await res.text();
      if (short.startsWith('https://tinyurl.com')) {
        shareUrl = short.trim();
      }
    }
  } catch (e) { /* fallback */ }

  const body = encodeURIComponent(`Hi,\r\n\r\nPlease find the link below:\r\n\r\n${displayName}\r\n${shareUrl} \r\n\r\nBest regards,`);

  window.location.href = `mailto:?subject=${subject}&body=${body}`;
  setShowSharePanel(false);
};

  // Opens Outlook Web (OWA) with proper HTML hyperlink
  // const openOutlookWebShare = (doc: IDocument, encodedUrl: string): void => {
  //   const displayName = doc.name;
  //   const subject = encodeURIComponent(`Sharing: ${displayName}`);
  //   const htmlBody = encodeURIComponent(
  //     `<p>Hi,</p>` +
  //     `<p>Please find the link below:</p>` +
  //     `<p><b>${displayName}</b></p>` +
  //     `<p><a href="${encodedUrl}" style="color:#0078d4;text-decoration:underline;">Click here to view</a></p>` +
  //     `<p>Kindly let me know if you face any access issues.</p>` +
  //     `<p>Best regards,</p>`
  //   );
  //   const owaUrl = `https://outlook.office365.com/mail/deeplink/compose?subject=${subject}&body=${htmlBody}&ishtml=true`;
  //   window.open(owaUrl, '_blank', 'noopener,noreferrer');
  //   setShowSharePanel(false);
  // };

  // ─── END SHARE FUNCTIONS ───────────────────────────────────────────────────

  const handleSelectAllChange = (): void => {
    const newSelectAll = !selectAllChecked;
    setSelectAllChecked(newSelectAll);

    if (newSelectAll) {
      setCheckedItems(new Set(currentDocuments.map((doc: IDocument) => String(doc.id))));
    } else {
      setCheckedItems(new Set());
    }
  };

  const handleCheckboxChange = (documentId: string): void => {
    const newCheckedItems = new Set(checkedItems);
    if (checkedItems.has(documentId)) {
      newCheckedItems.delete(documentId);
    } else {
      newCheckedItems.add(documentId);
    }
    setCheckedItems(newCheckedItems);
    setSelectAllChecked(newCheckedItems.size === currentDocuments.length && currentDocuments.length > 0);
  };

  const handleViewAll = (folderName: string): void => {
    const folder = foldersWithDocuments.find((f: IFolderWithDocuments) => f.name === folderName);
    if (folder && folder.allDocuments.length > 0) {
      setCurrentDocuments(folder.allDocuments);
      setCurrentFolder(folderName);
    }
  };

  const handleBackClick = (): void => {
    setCurrentFolder('');
    setCurrentDocuments([]);
    setCheckedItems(new Set());
    setSelectAllChecked(false);
  };

  const handleDocumentActions = (event: React.MouseEvent<HTMLElement>, doc: IDocument): void => {
    event.preventDefault();
    event.stopPropagation();
    setSelectedDocument(doc);
    setContextMenuTarget(event.currentTarget as HTMLElement);
  };

  const dismissContextMenu = (): void => {
    setContextMenuTarget(null);
    setSelectedDocument(null);
  };

  const getContextMenuItems = (): IContextualMenuItem[] => {
    if (!selectedDocument) return [];

    return [
      {
        key: 'view',
        text: 'View',
        iconProps: { iconName: 'View' },
        onClick: () => {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          handleDocumentClick(selectedDocument);
          dismissContextMenu();
        }
      },
      {
        key: 'share',
        text: 'Share',
        iconProps: { iconName: 'Share' },
        onClick: () => {
          handleShareDocument(selectedDocument);
          dismissContextMenu();
        }
      },
      {
        key: 'export',
        text: selectedDocument.fileType === 'url' ? 'Open Link' : 'Export',
        iconProps: { iconName: selectedDocument.fileType === 'url' ? 'Globe' : 'Download' },
        onClick: () => {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          handleDownloadDocument(selectedDocument);
          dismissContextMenu();
        }
      },
      {
        key: 'copyLink',
        text: 'Copy link',
        iconProps: { iconName: 'Link' },
        onClick: () => {
          if (selectedDocument.serverRelativeUrl) {
            copyToClipboard(selectedDocument.serverRelativeUrl, selectedDocument.name);
          }
          dismissContextMenu();
        }
      }
    ];
  };

  // ─── SHARE PANEL JSX ──────────────────────────────────────────────────────
  const renderSharePanel = (): JSX.Element | null => {
    if (!showSharePanel || !shareDoc) return null;

    const encodedUrl = getEncodedUrl(shareDoc);

    return (
      <div style={{
        position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
        backgroundColor: 'rgba(0,0,0,0.4)', zIndex: 9999,
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      }} onClick={() => setShowSharePanel(false)}>
        <div style={{
          background: '#fff', borderRadius: '12px', padding: '24px',
          width: '380px', boxShadow: '0 8px 32px rgba(0,0,0,0.18)'
        }} onClick={(e) => e.stopPropagation()}>

          {/* Header */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, fontSize: '16px' }}>Share</span>
            <button onClick={() => setShowSharePanel(false)}
              style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '18px', color: '#605e5c' }}>✕</button>
          </div>

          <p style={{ fontSize: '13px', color: '#605e5c', marginBottom: '20px', wordBreak: 'break-all' }}>
            {shareDoc.name}
          </p>

          {/* Share icons */}
          <div style={{ display: 'flex', gap: '16px', justifyContent: 'center', marginBottom: '24px', flexWrap: 'wrap' }}>

            {/* Outlook Desktop */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }}
              onClick={() => { void openOutlookAppShare(shareDoc, encodedUrl); }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#0078d4', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="OutlookLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Outlook<br />App</span>
            </div>

            {/* Outlook Web */}
            {/* <div style={{ textAlign: 'center', cursor: 'pointer' }}
              onClick={() => openOutlookWebShare(shareDoc, encodedUrl)}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#005a9e', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="OutlookLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Outlook<br />Web</span>
            </div> */}

            {/* Teams */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://teams.microsoft.com/share?href=${encodeURIComponent(encodedUrl)}&msgText=${encodeURIComponent(`Check out: ${shareDoc.name}`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#6264a7', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="TeamsLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Teams</span>
            </div>

            {/* WhatsApp */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://wa.me/?text=${encodeURIComponent(`${shareDoc.name}\n${encodedUrl}`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#25d366', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="Chat" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>WhatsApp</span>
            </div>

            {/* Gmail */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://mail.google.com/mail/?view=cm&su=${encodeURIComponent(`Sharing: ${shareDoc.name}`)}&body=${encodeURIComponent(`Hi,\n\nPlease find the link below:\n${shareDoc.name}\n${encodedUrl}\n\nBest regards`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#ea4335', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="Mail" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Gmail</span>
            </div>
          </div>

          {/* Copy Link */}
          <div style={{
            display: 'flex', alignItems: 'center', gap: '10px',
            border: '1px solid #edebe9', borderRadius: '6px', padding: '10px 12px'
          }}>
            <span style={{ flex: 1, fontSize: '12px', color: '#605e5c', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {encodedUrl}
            </span>
            <PrimaryButton text="Copy" onClick={() => {
              copyToClipboard(encodedUrl, shareDoc.name);
              setShowSharePanel(false);
            }} styles={{ root: { minWidth: '60px', height: '28px', padding: '0 12px' } }} />
          </div>
        </div>
      </div>
    );
  };
  // ─── END SHARE PANEL ──────────────────────────────────────────────────────

  if (isLoading) {
    return (
      <div className={styles.documentLibrary}>
        <div className={styles.mainHeader}>
          <h2 className={styles.mainTitle}>{pageTitle}</h2>
        </div>
        <div style={{
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '60px 20px',
          textAlign: 'center'
        }}>
          <Icon iconName="DocumentSet" style={{ fontSize: '48px', color: '#0078d4', marginBottom: '16px', opacity: 0.8 }} />
          <p style={{ fontSize: '16px', color: '#605e5c', margin: 0 }}>Loading documents...</p>
        </div>
      </div>
    );
  }

  if (foldersWithDocuments.length === 0) {
    return (
      <div className={styles.documentLibrary}>
        {message && (
          <MessageBar messageBarType={message.includes('❌') ? MessageBarType.error : MessageBarType.warning} isMultiline={false}>
            {message}
          </MessageBar>
        )}
        <div className={styles.mainHeader}>
          <h2 className={styles.mainTitle}>{pageTitle}</h2>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '60px 20px', textAlign: 'center' }}>
          <Icon iconName="DocumentSet" style={{ fontSize: '64px', color: '#a19f9d', marginBottom: '24px' }} />
          <h3 style={{ fontSize: '24px', fontWeight: 600, color: '#323130', margin: '0 0 16px 0' }}>No documents found</h3>
          <p style={{ fontSize: '16px', color: '#605e5c', margin: '8px 0', maxWidth: '400px', lineHeight: 1.5 }}>
            There are no documents in this folder.
          </p>
        </div>
      </div>
    );
  }

  if (currentFolder) {
    return (
      <div className={styles.documentLibrary}>
        {message && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>{message}</MessageBar>
        )}

        <div className={styles.header}>
          <div className={styles.headerContent}>
            <Icon iconName="ChevronLeft" className={styles.backIcon} onClick={handleBackClick} style={{ cursor: 'pointer', marginRight: '12px' }} />
            <h2 className={styles.mainTitle}>{pageTitle}</h2>
          </div>
        </div>

        <div className={styles.documentsSection}>
          <h3 className={styles.sectionTitle}>{currentFolder}</h3>
          <p style={{ color: '#666', marginBottom: '20px', fontSize: '16px' }}>
            Below are documents related to {currentFolder}.
          </p>

          {currentDocuments.length === 0 ? (
            <div className={styles.noDocuments}>
              <Icon iconName="DocumentSet" className={styles.noDocumentsIcon} />
              <p>No documents found in {currentFolder} folder.</p>
            </div>
          ) : (
            <>
              <div className={styles.documentsTable}>
                <div className={styles.tableHeader}>
                  <div className={styles.headerCell}>
                    <div className={styles.checkbox} onClick={handleSelectAllChange} role="checkbox" tabIndex={0} aria-checked={selectAllChecked} aria-label="Select all documents">
                      {selectAllChecked && <Icon iconName="CheckMark" className={styles.checkIcon} />}
                    </div>
                    <span>Name</span>
                    <Icon iconName="ChevronDown" className={styles.sortIcon} />
                  </div>
                  <div className={styles.headerCell}>
                    <span>Modified</span>
                    <Icon iconName="ChevronDown" className={styles.sortIcon} />
                  </div>
                  <div className={styles.headerCell}>
                    <span>Actions</span>
                  </div>
                </div>

                <div className={styles.tableBody}>
                  {currentDocuments.map((doc: IDocument) => {
                    const isChecked = checkedItems.has(String(doc.id));
                    return (
                      <div key={String(doc.id)} className={`${styles.tableRow} ${isChecked ? styles.selected : ''}`}>
                        <div className={styles.nameCell}>
                          <div className={styles.checkbox} onClick={() => handleCheckboxChange(String(doc.id))} role="checkbox" tabIndex={0} aria-checked={isChecked}>
                            {isChecked && <Icon iconName="CheckMark" className={styles.checkIcon} />}
                          </div>
                          <Icon iconName={doc.iconName} className={styles.fileIcon} />
                          <span className={styles.fileName} onClick={() => handleDocumentClick(doc)} style={{ cursor: 'pointer', color: '#000' }}>
                            {doc.name}
                          </span>
                        </div>
                        <div className={styles.dataCell}>{doc.modified}</div>
                        <div className={styles.actionsCell}>
                          <IconButton
                            iconProps={{ iconName: 'MoreVertical' }}
                            className={styles.moreButton}
                            onClick={(event: React.MouseEvent<HTMLButtonElement>) => handleDocumentActions(event, doc)}
                            ariaLabel={`More actions for ${doc.name}`}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </>
          )}
        </div>

        {contextMenuTarget && (
          <ContextualMenu items={getContextMenuItems()} target={contextMenuTarget} onDismiss={dismissContextMenu} directionalHint={6} />
        )}

        {renderSharePanel()}
      </div>
    );
  }

  return (
    <div className={styles.documentLibrary}>
      {message && (
        <MessageBar messageBarType={message.includes('Failed') || message.includes('❌') ? MessageBarType.error : MessageBarType.success} isMultiline={false}>
          {message}
        </MessageBar>
      )}

      <div className={styles.mainHeader}>
        <h2 className={styles.mainTitle}>{pageTitle}</h2>
      </div>

      <div style={{ fontSize: '16px', color: '#666', marginBottom: '30px' }}>
        <p>Below are documents related to {pageTitle}.</p>
      </div>

      <div className={styles.foldersContainer}>
        {foldersWithDocuments.map((folder: IFolderWithDocuments) => (
          <div key={folder.name} className={styles.folderSection}>
            <div className={styles.folderHeader}>
              <h3 className={styles.folderTitle}>{folder.name}</h3>
              <PrimaryButton className={styles.viewAllButton} text="View all" onClick={() => handleViewAll(folder.name)} />
            </div>

            <div className={styles.documentsGrid}>
              {folder.documents.map((doc: IDocument) => (
                <div key={String(doc.id)} className={styles.documentCard}>
                  <div className={styles.cardContent}>
                    <h4 className={styles.documentTitle} onClick={() => handleDocumentClick(doc)} style={{ cursor: 'pointer', color: '#000' }}>
                      {doc.name}
                    </h4>
                    <p className={styles.documentMeta}>Modified {doc.modified}</p>
                    <p className={styles.documentDescription}>{doc.description}</p>
                  </div>

                  <div className={styles.cardActions}>
                    <button className={styles.cardActionButton} onClick={() => handleDownloadDocument(doc)}>
                      <Icon iconName={doc.fileType === 'url' ? 'Globe' : 'Download'} className={styles.actionIcon} />
                      {doc.fileType === 'url' ? 'Open Link' : 'Export'}
                    </button>
                    <button className={styles.cardActionButton} onClick={() => handleShareDocument(doc)}>
                      <Icon iconName="Share" className={styles.actionIcon} />
                      Share
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      {renderSharePanel()}
    </div>
  );
};

export default DocumentLibrary;