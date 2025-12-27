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
        const folderUrl = `${baseUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')?$expand=Folders,Files`;

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

  // const getLibraryTitleFromProps = (): string => {
  //   const parts = props.listName.split('/').filter(p => p.length > 0);
  //   return parts[0] || 'Documents';
  // };

  // ✅ FINAL FIX: Extract library name from actual folder path
const getAllFolderMetadata = async (folderPath: string): Promise<Record<string, any>> => {
  const baseUrl = props.context.pageContext.web.absoluteUrl;
  
  // ✅ Extract library name from the folder path
  // Example path: /sites/MrkedCapitalIntranet/BullWealth Documents/Advisory Group/...
  const pathParts = folderPath.split('/').filter(p => p.length > 0);
  
  let libraryTitle = 'Documents';
  
  // Find the library name (usually 3rd segment after /sites/sitename/)
  if (pathParts.length >= 3) {
    // Skip 'sites' and site name, get library
    libraryTitle = pathParts[2];
  } else if (pathParts.length >= 1) {
    libraryTitle = pathParts[0];
  }

  console.log('📡 Folder path:', folderPath);
  console.log('📡 Extracted library name:', libraryTitle);

  try {
    const itemsUrl =
      `${baseUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitle)}')/items?` +
      `$select=FileRef,FileLeafRef,OrderBy,Editor/Title,Author/Title,FSObjType&` +
      `$expand=Editor,Author&` +
      `$filter=FSObjType eq 0&` +
      `$top=5000`;

    console.log('📡 Fetching from:', itemsUrl);
    
    const res = await props.context.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);
    
    if (!res.ok) {
      console.error('❌ Metadata fetch failed:', res.status, res.statusText);
      const errorText = await res.text();
      console.error('❌ Response:', errorText);
      return {};
    }

    const data = await res.json();
    const lookup: Record<string, any> = {};
    
    console.log('📥 Total items received:', data.value.length);
    
    if (data.value.length === 0) {
      console.warn('⚠️ No items found in library!');
      return {};
    }
    
    for (const item of data.value) {
      if (!item.FileRef) continue;
      
      let editorName = 'Unknown';
      if (item.Editor && item.Editor.Title) {
        editorName = item.Editor.Title;
      } else if (item.Author && item.Author.Title) {
        editorName = item.Author.Title;
      }
      
      // Clean username
      if (editorName.includes('@')) {
        editorName = editorName.split('@')[0];
      }
      if (editorName.includes('|')) {
        const parts = editorName.split('|');
        editorName = parts[parts.length - 1].split('@')[0];
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
        EditorTitle: editorName,
        FileRef: item.FileRef
      };
      
      // ✅ Store with lowercase key
      const key = item.FileRef.toLowerCase();
      lookup[key] = metadata;
      
      // Log first 5 items
      if (Object.keys(lookup).length <= 5) {
        console.log(`📄 Key: ${key}`);
        console.log(`   OrderBy: ${orderByValue}, ModifiedBy: ${editorName}`);
      }
    }
    
    console.log(`✅ Created ${Object.keys(lookup).length} lookup entries`);
    console.log(`✅ Sample keys:`, Object.keys(lookup).slice(0, 3));
    
    return lookup;
  } catch (error) {
    console.error('❌ Error fetching metadata:', error);
    return {};
  }
};


// ✅ FIXED: Match by ServerRelativeUrl
const getDocumentsWithRealUsers = async (
  files: any[],
  metadataLookup: Record<string, any>
): Promise<IDocument[]> => {
  const documents: IDocument[] = [];

  console.log(`🔍 Processing ${files.length} files with ${Object.keys(metadataLookup).length} metadata keys`);

  for (const file of files) {
    const fileName = file.Name || 'Unknown';
    const serverRelativeUrl = file.ServerRelativeUrl || '';
    
    // ✅ Match using lowercase ServerRelativeUrl (which is the FileRef)
    const lookupKey = serverRelativeUrl.toLowerCase();
    const itemMeta = metadataLookup[lookupKey];

    let modifiedBy = 'Unknown';
    let orderBy: number = 999;

    if (itemMeta) {
      modifiedBy = itemMeta.EditorTitle || 'Unknown';
      orderBy = itemMeta.OrderBy !== undefined ? itemMeta.OrderBy : 999;
      
      console.log(`✅ ${fileName}: OrderBy=${orderBy}, ModifiedBy=${modifiedBy}`);
    } else {
      console.warn(`⚠️ No metadata for: ${fileName}`);
      console.warn(`   Looking for key: ${lookupKey}`);
    }

    const fileType = fileName.split('.').pop() || 'file';
    const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
    const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

    const documentUrl = `${window.location.protocol}//${window.location.host}${serverRelativeUrl}`;
    const description = fileName.replace(/\.[^/.]+$/, '') || 'No description';

    documents.push({
      id: file.UniqueId || Math.random().toString(),
      name: fileName.replace(/\.[^/.]+$/, ''),
      fileType,
      modified: formatDate(modifiedDate),
      modifiedBy,
      serverRelativeUrl: documentUrl,
      downloadUrl: serverRelativeUrl,
      iconName: getFileIcon(fileType),
      description,
      createdDate: formatDate(createdDate),
      modifiedTimestamp: modifiedDate.getTime(),
      createdTimestamp: createdDate.getTime(),
      orderBy
    } as any);
  }

  return documents;
};




  // ✅ OPTIMIZED: Single metadata fetch for entire folder structure
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

      // ✅ PERFORMANCE: Get ALL metadata in ONE API call
      const allMetadata = await getAllFolderMetadata(folderPath);

     // const pathParts = props.listName.split('/');
      //let libraryName = pathParts[0] || 'Documents';

      const folderGroups: { [key: string]: { documents: IDocument[], folderPath: string, orderBy?: number } } = {};

      // Process root files
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

      // ✅ OPTIMIZED: Process subfolders in parallel
      const subfolderPromises = subfolders.map(async (subfolder: any) => {
        const subfolderName = subfolder.Name;
        const subfolderPath = subfolder.ServerRelativeUrl;

        try {
          //const subfolderUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/Files?$select=Name,ServerRelativeUrl,UniqueId,TimeLastModified,TimeCreated`;
          const subfolderUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/Files?$select=Name,ServerRelativeUrl,UniqueId,TimeLastModified,TimeCreated,FileLeafRef`;

          const subfolderResponse = await props.context.spHttpClient.get(
            subfolderUrl,
            SPHttpClient.configurations.v1
          );

          if (subfolderResponse.ok) {
            const subfolderData = await subfolderResponse.json();
            const subfolderFiles = subfolderData.value || [];

            if (subfolderFiles.length > 0) {
              // ✅ Use pre-fetched metadata (already includes subfolders)
              const mappedDocuments = await getDocumentsWithRealUsers(subfolderFiles, allMetadata);
              const sortedDocuments = sortDocumentsByOrderOnly(mappedDocuments);

              // Get folder OrderBy from metadata
              let folderOrderBy: number | undefined = undefined;
              const folderKey = subfolderPath.toLowerCase();
              if (allMetadata[folderKey] && allMetadata[folderKey].OrderBy !== undefined) {
                folderOrderBy = parseInt(String(allMetadata[folderKey].OrderBy), 10);
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
          console.error(`Error processing subfolder ${subfolderName}:`, error);
        }
        return null;
      });

      // ✅ Wait for all subfolders in parallel
      const resolvedSubfolders = await Promise.all(subfolderPromises);

      // Add valid subfolders to groups
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
        foldersWithDocs.push({
          name: folderName,
          documents: sortedDocs.slice(0, 4),
          allDocuments: sortedDocs,
          folderPath: folderGroups[folderName].folderPath,
          orderBy: folderGroups[folderName].orderBy
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
      console.error('Error in processFolderData:', error);
      setMessage(`❌ Error processing folder data: ${error.message || 'Unknown error'}`);
      setTimeout(() => setMessage(''), 8000);
      setFoldersWithDocuments([]);
    }
  };

  // const mapFileToDocument = (file: any): IDocument => {
  //   const fileName: string = file.Name || file.LeafRef || 'Unknown';
  //   const fileType: string = fileName.split('.').pop() || 'file';

  //   let modifiedBy = 'Unknown';
  //   let orderBy: number = 999;

  //   if (file.ModifiedBy && file.ModifiedBy.Title) {
  //     modifiedBy = file.ModifiedBy.Title.trim();
  //   } else if (file.Author && file.Author.Title) {
  //     modifiedBy = file.Author.Title.trim();
  //   }

  //   if (modifiedBy !== 'Unknown' && modifiedBy.includes('@')) {
  //     modifiedBy = modifiedBy.split('@')[0];
  //   }

  //   if (file.OrderBy !== null && file.OrderBy !== undefined) {
  //     const parsed = parseInt(file.OrderBy, 10);
  //     if (!isNaN(parsed)) {
  //       orderBy = parsed;
  //     }
  //   }

  //   let documentUrl = '#';
  //   if (file.ServerRelativeUrl) {
  //     documentUrl = `${window.location.protocol}//${window.location.host}${file.ServerRelativeUrl}`;
  //   }

  //   const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
  //   const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

  //   let description = fileName.replace(/\.[^/.]+$/, "") || 'No description available';

  //   return {
  //     id: file.UniqueId || Math.random().toString(),
  //     name: fileName.replace(/\.[^/.]+$/, ""),
  //     fileType: fileType,
  //     modified: formatDate(modifiedDate),
  //     modifiedBy: modifiedBy,
  //     serverRelativeUrl: documentUrl,
  //     downloadUrl: file.ServerRelativeUrl || '#',
  //     iconName: getFileIcon(fileType),
  //     description: description,
  //     createdDate: formatDate(createdDate),
  //     modifiedTimestamp: modifiedDate.getTime(),
  //     createdTimestamp: createdDate.getTime(),
  //     orderBy: orderBy
  //   } as any;
  // };

  const handleDocumentClick = (doc: IDocument): void => {
    if (!doc.serverRelativeUrl || doc.serverRelativeUrl === '#') {
      setMessage('❌ Document URL not available');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    try {
      window.open(doc.serverRelativeUrl, '_blank');
    } catch {
      setMessage('❌ Unable to open document. Please check your permissions.');
      setTimeout(() => setMessage(''), 5000);
    }
  };

  const getFileIcon = (fileType: string): string => {
    const type = (fileType || '').toLowerCase();
    switch (type) {
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

  const handleDownloadDocument = (doc: IDocument): void => {
    if (doc.downloadUrl && doc.downloadUrl !== '#') {
      const downloadUrl = doc.downloadUrl.startsWith('http') ? doc.downloadUrl : `${window.location.protocol}//${window.location.host}${doc.downloadUrl}`;

      const link = window.document.createElement('a');
      link.href = downloadUrl;
      link.download = `${doc.name}.${doc.fileType}`;
      link.style.display = 'none';

      window.document.body.appendChild(link);
      link.click();
      window.document.body.removeChild(link);

      setMessage(`Downloading ${doc.name}...`);
      setTimeout(() => setMessage(''), 3000);
    } else {
      setMessage('Download not available for this document');
      setTimeout(() => setMessage(''), 3000);
    }
  };

  const handleShareDocument = (doc: IDocument): void => {
    if ((window as any).SP && (window as any).SP.UI && (window as any).SP.UI.ModalDialog) {
      try {
        const shareUrl = doc.serverRelativeUrl || window.location.href;
        const options = {
          url: `${props.context.pageContext.web.absoluteUrl}/_layouts/15/sharedialog.aspx?obj=${encodeURIComponent(shareUrl)}&ma=0`,
          title: 'Share Document',
          allowMaximize: false,
          showClose: true,
          width: 600,
          height: 650
        };
        (window as any).SP.UI.ModalDialog.showModalDialog(options);
        setMessage(`Opening SharePoint sharing for "${doc.name}"`);
        setTimeout(() => setMessage(''), 3000);
      } catch {
        fallbackShare(doc);
      }
    } else {
      fallbackShare(doc);
    }
  };

  const fallbackShare = (doc: IDocument): void => {
    const shareUrl = doc.serverRelativeUrl || window.location.href;
    const shareData = {
      title: doc.name,
      text: `Check out this document: ${doc.name}`,
      url: shareUrl
    };

    if (navigator.share) {
      navigator.share(shareData).catch(() => {
        copyToClipboard(shareUrl, doc.name);
      });
    } else {
      copyToClipboard(shareUrl, doc.name);
    }
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
        text: 'Export',
        iconProps: { iconName: 'Download' },
        onClick: () => {
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
          <Icon
            iconName="DocumentSet"
            style={{
              fontSize: '48px',
              color: '#0078d4',
              marginBottom: '16px',
              opacity: 0.8
            }}
          />
          <p style={{ fontSize: '16px', color: '#605e5c', margin: 0 }}>
            Loading documents...
          </p>
        </div>
      </div>
    );
  }

  if (foldersWithDocuments.length === 0) {
    return (
      <div className={styles.documentLibrary}>
        {message && (
          <MessageBar
            messageBarType={message.includes('❌') ? MessageBarType.error : MessageBarType.warning}
            isMultiline={false}
          >
            {message}
          </MessageBar>
        )}

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
          <Icon
            iconName="DocumentSet"
            style={{
              fontSize: '64px',
              color: '#a19f9d',
              marginBottom: '24px'
            }}
          />
          <h3 style={{
            fontSize: '24px',
            fontWeight: 600,
            color: '#323130',
            margin: '0 0 16px 0'
          }}>
            No documents found
          </h3>
          <p style={{
            fontSize: '16px',
            color: '#605e5c',
            margin: '8px 0',
            maxWidth: '400px',
            lineHeight: 1.5
          }}>
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
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            {message}
          </MessageBar>
        )}

        <div className={styles.header}>
          <div className={styles.headerContent}>
            <Icon
              iconName="ChevronLeft"
              className={styles.backIcon}
              onClick={handleBackClick}
              style={{ cursor: 'pointer', marginRight: '12px' }}
            />
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
                    <div
                      className={styles.checkbox}
                      onClick={handleSelectAllChange}
                      role="checkbox"
                      tabIndex={0}
                      aria-checked={selectAllChecked}
                      aria-label="Select all documents"
                    >
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
                    <span>Modified By</span>
                    <Icon iconName="ChevronDown" className={styles.sortIcon} />
                  </div>
                  <div className={styles.headerCell}>
                    <span>Actions</span>
                  </div>
                </div>

                <div className={styles.tableBody}>
                  {sortDocumentsByOrderOnly(currentDocuments).map((doc: IDocument) => {
                    const isChecked = checkedItems.has(String(doc.id));
                    return (
                      <div
                        key={String(doc.id)}
                        className={`${styles.tableRow} ${isChecked ? styles.selected : ''}`}
                      >
                        <div className={styles.nameCell}>
                          <div
                            className={styles.checkbox}
                            onClick={() => handleCheckboxChange(String(doc.id))}
                            role="checkbox"
                            tabIndex={0}
                            aria-checked={isChecked}
                          >
                            {isChecked && <Icon iconName="CheckMark" className={styles.checkIcon} />}
                          </div>
                          <Icon iconName={doc.iconName} className={styles.fileIcon} />
                          <span
                            className={styles.fileName}
                            onClick={() => handleDocumentClick(doc)}
                            style={{ cursor: 'pointer', color: '#000' }}
                          >
                            {doc.name}
                          </span>
                        </div>
                        <div className={styles.dataCell}>
                          {doc.modified}
                        </div>
                        <div className={styles.dataCell}>
                          {doc.modifiedBy}
                        </div>
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
          <ContextualMenu
            items={getContextMenuItems()}
            target={contextMenuTarget}
            onDismiss={dismissContextMenu}
            directionalHint={6}
          />
        )}
      </div>
    );
  }

  return (
    <div className={styles.documentLibrary}>
      {message && (
        <MessageBar
          messageBarType={message.includes('Failed') || message.includes('❌') ? MessageBarType.error : MessageBarType.success}
          isMultiline={false}
        >
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
              <PrimaryButton
                className={styles.viewAllButton}
                text="View all"
                onClick={() => handleViewAll(folder.name)}
              />
            </div>

            <div className={styles.documentsGrid}>
              {folder.documents
                .sort((a, b) => {
                  const aOrder = (a as any).orderBy !== undefined ? (a as any).orderBy : 999;
                  const bOrder = (b as any).orderBy !== undefined ? (b as any).orderBy : 999;
                  return aOrder - bOrder;
                })
                .map((doc: IDocument) => (
                  <div key={String(doc.id)} className={styles.documentCard}>
                    <div className={styles.cardContent}>
                      <h4
                        className={styles.documentTitle}
                        onClick={() => handleDocumentClick(doc)}
                        style={{ cursor: 'pointer', color: '#000' }}
                      >
                        {doc.name}
                      </h4>
                      <p className={styles.documentMeta}>
                        Modified {doc.modified}
                      </p>
                      <p className={styles.documentDescription}>
                        {doc.description}
                      </p>
                    </div>

                    <div className={styles.cardActions}>
                      <button
                        className={styles.cardActionButton}
                        onClick={() => handleDownloadDocument(doc)}
                      >
                        <Icon iconName="Download" className={styles.actionIcon} />
                        Export
                      </button>
                      <button
                        className={styles.cardActionButton}
                        onClick={() => handleShareDocument(doc)}
                      >
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
    </div>
  );
};

export default DocumentLibrary;
