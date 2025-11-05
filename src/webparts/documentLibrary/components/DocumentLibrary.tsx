/* eslint-disable @typescript-eslint/no-use-before-define */
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
  const [checkedItems, setCheckedItems] = useState<Set<number>>(new Set());
  const [selectAllChecked, setSelectAllChecked] = useState<boolean>(false);
  const [message, setMessage] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);


  useEffect(() => {
    if (!currentFolder) {
      void loadFolderStructure();
    }
  }, [props.listName]);


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


  // ✅ Extract Order By from filename
  const extractOrderFromFilename = (fileName: string): number | undefined => {
    if (fileName.includes('Appendix F')) return 1;
    if (fileName.includes('Appendix C')) return 2;
    if (fileName.includes('Sick Leave')) return 3;
    if (fileName.includes('Vacation')) return 4;
    if (fileName.match(/\d{8}\.pdf$/)) return 5;
    return undefined;
  };


  // ✅ Sort by Order By column ONLY
  const sortDocumentsByModified = (documents: IDocument[]): IDocument[] => {
    return [...documents].sort((a, b) => {
      const aOrder = (a as any).orderBy;
      const bOrder = (b as any).orderBy;
      
      if (aOrder !== undefined && aOrder !== null) {
        if (bOrder !== undefined && bOrder !== null) {
          return Number(aOrder) - Number(bOrder);
        }
        return -1;
      }
      if (bOrder !== undefined && bOrder !== null) {
        return 1;
      }
      
      return 0;
    });
  };


  const loadFolderStructure = async (): Promise<void> => {
    setIsLoading(true);
    console.log('=== LOADING FOLDER STRUCTURE ===');
    console.log('Target folder path:', props.listName);

    try {
      const baseUrl = props.context.pageContext.web.absoluteUrl;
      const pathParts = props.listName.split('/');
      const mainLibrary = pathParts[0];
      const targetFolder = pathParts.slice(1).join('/');
      
      console.log('Main library:', mainLibrary);
      console.log('Target folder:', targetFolder);

      await tryFolderPaths(baseUrl, mainLibrary, targetFolder);

    } catch (error: any) {
      console.error('❌ Error loading folder structure:', error);
      setMessage(`❌ Error loading folder "${props.listName}": ${error.message}`);
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
        console.log(`Trying folder path: ${folderPath}`);
        
        const folderUrl = `${baseUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')?$expand=Folders,Files`;
        
        const response = await props.context.spHttpClient.get(
          folderUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          console.log(`✅ Success with path: ${folderPath}`);
          await processFolderData(data, folderPath);
          return;
        }
      } catch (error) {
        console.log(`❌ Path failed: ${folderPath}`, error);
        continue;
      }
    }

    console.error('❌ All folder paths failed');
    setMessage(`⚠️ Could not find folder "${props.listName}"`);
    setTimeout(() => setMessage(''), 10000);
    setFoldersWithDocuments([]);
  };


  const getDocumentsWithRealUsers = async (files: any[], folderPath: string): Promise<IDocument[]> => {
    const documents: IDocument[] = [];
    
    try {
      const pathParts = props.listName.split('/');
      let libraryName = pathParts[0] || 'Documents';
      
      if (libraryName.toLowerCase() === 'documents') {
        libraryName = 'Documents';
      }
      
      console.log(`🔍 Querying library: "${libraryName}" for ${files.length} files`);
      
      const listItemsUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/items?$select=Id,Title,FileLeafRef,FileRef,File_x0020_Type,Modified,Created,Author/Title,Author/Name,Editor/Title,Editor/Name,FileDirRef,EncodedAbsUrl&$expand=Author,Editor&$filter=FSObjType eq 0&$top=2000&$nocache=${Date.now()}`;
      
      const listResponse = await props.context.spHttpClient.get(listItemsUrl, SPHttpClient.configurations.v1);
      
      if (listResponse.ok) {
        const listData = await listResponse.json();
        const listItems = listData.value || [];
        
        console.log(`📋 Found ${listItems.length} list items in library`);
        
        for (const file of files) {
          const fileName = file.Name || file.LeafName || 'Unknown';
          
          const matchingListItem = listItems.find((item: any) => 
            item.FileLeafRef === fileName ||
            item.FileRef === file.ServerRelativeUrl ||
            (item.EncodedAbsUrl && item.EncodedAbsUrl.includes(fileName))
          );
          
          let modifiedBy = 'System Account';
          
          if (matchingListItem) {
            if (matchingListItem.Editor && matchingListItem.Editor.Title) {
              modifiedBy = matchingListItem.Editor.Title;
              console.log(`✅ Found REAL user for ${fileName}: ${modifiedBy}`);
            } else if (matchingListItem.Author && matchingListItem.Author.Title) {
              modifiedBy = matchingListItem.Author.Title;
              console.log(`✅ Found REAL user (Author) for ${fileName}: ${modifiedBy}`);
            }
            
            if (modifiedBy !== 'System Account') {
              modifiedBy = modifiedBy.replace(/@.*$/, '');
              
              if (modifiedBy.toLowerCase().includes('system') || 
                  modifiedBy.toLowerCase().includes('sharepoint') ||
                  modifiedBy === '' ||
                  modifiedBy.startsWith('i:0#') ||
                  modifiedBy.includes('|membership|')) {
                modifiedBy = 'System Account';
              }
            }
          } else {
            console.log(`⚠️ No list item found for: ${fileName}`);
            
            if (file.Editor && file.Editor.Title) {
              modifiedBy = file.Editor.Title.replace(/@.*$/, '');
            } else if (file.ModifiedBy && file.ModifiedBy.Title) {
              modifiedBy = file.ModifiedBy.Title.replace(/@.*$/, '');
            } else if (file.Author && file.Author.Title) {
              modifiedBy = file.Author.Title.replace(/@.*$/, '');
            }
          }

          const fileType = fileName.split('.').pop() || 'file';
          const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
          const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

          let documentUrl = '#';
          if (file.ServerRelativeUrl) {
            documentUrl = `${window.location.protocol}//${window.location.host}${file.ServerRelativeUrl}`;
          }

          let description = '';
          if (matchingListItem && matchingListItem.Title && matchingListItem.Title !== fileName) {
            description = matchingListItem.Title;
          } else if (file.Description) {
            description = file.Description;
          } else {
            description = fileName.replace(/\.[^/.]+$/, "") || 'No description available';
          }

          console.log(`📄 Final: ${fileName} -> Modified by: ${modifiedBy}`);

          documents.push({
            id: file.UniqueId || Math.random(),
            name: fileName.replace(/\.[^/.]+$/, ""),
            fileType: fileType,
            modified: formatDate(modifiedDate),
            modifiedBy: modifiedBy,
            serverRelativeUrl: documentUrl,
            downloadUrl: file.ServerRelativeUrl || '#',
            iconName: getFileIcon(fileType),
            description: description,
            createdDate: formatDate(createdDate),
            modifiedTimestamp: modifiedDate.getTime(),
            createdTimestamp: createdDate.getTime(),
            orderBy: extractOrderFromFilename(fileName)  // ✅ ADD ORDER BY
          });
        }
      } else {
        console.error(`❌ Failed to get list items. Status: ${listResponse.status}`);
        return files.map((file: any) => mapFileToDocument(file));
      }
    } catch (error) {
      console.error('❌ Error in getDocumentsWithRealUsers:', error);
      return files.map((file: any) => mapFileToDocument(file));
    }
    
    return documents;
  };


  const processFolderData = async (data: any, folderPath: string): Promise<void> => {
    console.log('=== PROCESSING FOLDER DATA ===');
    
    const subfolders = data.Folders || [];
    const files = data.Files || [];
    
    console.log(`Found ${subfolders.length} subfolders and ${files.length} files`);

    if (subfolders.length === 0 && files.length === 0) {
      setMessage(`⚠️ Folder "${props.listName}" is empty`);
      setTimeout(() => setMessage(''), 8000);
      setFoldersWithDocuments([]);
      return;
    }

    const folderGroups: { [key: string]: { documents: IDocument[], folderPath: string, orderBy?: number } } = {};

    if (files.length > 0) {
      const mappedDocuments = await getDocumentsWithRealUsers(files, folderPath);
      const sortedDocuments = sortDocumentsByModified(mappedDocuments);
      
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
      
      console.log(`✅ Using display name: "${displayName}" for folder with ${files.length} files`);
      
      folderGroups[displayName] = {
        documents: sortedDocuments,
        folderPath: folderPath
      };
    }

    for (const subfolder of subfolders) {
      const subfolderName = subfolder.Name;
      const subfolderPath = subfolder.ServerRelativeUrl;
      
      console.log(`Processing subfolder: ${subfolderName} at ${subfolderPath}`);
      
      try {
        const subfolderUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/Files?$select=Name,ServerRelativeUrl,UniqueId,TimeLastModified,TimeCreated&$nocache=${Date.now()}`;
        
        const subfolderResponse = await props.context.spHttpClient.get(
          subfolderUrl,
          SPHttpClient.configurations.v1
        );

        if (subfolderResponse.ok) {
          const subfolderData = await subfolderResponse.json();
          const subfolderFiles = subfolderData.value || [];
          
          console.log(`  → Found ${subfolderFiles.length} files in ${subfolderName}`);
          
          if (subfolderFiles.length > 0) {
            const mappedDocuments = await getDocumentsWithRealUsers(subfolderFiles, subfolderPath);
            const sortedDocuments = sortDocumentsByModified(mappedDocuments);
            
            let folderOrderBy: number | undefined = undefined;
            
            try {
              const folderPropsUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/ListItemAllFields`;
              const folderPropsResponse = await props.context.spHttpClient.get(folderPropsUrl, SPHttpClient.configurations.v1);
              
              if (folderPropsResponse.ok) {
                const folderProps = await folderPropsResponse.json();
                console.log(`📊 Folder props for ${subfolderName}:`, folderProps);
                
                if (folderProps.Order_x0020_By !== null && folderProps.Order_x0020_By !== undefined) {
                  folderOrderBy = parseInt(folderProps.Order_x0020_By, 10);
                } else if (folderProps.OrderBy !== null && folderProps.OrderBy !== undefined) {
                  folderOrderBy = parseInt(folderProps.OrderBy, 10);
                } else if (folderProps.Order !== null && folderProps.Order !== undefined) {
                  folderOrderBy = parseInt(folderProps.Order, 10);
                }
                
                console.log(`📂 Folder "${subfolderName}" Order By: ${folderOrderBy}`);
              } else {
                console.log(`⚠️ Could not get ListItemAllFields for folder ${subfolderName}, status: ${folderPropsResponse.status}`);
              }
            } catch (err) {
              console.error(`❌ Error getting Order By for folder ${subfolderName}:`, err);
            }
            
            if (folderOrderBy === undefined) {
              console.log(`⚠️ Using fallback order for ${subfolderName}`);
              const orderMap: { [key: string]: number } = {
                'Policies': 1,
                'Templates': 2,
                'Training Material': 3,
                'Operational Procedures': 4
              };
              folderOrderBy = orderMap[subfolderName];
            }
            
            folderGroups[subfolderName] = {
              documents: sortedDocuments,
              folderPath: subfolderPath,
              orderBy: folderOrderBy
            };
          }
        }
      } catch (subError) {
        console.error(`❌ Error loading subfolder ${subfolderName}:`, subError);
      }
    }

    console.log('Final folder groups:', Object.keys(folderGroups));

    const foldersWithDocs: IFolderWithDocuments[] = [];

    for (const folderName of Object.keys(folderGroups)) {
      foldersWithDocs.push({
        name: folderName,
        documents: folderGroups[folderName].documents.slice(0, 4),
        allDocuments: folderGroups[folderName].documents,
        folderPath: folderGroups[folderName].folderPath,
        orderBy: folderGroups[folderName].orderBy
      });
    }

    foldersWithDocs.sort((a, b) => {
      if (a.orderBy !== undefined && b.orderBy !== undefined) {
        return a.orderBy - b.orderBy;
      }
      if (a.orderBy !== undefined) return -1;
      if (b.orderBy !== undefined) return 1;
      
      return a.name.localeCompare(b.name);
    });

    console.log('📂 Folder display order:', foldersWithDocs.map(f => `${f.name} (Order: ${f.orderBy})`));

    if (foldersWithDocs.length === 0) {
      console.log('❌ No folders with documents found');
      setFoldersWithDocuments([]);
      setMessage(`⚠️ No documents found in "${props.listName}"`);
      setTimeout(() => setMessage(''), 8000);
    } else {
      console.log(`✅ SUCCESS: Created ${foldersWithDocs.length} folder sections`);
      setFoldersWithDocuments(foldersWithDocs);
      const totalDocs = foldersWithDocs.reduce((sum, folder) => sum + folder.allDocuments.length, 0);
      setMessage(`✅ Loaded ${totalDocs} documents from ${foldersWithDocs.length} folders`);
      setTimeout(() => setMessage(''), 5000);
    }
  };


  const mapFileToDocument = (file: any): IDocument => {
    const fileName: string = file.Name || file.LeafName || 'Unknown';
    const fileType: string = fileName.split('.').pop() || 'file';
    
    let modifiedBy = 'System Account';
    if (file.ModifiedBy && file.ModifiedBy.Title) {
      modifiedBy = file.ModifiedBy.Title.replace(/@.*$/, '');
    } else if (file.Author && file.Author.Title) {
      modifiedBy = file.Author.Title.replace(/@.*$/, '');
    }

    let documentUrl = '#';
    if (file.ServerRelativeUrl) {
      documentUrl = `${window.location.protocol}//${window.location.host}${file.ServerRelativeUrl}`;
    }

    const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
    const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

    let description = '';
    if (file.Description) {
      description = file.Description;
    } else if (file.ListItemAllFields && file.ListItemAllFields.Description) {
      description = file.ListItemAllFields.Description;
    } else {
      description = fileName.replace(/\.[^/.]+$/, "") || 'No description available';
    }

    return {
      id: file.UniqueId || Math.random(),
      name: fileName.replace(/\.[^/.]+$/, ""),
      fileType: fileType,
      modified: formatDate(modifiedDate),
      modifiedBy: modifiedBy,
      serverRelativeUrl: documentUrl,
      downloadUrl: file.ServerRelativeUrl || '#',
      iconName: getFileIcon(fileType),
      description: description,
      createdDate: formatDate(createdDate),
      modifiedTimestamp: modifiedDate.getTime(),
      createdTimestamp: createdDate.getTime(),
      orderBy: extractOrderFromFilename(fileName)  // ✅ ADD ORDER BY
    };
  };


  const handleDocumentClick = (doc: IDocument): void => {
    if (!doc.serverRelativeUrl || doc.serverRelativeUrl === '#') {
      setMessage('❌ Document URL not available');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    try {
      console.log('Opening document URL:', doc.serverRelativeUrl);
      window.open(doc.serverRelativeUrl, '_blank');
    } catch (error) {
      console.error('Failed to open document:', error);
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
      } catch (error) {
        console.error('SharePoint sharing failed:', error);
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
      navigator.share(shareData).catch((error) => {
        console.error('Web Share API failed:', error);
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
      }).catch((err) => {
        console.error('Failed to copy link:', err);
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
      } catch (err) {
        console.error('Fallback copy failed:', err);
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
      setCheckedItems(new Set(currentDocuments.map((doc: IDocument) => doc.id)));
    } else {
      setCheckedItems(new Set());
    }
  };


  const handleCheckboxChange = (documentId: number): void => {
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
          <h2 className={styles.mainTitle}>{props.title}</h2>
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
            Loading documents from "{props.listName}"...
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
          <h2 className={styles.mainTitle}>{props.title}</h2>
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
            There are no documents in "{props.listName}" folder.
          </p>
          <p style={{
            fontSize: '16px',
            color: '#605e5c',
            margin: '8px 0',
            maxWidth: '400px',
            lineHeight: 1.5
          }}>
            Documents will appear here once they are uploaded to the SharePoint library.
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
            />
            <h2 className={styles.mainTitle}>{props.title}</h2>
          </div>
        </div>

        <div className={styles.documentsSection}>
          <h3 className={styles.sectionTitle}>{currentFolder}</h3>
          
          {currentDocuments.length === 0 ? (
            <div className={styles.noDocuments}>
              <Icon iconName="DocumentSet" className={styles.noDocumentsIcon} />
              <p>No documents found in {currentFolder} folder.</p>
            </div>
          ) : (
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
                {currentDocuments.map((doc: IDocument) => {
                  const isChecked = checkedItems.has(doc.id);
                  return (
                    <div 
                      key={doc.id} 
                      className={`${styles.tableRow} ${isChecked ? styles.selected : ''}`}
                    >
                      <div className={styles.nameCell}>
                        <div 
                          className={styles.checkbox}
                          onClick={() => handleCheckboxChange(doc.id)}
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
        <h2 className={styles.mainTitle}>{props.title}</h2>
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
              {folder.documents.map((doc: IDocument) => (
                <div key={doc.id} className={styles.documentCard}>
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
