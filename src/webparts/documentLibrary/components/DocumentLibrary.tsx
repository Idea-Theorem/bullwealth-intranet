import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './DocumentLibrary.module.scss';
import { IDocumentLibraryProps, IDocument } from './IDocumentLibraryProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { ContextualMenu, IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

interface IFolderWithDocuments {
  name: string;
  documents: IDocument[];
  allDocuments: IDocument[];
  folderPath: string;
}

const DocumentLibrary: React.FC<IDocumentLibraryProps> = (props) => {
  const [foldersWithDocuments, setFoldersWithDocuments] = useState<IFolderWithDocuments[]>([]);
  const [currentFolder, setCurrentFolder] = useState<string>('');
  const [currentDocuments, setCurrentDocuments] = useState<IDocument[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedDocument, setSelectedDocument] = useState<IDocument | null>(null);
  const [contextMenuTarget, setContextMenuTarget] = useState<HTMLElement | null>(null);
  const [checkedItems, setCheckedItems] = useState<Set<number>>(new Set());
  const [selectAllChecked, setSelectAllChecked] = useState<boolean>(false);
  const [message, setMessage] = useState<string>('');

  useEffect(() => {
    if (!currentFolder) {
      // eslint-disable-next-line no-void, @typescript-eslint/no-use-before-define
      void loadFolderStructure();
    }
  }, [props.listName]);

  useEffect(() => {
    if (currentFolder) {
      setCheckedItems(new Set());
      setSelectAllChecked(false);
    }
  }, [currentFolder]);

  // ✅ FIX 1: Correct date formatting that handles timezone properly
  const formatDate = (date: Date): string => {
    // Ensure we get the local date representation
    return new Intl.DateTimeFormat('en-US', {
      year: 'numeric',
      month: 'short', 
      day: 'numeric'
    }).format(date);
  };

  // ✅ FIX 3: Enhanced sorting using raw timestamps for accuracy
  const sortDocumentsByModified = (documents: IDocument[]): IDocument[] => {
    return [...documents].sort((a, b) => {
      const timeA = (a as any).modifiedTimestamp || 0;
      const timeB = (b as any).modifiedTimestamp || 0;
      return timeB - timeA; // Newest first
    });
  };

  const loadFolderStructure = async (): Promise<void> => {
    setLoading(true);
    console.log('=== LOADING FOLDER STRUCTURE ===');
    console.log('Target folder path:', props.listName);

    try {
      const baseUrl = props.context.pageContext.web.absoluteUrl;
      
      // Parse the folder path - Extract main library and target folder
      const pathParts = props.listName.split('/');
      const mainLibrary = pathParts[0];
      const targetFolder = pathParts.slice(1).join('/');
      
      console.log('Main library:', mainLibrary);
      console.log('Target folder:', targetFolder);

      // Try multiple folder path formats
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      await tryFolderPaths(baseUrl, mainLibrary, targetFolder);

    } catch (error: any) {
      console.error('❌ Error loading folder structure:', error);
      setMessage(`❌ Error loading folder "${props.listName}": ${error.message}. Using demo data.`);
      setTimeout(() => setMessage(''), 8000);
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      createDemoData();
    } finally {
      setLoading(false);
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
          // eslint-disable-next-line @typescript-eslint/no-use-before-define
          await processFolderData(data, folderPath);
          return;
        }
      } catch (error) {
        console.log(`❌ Path failed: ${folderPath}`, error);
        continue;
      }
    }

    // If all paths failed, use demo data
    console.error('❌ All folder paths failed');
    setMessage(`❌ Could not find folder "${props.listName}". Using demo data.`);
    setTimeout(() => setMessage(''), 10000);
    // eslint-disable-next-line @typescript-eslint/no-use-before-define
    createDemoData();
  };

  const processFolderData = async (data: any, folderPath: string): Promise<void> => {
    console.log('=== PROCESSING FOLDER DATA ===');
    
    const subfolders = data.Folders || [];
    const files = data.Files || [];
    
    console.log(`Found ${subfolders.length} subfolders and ${files.length} files`);

    if (subfolders.length === 0 && files.length === 0) {
      setMessage(`⚠️ Folder "${props.listName}" is empty. Using demo data.`);
      setTimeout(() => setMessage(''), 8000);
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      createDemoData();
      return;
    }

    // Process folders and their contents
    const folderGroups: { [key: string]: { documents: IDocument[], folderPath: string } } = {};

    // Process files in the root folder (if any)
    if (files.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      const mappedDocuments = files.map((file: any) => mapFileToDocument(file));
      const sortedDocuments = sortDocumentsByModified(mappedDocuments);
      
      folderGroups['Root'] = {
        documents: sortedDocuments,
        folderPath: folderPath
      };
    }

    // Process each subfolder
    // Process each subfolder
for (const subfolder of subfolders) {
  const subfolderName = subfolder.Name;
  const subfolderPath = subfolder.ServerRelativeUrl;
  
  console.log(`Processing subfolder: ${subfolderName} at ${subfolderPath}`);
  
  try {
    // ✅ METHOD 1: Try direct Description field first
    const subfolderUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/Files?$select=Name,ServerRelativeUrl,UniqueId,TimeLastModified,TimeCreated,Modified,Created,Length,Description,ModifiedBy/Title,Author/Title&$expand=ModifiedBy,Author&$nocache=${Date.now()}`;
    
    let subfolderResponse = await props.context.spHttpClient.get(
      subfolderUrl,
      SPHttpClient.configurations.v1
    );

    // ✅ METHOD 2: If direct Description doesn't work, try ListItem approach
    if (!subfolderResponse.ok) {
      console.log('Trying ListItem approach for custom columns...');
      // ✅ ENHANCED: Get files with custom columns using ListItem API
const subfolderUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(subfolderPath)}')/Files?$select=Name,ServerRelativeUrl,UniqueId,TimeLastModified,TimeCreated,ListItemAllFields/Description,ListItemAllFields/ID&$expand=ListItemAllFields`;
      
      subfolderResponse = await props.context.spHttpClient.get(
        subfolderUrl,
        SPHttpClient.configurations.v1
      );
    }

    if (subfolderResponse.ok) {
      const subfolderData = await subfolderResponse.json();
      const subfolderFiles = subfolderData.value || [];
      
      console.log(`  → Found ${subfolderFiles.length} files in ${subfolderName}`);
      
      if (subfolderFiles.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        const mappedDocuments = subfolderFiles.map((file: any) => mapFileToDocument(file));
        const sortedDocuments = sortDocumentsByModified(mappedDocuments);
        
        folderGroups[subfolderName] = {
          documents: sortedDocuments,
          folderPath: subfolderPath
        };
      }
    }
  } catch (subError) {
    console.error(`❌ Error loading subfolder ${subfolderName}:`, subError);
  }
}


    console.log('Final folder groups:', Object.keys(folderGroups));

    // Create display structure
    const foldersWithDocs: IFolderWithDocuments[] = [];
    
    for (const folderName of Object.keys(folderGroups)) {
      foldersWithDocs.push({
        name: folderName,
        documents: folderGroups[folderName].documents.slice(0, 4), // Show top 4 newest documents
        allDocuments: folderGroups[folderName].documents, // All documents (already sorted)
        folderPath: folderGroups[folderName].folderPath
      });
    }

    if (foldersWithDocs.length === 0) {
      console.log('❌ No folders with documents found');
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      createDemoData();
    } else {
      console.log(`✅ SUCCESS: Created ${foldersWithDocs.length} folder sections`);
      setFoldersWithDocuments(foldersWithDocs);
      
      const totalDocs = foldersWithDocs.reduce((sum, folder) => sum + folder.allDocuments.length, 0);
      setMessage(`✅ Successfully loaded ${totalDocs} documents from ${foldersWithDocs.length} folders in "${props.listName}"`);
      setTimeout(() => setMessage(''), 5000);
    }
  };

  // ✅ FIX 2 & 4: Enhanced mapFileToDocument with correct date handling and timestamps
  // ✅ FIXED: Map file to document with proper Description handling
const mapFileToDocument = (file: any): IDocument => {
  const fileName: string = file.Name || file.LeafName || 'Unknown';
  const fileType: string = fileName.split('.').pop() || 'file';
  
  let modifiedBy = 'System Account';
  if (file.ModifiedBy && file.ModifiedBy.Title) {
    modifiedBy = file.ModifiedBy.Title;
  } else if (file.Author && file.Author.Title) {
    modifiedBy = file.Author.Title;
  }

  let documentUrl = '#';
  if (file.ServerRelativeUrl) {
    documentUrl = `${window.location.protocol}//${window.location.host}${file.ServerRelativeUrl}`;
  }

  const modifiedDate = file.TimeLastModified ? new Date(file.TimeLastModified) : new Date();
  const createdDate = file.TimeCreated ? new Date(file.TimeCreated) : modifiedDate;

  // ✅ CRITICAL: Extract Description from the correct field
  let description = '';
  if (file.Description) {
    // Direct Description field
    description = file.Description;
  } else if (file.ListItemAllFields && file.ListItemAllFields.Description) {
    // Description from ListItem fields
    description = file.ListItemAllFields.Description;
  } else {
    // Fallback to filename without extension
    description = fileName.replace(/\.[^/.]+$/, "") || 'No description available';
  }

  console.log(`File: ${fileName}, Description: ${description}`);

  return {
    id: file.UniqueId || Math.random(),
    name: fileName.replace(/\.[^/.]+$/, ""),
    fileType: fileType,
    modified: formatDate(modifiedDate),
    modifiedBy: modifiedBy,
    serverRelativeUrl: documentUrl,
    downloadUrl: file.ServerRelativeUrl || '#',
    // eslint-disable-next-line @typescript-eslint/no-use-before-define
    iconName: getFileIcon(fileType),
    // ✅ USE: Real Description from SharePoint column
    description: description,
    createdDate: formatDate(createdDate),
    modifiedTimestamp: modifiedDate.getTime(),
    createdTimestamp: createdDate.getTime()
  };
};


  const handleDocumentClick = (doc: IDocument): void => {
    console.log('=== OPENING DOCUMENT ===');
    console.log('Document name:', doc.name);
    console.log('Document URL:', doc.serverRelativeUrl);

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

  const createDemoData = (): void => {
    console.log('=== CREATING DEMO DATA ===');
    
    const baseUrl = props.context.pageContext.web.absoluteUrl;
    
    // ✅ REALISTIC DEMO DATA with correct dates and proper timestamps
    const currentTime = new Date().getTime();
    const oneDayMs = 24 * 60 * 60 * 1000; // One day in milliseconds

    const policiesDocuments: IDocument[] = [
      {
        id: 1,
        name: 'BCMI Cross Trade Policy',
        fileType: 'pdf',
        modified: formatDate(new Date(currentTime - (6 * oneDayMs))), // ✅ 6 days ago (matches your screenshot)
        modifiedBy: 'Koteshwar Rao M',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B3550E012-FEF7-42B4-84D6-8DB73C3D4283%7D&file=BCMI%20Cross%20Trade%20Policy.pdf&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Policies/BCMI%20Cross%20Trade%20Policy.pdf`,
        iconName: 'PDF',
        description: 'BCMI Cross Trade Policy – October 18, 2023', // ✅ Real description from screenshot
        createdDate: formatDate(new Date(currentTime - (6 * oneDayMs))),
        modifiedTimestamp: currentTime - (6 * oneDayMs),
        createdTimestamp: currentTime - (6 * oneDayMs)
      },
      {
        id: 2,
        name: 'BCMI Policies & Procedures Manual',
        fileType: 'pdf',
        modified: formatDate(new Date(currentTime - (1000))), // ✅ A few seconds ago
        modifiedBy: 'Koteshwar Rao M',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B4661F123-GFG8-43C5-95E7-9EC84D4E5394%7D&file=BCMI%20Policies%20Procedures%20Manual.pdf&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Policies/BCMI%20Policies%20Procedures%20Manual.pdf`,
        iconName: 'PDF',
        description: 'BCMI Policies & Procedures Manual', // ✅ Real description from screenshot
        createdDate: formatDate(new Date(currentTime - (1000))),
        modifiedTimestamp: currentTime - (1000),
        createdTimestamp: currentTime - (1000)
      },
      {
        id: 3,
        name: 'Compliance Manual Deck Document',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (2000))), // ✅ A few seconds ago
        modifiedBy: 'Koteshwar Rao M',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B5772G234-HGH9-44D6-96F8-AED95E5F6405%7D&file=Compliance%20Manual%20Deck%20Document.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Policies/Compliance%20Manual%20Deck%20Document.docx`,
        iconName: 'WordDocument',
        description: 'Testing', // ✅ Real description from screenshot
        createdDate: formatDate(new Date(currentTime - (2000))),
        modifiedTimestamp: currentTime - (2000),
        createdTimestamp: currentTime - (2000)
      },
      {
        id: 4,
        name: 'Cross Trade Policy',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (6 * oneDayMs))), // ✅ 6 days ago (matches screenshot)
        modifiedBy: 'Koteshwar Rao M',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B6883H345-IHI0-55E7-A7G9-BFE06F6G7516%7D&file=Cross%20Trade%20Policy.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Policies/Cross%20Trade%20Policy.docx`,
        iconName: 'WordDocument',
        description: 'Cross Trade Policy Document',
        createdDate: formatDate(new Date(currentTime - (6 * oneDayMs))),
        modifiedTimestamp: currentTime - (6 * oneDayMs),
        createdTimestamp: currentTime - (6 * oneDayMs)
      }
    ];

    const processDocuments: IDocument[] = [
      {
        id: 5,
        name: 'Process Policy Updated',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (oneDayMs))), // 1 day ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B7994I456-JIJ1-66F8-B8H0-CG17G7H8627%7D&file=Process%20Policy%20Updated.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Process/Process%20Policy%20Updated.docx`,
        iconName: 'WordDocument',
        description: 'Updated Process Policy Guidelines',
        createdDate: formatDate(new Date(currentTime - (oneDayMs))),
        modifiedTimestamp: currentTime - (oneDayMs),
        createdTimestamp: currentTime - (oneDayMs)
      },
      {
        id: 6,
        name: 'Standard Operating Procedures',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (2 * oneDayMs))), // 2 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B8005J567-KJK2-77G9-C9I1-DH28H8I9738%7D&file=Standard%20Operating%20Procedures.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Process/Standard%20Operating%20Procedures.docx`,
        iconName: 'WordDocument',
        description: 'Standard Operating Procedures Manual',
        createdDate: formatDate(new Date(currentTime - (2 * oneDayMs))),
        modifiedTimestamp: currentTime - (2 * oneDayMs),
        createdTimestamp: currentTime - (2 * oneDayMs)
      },
      {
        id: 7,
        name: 'Process Guidelines',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (3 * oneDayMs))), // 3 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B9116K678-LKL3-88H0-D0J2-EI39I9J0849%7D&file=Process%20Guidelines.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Process/Process%20Guidelines.docx`,
        iconName: 'WordDocument',
        description: 'Process Guidelines Documentation',
        createdDate: formatDate(new Date(currentTime - (3 * oneDayMs))),
        modifiedTimestamp: currentTime - (3 * oneDayMs),
        createdTimestamp: currentTime - (3 * oneDayMs)
      },
      {
        id: 8,
        name: 'Process Review Checklist',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (4 * oneDayMs))), // 4 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B0227L789-MLM4-99I1-E1K3-FJ40J0K1950%7D&file=Process%20Review%20Checklist.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Process/Process%20Review%20Checklist.docx`,
        iconName: 'WordDocument',
        description: 'Process Review and Audit Checklist',
        createdDate: formatDate(new Date(currentTime - (4 * oneDayMs))),
        modifiedTimestamp: currentTime - (4 * oneDayMs),
        createdTimestamp: currentTime - (4 * oneDayMs)
      }
    ];

    const manualDocuments: IDocument[] = [
      {
        id: 9,
        name: 'Training Manual Latest',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (500))), // Few seconds ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B1338M890-NMN5-00J2-F2L4-GK51K1L2061%7D&file=Training%20Manual%20Latest.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Manual/Training%20Manual%20Latest.docx`,
        iconName: 'WordDocument',
        description: 'Latest Training Manual Version',
        createdDate: formatDate(new Date(currentTime - (500))),
        modifiedTimestamp: currentTime - (500),
        createdTimestamp: currentTime - (500)
      },
      {
        id: 10,
        name: 'User Guide Documentation',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (2 * oneDayMs))), // 2 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B2449N901-OPO6-11K3-G3M5-HL62L2M3172%7D&file=User%20Guide%20Documentation.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Manual/User%20Guide%20Documentation.docx`,
        iconName: 'WordDocument',
        description: 'Complete User Guide and Documentation',
        createdDate: formatDate(new Date(currentTime - (2 * oneDayMs))),
        modifiedTimestamp: currentTime - (2 * oneDayMs),
        createdTimestamp: currentTime - (2 * oneDayMs)
      },
      {
        id: 11,
        name: 'Reference Manual',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (5 * oneDayMs))), // 5 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B3560O012-PQP7-22L4-H4N6-IM73M3N4283%7D&file=Reference%20Manual.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Manual/Reference%20Manual.docx`,
        iconName: 'WordDocument',
        description: 'Complete Reference Manual',
        createdDate: formatDate(new Date(currentTime - (5 * oneDayMs))),
        modifiedTimestamp: currentTime - (5 * oneDayMs),
        createdTimestamp: currentTime - (5 * oneDayMs)
      },
      {
        id: 12,
        name: 'Quick Start Guide',
        fileType: 'docx',
        modified: formatDate(new Date(currentTime - (7 * oneDayMs))), // 7 days ago
        modifiedBy: 'System Account',
        serverRelativeUrl: `${baseUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B4671P123-QRQ8-33M5-I5O7-JN84N4O5394%7D&file=Quick%20Start%20Guide.docx&action=default`,
        downloadUrl: `${baseUrl}/${encodeURIComponent(props.listName)}/Manual/Quick%20Start%20Guide.docx`,
        iconName: 'WordDocument',
        description: 'Quick Start Guide for New Users',
        createdDate: formatDate(new Date(currentTime - (7 * oneDayMs))),
        modifiedTimestamp: currentTime - (7 * oneDayMs),
        createdTimestamp: currentTime - (7 * oneDayMs)
      }
    ];

    // ✅ Sort all demo data by modified timestamp (newest first)
    const sortedPolicies = sortDocumentsByModified(policiesDocuments);
    const sortedProcess = sortDocumentsByModified(processDocuments);
    const sortedManual = sortDocumentsByModified(manualDocuments);

    const demoFolders: IFolderWithDocuments[] = [
      {
        name: 'Policies',
        documents: sortedPolicies.slice(0, 4), // Show top 4 newest
        allDocuments: sortedPolicies, // All documents (already sorted)
        folderPath: `${props.listName}/Policies`
      },
      {
        name: 'Process',
        documents: sortedProcess.slice(0, 4),
        allDocuments: sortedProcess,
        folderPath: `${props.listName}/Process`
      },
      {
        name: 'Manual',
        documents: sortedManual.slice(0, 4),
        allDocuments: sortedManual,
        folderPath: `${props.listName}/Manual`
      }
    ];

    setFoldersWithDocuments(demoFolders);
    setMessage(`📋 Loaded demo data for "${props.listName}" with correct dates and descriptions (sorted by newest first)`);
    setTimeout(() => setMessage(''), 5000);
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
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        fallbackShare(doc);
      }
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
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
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        copyToClipboard(shareUrl, doc.name);
      });
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
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

  if (loading) {
    return (
      <div className={styles.documentLibrary}>
        <div className={styles.loading}>
          <Spinner label={`Loading documents from "${props.listName}"...`} />
        </div>
      </div>
    );
  }

  // Table view when "View all" is clicked
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

  // Card view - main library view
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
                      {/* ✅ FIXED: Display real description instead of placeholder */}
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
