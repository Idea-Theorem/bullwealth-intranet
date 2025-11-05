/* eslint-disable no-prototype-builtins */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps, IDocumentCategory } from './IDocumentsProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';

export interface IDocumentsState {
  categories: IDocumentCategory[];
  editMode: boolean;
  uploadingFor?: string;
  loading: boolean;
}

export default class Documents extends React.Component<IDocumentsProps, IDocumentsState> {
  constructor(props: IDocumentsProps) {
    super(props);
    this.state = {
      categories: props.categories,
      editMode: false,
      uploadingFor: undefined,
      loading: props.isDynamicMode
    };
  }

  public componentDidMount(): void {
    if (this.props.isDynamicMode) {
      this.fetchFoldersFromLibrary();
    }
  }

  public componentDidUpdate(prevProps: IDocumentsProps): void {
    if (JSON.stringify(prevProps.categories) !== JSON.stringify(this.props.categories)) {
      this.setState({ categories: this.props.categories });
    }
    
    if (prevProps.isDynamicMode !== this.props.isDynamicMode ||
        prevProps.folderPath !== this.props.folderPath ||
        prevProps.documentLibraryName !== this.props.documentLibraryName) {
      if (this.props.isDynamicMode) {
        this.fetchFoldersFromLibrary();
      }
    }
  }

  private fetchFoldersFromLibrary = async (): Promise<void> => {
    this.setState({ loading: true });
    
    try {
      const { context, documentLibraryName, folderPath, sitePageBasePath } = this.props;
      const siteUrl = context.pageContext.web.absoluteUrl;
      const sitePath = context.pageContext.web.serverRelativeUrl;
      
      const cleanLibraryName = documentLibraryName.trim();
      const cleanFolderPath = folderPath.trim();
      const fullPath = `${sitePath}/${cleanLibraryName}/${cleanFolderPath}`.replace(/\/+/g, '/');
      
      console.log('=== FOLDER FETCH DEBUG ===');
      console.log('Site URL:', siteUrl);
      console.log('Site Path:', sitePath);
      console.log('Full Path:', fullPath);
      
      // ✅ STEP 1: Get all folders
      const foldersApiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@path)/Folders?@path='${encodeURIComponent(fullPath)}'&$select=Name,ServerRelativeUrl,ItemCount`;
      
      console.log('Folders API URL:', foldersApiUrl);
      
      const foldersResponse: SPHttpClientResponse = await context.spHttpClient.get(
        foldersApiUrl,
        SPHttpClient.configurations.v1
      );

      console.log('Folders Response status:', foldersResponse.status);

      if (!foldersResponse.ok) {
        const errorText = await foldersResponse.text();
        console.error('❌ API Error - Status:', foldersResponse.status);
        console.error('Error response:', errorText);
        alert(`Failed to load folders.\nStatus: ${foldersResponse.status}\n\nCheck browser console for details.`);
        this.setState({ categories: [], loading: false });
        return;
      }

      const foldersData = await foldersResponse.json();
      console.log('Folders data:', foldersData);
      
      if (!foldersData.value || !Array.isArray(foldersData.value) || foldersData.value.length === 0) {
        console.warn('⚠️ No folders found');
        this.setState({ categories: [], loading: false });
        return;
      }

      console.log(`Found ${foldersData.value.length} folders:`, foldersData.value.map((f: any) => f.Name));
      
      // ✅ STEP 2: Get Order By metadata for ALL folders - ENHANCED
const parentFolderPath = fullPath.replace(/\/$/, ''); // Remove trailing slash
console.log('📋 Parent folder path for matching:', parentFolderPath);

// Get ALL folders first (no filter), then match in code
const listItemsApiUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cleanLibraryName)}')/items?$select=FileLeafRef,FileDirRef,Order_x0020_By,FSObjType&$filter=FSObjType eq 1&$top=5000`;

console.log('📋 List Items API URL:', listItemsApiUrl);

const listItemsResponse = await context.spHttpClient.get(
  listItemsApiUrl,
  SPHttpClient.configurations.v1
);

let folderOrderMap: { [key: string]: number } = {};

if (listItemsResponse.ok) {
  const listItemsData = await listItemsResponse.json();
  console.log(`📋 Total folders in library: ${listItemsData.value?.length || 0}`);
  
  if (listItemsData.value && Array.isArray(listItemsData.value)) {
    // Log first 3 items to see what paths look like
    console.log('📁 Sample folder items:', 
      listItemsData.value.slice(0, 3).map((item: any) => ({
        name: item.FileLeafRef,
        path: item.FileDirRef,
        orderBy: item.Order_x0020_By
      }))
    );
    
    // Match folders by checking if their parent path matches
    listItemsData.value.forEach((item: any) => {
      const folderName = item.FileLeafRef;
      const folderPath = item.FileDirRef;
      
      // Try multiple matching strategies
      const isMatch = 
        folderPath === parentFolderPath ||                    // Exact match
        folderPath.toLowerCase() === parentFolderPath.toLowerCase() ||  // Case insensitive
        folderPath.endsWith('/' + cleanFolderPath) ||         // Ends with target folder
        folderPath.endsWith(cleanFolderPath);                 // Without slash
      
      if (isMatch) {
        const orderBy = item.Order_x0020_By;
        if (orderBy !== null && orderBy !== undefined) {
          folderOrderMap[folderName] = parseInt(orderBy, 10);
          console.log(`✅ MATCHED: "${folderName}" at "${folderPath}" -> Order By: ${orderBy}`);
        } else {
          console.log(`⚠️ MATCHED but no Order By: "${folderName}" at "${folderPath}"`);
        }
      }
    });
    
    console.log('📊 Final folder order map:', folderOrderMap);
    
    if (Object.keys(folderOrderMap).length === 0) {
      console.warn('⚠️ WARNING: No folders matched! Expected path:', parentFolderPath);
      console.warn('⚠️ Sample actual paths:', 
        listItemsData.value.slice(0, 3).map((i: any) => i.FileDirRef)
      );
    }
  }
} else {
  const errorText = await listItemsResponse.text();
  console.error('❌ List items API failed:', listItemsResponse.status);
  console.error('Error:', errorText);
}


      // ✅ STEP 3: Create categories with Order By values
      const dynamicCategories: IDocumentCategory[] = await Promise.all(
        foldersData.value.map(async (folder: any) => {
          const folderName = folder.Name;
          const folderServerRelativeUrl = folder.ServerRelativeUrl;
          
          let orderBy: number | undefined = undefined;
          if (folderOrderMap.hasOwnProperty(folderName)) {
            const mappedValue = folderOrderMap[folderName];
            orderBy = mappedValue >= 0 ? mappedValue : undefined;
          }
          
          console.log(`📂 Processing folder: "${folderName}", Order By: ${orderBy !== undefined ? orderBy : 'none'}, Path: ${folderServerRelativeUrl}`);
          
          const imageUrl = await this.checkForCoverImage(folderServerRelativeUrl);
          
          const cleanFolderName = folderName
            .replace(/\s+/g, '-')
            .replace(/&/g, 'and')
            .replace(/[^a-zA-Z0-9-]/g, '')
            .replace(/--+/g, '-')
            .replace(/^-|-$/g, '');

          const normalizedBasePath = sitePageBasePath.startsWith('/') 
            ? sitePageBasePath 
            : `/${sitePageBasePath}`;
          
          const pageUrl = normalizedBasePath.startsWith('/sites/') 
            ? `${siteUrl.split('/sites/')[0]}${normalizedBasePath}${cleanFolderName}.aspx`
            : `${siteUrl}${normalizedBasePath}${cleanFolderName}.aspx`;

          return {
            id: folderName,
            title: folderName,
            folderName: folderName,
            imageData: imageUrl || '',
            libraryUrl: `${siteUrl}/${cleanLibraryName}/Forms/AllItems.aspx?id=${encodeURIComponent(folderServerRelativeUrl)}`,
            viewAllUrl: '',
            pageUrl: pageUrl,
            viewDocumentsText: 'View Documents',
            orderBy: orderBy
          };
        })
      );

      // ✅ STEP 4: Sort categories by Order By value
      const sortedCategories = dynamicCategories.sort((a, b) => {
        if (a.orderBy !== undefined && b.orderBy !== undefined) {
          return a.orderBy - b.orderBy;
        }
        if (a.orderBy !== undefined) return -1;
        if (b.orderBy !== undefined) return 1;
        return a.title.localeCompare(b.title);
      });

      console.log('📂 Sorted categories:', sortedCategories.map(c => `${c.title} (Order: ${c.orderBy !== undefined ? c.orderBy : 'none'})`));
      this.setState({ categories: sortedCategories, loading: false });

    } catch (error) {
      console.error('❌ Exception:', error);
      alert(`Error loading folders: ${error instanceof Error ? error.message : String(error)}`);
      this.setState({ categories: [], loading: false });
    }
  };

  private checkForCoverImage = async (folderUrl: string): Promise<string | null> => {
  try {
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    const imageExtensions = ['jpg', 'jpeg', 'png', 'gif'];
    
    for (const ext of imageExtensions) {
      try {
        const imageServerRelativeUrl = folderUrl + '/cover.' + ext;
        const checkUrl = siteUrl + '/_api/web/GetFileByServerRelativeUrl(\'' + imageServerRelativeUrl + '\')';
        
        const response = await this.props.context.spHttpClient.get(
          checkUrl, 
          SPHttpClient.configurations.v1
        );
        
        if (response.ok) {
          return siteUrl + imageServerRelativeUrl;
        }
      } catch (err) {
        continue;
      }
    }
    return null;
  } catch (error) {
    console.error('Error checking for cover image:', error);
    return null;
  }
};

private handleCategoryClick = (category: IDocumentCategory): void => {
  if (this.props.isDynamicMode && category.pageUrl) {
    window.location.href = category.pageUrl;
  } else if (category.libraryUrl && category.libraryUrl !== '') {
    const url = this.formatUrl(category.libraryUrl);
    window.open(url, '_self');
  }
};


  private formatUrl = (url: string): string => {
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      return `https://${url}`;
    }
    return url;
  };

  private isValidUrl = (url: string): boolean => {
    try {
      const formattedUrl = this.formatUrl(url);
      new URL(formattedUrl);
      return true;
    } catch {
      return false;
    }
  };

  private handleViewDocuments = (category: IDocumentCategory, event: React.MouseEvent): void => {
    event.stopPropagation();
    event.preventDefault();
    
    if (this.props.isDynamicMode && category.pageUrl) {
      window.location.href = category.pageUrl;
    } else if (category.viewAllUrl && category.viewAllUrl !== '') {
      const url = this.formatUrl(category.viewAllUrl);
      window.open(url, '_self');
    } else if (category.libraryUrl && category.libraryUrl !== '') {
      let url = this.formatUrl(category.libraryUrl);
      if (url.includes('sharepoint.com') && !url.includes('/Forms/AllItems.aspx')) {
        url = url.replace(/\/$/, '') + '/Forms/AllItems.aspx';
      }
      window.open(url, '_self');
    }
  };

  private handleSharePointUpload = async (categoryId: string): Promise<void> => {
    this.setState({ uploadingFor: categoryId });

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.multiple = false;
    
    fileInput.onchange = async (e: Event) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        try {
          const uploadedImageUrl = await this.uploadImageToSharePoint(file);
          
          const updatedCategories = this.state.categories.map(cat =>
            cat.id === categoryId 
              ? { ...cat, imageData: uploadedImageUrl }
              : cat
          );
          
          this.setState({ categories: updatedCategories, uploadingFor: undefined });
          this.props.onCategoriesUpdate(updatedCategories);
          
        } catch (error) {
          console.error('Upload failed:', error);
          this.setState({ uploadingFor: undefined });
          alert('Upload failed. Please try again.');
        }
      } else {
        this.setState({ uploadingFor: undefined });
      }
    };

    fileInput.oncancel = () => {
      this.setState({ uploadingFor: undefined });
    };

    fileInput.click();
  };

  private uploadImageToSharePoint = async (file: File): Promise<string> => {
    const fileName = `document-category-${Date.now()}-${file.name}`;
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    const uploadUrl = `${siteUrl}/_api/web/lists/getbytitle('Site Assets')/RootFolder/Files/Add(url='${fileName}',overwrite=true)`;

    const arrayBuffer = await file.arrayBuffer();
    
    const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(
      uploadUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: arrayBuffer
      }
    );

    if (response.ok) {
      const result = await response.json();
      return result.d.ServerRelativeUrl.startsWith('/') 
        ? `${siteUrl}${result.d.ServerRelativeUrl}`
        : result.d.ServerRelativeUrl;
    } else {
      throw new Error(`Upload failed: ${response.statusText}`);
    }
  };

  private handleImageUpload = (categoryId: string): void => {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
          const imageData = event.target?.result as string;
          const updatedCategories = this.state.categories.map(cat =>
            cat.id === categoryId 
              ? { ...cat, imageData: imageData }
              : cat
          );
          this.setState({ categories: updatedCategories });
          this.props.onCategoriesUpdate(updatedCategories);
        };
        reader.readAsDataURL(file);
      }
    };
    fileInput.click();
  };

  private handleDeleteCategory = (categoryId: string, event: React.MouseEvent): void => {
    event.stopPropagation();
    if (confirm('Are you sure you want to delete this category?')) {
      const updatedCategories = this.state.categories.filter(cat => cat.id !== categoryId);
      this.setState({ categories: updatedCategories });
      this.props.onCategoriesUpdate(updatedCategories);
    }
  };

  private getDefaultImage = (title: string): string => {
    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#e67e22'];
    const colorIndex = title.length % colors.length;
    const color = colors[colorIndex];
    
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 250'%3E%3Crect width='400' height='250' fill='${encodeURIComponent(color)}'/%3E%3Cg fill='white'%3E%3Crect x='150' y='80' width='100' height='80' rx='5' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='170,100 170,140 230,140' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='180,110 210,110' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,120 220,120' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,130 200,130' fill='none' stroke='white' stroke-width='2'/%3E%3C/g%3E%3C/svg%3E`;
  };

  public render(): React.ReactElement<IDocumentsProps> {
    const { columnsPerRow, isDynamicMode } = this.props;
    const { categories, editMode, uploadingFor, loading } = this.state;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    if (loading) {
      return (
        <div className={styles.documents}>
          <div>
            <Spinner label="Loading resources..." />
          </div>
        </div>
      );
    }

    return (
      <div className={styles.documents}>
        <div className={`${styles.categoriesGrid} ${gridClassName}`}>
          {categories.map((category) => (
            <div 
              key={category.id} 
              className={`${styles.categoryCard} ${!this.isValidUrl(category.libraryUrl) ? styles.invalidUrl : ''}`}
              onClick={() => !editMode && this.handleCategoryClick(category)}
            >
              {editMode && !isDynamicMode && (
                <button
                  className={styles.deleteButton}
                  onClick={(e) => this.handleDeleteCategory(category.id, e)}
                  title="Delete category"
                >
                  ×
                </button>
              )}
              <div className={styles.imageContainer}>
                <img 
                  src={category.imageData || this.getDefaultImage(category.title)} 
                  alt={category.title}
                  className={styles.categoryImage}
                />
                {editMode && !isDynamicMode && (
                  <div className={styles.editOverlay}>
                    <button 
                      className={styles.uploadButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        void this.handleSharePointUpload(category.id);
                      }}
                      disabled={uploadingFor === category.id}
                    >
                      {uploadingFor === category.id ? '⏳ Uploading...' : '📷 Upload to SharePoint'}
                    </button>
                    <button 
                      className={styles.uploadButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        this.handleImageUpload(category.id);
                      }}
                      style={{ marginTop: '5px', fontSize: '12px' }}
                    >
                      📁 Upload as Base64
                    </button>
                  </div>
                )}
              </div>
              <div className={styles.categoryInfo}>
                <h3 className={styles.categoryTitle}>{category.title}</h3>
                <a 
                  href="#"
                  className={styles.viewAllLink}
                  onClick={(e) => this.handleViewDocuments(category, e)}
                >
                  {category.viewDocumentsText || 'View Documents'}
                </a>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
