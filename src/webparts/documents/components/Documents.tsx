import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps, IDocumentCategory } from './IDocumentsProps';
<<<<<<< Updated upstream
=======
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
>>>>>>> Stashed changes

export interface IDocumentsState {
  categories: IDocumentCategory[];
  editMode: boolean;
  uploadingFor?: string; // Track which category is uploading
}

export default class Documents extends React.Component<IDocumentsProps, IDocumentsState> {
  constructor(props: IDocumentsProps) {
    super(props);
    this.state = {
      categories: props.categories,
      editMode: false,
      uploadingFor: undefined
    };
  }

  public componentDidUpdate(prevProps: IDocumentsProps): void {
    if (JSON.stringify(prevProps.categories) !== JSON.stringify(this.props.categories)) {
      this.setState({ categories: this.props.categories });
    }
  }
  
  private handleCategoryClick = (category: IDocumentCategory): void => {
    if (category.libraryUrl && category.libraryUrl !== '') {
      const url = this.formatUrl(category.libraryUrl);
      window.open(url, '_self');
    }
  }

  private formatUrl = (url: string): string => {
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      return `https://${url}`;
    }
    return url;
  }

  private isValidUrl = (url: string): boolean => {
    try {
      const formattedUrl = this.formatUrl(url);
      new URL(formattedUrl);
      return true;
    } catch {
      return false;
    }
  }

  private handleViewDocuments = (category: IDocumentCategory, event: React.MouseEvent): void => {
    event.stopPropagation();
    event.preventDefault();
    
    if (category.viewAllUrl && category.viewAllUrl !== '') {
      const url = this.formatUrl(category.viewAllUrl);
      window.open(url, '_self');
    } else if (category.libraryUrl && category.libraryUrl !== '') {
      let url = this.formatUrl(category.libraryUrl);
      if (url.includes('sharepoint.com') && !url.includes('/Forms/AllItems.aspx')) {
        url = url.replace(/\/$/, '') + '/Forms/AllItems.aspx';
      }
      window.open(url, '_self');
    }
  }

  // NEW: SharePoint native upload functionality
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
          // Upload to SharePoint
          const uploadedImageUrl = await this.uploadImageToSharePoint(file);
          
          // Update category with SharePoint URL
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
  }

  // Upload image to SharePoint Site Assets library
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
  }

  // Fallback: Base64 upload (existing functionality)
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
  }

  private handleDeleteCategory = (categoryId: string, event: React.MouseEvent): void => {
    event.stopPropagation();
    if (confirm('Are you sure you want to delete this category?')) {
      const updatedCategories = this.state.categories.filter(cat => cat.id !== categoryId);
      this.setState({ categories: updatedCategories });
      this.props.onCategoriesUpdate(updatedCategories);
    }
  }

  private getDefaultImage = (title: string): string => {
    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#e67e22'];
    const colorIndex = title.length % colors.length;
    const color = colors[colorIndex];
    
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 250'%3E%3Crect width='400' height='250' fill='${encodeURIComponent(color)}'/%3E%3Cg fill='white'%3E%3Crect x='150' y='80' width='100' height='80' rx='5' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='170,100 170,140 230,140' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='180,110 210,110' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,120 220,120' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,130 200,130' fill='none' stroke='white' stroke-width='2'/%3E%3C/g%3E%3C/svg%3E`;
  }

<<<<<<< Updated upstream
  private toggleEditMode = (): void => {
    this.setState({ editMode: !this.state.editMode });
  }

=======
>>>>>>> Stashed changes
  public render(): React.ReactElement<IDocumentsProps> {
    const { columnsPerRow } = this.props;
    const { categories, editMode, uploadingFor } = this.state;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    return (
      <div className={styles.documents}>
        <div className={`${styles.categoriesGrid} ${gridClassName}`}>
          {categories.map((category) => (
            <div 
              key={category.id} 
              className={`${styles.categoryCard} ${!this.isValidUrl(category.libraryUrl) ? styles.invalidUrl : ''}`}
              onClick={() => !editMode && this.handleCategoryClick(category)}
            >
              {editMode && (
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
                {editMode && (
                  <div className={styles.editOverlay}>
                    <button 
                      className={styles.uploadButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        // eslint-disable-next-line @typescript-eslint/no-floating-promises
                        this.handleSharePointUpload(category.id);
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
                  {/* NEW: Use custom text or fallback to "View Documents" */}
                  {category.viewDocumentsText || 'View Documents'}
                </a>
              </div>
            </div>
          ))}
        </div>
<<<<<<< Updated upstream
        
        <button 
          className={styles.editModeToggle}
          onClick={this.toggleEditMode}
          title={editMode ? "Done editing" : "Edit categories"}
        >
          {editMode ? '✓' : '✏️'}
        </button>
=======
>>>>>>> Stashed changes
      </div>
    );
  }
}
