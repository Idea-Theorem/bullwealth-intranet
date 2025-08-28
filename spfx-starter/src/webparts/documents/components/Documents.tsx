import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps, IDocumentCategory } from './IDocumentsProps';

export interface IDocumentsState {
  categories: IDocumentCategory[];
  editMode: boolean;
}

export default class Documents extends React.Component<IDocumentsProps, IDocumentsState> {
  constructor(props: IDocumentsProps) {
    super(props);
    this.state = {
      categories: props.categories,
      editMode: false
    };
  }

  public componentDidUpdate(prevProps: IDocumentsProps): void {
    if (JSON.stringify(prevProps.categories) !== JSON.stringify(this.props.categories)) {
      this.setState({ categories: this.props.categories });
    }
  }
  
  private handleCategoryClick = (category: IDocumentCategory): void => {
    if (category.libraryUrl && category.libraryUrl !== '') {
      window.open(category.libraryUrl, '_blank');
    } else if (category.libraryName) {
      const libraryUrl = `${this.props.context.pageContext.web.absoluteUrl}/${category.libraryName.replace(/\s+/g, '')}`;
      window.open(libraryUrl, '_blank');
    }
  }

  private handleViewDocuments = (category: IDocumentCategory, event: React.MouseEvent): void => {
    event.stopPropagation();
    event.preventDefault();
    
    if (category.libraryName) {
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      const libraryPath = category.libraryName.replace(/\s+/g, '');
      const viewAllUrl = `${siteUrl}/${libraryPath}/Forms/AllItems.aspx`;
      window.open(viewAllUrl, '_blank');
    } else if (category.viewAllUrl && category.viewAllUrl !== '') {
      window.open(category.viewAllUrl, '_blank');
    } else if (category.libraryUrl && category.libraryUrl !== '') {
      window.open(category.libraryUrl, '_blank');
    }
  }

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

  private toggleEditMode = (): void => {
    this.setState({ editMode: !this.state.editMode });
  }

  public render(): React.ReactElement<IDocumentsProps> {
    const { columnsPerRow } = this.props;
    const { categories, editMode } = this.state;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    return (
      <div className={styles.documents}>
        <div className={`${styles.categoriesGrid} ${gridClassName}`}>
          {categories.map((category) => (
            <div 
              key={category.id} 
              className={styles.categoryCard}
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
                        this.handleImageUpload(category.id);
                      }}
                    >
                      📷 Upload Image
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
                  View Documents
                </a>
              </div>
            </div>
          ))}
        </div>
        
        <button 
          className={styles.editModeToggle}
          onClick={this.toggleEditMode}
          title={editMode ? "Done editing" : "Edit categories"}
        >
          {editMode ? '✓' : '✏️'}
        </button>
      </div>
    );
  }
}