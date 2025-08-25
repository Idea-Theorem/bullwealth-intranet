import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps, IDocumentCategory } from './IDocumentsProps';

export default class Documents extends React.Component<IDocumentsProps> {
  
  private handleCategoryClick = (category: IDocumentCategory): void => {
    if (category.documentUrl && category.documentUrl !== '#') {
      window.open(category.documentUrl, '_blank');
    }
  }

  private handleViewAll = (category: IDocumentCategory, event: React.MouseEvent): void => {
    event.stopPropagation();
    if (category.viewAllUrl && category.viewAllUrl !== '#') {
      window.open(category.viewAllUrl, '_blank');
    }
  }

  private getDefaultImage = (title: string): string => {
    // Generate different colored placeholders based on title
    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#e67e22'];
    const colorIndex = title.length % colors.length;
    const color = colors[colorIndex];
    
    // Create an SVG placeholder
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 250'%3E%3Crect width='400' height='250' fill='${encodeURIComponent(color)}'/%3E%3Cg fill='white'%3E%3Crect x='150' y='80' width='100' height='80' rx='5' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='170,100 170,140 230,140' fill='none' stroke='white' stroke-width='3'/%3E%3Cpolyline points='180,110 210,110' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,120 220,120' fill='none' stroke='white' stroke-width='2'/%3E%3Cpolyline points='180,130 200,130' fill='none' stroke='white' stroke-width='2'/%3E%3C/g%3E%3C/svg%3E`;
  }

  public render(): React.ReactElement<IDocumentsProps> {
    const { title, categories, columnsPerRow } = this.props;
    
    //const gridClassName = styles[`columns${columnsPerRow}`] || styles.columns4;
const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;
    return (
      <div className={styles.documents}>
        <div className={styles.header}>
          <h2 className={styles.title}>{title}</h2>
          <p className={styles.subtitle}>
            Below is various documents, training and material for BullWealth branch
          </p>
        </div>
        
        <div className={`${styles.categoriesGrid} ${gridClassName}`}>
          {categories.map((category) => (
            <div 
              key={category.id} 
              className={styles.categoryCard}
              onClick={() => this.handleCategoryClick(category)}
            >
              <div className={styles.imageContainer}>
                <img 
                  src={category.imageUrl || this.getDefaultImage(category.title)} 
                  alt={category.title}
                  className={styles.categoryImage}
                />
                <div className={styles.overlay}>
                  <button 
                    className={styles.viewButton}
                    onClick={(e) => this.handleViewAll(category, e)}
                    aria-label={`View all ${category.title} documents`}
                  >
                    View Documents
                  </button>
                </div>
              </div>
              <div className={styles.categoryInfo}>
                <h3 className={styles.categoryTitle}>{category.title}</h3>
                <a 
                  href={category.viewAllUrl}
                  className={styles.viewAllLink}
                  onClick={(e) => this.handleViewAll(category, e)}
                >
                  View Documents
                </a>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}