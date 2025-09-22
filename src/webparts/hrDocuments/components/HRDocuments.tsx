import * as React from 'react';
import styles from './HRDocuments.module.scss';
import { IHRDocumentsProps, IHRDocument } from './IHRDocumentsProps';

export interface IHRDocumentsState {
  documents: IHRDocument[];
}

export default class HRDocuments extends React.Component<IHRDocumentsProps, IHRDocumentsState> {

  constructor(props: IHRDocumentsProps) {
    super(props);
    this.state = {
      documents: props.documents
    };
  }

  public componentDidUpdate(prevProps: IHRDocumentsProps): void {
    if (prevProps.documents !== this.props.documents) {
      this.setState({ documents: this.props.documents });
    }
  }

  private handleDocumentClick = (document: IHRDocument): void => {
    if (document.documentUrl && document.documentUrl !== '#') {
      window.open(document.documentUrl, '_blank');
    }
  }

  private getDefaultIcon = (type: 'word' | 'pdf' | 'video' | 'custom'): React.ReactElement => {
    if (type === 'word') {
      return (
        <svg className={styles.wordIcon} viewBox="0 0 32 32" fill="none">
          <rect width="32" height="32" rx="4" fill="#2B579A"/>
          <text x="16" y="22" textAnchor="middle" fill="white" fontSize="14" fontWeight="bold">W</text>
        </svg>
      );
    } else if (type === 'pdf') {
      return (
        <svg className={styles.pdfIcon} viewBox="0 0 32 32" fill="none">
          <rect width="32" height="32" rx="4" fill="#DC2626"/>
          <text x="16" y="22" textAnchor="middle" fill="white" fontSize="10" fontWeight="bold">PDF</text>
        </svg>
      );
    } else if (type === 'video') {
      return (
        <svg className={styles.videoIcon} viewBox="0 0 32 32" fill="none">
          <rect width="32" height="32" rx="4" fill="#7C3AED"/>
          <text x="16" y="22" textAnchor="middle" fill="white" fontSize="10" fontWeight="bold">MP4</text>
        </svg>
      );
    }
    return (
      <svg className={styles.defaultIcon} viewBox="0 0 32 32" fill="none">
        <rect width="32" height="32" rx="4" fill="#6B7280"/>
        <text x="16" y="22" textAnchor="middle" fill="white" fontSize="14">📄</text>
      </svg>
    );
  }

  public render(): React.ReactElement<IHRDocumentsProps> {
    const { title, columnsPerRow, showDate } = this.props;
    const { documents } = this.state;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    return (
      <div className={styles.hrDocuments}>
        <div className={styles.header}>
          <h2 className={styles.title}>{title}</h2>
        </div>

        <div className={`${styles.documentsGrid} ${gridClassName}`}>
          {documents.map((document) => (
            <div key={document.id} className={styles.documentCard}>
              <div 
                className={styles.iconContainer}
                onClick={() => this.handleDocumentClick(document)}
              >
                {document.iconData ? (
                  <img 
                    src={document.iconData} 
                    alt={document.title}
                    className={styles.customIcon}
                  />
                ) : (
                  this.getDefaultIcon(document.iconType)
                )}
              </div>

              <div className={styles.documentInfo}>
                <h3 
                  className={styles.documentTitle}
                  onClick={() => this.handleDocumentClick(document)}
                >
                  {document.title || 'Untitled Document'}
                </h3>
                
                {showDate && (
                  <p className={styles.documentDate}>
                    {document.author} · {document.date}
                  </p>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
