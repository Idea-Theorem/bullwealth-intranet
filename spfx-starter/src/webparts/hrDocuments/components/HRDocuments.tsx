import * as React from 'react';
import styles from './HRDocuments.module.scss';
import { IHRDocumentsProps, IHRDocument } from './IHRDocumentsProps';

export interface IHRDocumentsState {
  documents: IHRDocument[];
  showUploadModal: boolean;
  selectedDocumentId: string | null;
  editMode: boolean;
}

export default class HRDocuments extends React.Component<IHRDocumentsProps, IHRDocumentsState> {
  private fileInputRef = React.createRef<HTMLInputElement>();

  constructor(props: IHRDocumentsProps) {
    super(props);
    this.state = {
      documents: props.documents,
      showUploadModal: false,
      selectedDocumentId: null,
      editMode: false
    };
  }

  public componentDidUpdate(prevProps: IHRDocumentsProps): void {
    if (prevProps.documents !== this.props.documents) {
      this.setState({ documents: this.props.documents });
    }
  }

  private handleAddDocument = (): void => {
    const newDocument: IHRDocument = {
      id: `doc-${Date.now()}`,
      title: 'New Document',
      date: new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      }),
      iconData: '',
      iconType: 'word',
      documentUrl: '#'
    };

    const updatedDocuments = [...this.state.documents, newDocument];
    this.setState({ documents: updatedDocuments });
    this.props.onDocumentsUpdate(updatedDocuments);
  }

  private handleDeleteDocument = (id: string): void => {
    const updatedDocuments = this.state.documents.filter(doc => doc.id !== id);
    this.setState({ documents: updatedDocuments });
    this.props.onDocumentsUpdate(updatedDocuments);
  }

  private handleUpdateDocument = (id: string, updates: Partial<IHRDocument>): void => {
    const updatedDocuments = this.state.documents.map(doc =>
      doc.id === id ? { ...doc, ...updates } : doc
    );
    this.setState({ documents: updatedDocuments });
    this.props.onDocumentsUpdate(updatedDocuments);
  }

  private handleIconUpload = (documentId: string): void => {
    this.setState({ selectedDocumentId: documentId }, () => {
      this.fileInputRef.current?.click();
    });
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (!file || !this.state.selectedDocumentId) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const iconData = e.target?.result as string;
      this.handleUpdateDocument(this.state.selectedDocumentId!, {
        iconData: iconData,
        iconType: 'custom'
      });
      this.setState({ selectedDocumentId: null });
    };
    reader.readAsDataURL(file);
    
    // Reset input
    event.target.value = '';
  }

  private handleDocumentUpload = (documentId: string): void => {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx';
    fileInput.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        const fileUrl = URL.createObjectURL(file);
        const fileType = file.name.endsWith('.pdf') ? 'pdf' : 'word';
        this.handleUpdateDocument(documentId, {
          documentUrl: fileUrl,
          iconType: fileType as 'word' | 'pdf' | 'custom',
          title: file.name.replace(/\.[^/.]+$/, '')
        });
      }
    };
    fileInput.click();
  }

  private handleDocumentClick = (document: IHRDocument): void => {
    if (document.documentUrl && document.documentUrl !== '#') {
      window.open(document.documentUrl, '_blank');
    }
  }

  private getDefaultIcon = (type: 'word' | 'pdf' | 'custom'): React.ReactElement => {
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
    }
    return (
      <svg className={styles.defaultIcon} viewBox="0 0 32 32" fill="none">
        <rect width="32" height="32" rx="4" fill="#6B7280"/>
        <text x="16" y="22" textAnchor="middle" fill="white" fontSize="14">📄</text>
      </svg>
    );
  }

  private toggleEditMode = (): void => {
    this.setState({ editMode: !this.state.editMode });
  }

  public render(): React.ReactElement<IHRDocumentsProps> {
    const { title, columnsPerRow, showDate, allowUpload } = this.props;
    const { documents, editMode } = this.state;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    return (
      <div className={styles.hrDocuments}>
        <div className={styles.header}>
          <h2 className={styles.title}>{title}</h2>
          {allowUpload && (
            <div className={styles.controls}>
              <button 
                className={styles.editButton}
                onClick={this.toggleEditMode}
              >
                {editMode ? 'Done' : 'Edit'}
              </button>
              {editMode && (
                <button 
                  className={styles.addButton}
                  onClick={this.handleAddDocument}
                >
                  + Add Document
                </button>
              )}
            </div>
          )}
        </div>

        <div className={`${styles.documentsGrid} ${gridClassName}`}>
          {documents.map((document) => (
            <div key={document.id} className={styles.documentCard}>
              {editMode && (
                <div className={styles.editControls}>
                  <button
                    className={styles.deleteButton}
                    onClick={() => this.handleDeleteDocument(document.id)}
                    aria-label="Delete document"
                  >
                    ×
                  </button>
                </div>
              )}
              
              <div 
                className={styles.iconContainer}
                onClick={() => !editMode && this.handleDocumentClick(document)}
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
                
                {editMode && (
                  <div className={styles.iconOverlay}>
                    <button
                      className={styles.uploadIconButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        this.handleIconUpload(document.id);
                      }}
                    >
                      📷 Change Icon
                    </button>
                  </div>
                )}
              </div>

              <div className={styles.documentInfo}>
                {editMode ? (
                  <input
                    className={styles.titleInput}
                    value={document.title}
                    onChange={(e) => this.handleUpdateDocument(document.id, { title: e.target.value })}
                    placeholder="Document title"
                  />
                ) : (
                  <h3 
                    className={styles.documentTitle}
                    onClick={() => this.handleDocumentClick(document)}
                  >
                    {document.title}
                  </h3>
                )}
                
                {showDate && (
                  <p className={styles.documentDate}>
                    John Doe · {document.date}
                  </p>
                )}

                {editMode && (
                  <button
                    className={styles.uploadDocButton}
                    onClick={() => this.handleDocumentUpload(document.id)}
                  >
                    📎 Upload Document
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>

        <input
          ref={this.fileInputRef}
          type="file"
          accept="image/*"
          onChange={this.handleFileSelect}
          style={{ display: 'none' }}
        />
      </div>
    );
  }
}