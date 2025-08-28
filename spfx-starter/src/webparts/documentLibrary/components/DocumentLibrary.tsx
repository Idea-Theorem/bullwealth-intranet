import * as React from 'react';
import styles from './DocumentLibrary.module.scss';
import { IDocumentLibraryProps, IDocument } from './IDocumentLibraryProps';

export interface IDocumentLibraryState {
  documents: IDocument[];
  selectedDocuments: string[];
  showActionsMenu: string | null;
  sortColumn: string;
  sortDirection: 'asc' | 'desc';
  isLoading: boolean;
  error: string | null;
}

export default class DocumentLibrary extends React.Component<IDocumentLibraryProps, IDocumentLibraryState> {
  private checkboxRef = React.createRef<HTMLInputElement>();

  constructor(props: IDocumentLibraryProps) {
    super(props);
    this.state = {
      documents: [],
      selectedDocuments: [],
      showActionsMenu: null,
      sortColumn: 'modified',
      sortDirection: 'desc',
      isLoading: true,
      error: null
    };
  }

  public componentDidMount(): void {
    void this.loadDocuments();
  }

  public componentDidUpdate(): void {
    // Set indeterminate state for checkbox
    if (this.checkboxRef.current) {
      const someSelected = this.state.documents.some((doc: IDocument) => doc.isSelected);
      const allSelected = this.state.documents.length > 0 && this.state.documents.every((doc: IDocument) => doc.isSelected);
      this.checkboxRef.current.indeterminate = someSelected && !allSelected;
    }
  }

  private loadDocuments = async (): Promise<void> => {
    // Mock data for demonstration - replace with actual SharePoint API call
    const mockDocuments: IDocument[] = [
      {
        id: '1',
        name: 'Record Keep Policy',
        modified: 'July 5, 2025',
        modifiedBy: 'John Smith',
        fileType: 'pdf',
        fileUrl: '#',
        isSelected: false,
        createElement: function (): unknown {
          throw new Error('Function not implemented.');
        },
        body: undefined,
        execCommand: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        }
      },
      {
        id: '2',
        name: 'Cross Trade Policy',
        modified: 'July 5, 2025',
        modifiedBy: 'John Smith',
        fileType: 'pdf',
        fileUrl: '#',
        isSelected: false,
        createElement: function (): unknown {
          throw new Error('Function not implemented.');
        },
        execCommand: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        },
        body: undefined
      },
      {
        id: '3',
        name: 'Compliance Manual Deck',
        modified: 'July 5, 2025',
        modifiedBy: 'John Smith',
        fileType: 'pptx',
        fileUrl: '#',
        isSelected: false,
        createElement: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        },
        execCommand: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        },
        body: undefined
      }
    ];

    this.setState({ 
      documents: mockDocuments, 
      isLoading: false 
    });
  }

  private handleSelectAll = (checked: boolean): void => {
    const documents = this.state.documents.map((doc: IDocument) => ({
      ...doc,
      isSelected: checked
    }));
    const selectedDocuments = checked ? documents.map((doc: IDocument) => doc.id) : [];
    this.setState({ documents, selectedDocuments });
  }

  private handleSelectDocument = (id: string): void => {
    const documents = this.state.documents.map((doc: IDocument) =>
      doc.id === id ? { ...doc, isSelected: !doc.isSelected } : doc
    );
    const selectedDocuments = documents
      .filter((doc: IDocument) => doc.isSelected)
      .map((doc: IDocument) => doc.id);
    this.setState({ documents, selectedDocuments });
  }

  private handleSort = (column: string): void => {
    const { sortColumn, sortDirection, documents } = this.state;
    const newDirection = sortColumn === column && sortDirection === 'asc' ? 'desc' : 'asc';
    
    const sortedDocuments = [...documents].sort((a: IDocument, b: IDocument) => {
      const aVal = (a as any)[column];
      const bVal = (b as any)[column];
      
      if (newDirection === 'asc') {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });

    this.setState({
      documents: sortedDocuments,
      sortColumn: column,
      sortDirection: newDirection
    });
  }

  private handleActionsClick = (documentId: string): void => {
    this.setState({ 
      showActionsMenu: this.state.showActionsMenu === documentId ? null : documentId 
    });
  }

  private handleAction = (action: string, documentId: string): void => {
    const document = this.state.documents.find((doc: IDocument) => doc.id === documentId);
    if (!document) return;

    switch (action) {
      case 'view':
        this.viewDocument(document);
        break;
      case 'share':
        this.shareDocument(document);
        break;
      case 'export':
        this.exportDocument(document);
        break;
      case 'copyLink':
        this.copyLink(document);
        break;
      default:
        break;
    }
    this.setState({ showActionsMenu: null });
  }

  private viewDocument = (document: IDocument): void => {
    // For uploaded files (blob URLs), open in new tab
    // For SharePoint files, navigate to the file
    if (document.fileUrl.startsWith('blob:')) {
      window.open(document.fileUrl, '_blank');
    } else if (document.fileUrl !== '#') {
      window.open(document.fileUrl, '_blank');
    } else {
      alert(`Viewing: ${document.name}`);
    }
  }

  private shareDocument = (document: IDocument): void => {
    const shareText = `Check out this document: ${document.name}`;
    
    if (navigator.share) {
      void navigator.share({
        title: document.name,
        text: shareText,
        url: document.fileUrl.startsWith('blob:') ? window.location.href : document.fileUrl
      }).catch((error) => {
        console.log('Share failed:', error);
        this.copyLink(document);
      });
    } else {
      this.copyLink(document);
      alert('Link copied to clipboard! You can now share it.');
    }
  }

  private exportDocument = (document: IDocument): void => {
    // Create a download link for the document
    const link = window.document.createElement('a');
    link.href = document.fileUrl;
    link.download = document.name;
    link.style.display = 'none';
    window.document.body.appendChild(link);
    link.click();
    window.document.body.removeChild(link);
  }

  private copyLink = (document: IDocument): void => {
    const linkToCopy = document.fileUrl.startsWith('blob:') 
      ? `${window.location.href}#${document.name}`
      : document.fileUrl;
      
    // eslint-disable-next-line no-void
    void navigator.clipboard.writeText(linkToCopy).then(() => {
      alert('Link copied to clipboard!');
    }).catch((error) => {
      console.error('Failed to copy link:', error);
      // Fallback method
      const textArea = document.createElement('textarea') as HTMLTextAreaElement;
      textArea.value = linkToCopy;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand('copy');
      document.body.removeChild(textArea);
      alert('Link copied to clipboard!');
    });
  }

  private handleUpload = (): void => {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.multiple = true;
    fileInput.accept = '.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt';
    fileInput.onchange = async (e: Event) => {
      const target = e.target as HTMLInputElement;
      const files = target.files;
      if (files && files.length > 0) {
        this.processUploadedFiles(files);
      }
    };
    fileInput.click();
  }

  private processUploadedFiles = (files: FileList): void => {
    const newDocuments: IDocument[] = [];
    
    Array.from(files).forEach((file, index) => {
      const fileExtension = file.name.split('.').pop() || 'unknown';
      const newDoc: IDocument = {
        id: `uploaded-${Date.now()}-${index}`,
        name: file.name,
        modified: new Date().toLocaleDateString('en-US', {
          year: 'numeric',
          month: 'long',
          day: 'numeric'
        }),
        modifiedBy: this.props.userDisplayName || 'Current User',
        fileType: fileExtension,
        fileUrl: URL.createObjectURL(file),
        isSelected: false,
        execCommand: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        },
        body: undefined,
        createElement: function (arg0: string): unknown {
          throw new Error('Function not implemented.');
        }
      };
      newDocuments.push(newDoc);
    });

    // Add new documents to the existing list
    this.setState(prevState => ({
      documents: [...prevState.documents, ...newDocuments],
      isLoading: false
    }));
  }

  private handleFileClick = (document: IDocument, event: React.MouseEvent): void => {
    event.stopPropagation();
    // Toggle the actions menu for this document
    this.setState({ 
      showActionsMenu: this.state.showActionsMenu === document.id ? null : document.id 
    });
  }

  private getFileIcon = (fileType: string): string => {
    const icons: { [key: string]: string } = {
      pdf: '📄',
      doc: '📝',
      docx: '📝',
      xls: '📊',
      xlsx: '📊',
      ppt: '📊',
      pptx: '📊',
      default: '📎'
    };
    return icons[fileType.toLowerCase()] || icons.default;
  }

  public render(): React.ReactElement<IDocumentLibraryProps> {
    const {
      title,
      description,
      libraryName,
      showUploadButton,
      showActions
    } = this.props;

    const {
      documents,
      showActionsMenu,
      sortColumn,
      sortDirection,
      isLoading,
      error
    } = this.state;

    const allSelected = documents.length > 0 && documents.every((doc: IDocument) => doc.isSelected);

    return (
      <div className={styles.documentLibrary}>
        <div className={styles.header}>
          <button className={styles.backButton} aria-label="Go back">
            ←
          </button>
          <h1 className={styles.title}>{title}</h1>
        </div>
        
        <p className={styles.description}>{description}</p>

        <div className={styles.libraryHeader}>
          <h2 className={styles.libraryName}>{libraryName}</h2>
          {showUploadButton && (
            <button className={styles.uploadButton} onClick={this.handleUpload}>
              📤 Upload
            </button>
          )}
        </div>

        {isLoading ? (
          <div className={styles.loading}>Loading documents...</div>
        ) : error ? (
          <div className={styles.error}>{error}</div>
        ) : (
          <div className={styles.documentTable}>
            <div className={styles.tableHeader}>
              <div className={styles.headerCell}>
                <input
                  ref={this.checkboxRef}
                  type="checkbox"
                  checked={allSelected}
                  onChange={(e) => this.handleSelectAll(e.target.checked)}
                  aria-label="Select all documents"
                />
              </div>
              <div 
                className={styles.headerCell}
                onClick={() => this.handleSort('name')}
              >
                Name {sortColumn === 'name' && (
                  <span className={styles.sortIcon}>
                    {sortDirection === 'asc' ? '↑' : '↓'}
                  </span>
                )}
              </div>
              <div 
                className={styles.headerCell}
                onClick={() => this.handleSort('modified')}
              >
                Modified {sortColumn === 'modified' && (
                  <span className={styles.sortIcon}>
                    {sortDirection === 'asc' ? '↑' : '↓'}
                  </span>
                )}
              </div>
              <div 
                className={styles.headerCell}
                onClick={() => this.handleSort('modifiedBy')}
              >
                Modified By {sortColumn === 'modifiedBy' && (
                  <span className={styles.sortIcon}>
                    {sortDirection === 'asc' ? '↑' : '↓'}
                  </span>
                )}
              </div>
              {showActions && (
                <div className={styles.headerCell}>Actions</div>
              )}
            </div>

            <div className={styles.tableBody}>
              {documents.map((document: IDocument) => (
                <div 
                  key={document.id} 
                  className={`${styles.tableRow} ${document.isSelected ? styles.selected : ''}`}
                >
                  <div className={styles.cell}>
                    <input
                      type="checkbox"
                      checked={document.isSelected}
                      onChange={() => this.handleSelectDocument(document.id)}
                      aria-label={`Select ${document.name}`}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                  <div 
                    className={`${styles.cell} ${styles.clickableCell}`}
                    onClick={(e) => this.handleFileClick(document, e)}
                  >
                    <span className={styles.fileIcon}>
                      {this.getFileIcon(document.fileType)}
                    </span>
                    <span className={styles.fileName}>{document.name}</span>
                  </div>
                  <div className={styles.cell}>{document.modified}</div>
                  <div className={styles.cell}>{document.modifiedBy}</div>
                  {showActions && (
                    <div className={styles.cell}>
                      <button
                        className={styles.actionsButton}
                        onClick={(e) => {
                          e.stopPropagation();
                          this.handleActionsClick(document.id);
                        }}
                        aria-label="Actions menu"
                      >
                        •••
                      </button>
                      {showActionsMenu === document.id && (
                        <div className={styles.actionsMenu}>
                          <button onClick={() => this.handleAction('view', document.id)}>
                            👁️ View
                          </button>
                          <button onClick={() => this.handleAction('share', document.id)}>
                            🔗 Share
                          </button>
                          <button onClick={() => this.handleAction('export', document.id)}>
                            ⬇️ Export
                          </button>
                          <button onClick={() => this.handleAction('copyLink', document.id)}>
                            📋 Copy link
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  }
}