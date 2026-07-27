import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  IPropertyPaneGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import HRDocuments from './components/HRDocuments';
import { IHRDocumentsProps } from './components/IHRDocumentsProps';

export interface IHRDocument {
  id: string;
  title: string;
  date: string;
  author: string;
  iconData: string;
  iconType: 'word' | 'pdf' | 'video' | 'custom';
  documentUrl: string;
  documentData?: string; // Base64 encoded file data
}

export interface IHRDocumentsWebPartProps {
  title: string;
  documents: string;
  columnsPerRow: number;
  showDate: boolean;
  expandedDocuments: string;
}

export default class HRDocumentsWebPart extends BaseClientSideWebPart<IHRDocumentsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _isUploading: boolean = false;

  protected onInit(): Promise<void> {
    if (!this.properties.title) {
      this.properties.title = 'Common HR Documents';
    }
    if (!this.properties.columnsPerRow) {
      this.properties.columnsPerRow = 4;
    }
    if (!this.properties.documents) {
      this.properties.documents = JSON.stringify(this.getDefaultDocuments());
    }
    if (!this.properties.expandedDocuments) {
      this.properties.expandedDocuments = JSON.stringify([0]);
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  public render(): void {
    let documents: IHRDocument[] = [];
    
    try {
      documents = this.properties.documents ? JSON.parse(this.properties.documents) : this.getDefaultDocuments();
    } catch {
      documents = this.getDefaultDocuments();
    }

    const element: React.ReactElement<IHRDocumentsProps> = React.createElement(
      HRDocuments,
      {
        title: this.properties.title || 'Common HR Documents',
        documents: documents,
        columnsPerRow: this.properties.columnsPerRow || 4,
        showDate: this.properties.showDate !== false,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        onDocumentsUpdate: (updatedDocs: IHRDocument[]) => {
          this.properties.documents = JSON.stringify(updatedDocs);
          this.context.propertyPane.refresh();
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getDefaultDocuments(): IHRDocument[] {
    return [
      {
        id: '1',
        title: 'Employee Handbook',
        date: 'September 22, 2025',
        author: 'HR Admin',
        iconData: '',
        iconType: 'word',
        documentUrl: '#'
      }
    ];
  }

  private getCurrentDocuments(): IHRDocument[] {
    try {
      return this.properties.documents ? JSON.parse(this.properties.documents) : [];
    } catch {
      return [];
    }
  }

  private getExpandedDocuments(): number[] {
    try {
      return this.properties.expandedDocuments ? JSON.parse(this.properties.expandedDocuments) : [];
    } catch {
      return [];
    }
  }

  private toggleDocumentExpansion = (index: number): void => {
    const expanded = this.getExpandedDocuments();
    const isExpanded = expanded.includes(index);
    
    let newExpanded: number[];
    if (isExpanded) {
      newExpanded = expanded.filter(i => i !== index);
    } else {
      newExpanded = [...expanded, index];
    }
    
    this.properties.expandedDocuments = JSON.stringify(newExpanded);
    this.context.propertyPane.refresh();
  }

  private addNewDocument = (): void => {
    const documents = this.getCurrentDocuments();
    const newDoc: IHRDocument = {
      id: `doc-${Date.now()}`,
      title: 'New Document',
      date: new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      }),
      author: this.context.pageContext.user.displayName,
      iconData: '',
      iconType: 'word',
      documentUrl: '#'
    };
    
    documents.push(newDoc);
    this.properties.documents = JSON.stringify(documents);
    
    // Auto-expand the new document
    const expanded = this.getExpandedDocuments();
    const newIndex = documents.length - 1;
    if (!expanded.includes(newIndex)) {
      expanded.push(newIndex);
      this.properties.expandedDocuments = JSON.stringify(expanded);
    }
    
    this.context.propertyPane.refresh();
    this.render();
  }

  private deleteDocument = (index: number): void => {
    const documents = this.getCurrentDocuments();
    if (index >= 0 && index < documents.length) {
      documents.splice(index, 1);
      this.properties.documents = JSON.stringify(documents);
      
      // Update expanded documents indices
      const expanded = this.getExpandedDocuments()
        .filter(i => i !== index)
        .map(i => i > index ? i - 1 : i);
      this.properties.expandedDocuments = JSON.stringify(expanded);
      
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private updateDocument = (index: number, updates: Partial<IHRDocument>): void => {
    const documents = this.getCurrentDocuments();
    if (index >= 0 && index < documents.length) {
      documents[index] = { ...documents[index], ...updates };
      this.properties.documents = JSON.stringify(documents);
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  // Simple file upload - converts to base64 and stores in web part properties
  private uploadDocument = (documentIndex: number): Promise<void> => {
    if (this._isUploading) {
      alert('Upload in progress. Please wait...');
      return Promise.resolve();
    }

    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf,.doc,.docx,.mp4,.avi,.mov,.wmv,.txt';
      
      input.onchange = async (e: Event) => {
        const target = e.target as HTMLInputElement;
        const file = target.files?.[0];
        if (file) {
          this._isUploading = true;
          this.context.propertyPane.refresh();
          
          try {
            // Convert file to base64
            const base64Data = await this.fileToBase64(file);
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            const fileType = this.getFileType(file.name);
            
            // Create blob URL for document access
            const blob = new Blob([file], { type: file.type });
            const blobUrl = URL.createObjectURL(blob);
            
            this.updateDocument(documentIndex, {
              title: fileName,
              documentUrl: blobUrl,
              documentData: base64Data,
              date: new Date().toLocaleDateString('en-US', { 
                year: 'numeric', 
                month: 'long', 
                day: 'numeric' 
              }),
              iconType: fileType
            });
            
            alert('Document uploaded successfully!');
          } catch (error) {
            console.error('Error uploading document:', error);
            alert(`Error uploading document: ${error.message || 'Please try again.'}`);
          } finally {
            this._isUploading = false;
            this.context.propertyPane.refresh();
            resolve();
          }
        }
      };
      
      input.click();
    });
  }

  // Simple icon upload - converts to base64
  private uploadIcon = (documentIndex: number): Promise<void> => {
    if (this._isUploading) {
      alert('Upload in progress. Please wait...');
      return Promise.resolve();
    }

    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'image/*';
      
      input.onchange = async (e: Event) => {
        const target = e.target as HTMLInputElement;
        const file = target.files?.[0];
        if (file) {
          this._isUploading = true;
          this.context.propertyPane.refresh();
          
          try {
            // Convert image to base64
            const base64Data = await this.fileToBase64(file);
            
            this.updateDocument(documentIndex, {
              iconData: base64Data,
              iconType: 'custom'
            });
            
            alert('Icon uploaded successfully!');
          } catch (error) {
            console.error('Error uploading icon:', error);
            alert(`Error uploading icon: ${error.message || 'Please try again.'}`);
          } finally {
            this._isUploading = false;
            this.context.propertyPane.refresh();
            resolve();
          }
        }
      };
      
      input.click();
    });
  }

  // Helper function to convert file to base64
  private fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === 'string') {
          resolve(reader.result);
        } else {
          reject(new Error('Failed to read file'));
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }

  // Helper function to detect file type
  private getFileType = (fileName: string): 'word' | 'pdf' | 'video' | 'custom' => {
    const extension = fileName.toLowerCase();
    if (extension.endsWith('.pdf')) return 'pdf';
    if (extension.endsWith('.doc') || extension.endsWith('.docx')) return 'word';
    if (extension.endsWith('.mp4') || extension.endsWith('.avi') || extension.endsWith('.mov') || extension.endsWith('.wmv')) return 'video';
    return 'custom';
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': environmentMessage = 'Office'; break;
            case 'Outlook': environmentMessage = 'Outlook'; break;
            case 'Teams': environmentMessage = 'Teams'; break;
            default: environmentMessage = 'SharePoint';
          }
          return environmentMessage;
        });
    }
    return Promise.resolve('SharePoint');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath.startsWith('doc_')) {
      const parts = propertyPath.split('_');
      const index = parseInt(parts[1]);
      const field = parts[2];
      
      if (!isNaN(index)) {
        const documents = this.getCurrentDocuments();
        if (index < documents.length) {
          if (field === 'title') {
            documents[index].title = newValue;
          } else if (field === 'author') {
            documents[index].author = newValue;
          } else if (field === 'date') {
            documents[index].date = newValue;
          } else if (field === 'iconType') {
            documents[index].iconType = newValue;
          }
          this.properties.documents = JSON.stringify(documents);
          this.render();
        }
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const documents = this.getCurrentDocuments();
    const expandedDocuments = this.getExpandedDocuments();

    const documentGroups: IPropertyPaneGroup[] = [];

    documents.forEach((doc, index) => {
      const isExpanded = expandedDocuments.includes(index);
      
      documentGroups.push({
        groupName: `Document ${index + 1} - ${doc.title || 'Untitled'}`,
        groupFields: [
          PropertyPaneButton(`expand_${index}`, {
            text: isExpanded ? '🔽 Hide Details' : '▶️ Show Details',
            buttonType: PropertyPaneButtonType.Normal,
            onClick: () => this.toggleDocumentExpansion(index)
          })
        ]
      });

      if (isExpanded) {
        documentGroups.push({
          groupName: `📝 Document ${index + 1} Settings`,
          groupFields: [
            PropertyPaneTextField(`doc_${index}_title`, {
              label: 'Document Title',
              value: doc.title,
              placeholder: 'Enter document title'
            }),
            PropertyPaneTextField(`doc_${index}_author`, {
              label: 'Author',
              value: doc.author,
              placeholder: 'Enter author name'
            }),
            PropertyPaneTextField(`doc_${index}_date`, {
              label: 'Date',
              value: doc.date,
              placeholder: 'Enter date'
            }),
            PropertyPaneDropdown(`doc_${index}_iconType`, {
              label: 'Document Type',
              options: [
                { key: 'word', text: 'Word Document' },
                { key: 'pdf', text: 'PDF Document' },
                { key: 'video', text: 'Video File' },
                { key: 'custom', text: 'Custom Icon' }
              ],
              selectedKey: doc.iconType
            })
          ]
        });

        documentGroups.push({
          groupName: `📤 Document ${index + 1} Upload`,
          groupFields: [
            PropertyPaneButton(`upload_icon_${index}`, {
              text: this._isUploading ? 'Uploading...' : '📷 Upload Icon',
              buttonType: PropertyPaneButtonType.Normal,
              disabled: this._isUploading,
              onClick: () => this.uploadIcon(index)
            }),
            PropertyPaneButton(`upload_doc_${index}`, {
              text: this._isUploading ? 'Uploading...' : '📎 Upload Document',
              buttonType: PropertyPaneButtonType.Normal,
              disabled: this._isUploading,
              onClick: () => this.uploadDocument(index)
            }),
            PropertyPaneButton(`delete_${index}`, {
              text: '🗑️ Delete Document',
              buttonType: PropertyPaneButtonType.Primary,
              onClick: () => this.deleteDocument(index)
            })
          ]
        });
      }
    });

    return {
      pages: [
        {
          header: { description: 'Configure HR Documents Display' },
          groups: [
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('title', { label: 'Section Title' }),
                PropertyPaneSlider('columnsPerRow', {
                  label: 'Columns per Row',
                  min: 2, max: 6, value: this.properties.columnsPerRow || 4,
                  showValue: true, step: 1
                }),
                PropertyPaneToggle('showDate', {
                  label: 'Show Author & Date',
                  checked: this.properties.showDate !== false,
                  onText: 'Yes', offText: 'No'
                })
              ]
            },
            {
              groupName: 'Document Management',
              groupFields: [
                PropertyPaneButton('addDocument', {
                  text: '+ Add New Document',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.addNewDocument
                })
              ]
            },
            ...documentGroups
          ]
        }
      ]
    };
  }
}
