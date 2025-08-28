/* eslint-disable no-void */
import * as React from 'react';
import styles from './CompliancePolicies.module.scss';
import { ICompliancePoliciesProps, IPolicyDocument } from './ICompliancePoliciesProps';

export default class CompliancePolicies extends React.Component<ICompliancePoliciesProps> {
  
  private handleViewAll = (): void => {
    const { viewAllUrl } = this.props;
    if (viewAllUrl && viewAllUrl !== '#') {
      window.open(viewAllUrl, '_blank');
    }
  }

  private handleExport = (policy: IPolicyDocument): void => {
    if (policy.documentUrl && policy.documentUrl !== '#') {
      window.open(policy.documentUrl, '_blank');
    }
  }

  private handleShare = (policy: IPolicyDocument): void => {
    // Create share URL
    const shareUrl = policy.documentUrl || window.location.href;
    
    // Check if Web Share API is available
    if (navigator.share) {
      // eslint-disable-next-line no-void
      void navigator.share({
        title: policy.title,
        text: policy.description,
        url: shareUrl
      }).catch((error) => {
        console.log('Error sharing:', error);
        this.copyToClipboard(shareUrl);
      });
    } else {
      // Fallback to copying URL to clipboard
      this.copyToClipboard(shareUrl);
    }
  }

  private copyToClipboard = (text: string): void => {
    void navigator.clipboard.writeText(text).then(() => {
      alert('Link copied to clipboard!');
    }).catch((err) => {
      console.error('Failed to copy:', err);
      // Fallback for older browsers
      const textArea = document.createElement('textarea');
      textArea.value = text;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand('copy');
      document.body.removeChild(textArea);
      alert('Link copied to clipboard!');
    });
  }

  public render(): React.ReactElement<ICompliancePoliciesProps> {
    const {
      //sectionTitle,
      categoryTitle,
      viewAllText,
      policies,
      columnsPerRow,
      showExportButton,
      showShareButton
    } = this.props;
    
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns3;

    return (
      <div className={styles.compliancePolicies}>
        <div className={styles.container}>
          {/* <h2 className={styles.sectionTitle}>{sectionTitle}</h2> */}
          
          <div className={styles.categoryHeader}>
            <h3 className={styles.categoryTitle}>{categoryTitle}</h3>
            <button 
              className={styles.viewAllButton}
              onClick={this.handleViewAll}
            >
              {viewAllText}
            </button>
          </div>
          
          <div className={`${styles.policiesGrid} ${gridClassName}`}>
            {policies.map((policy) => (
              <div key={policy.id} className={styles.policyCard}>
                <h4 className={styles.policyTitle}>{policy.title}</h4>
                
                <div className={styles.policyDate}>
                  {policy.dateAdded || policy.dateModified || ''}
                </div>
                
                <p className={styles.policyDescription}>
                  {policy.description}
                </p>
                
                <div className={styles.policyActions}>
                  {showExportButton && (
                    <button 
                      className={styles.actionButton}
                      onClick={() => this.handleExport(policy)}
                      aria-label={`Export ${policy.title}`}
                    >
                      <svg 
                        className={styles.icon}
                        width="16" 
                        height="16" 
                        viewBox="0 0 16 16" 
                        fill="none"
                      >
                        <path 
                          d="M14 10v3.5a.5.5 0 01-.5.5h-11a.5.5 0 01-.5-.5V10M11 7L8 4m0 0L5 7m3-7v10" 
                          stroke="currentColor" 
                          strokeWidth="1.5" 
                          strokeLinecap="round" 
                          strokeLinejoin="round"
                        />
                      </svg>
                      Export
                    </button>
                  )}
                  
                  {showShareButton && (
                    <button 
                      className={styles.actionButton}
                      onClick={() => this.handleShare(policy)}
                      aria-label={`Share ${policy.title}`}
                    >
                      <svg 
                        className={styles.icon}
                        width="16" 
                        height="16" 
                        viewBox="0 0 16 16" 
                        fill="none"
                      >
                        <path 
                          d="M12 5.5C13.1046 5.5 14 4.60457 14 3.5C14 2.39543 13.1046 1.5 12 1.5C10.8954 1.5 10 2.39543 10 3.5C10 3.70873 10.0315 3.91019 10.0896 4.09896L6.41043 6.34896C6.01417 5.94119 5.46629 5.68571 4.85714 5.68571C3.69391 5.68571 2.75 6.62963 2.75 7.79286C2.75 8.95608 3.69391 9.9 4.85714 9.9C5.46629 9.9 6.01417 9.64452 6.41043 9.23675L10.0896 11.4868C10.0315 11.6755 10 11.877 10 12.0857C10 13.1903 10.8954 14.0857 12 14.0857C13.1046 14.0857 14 13.1903 14 12.0857C14 10.9812 13.1046 10.0857 12 10.0857C11.3908 10.0857 10.843 10.3412 10.4468 10.749L6.76758 8.49896C6.82568 8.31019 6.85714 8.10873 6.85714 7.9C6.85714 7.69127 6.82568 7.48981 6.76758 7.30104L10.4468 5.05104C10.843 5.45881 11.3908 5.71429 12 5.71429V5.5Z" 
                          fill="currentColor"
                        />
                      </svg>
                      Share
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }
}