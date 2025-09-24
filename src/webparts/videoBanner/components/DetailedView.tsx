import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DetailedView.module.scss';

export interface IDetailedViewProps {
  messageData: any;
  onBack: () => void;
  context: WebPartContext;
}

export default class DetailedView extends React.Component<IDetailedViewProps, {}> {
  
  private formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-GB', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  }

  public render(): React.ReactElement<IDetailedViewProps> {
    const { messageData, onBack } = this.props;

    if (!messageData) {
      return (
        <div className={styles.detailPage}>
          <div className={styles.container}>
            <div className={styles.error}>
              <h2>Message not found</h2>
              <button className={styles.backHomeButton} onClick={onBack}>
                ← Back to Home
              </button>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.detailPage}>
        <div className={styles.container}>
          {/* Header with Back to Home Button */}
          <div className={styles.header}>
            <button className={styles.backHomeButton} onClick={onBack}>
              ← Back to Home
            </button>
          </div>

          {/* Content Area */}
          <div className={styles.content}>
            {/* Published Badge */}
            <div className={styles.publishedBadge}>
              PUBLISHED ON {this.formatDate(messageData.PublishedDate || messageData.Created).toUpperCase()}
            </div>

            {/* Title */}
            <h1 className={styles.title}>
              {messageData.Title}
            </h1>

            {/* Featured Image */}
            {messageData.FeaturedImageUrl && (
              <div className={styles.imageContainer}>
                <img 
                  src={messageData.FeaturedImageUrl}
                  alt="Featured image"
                  className={styles.featuredImage}
                  onError={(e) => {
                    console.log('Detail view image failed to load');
                    (e.target as HTMLImageElement).style.display = 'none';
                  }}
                />
              </div>
            )}
            
            {/* Message Content */}
            <div 
              className={styles.messageContent}
              dangerouslySetInnerHTML={{ __html: messageData.Content }}
            />

            {/* Author Information */}
            <div className={styles.authorInfo}>
              <strong>Published by:</strong>
              <span className={styles.authorName}>
                {messageData.Author?.Title || 'Unknown Author'}
              </span>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
