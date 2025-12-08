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
  };

  private handleViewAttachment = (): void => {
    const { messageData } = this.props;

    if (messageData.AttachmentUrl && typeof messageData.AttachmentUrl === 'string') {
      window.open(messageData.AttachmentUrl, '_blank', 'noopener,noreferrer');
      return;
    }

    if (messageData.AttachmentServerRelativeUrl && typeof messageData.AttachmentServerRelativeUrl === 'string') {
      const fullUrl =
        `${window.location.protocol}//${window.location.host}${messageData.AttachmentServerRelativeUrl}`;
      window.open(fullUrl, '_blank', 'noopener,noreferrer');
      return;
    }

    alert('No attachment available for this message.');
  };

  public render(): React.ReactElement<IDetailedViewProps> {
    const { messageData, onBack } = this.props;

    if (!messageData) {
      return (
        <div className={styles.detailPage}>
          <div className={styles.container}>
            <div className={styles.error}>
              <h2>Message not found</h2>
              <button
                className={styles.backHomeButton}
                onClick={onBack}
                type="button"
              >
                ← Back to Home
              </button>
            </div>
          </div>
        </div>
      );
    }

    const publishedDate = messageData.PublishedDate || messageData.Created;

    const hasAttachment =
      !!messageData.AttachmentUrl ||
      !!messageData.AttachmentServerRelativeUrl;

    return (
      <div className={styles.detailPage}>
        <div className={styles.container}>
          <div className={styles.header}>
            <button
              className={styles.backHomeButton}
              onClick={onBack}
              type="button"
            >
              ← Back to Home
            </button>
          </div>

          <div className={styles.content}>
            {publishedDate && (
              <div className={styles.publishedBadge}>
                PUBLISHED ON {this.formatDate(publishedDate).toUpperCase()}
              </div>
            )}

            <h1 className={styles.title}>
              {messageData.Title}
            </h1>

            {messageData.FeaturedImageUrl && (
              <div className={styles.imageContainer}>
                <img
                  src={messageData.FeaturedImageUrl}
                  alt="Featured image"
                  className={styles.featuredImage}
                  onError={(e) => {
                    (e.target as HTMLImageElement).style.display = 'none';
                  }}
                />
              </div>
            )}

            <div
              className={styles.messageContent}
              dangerouslySetInnerHTML={{ __html: messageData.Content }}
            />

            {hasAttachment && (
              <button
                className={styles.viewAttachmentButton}
                onClick={this.handleViewAttachment}
                type="button"
              >
                View Attachment
              </button>
            )}

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
