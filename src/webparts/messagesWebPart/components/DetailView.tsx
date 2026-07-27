/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { 
  Stack, 
  Text, 
  Spinner, 
  MessageBar, 
  MessageBarType,
  IconButton,
  Separator,
  DefaultButton
} from '@fluentui/react';
import { IMessageItem } from './IMessagesProps';
import styles from './Messages.module.scss';

export interface IDetailViewProps {
  messageId: number;
  messages: IMessageItem[];
  onBack: () => void;
  context: WebPartContext;
  listName: string;
}

const DetailView: React.FC<IDetailViewProps> = ({ 
  messageId, 
  messages, 
  onBack, 
  context, 
  listName 
}) => {
  const [messageItem, setMessageItem] = useState<IMessageItem | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    loadMessageDetail();
  }, [messageId]);

  const loadMessageDetail = async (): Promise<void> => {
    try {
      setLoading(true);
      
      // First try to find in existing messages
      const existingMessage = messages.find(m => m.Id === messageId);
      if (existingMessage) {
        setMessageItem(existingMessage);
        setLoading(false);
        return;
      }

      // If not found, fetch from SharePoint (removed Category from select)
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${messageId})?$expand=Author&$select=Id,Title,Content,PublishedDate,FeaturedImage,Author/Title`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const item: any = await response.json();
        
        let featuredImage = null;
        if (item.FeaturedImage) {
          if (typeof item.FeaturedImage === 'string') {
            try {
              featuredImage = JSON.parse(item.FeaturedImage);
            } catch {
              featuredImage = { Url: item.FeaturedImage };
            }
          } else {
            featuredImage = item.FeaturedImage;
          }
        }
        
        const messageData: IMessageItem = {
          ...item,
          Year: new Date(item.PublishedDate).getFullYear(),
          Category: 'General', // Default since no category field
          FeaturedImage: featuredImage
        };
        
        setMessageItem(messageData);
      } else {
        setError('Failed to load message details');
      }
    } catch (err) {
      setError('Error loading message: ' + (err as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  };

  const handlePrint = (): void => {
    window.print();
  };

  const handleShare = (): void => {
    if (navigator.share && messageItem) {
      navigator.share({
        title: messageItem.Title,
        text: messageItem.Title,
        url: window.location.href
      });
    } else {
      // Fallback to copy URL
      navigator.clipboard.writeText(window.location.href);
    }
  };

  const handleDownload = (): void => {
    if (!messageItem) return;
    
    const element = document.createElement('a');
    const file = new Blob([`
      <html>
        <head><title>${messageItem.Title}</title></head>
        <body>
          <h1>${messageItem.Title}</h1>
          <p><strong>Published:</strong> ${formatDate(messageItem.PublishedDate)}</p>
          <div>${messageItem.Content}</div>
        </body>
      </html>
    `], {type: 'text/html'});
    
    element.href = URL.createObjectURL(file);
    element.download = `${messageItem.Title.replace(/[^a-z0-9]/gi, '_')}.html`;
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  if (loading) {
    return (
      <div className={styles.detailView}>
        <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
          <Spinner label="Loading message details..." />
        </Stack>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.detailView}>
        <div className={styles.detailContainer}>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
          <div className={styles.centerContent}>
            <DefaultButton text="Back to Messages" onClick={onBack} />
          </div>
        </div>
      </div>
    );
  }

  if (!messageItem) {
    return (
      <div className={styles.detailView}>
        <div className={styles.detailContainer}>
          <MessageBar messageBarType={MessageBarType.warning}>
            Message not found
          </MessageBar>
          <div className={styles.centerContent}>
            <DefaultButton text="Back to Messages" onClick={onBack} />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.detailView}>
      <div className={styles.detailContainer}>
        {/* Header with Back Button */}
        <Stack horizontal verticalAlign="center" className={styles.detailHeader}>
          <IconButton
            iconProps={{ iconName: 'Back' }}
            title="Go back"
            onClick={onBack}
            className={styles.backButton}
          />
          <Text variant="medium" className={styles.breadcrumb}>
            Messages / Details
          </Text>
        </Stack>

        {/* Centered Content */}
        <div className={styles.centerContent}>
          {/* Document Header */}
          <div className={styles.documentHeader}>
            <div className={styles.statusBadge}>
              Published on {formatDate(messageItem.PublishedDate)}
            </div>
            
            <Text as="h1" variant="xxLarge" className={styles.detailTitle}>
              {messageItem.Title}
            </Text>
          </div>

          <Separator className={styles.separator} />

          {/* Document Content */}
          <div className={styles.documentContent}>
            <div className={styles.contentBody}>
              <div 
                dangerouslySetInnerHTML={{ __html: messageItem.Content }}
                className={styles.htmlContent}
              />
            </div>

            {/* Author Information */}
            <div className={styles.authorSection}>
              <Text variant="medium" className={styles.authorLabel}>
                Published by:
              </Text>
              <Text variant="medium" className={styles.authorName}>
                {messageItem.Author?.Title || 'System Administrator'}
              </Text>
            </div>
          </div>

          {/* Action Buttons */}
          <div className={styles.actionButtons}>
            <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="center">
              <DefaultButton
                iconProps={{ iconName: 'Print' }}
                text="Print"
                onClick={handlePrint}
                className={styles.actionButton}
              />
              <DefaultButton
                iconProps={{ iconName: 'Share' }}
                text="Share"
                onClick={handleShare}
                className={styles.actionButton}
              />
              <DefaultButton
                iconProps={{ iconName: 'Download' }}
                text="Download"
                onClick={handleDownload}
                className={styles.actionButton}
              />
            </Stack>
          </div>
        </div>
      </div>
    </div>
  );
};

export default DetailView;
