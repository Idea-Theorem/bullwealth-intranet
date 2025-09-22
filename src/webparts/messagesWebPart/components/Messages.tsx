/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { 
  Stack, 
  Spinner, 
  MessageBar, 
  MessageBarType,
  DefaultButton
} from '@fluentui/react';
import { IMessagesProps, IMessageItem, IGroupedMessages, ViewMode } from './IMessagesProps';
import CardView from './CardView';
import DetailView from './DetailView';
import styles from './Messages.module.scss';

const Messages: React.FC<IMessagesProps> = (props) => {
  const [messages, setMessages] = useState<IMessageItem[]>([]);
  const [groupedMessages, setGroupedMessages] = useState<IGroupedMessages>({});
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [viewMode, setViewMode] = useState<ViewMode>(ViewMode.CardView);
  const [selectedMessageId, setSelectedMessageId] = useState<number | null>(null);

  useEffect(() => {
    loadMessages();
  }, []);

  // Move processImageField outside of component scope to fix 'this' reference
  const processImageField = (imageField: any, contextUrl: string): any => {
    console.log('Processing image field:', imageField);
    
    if (!imageField) {
      console.log('No image field provided');
      return null;
    }

    try {
      let imageData = imageField;
      
      // If it's a string, try to parse it as JSON
      if (typeof imageField === 'string') {
        try {
          imageData = JSON.parse(imageField);
          console.log('Parsed JSON imageData:', imageData);
        } catch (parseError) {
          console.log('Failed to parse as JSON, treating as string:', parseError);
          // If it's a direct URL string
          if (imageField.startsWith('http')) {
            return { Url: imageField };
          }
          return null;
        }
      }

      // Handle array format (SharePoint Online sometimes returns arrays)
      if (Array.isArray(imageData)) {
        console.log('Image data is array:', imageData);
        if (imageData.length > 0) {
          const firstImage = imageData[0];
          return {
            Url: firstImage.serverUrl || 
                 firstImage.serverRelativeUrl || 
                 firstImage.thumbnailUrl || 
                 firstImage.url ||
                 firstImage.Url ||
                 (firstImage.id ? `${contextUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(firstImage.id)}` : null)
          };
        }
      }

      // Handle object format
      if (typeof imageData === 'object' && imageData !== null) {
        console.log('Image data is object:', imageData);
        return {
          Url: imageData.serverUrl || 
               imageData.serverRelativeUrl || 
               imageData.thumbnailUrl ||
               imageData.url ||
               imageData.Url ||
               (imageData.id ? `${contextUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(imageData.id)}` : null)
        };
      }

      console.log('Unable to process image data format');
      return null;
    } catch (error) {
      console.error('Error processing image field:', error);
      return null;
    }
  };

  const loadMessages = async (): Promise<void> => {
    try {
      setLoading(true);
      setError('');
      
      console.log('Loading from list: Archived-Messages');
      console.log('Site URL:', props.context.pageContext.web.absoluteUrl);
      
      // Build the API URL for Archived-Messages list (removed Category from select)
      const apiUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?$expand=Author&$select=Id,Title,Content,PublishedDate,FeaturedImage,Author/Title&$orderby=Id desc&$top=100`;
      
      console.log('API URL:', apiUrl);

      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      console.log('Response status:', response.status);

      if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error:', errorText);
        setError(`Error loading messages: HTTP ${response.status}. Make sure 'Archived-Messages' list exists.`);
        return;
      }

      const data = await response.json();
      console.log('Raw API response:', data);
      console.log('Number of items received:', data.value ? data.value.length : 0);
      
      if (data.value && data.value.length > 0) {
        console.log('First item raw data:', data.value[0]);
        console.log('First item FeaturedImage field:', data.value[0].FeaturedImage);
      }

      const items: IMessageItem[] = data.value.map((item: any) => {
        console.log(`\n--- Processing Item ${item.Id} ---`);
        console.log('Title:', item.Title);
        console.log('Raw FeaturedImage:', item.FeaturedImage);
        
        // Fixed: Call processImageField as a regular function, not with 'this'
        const processedImage = processImageField(item.FeaturedImage, props.context.pageContext.web.absoluteUrl);
        console.log('Processed image:', processedImage);
        
        const messageItem: IMessageItem = {
          Id: item.Id,
          Title: item.Title || 'Untitled',
          Content: item.Content || '',
          PublishedDate: item.PublishedDate || item.Created || new Date().toISOString(),
          Category: 'General', // Default category since it's removed from list
          Year: new Date(item.PublishedDate || item.Created || new Date()).getFullYear(),
          FeaturedImage: processedImage,
          Author: item.Author || { Title: 'Unknown' }
        };
        
        console.log('Final message item:', messageItem);
        return messageItem;
      });
      
      console.log('\n--- Final processed items ---');
      console.log('Total items processed:', items.length);
      items.forEach(item => {
        console.log(`Item ${item.Id}: Image = ${item.FeaturedImage ? item.FeaturedImage.Url : 'No image'}`);
      });
      
      setMessages(items);
      groupMessagesByYear(items);
      
    } catch (err) {
      console.error('Error in loadMessages:', err);
      setError('Error loading messages: ' + (err as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const groupMessagesByYear = (items: IMessageItem[]): void => {
    const grouped = items.reduce((acc: IGroupedMessages, item: IMessageItem) => {
      const year = item.Year.toString();
      if (!acc[year]) {
        acc[year] = [];
      }
      acc[year].push(item);
      return acc;
    }, {});
    
    console.log('Grouped messages by year:', grouped);
    setGroupedMessages(grouped);
  };

  const handleReadMore = (messageId: number): void => {
    setSelectedMessageId(messageId);
    setViewMode(ViewMode.DetailView);
  };

  const handleBackToCards = (): void => {
    setViewMode(ViewMode.CardView);
    setSelectedMessageId(null);
  };

  const handleRetry = (): void => {
    loadMessages();
  };

  if (loading) {
    return (
      <div className={styles.messages}>
        <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
          <Spinner label="Loading messages from Archived-Messages..." />
        </Stack>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.messages}>
        <MessageBar messageBarType={MessageBarType.error}>
          <div>
            <div><strong>Error:</strong> {error}</div>
            <div style={{ marginTop: '10px' }}>
              <strong>Troubleshooting:</strong>
              <ul>
                <li>Verify 'Archived-Messages' list exists</li>
                <li>Check list has 'FeaturedImage' column</li>
                <li>Ensure you have read permissions</li>
              </ul>
            </div>
          </div>
        </MessageBar>
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <DefaultButton text="Retry Loading" onClick={handleRetry} iconProps={{ iconName: 'Refresh' }} />
        </div>
      </div>
    );
  }

  if (messages.length === 0) {
    return (
      <div className={styles.messages}>
        <MessageBar messageBarType={MessageBarType.info}>
          No messages found in 'Archived-Messages' list. Please add some items.
        </MessageBar>
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <DefaultButton text="Reload" onClick={handleRetry} iconProps={{ iconName: 'Refresh' }} />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.messages}>
      {viewMode === ViewMode.CardView && (
        <CardView 
          title={props.title}
          groupedMessages={groupedMessages}
          onReadMore={handleReadMore}
          columnsPerRow={props.columnsPerRow || 4}
        />
      )}
      
      {viewMode === ViewMode.DetailView && selectedMessageId && (
        <DetailView 
          messageId={selectedMessageId}
          messages={messages}
          onBack={handleBackToCards}
          context={props.context}
          listName="Archived-Messages"
        />
      )}
    </div>
  );
};

export default Messages;
