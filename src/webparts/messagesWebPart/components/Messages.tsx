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

  // Enhanced URL parameter handling for VideoBanner integration
  useEffect(() => {
    // Check for URL parameters to show specific message detail
    const urlParams = new URLSearchParams(window.location.search);
    const messageId = urlParams.get('messageId');
    const view = urlParams.get('view');
    
    console.log('🔗 URL Parameters detected:', { messageId, view });
    
    if (messageId && view === 'detail') {
      const id = parseInt(messageId);
      console.log('📖 Opening detail view for message ID:', id);
      setSelectedMessageId(id);
      setViewMode(ViewMode.DetailView);
    }
    
    loadMessages();
  }, []);

  // Enhanced image URL extraction with multiple SharePoint approaches
  const getImageUrl = async (itemId: number, imageFieldValue: any, siteUrl: string): Promise<string | null> => {
    console.log(`🖼️ Getting image URL for item ${itemId}`);
    console.log('Raw image field:', imageFieldValue);

    if (!imageFieldValue) {
      console.log('No image field value provided');
      return null;
    }

    try {
      let imageData: any;

      // Parse JSON if string
      if (typeof imageFieldValue === 'string') {
        try {
          imageData = JSON.parse(imageFieldValue);
        } catch (parseError) {
          console.log('Failed to parse image field as JSON');
          return null;
        }
      } else {
        imageData = imageFieldValue;
      }

      if (imageData && imageData.fileName) {
        console.log('Found fileName:', imageData.fileName);

        // Method 1: SharePoint RenderListDataAsStream (most reliable for Image columns)
        try {
          const renderApiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/RenderListDataAsStream`;
          
          const renderResponse = await props.context.spHttpClient.post(
            renderApiUrl,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata'
              },
              body: JSON.stringify({
                parameters: {
                  ViewXml: `<View><Query><Where><Eq><FieldRef Name='ID'/><Value Type='Number'>${itemId}</Value></Eq></Where></Query></View>`,
                  RenderOptions: 1
                }
              })
            }
          );

          if (renderResponse.ok) {
            const renderData = await renderResponse.json();
            console.log('🎨 RenderListDataAsStream success:', renderData);
            
            if (renderData.Row && renderData.Row.length > 0) {
              const row = renderData.Row[0];
              if (row.FeaturedImage) {
                const imgMatch = row.FeaturedImage.match(/src="([^"]+)"/);
                if (imgMatch) {
                  const imageUrl = imgMatch[1];
                  console.log('✅ Found rendered image URL:', imageUrl);
                  return imageUrl;
                }
              }
            }
          }
        } catch (renderError) {
          console.log('⚠️ RenderListDataAsStream failed, trying alternatives');
        }

        // Method 2: FieldValuesAsHtml approach
        try {
          const fieldHtmlUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items(${itemId})/FieldValuesAsHtml/FeaturedImage`;
          
          const fieldResponse = await props.context.spHttpClient.get(
            fieldHtmlUrl,
            SPHttpClient.configurations.v1
          );

          if (fieldResponse.ok) {
            const fieldData = await fieldResponse.json();
            console.log('📄 FieldValuesAsHtml response:', fieldData);
            
            if (fieldData.value) {
              const imgMatch = fieldData.value.match(/src="([^"]+)"/);
              if (imgMatch) {
                console.log('✅ Found HTML field image URL:', imgMatch[1]);
                return imgMatch[1];
              }
            }
          }
        } catch (fieldError) {
          console.log('⚠️ FieldValuesAsHtml failed');
        }

        // Method 3: Direct URL construction with validation
        const possibleUrls = [
          `${siteUrl}/Lists/Archived-Messages/Attachments/${itemId}/${imageData.fileName}`,
          `${siteUrl}/SiteAssets/${imageData.fileName}`,
          `${siteUrl}/PublishingImages/${imageData.fileName}`,
          `${siteUrl}/_layouts/15/getpreview.ashx?path=${siteUrl}/Lists/Archived-Messages/Attachments/${itemId}/${imageData.fileName}`
        ];

        for (const testUrl of possibleUrls) {
          try {
            // Test if URL is accessible
            const response = await fetch(testUrl, { method: 'HEAD' });
            if (response.ok) {
              console.log('✅ Found working direct URL:', testUrl);
              return testUrl;
            }
          } catch {
            console.log(`❌ URL failed: ${testUrl}`);
          }
        }

        // Method 4: Attachment API approach
        try {
          const attachmentUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items(${itemId})/AttachmentFiles('${imageData.fileName}')/$value`;
          
          const attachmentResponse = await props.context.spHttpClient.get(
            attachmentUrl,
            SPHttpClient.configurations.v1
          );

          if (attachmentResponse.ok) {
            console.log('✅ Found as attachment:', attachmentUrl);
            return attachmentUrl;
          }
        } catch {
          console.log('⚠️ Attachment API failed');
        }
      }

      console.log('❌ No working image URL found for item', itemId);
      return null;

    } catch (error) {
      console.error('💥 Error getting image URL:', error);
      return null;
    }
  };

  const loadMessages = async (): Promise<void> => {
    try {
      setLoading(true);
      setError('');
      
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      console.log('🔍 Loading messages from Archived-Messages list');
      console.log('Site URL:', siteUrl);
      
      // Enhanced API call with better field selection
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?$expand=Author&$select=Id,Title,Content,PublishedDate,Created,FeaturedImage,Author/Title&$orderby=Created desc&$top=200`;
      
      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('❌ API Error:', errorText);
        setError(`Error loading messages: HTTP ${response.status}. Please check if 'Archived-Messages' list exists.`);
        return;
      }

      const data = await response.json();
      console.log(`📊 API Success: Received ${data.value?.length || 0} items`);
      
      if (!data.value || data.value.length === 0) {
        console.log('⚠️ No items found in Archived-Messages list');
        setMessages([]);
        setGroupedMessages({});
        return;
      }

      // Process items with enhanced image handling
      const processedItems: IMessageItem[] = [];
      
      for (const item of data.value) {
        console.log(`\n📝 Processing: "${item.Title}" (ID: ${item.Id})`);
        
        // Enhanced date handling
        const publishedDate = item.PublishedDate || item.Created || new Date().toISOString();
        const year = new Date(publishedDate).getFullYear();
        
        console.log(`📅 Date: ${publishedDate} → Year: ${year}`);
        
        // Get image URL with enhanced error handling
        let imageUrl: string | null = null;
        if (item.FeaturedImage) {
          try {
            imageUrl = await getImageUrl(item.Id, item.FeaturedImage, siteUrl);
          } catch (imageError) {
            console.error(`❌ Image processing failed for item ${item.Id}:`, imageError);
          }
        }
        
        const processedItem: IMessageItem = {
          Id: item.Id,
          Title: item.Title || 'Untitled Message',
          Content: item.Content || '',
          PublishedDate: publishedDate,
          Category: 'Newsletter', // Enhanced category
          Year: year,
          FeaturedImage: imageUrl ? { 
            Url: imageUrl, 
            Description: item.Title || 'Featured Image'
          } : undefined,
          Author: item.Author || { Title: 'System' }
        };
        
        processedItems.push(processedItem);
        
        console.log(`✅ Item processed: Image=${imageUrl ? '✓' : '✗'}`);
      }
      
      // Fixed: Use ES5 compatible array creation from Set
      const uniqueYears = processedItems.map(i => i.Year);
      const uniqueYearsSet = new Set(uniqueYears);
      const uniqueYearsArray: number[] = [];
      uniqueYearsSet.forEach(year => uniqueYearsArray.push(year));
      
      console.log('\n📈 PROCESSING SUMMARY:');
      console.log(`Total items processed: ${processedItems.length}`);
      console.log(`Items with images: ${processedItems.filter(i => i.FeaturedImage).length}`);
      console.log(`Years covered: ${uniqueYearsArray.sort().join(', ')}`);
      
      setMessages(processedItems);
      groupMessagesByYear(processedItems);
      
    } catch (err) {
      console.error('💥 Critical error in loadMessages:', err);
      setError('Critical error loading messages: ' + (err as Error).message);
    } finally {
      setLoading(false);
    }
  };

  // Enhanced grouping with better year handling
  const groupMessagesByYear = (items: IMessageItem[]): void => {
    const grouped = items.reduce((acc: IGroupedMessages, item: IMessageItem) => {
      const year = item.Year.toString();
      if (!acc[year]) {
        acc[year] = [];
      }
      acc[year].push(item);
      return acc;
    }, {});
    
    console.log('📊 Grouped by year:', Object.keys(grouped).map(year => `${year}: ${grouped[year].length} items`));
    setGroupedMessages(grouped);
  };

  const handleReadMore = (messageId: number): void => {
    console.log('📖 Opening detail view for message:', messageId);
    setSelectedMessageId(messageId);
    setViewMode(ViewMode.DetailView);
  };

  const handleBackToCards = (): void => {
    console.log('⬅️ Returning to card view');
    setViewMode(ViewMode.CardView);
    setSelectedMessageId(null);
    
    // Clear URL parameters when going back
    const url = new URL(window.location.href);
    url.searchParams.delete('messageId');
    url.searchParams.delete('view');
    window.history.replaceState({}, '', url.toString());
  };

  const handleRetry = (): void => {
    console.log('🔄 Retrying load messages...');
    loadMessages();
  };

  if (loading) {
    return (
      <div className={styles.messages}>
        <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
          <Spinner label="Loading newsletters and messages..." />
          <p style={{ marginTop: '10px', color: '#666' }}>
            Fetching data from Archived-Messages list...
          </p>
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
              <strong>Troubleshooting Tips:</strong>
              <ul style={{ marginTop: '5px' }}>
                <li>Verify 'Archived-Messages' list exists in this site</li>
                <li>Check you have read permissions to the list</li>
                <li>Ensure list has Title, Content, and FeaturedImage columns</li>
              </ul>
            </div>
          </div>
        </MessageBar>
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <DefaultButton 
            text="Retry Loading" 
            onClick={handleRetry} 
            iconProps={{ iconName: 'Refresh' }} 
          />
        </div>
      </div>
    );
  }

  if (messages.length === 0) {
    return (
      <div className={styles.messages}>
        <MessageBar messageBarType={MessageBarType.info}>
          No messages found in 'Archived-Messages' list. Add some newsletter items to get started!
        </MessageBar>
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <DefaultButton 
            text="Reload" 
            onClick={handleRetry} 
            iconProps={{ iconName: 'Refresh' }} 
          />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.messages}>
      {viewMode === ViewMode.CardView && (
        <CardView 
          title={props.title || 'Newsletters and Messages from CEO'}
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
