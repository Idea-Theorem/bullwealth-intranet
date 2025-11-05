/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './VideoBanner.module.scss';
import { IVideoBannerProps } from './IVideoBannerProps';
import DetailedView from './DetailedView';

export interface IVideoBannerState {
  showModal: boolean;
  showDetailView: boolean;
  latestMessage: any;
  loading: boolean;
}

export default class VideoBanner extends React.Component<IVideoBannerProps, IVideoBannerState> {
  
  constructor(props: IVideoBannerProps) {
    super(props);
    this.state = {
      showModal: false,
      showDetailView: false,
      latestMessage: null,
      loading: true
    };
  }

  public componentDidMount(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.loadLatestMessage();
  }

  private getImageUrl = async (itemId: number, imageFieldValue: any, siteUrl: string): Promise<string | null> => {
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

        // Method 1: RenderListDataAsStream (most reliable)
        try {
          const renderApiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/RenderListDataAsStream`;
          
          const renderResponse = await this.props.context.spHttpClient.post(
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
          
          const fieldResponse = await this.props.context.spHttpClient.get(
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
          
          const attachmentResponse = await this.props.context.spHttpClient.get(
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
  }

  private loadLatestMessage = async (): Promise<void> => {
    try {
      this.setState({ loading: true });
      
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      console.log('🔍 Loading latest published message from Archived-Messages...');
      
      const today = new Date();
      const todayISOString = today.toISOString();
      
      console.log('📅 Current date for filtering:', todayISOString);
      
      // Query latest message
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?` +
        `$expand=Author&` +
        `$select=Id,Title,Content,PublishedDate,Created,FeaturedImage,NewsletterVideo,Author/Title&` +
        `$filter=PublishedDate le datetime'${todayISOString}'&` +
        `$orderby=PublishedDate desc&` +
        `$top=1`;
      
      console.log('🔗 API URL:', apiUrl);
      
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        console.log('📊 API Response:', data);
        
        if (data.value && data.value.length > 0) {
          const latestItem = data.value[0];
          console.log('📄 Latest published item:', {
            Id: latestItem.Id,
            Title: latestItem.Title,
            PublishedDate: latestItem.PublishedDate
          });
          
          // ✅ Get image URL using proven working method
          let featuredImageUrl: string | null = null;
          
          if (latestItem.FeaturedImage) {
            try {
              featuredImageUrl = await this.getImageUrl(latestItem.Id, latestItem.FeaturedImage, siteUrl);
            } catch (imageError) {
              console.error(`❌ Image processing failed for item ${latestItem.Id}:`, imageError);
            }
          }
          
          // ✅ Get video URL
          let videoUrl: string | null = null;
          
          if (latestItem.NewsletterVideo) {
            console.log('🎬 Raw video data:', latestItem.NewsletterVideo);
            
            if (typeof latestItem.NewsletterVideo === 'string') {
              videoUrl = latestItem.NewsletterVideo.trim();
            } else if (typeof latestItem.NewsletterVideo === 'object') {
              videoUrl = latestItem.NewsletterVideo.Url || latestItem.NewsletterVideo.url || null;
              if (videoUrl && typeof videoUrl === 'string') {
                videoUrl = videoUrl.trim();
              }
            }
            
            console.log('✅ Extracted video URL:', videoUrl);
          }
          
          console.log('🎯 FINAL RESULT:', {
            featuredImageUrl,
            videoUrl,
            hasImage: !!featuredImageUrl,
            hasVideo: !!videoUrl
          });
          
          this.setState({ 
            latestMessage: {
              ...latestItem,
              FeaturedImageUrl: featuredImageUrl,
              VideoUrl: videoUrl
            }
          });
        } else {
          console.log('⚠️ No published messages found');
          this.setState({ latestMessage: null });
        }
      } else {
        console.error('❌ API Error - Status:', response.status);
        const errorText = await response.text();
        console.error('Error details:', errorText);
      }
    } catch (error) {
      console.error('💥 Load error:', error);
    } finally {
      this.setState({ loading: false });
    }
  }

  private handlePlayClick = (): void => {
    const { latestMessage } = this.state;
    const videoUrl: string | undefined = latestMessage?.VideoUrl;
    
    if (!videoUrl || videoUrl.trim() === '') {
      console.log('⚠️ No video URL available');
      return;
    }

    this.setState({ showModal: true });
  }

  private handleReadMore = (): void => {
    console.log('📖 Opening detail view');
    this.setState({ showDetailView: true });
  }

  private handleBackFromDetail = (): void => {
    console.log('🏠 Back to home');
    this.setState({ showDetailView: false });
  }

  private handleCloseModal = (): void => {
    this.setState({ showModal: false });
  }

  private isYouTubeUrl = (url: string): boolean => {
    return !!(url && (url.indexOf('youtube.com') !== -1 || url.indexOf('youtu.be') !== -1));
  }

  private isStreamUrl = (url: string): boolean => {
    return !!(url && (url.indexOf('microsoftstream.com') !== -1 || url.indexOf('stream.microsoft.com') !== -1));
  }

  private getEmbedUrl = (url: string): string => {
    if (this.isYouTubeUrl(url)) {
      const videoId: string = this.extractYouTubeId(url);
      return `https://www.youtube.com/embed/${videoId}?autoplay=1&rel=0`;
    }
    
    if (this.isStreamUrl(url)) {
      return url.replace('/watch/', '/embed/');
    }
    
    return url;
  }

  private extractYouTubeId = (url: string): string => {
    const match = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([^&\n?#]+)/);
    return match ? match[1] : '';
  }

  private getDefaultThumbnail = (): string => {
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 450"%3E%3Crect width="800" height="450" fill="%23f0f0f0"/%3E%3Ctext x="400" y="200" text-anchor="middle" fill="%23999999" font-size="24" font-family="Arial"%3ELatest News%3C/text%3E%3Ctext x="400" y="250" text-anchor="middle" fill="%23999999" font-size="16" font-family="Arial"%3EImage Preview%3C/text%3E%3C/svg%3E';
  }

  private getDefaultBackground = (): string => {
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080"%3E%3Cdefs%3E%3ClinearGradient id="bg" x1="0%25" y1="0%25" x2="100%25" y2="100%25"%3E%3Cstop offset="0%25" style="stop-color:%234a90e2;stop-opacity:1" /%3E%3Cstop offset="100%25" style="stop-color:%237b68ee;stop-opacity:1" /%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width="1920" height="1080" fill="url(%23bg)" /%3E%3C/svg%3E';
  }

  private formatDate = (dateString: string | null | undefined): string | null => {
    if (!dateString) {
      return null;
    }

    try {
      const date = new Date(dateString);
      if (isNaN(date.getTime())) {
        return null;
      }

      return `Published ${date.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'short',
        year: 'numeric'
      })}`;
    } catch (error) {
      console.error('❌ Date error:', error);
      return null;
    }
  }

  private stripHtmlTags = (html: string): string => {
    if (!html) return '';
    const div: HTMLDivElement = document.createElement('div');
    div.innerHTML = html;
    return div.textContent || div.innerText || '';
  }

  private truncateText = (text: string, maxLength: number = 280): string => {
    if (!text) return '';
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '...';
  }

  public render(): React.ReactElement<IVideoBannerProps> {
    const { showModal, showDetailView, latestMessage, loading } = this.state;

    if (showDetailView) {
      return (
        <DetailedView 
          messageData={latestMessage}
          onBack={this.handleBackFromDetail}
          context={this.props.context}
        />
      );
    }

    const displayTitle: string = latestMessage ? latestMessage.Title : 'Latest News';
    const displayMessage: string = latestMessage 
      ? this.truncateText(this.stripHtmlTags(latestMessage.Content))
      : 'Stay updated with our latest announcements and news.';
    
    const displayDate: string | null = latestMessage 
      ? (this.formatDate(latestMessage.PublishedDate) || this.formatDate(latestMessage.Created))
      : null;
    
    const displayBackground: string = this.getDefaultBackground();
    const rightSideImage: string = latestMessage?.FeaturedImageUrl || this.getDefaultThumbnail();
    const videoUrl: string | undefined = latestMessage?.VideoUrl;
    const hasVideo: boolean = !!(videoUrl && videoUrl.trim() !== '');

    const backgroundStyle = {
      backgroundImage: `linear-gradient(rgba(0, 0, 0, 0.4), rgba(0, 0, 0, 0.4)), url(${displayBackground})`
    };

    if (loading) {
      return (
        <div className={styles.videoBanner} style={{ background: '#f5f5f5' }}>
          <div className={styles.overlay}>
            <div className={styles.contentWrapper}>
              <div className={styles.textContent}>
                <div className={styles.loadingSpinner}>
                  <div className={styles.spinner}></div>
                  <p>Loading latest news...</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.videoBanner} style={backgroundStyle}>
        <div className={styles.overlay}>
          <div className={styles.contentWrapper}>
            <div className={styles.textContent}>
              <h2 className={styles.title}>{displayTitle}</h2>
              
              {displayDate && (
                <div className={styles.publishedDate}>
                  {displayDate}
                </div>
              )}
              
              <p className={styles.message}>{displayMessage}</p>
              
              <button 
                className={styles.readMoreButton}
                onClick={this.handleReadMore}
                type="button"
              >
                READ MORE
              </button>
            </div>
            
            <div className={styles.videoSection}>
              <div className={styles.thumbnailContainer}>
                <img 
                  src={rightSideImage}
                  alt="Latest news featured image"
                  className={styles.thumbnail}
                  onError={(e) => {
                    console.log('⚠️ Image load failed, using default');
                    (e.target as HTMLImageElement).src = this.getDefaultThumbnail();
                  }}
                />
                {hasVideo && (
                  <button 
                    className={styles.playButton}
                    onClick={this.handlePlayClick}
                    aria-label="Play video"
                    type="button"
                  >
                    <svg 
                      className={styles.playIcon} 
                      viewBox="0 0 80 80" 
                      fill="none" 
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <circle cx="40" cy="40" r="40" fill="white" fillOpacity="0.9"/>
                      <path d="M32 28L52 40L32 52V28Z" fill="#333333"/>
                    </svg>
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {showModal && hasVideo && videoUrl && (
          <div className={styles.modal} onClick={this.handleCloseModal}>
            <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
              <button 
                className={styles.closeButton}
                onClick={this.handleCloseModal}
                aria-label="Close video"
                type="button"
              >
                ×
              </button>
              
              <div className={styles.videoContainer}>
                {this.isYouTubeUrl(videoUrl) || this.isStreamUrl(videoUrl) ? (
                  <iframe 
                    className={styles.videoFrame}
                    src={this.getEmbedUrl(videoUrl)}
                    allowFullScreen
                    allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                    title="Newsletter Video"
                  />
                ) : (
                  <video 
                    className={styles.modalVideo}
                    controls
                    autoPlay
                    src={videoUrl}
                  >
                    Your browser does not support the video tag.
                  </video>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }
}
