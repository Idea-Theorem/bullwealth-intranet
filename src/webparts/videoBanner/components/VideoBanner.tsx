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

  private loadLatestMessage = async (): Promise<void> => {
    try {
      this.setState({ loading: true });
      
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      console.log('🔍 Loading latest published message from Archived-Messages...');
      
      // Get current date for filtering
      const today = new Date();
      const todayISOString = today.toISOString();
      
      console.log('📅 Current date for filtering:', todayISOString);
      
      // Enhanced query: Get messages where PublishedDate <= today, ordered by PublishedDate DESC
      // This ensures we get the most recently published message that should be visible now
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?$expand=Author&$select=Id,Title,Content,PublishedDate,Created,FeaturedImage,NewsletterVideo,Author/Title&$filter=PublishedDate le datetime'${todayISOString}'&$orderby=PublishedDate desc&$top=1`;
      
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
            PublishedDate: latestItem.PublishedDate,
            Created: latestItem.Created,
            daysFromNow: this.getDaysFromNow(latestItem.PublishedDate)
          });
          
          // Enhanced featured image processing
          let featuredImageUrl: string | null = null;
          
          if (latestItem.FeaturedImage) {
            try {
              let imageData;
              if (typeof latestItem.FeaturedImage === 'string') {
                imageData = JSON.parse(latestItem.FeaturedImage);
              } else {
                imageData = latestItem.FeaturedImage;
              }
              
              console.log('🖼️ Processing image data:', imageData);
              
              if (imageData && imageData.fileName) {
                // Try SharePoint RenderListDataAsStream for proper image URLs
                try {
                  const renderUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/RenderListDataAsStream`;
                  
                  const renderResponse = await this.props.context.spHttpClient.post(
                    renderUrl,
                    SPHttpClient.configurations.v1,
                    {
                      headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata'
                      },
                      body: JSON.stringify({
                        parameters: {
                          ViewXml: `<View><Query><Where><Eq><FieldRef Name='ID'/><Value Type='Number'>${latestItem.Id}</Value></Eq></Where></Query></View>`,
                          RenderOptions: 1
                        }
                      })
                    }
                  );

                  if (renderResponse.ok) {
                    const renderData = await renderResponse.json();
                    console.log('🎨 Render data:', renderData);
                    
                    if (renderData.Row && renderData.Row.length > 0) {
                      const row = renderData.Row[0];
                      if (row.FeaturedImage) {
                        const imgMatch = row.FeaturedImage.match(/src="([^"]+)"/);
                        if (imgMatch) {
                          featuredImageUrl = imgMatch[1];
                          console.log('✅ Found rendered image URL:', featuredImageUrl);
                        }
                      }
                    }
                  }
                } catch (renderError) {
                  console.log('⚠️ Render API failed, trying direct URLs');
                  
                  // Fallback to direct URL construction
                  const directUrls = [
                    `${siteUrl}/Lists/Archived-Messages/Attachments/${latestItem.Id}/${imageData.fileName}`,
                    `${siteUrl}/SiteAssets/${imageData.fileName}`,
                    `${siteUrl}/_layouts/15/getpreview.ashx?path=${siteUrl}/Lists/Archived-Messages/Attachments/${latestItem.Id}/${imageData.fileName}`
                  ];

                  for (const testUrl of directUrls) {
                    try {
                      const response = await fetch(testUrl, { method: 'HEAD' });
                      if (response.ok) {
                        featuredImageUrl = testUrl;
                        console.log('✅ Found working direct URL:', featuredImageUrl);
                        break;
                      }
                    } catch {
                      continue;
                    }
                  }
                }
              }
            } catch (error) {
              console.error('❌ Image processing error:', error);
            }
          }
          
          // Extract video URL
          let videoUrl: string | null = null;
          if (latestItem.NewsletterVideo) {
            if (typeof latestItem.NewsletterVideo === 'string') {
              videoUrl = latestItem.NewsletterVideo.trim();
            } else if (latestItem.NewsletterVideo.Url) {
              videoUrl = latestItem.NewsletterVideo.Url;
            }
          }
          
          console.log('🎬 Final processing result:', {
            featuredImageUrl,
            videoUrl,
            publishedDate: latestItem.PublishedDate,
            isCurrentlyPublished: new Date(latestItem.PublishedDate) <= new Date()
          });
          
          this.setState({ 
            latestMessage: {
              ...latestItem,
              FeaturedImageUrl: featuredImageUrl,
              VideoUrl: videoUrl
            }
          });
        } else {
          console.log('⚠️ No published messages found for current date');
          
          // Fallback: Try to get the most recent message regardless of publish date
          const fallbackUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?$expand=Author&$select=Id,Title,Content,PublishedDate,Created,FeaturedImage,NewsletterVideo,Author/Title&$orderby=Created desc&$top=1`;
          
          console.log('🔄 Trying fallback query...');
          
          const fallbackResponse = await this.props.context.spHttpClient.get(
            fallbackUrl,
            SPHttpClient.configurations.v1
          );
          
          if (fallbackResponse.ok) {
            const fallbackData = await fallbackResponse.json();
            if (fallbackData.value && fallbackData.value.length > 0) {
              const fallbackItem = fallbackData.value[0];
              console.log('📄 Fallback item found:', fallbackItem.Title);
              
              this.setState({ 
                latestMessage: {
                  ...fallbackItem,
                  FeaturedImageUrl: null,
                  VideoUrl: null
                }
              });
            }
          }
        }
      } else {
        console.error('❌ API Error - Status:', response.status);
      }
    } catch (error) {
      console.error('💥 Load error:', error);
    } finally {
      this.setState({ loading: false });
    }
  }

  // Helper function to calculate days from now
  private getDaysFromNow = (dateString: string): number => {
    const date = new Date(dateString);
    const now = new Date();
    const diffTime = now.getTime() - date.getTime();
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }

  private handlePlayClick = (): void => {
    const { latestMessage } = this.state;
    const videoUrl: string | undefined = latestMessage?.VideoUrl || this.props.videoUrl;
    
    if (!videoUrl || videoUrl.trim() === '') {
      console.log('⚠️ No video URL available');
      return;
    }

    this.setState({ showModal: true });
  }

  private handleReadMore = (): void => {
    console.log('📖 Opening detail view as separate page');
    this.setState({ showDetailView: true });
  }

  private handleBackFromDetail = (): void => {
    console.log('🏠 Back to home page');
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
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 450"%3E%3Crect width="800" height="450" fill="%23f0f0f0"/%3E%3Ctext x="400" y="200" text-anchor="middle" fill="%23666666" font-size="24" font-family="Arial"%3ELatest News%3C/text%3E%3Ctext x="400" y="250" text-anchor="middle" fill="%23666666" font-size="16" font-family="Arial"%3EImage Preview%3C/text%3E%3C/svg%3E';
  }

  private getDefaultBackground = (): string => {
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080"%3E%3Cdefs%3E%3ClinearGradient id="bg" x1="0%25" y1="0%25" x2="100%25" y2="100%25"%3E%3Cstop offset="0%25" style="stop-color:%234a90e2;stop-opacity:1" /%3E%3Cstop offset="100%25" style="stop-color:%237b68ee;stop-opacity:1" /%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width="1920" height="1080" fill="url(%23bg)" /%3E%3C/svg%3E';
  }

  // Enhanced date formatting with current date awareness
  private formatDate = (dateString: string | null | undefined): string | null => {
    if (!dateString) {
      console.log('⚠️ No date provided');
      return null;
    }

    try {
      const date = new Date(dateString);
      if (isNaN(date.getTime())) {
        console.log('⚠️ Invalid date:', dateString);
        return null;
      }

      const formatted = `Published ${date.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'short',
        year: 'numeric'
      })}`;

      console.log('📅 Date formatted as:', formatted);
      return formatted;
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

  // Increased content length to show more lines
  private truncateText = (text: string, maxLength: number = 280): string => {
    if (!text) return '';
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '...';
  }

  public render(): React.ReactElement<IVideoBannerProps> {
    const { backgroundImageUrl } = this.props;
    const { showModal, showDetailView, latestMessage, loading } = this.state;

    // Show DetailedView as separate page-like view
    if (showDetailView) {
      return (
        <DetailedView 
          messageData={latestMessage}
          onBack={this.handleBackFromDetail}
          context={this.props.context}
        />
      );
    }

    // Main banner view
    const displayTitle: string = latestMessage ? latestMessage.Title : (this.props.title || 'Latest News');
    const displayMessage: string = latestMessage 
      ? this.truncateText(this.stripHtmlTags(latestMessage.Content))
      : (this.props.message || 'Stay updated with our latest announcements and news.');
    
    // Date processing with better fallback
    const displayDate: string | null = latestMessage 
      ? (this.formatDate(latestMessage.PublishedDate) || this.formatDate(latestMessage.Created))
      : null;
    
    console.log('🎯 Final render values:', {
      displayTitle,
      displayDate,
      messageLength: displayMessage.length,
      hasImage: !!latestMessage?.FeaturedImageUrl,
      publishedDate: latestMessage?.PublishedDate,
      isPublished: latestMessage ? new Date(latestMessage.PublishedDate || latestMessage.Created) <= new Date() : false
    });
    
    const displayBackground: string = backgroundImageUrl || this.getDefaultBackground();
    const rightSideImage: string = latestMessage?.FeaturedImageUrl || this.props.thumbnailUrl || this.getDefaultThumbnail();
    const videoUrl: string | undefined = latestMessage?.VideoUrl || this.props.videoUrl;
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
                  <p>Loading latest published news...</p>
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
                {this.props.buttonText || 'READ MORE'}
              </button>
            </div>
            
            <div className={styles.videoSection}>
              <div className={styles.thumbnailContainer}>
                <img 
                  src={rightSideImage}
                  alt="Latest news featured image"
                  className={styles.thumbnail}
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = this.getDefaultThumbnail();
                  }}
                />
                {hasVideo && (
                  <button 
                    className={styles.playButton}
                    onClick={this.handlePlayClick}
                    aria-label="Play newsletter video"
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
                    autoPlay={this.props.autoPlay || false}
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
