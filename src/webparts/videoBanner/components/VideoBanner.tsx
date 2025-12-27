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
  currentSlideIndex: number;
  userName: string;
}

export default class VideoBanner extends React.Component<IVideoBannerProps, IVideoBannerState> {
  
  private slideIntervalId: any = null;

  constructor(props: IVideoBannerProps) {
    super(props);
    this.state = {
      showModal: false,
      showDetailView: false,
      latestMessage: null,
      loading: true,
      currentSlideIndex: 0,
      userName: ''
    };
  }

  public componentDidMount(): void {
    this.loadLatestMessage();
    this.loadUserName();
    this.startAutoSlide();
  }

  public componentWillUnmount(): void {
    this.stopAutoSlide();
  }

  private loadUserName = async (): Promise<void> => {
    try {
      const user = this.props.context.pageContext.user;
      this.setState({ userName: user.displayName });
    } catch {
      this.setState({ userName: 'User' });
    }
  };

  private startAutoSlide = (): void => {
    const { backgroundImages, autoSlide, slideInterval } = this.props;
    
    if (autoSlide && backgroundImages && backgroundImages.length > 1) {
      this.stopAutoSlide();
      this.slideIntervalId = setInterval(() => {
        this.nextSlide();
      }, slideInterval * 1000);
    }
  };

  private stopAutoSlide = (): void => {
    if (this.slideIntervalId) {
      clearInterval(this.slideIntervalId);
      this.slideIntervalId = null;
    }
  };

  private nextSlide = (): void => {
    const { backgroundImages } = this.props;
    if (!backgroundImages || backgroundImages.length === 0) return;
    
    this.setState(prevState => ({
      currentSlideIndex: (prevState.currentSlideIndex + 1) % backgroundImages.length
    }));
  };

  private getImageUrl = async (itemId: number, imageFieldValue: any, siteUrl: string): Promise<string | null> => {
    if (!imageFieldValue) return null;
    try {
      let imageData: any = imageFieldValue;
      if (typeof imageFieldValue === 'string') {
        try {
          imageData = JSON.parse(imageFieldValue);
        } catch {
          return null;
        }
      }

      if (imageData && imageData.fileName) {
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
            if (renderData.Row && renderData.Row.length > 0) {
              const row = renderData.Row[0];
              if (row.FeaturedImage) {
                const imgMatch = row.FeaturedImage.match(/src="([^"]+)"/);
                if (imgMatch) {
                  return imgMatch[1];
                }
              }
            }
          }
        } catch {
          // ignore
        }
      }
      return null;
    } catch {
      return null;
    }
  };

  private loadLatestMessage = async (): Promise<void> => {
    try {
      this.setState({ loading: true });
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      const today = new Date().toISOString();
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items?` +
        `$expand=Author&` +
        `$select=Id,Title,Content,PublishedDate,Created,FeaturedImage,NewsletterVideo,Author/Title&` +
        `$filter=PublishedDate le datetime'${today}'&` +
        `$orderby=PublishedDate desc&` +
        `$top=1`;

      const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        this.setState({ latestMessage: null });
        return;
      }

      const data = await response.json();
      if (!data.value || data.value.length === 0) {
        this.setState({ latestMessage: null });
        return;
      }

      const latestItem = data.value[0];

      let featuredImageUrl: string | null = null;
      if (latestItem.FeaturedImage) {
        featuredImageUrl = await this.getImageUrl(latestItem.Id, latestItem.FeaturedImage, siteUrl);
      }

      let videoUrl: string | null = null;
      if (latestItem.NewsletterVideo) {
        if (typeof latestItem.NewsletterVideo === 'string') {
          videoUrl = latestItem.NewsletterVideo.trim();
        } else if (typeof latestItem.NewsletterVideo === 'object') {
          videoUrl = (latestItem.NewsletterVideo.Url || latestItem.NewsletterVideo.url || '').trim();
        }
      }

      let attachmentUrl: string | null = null;
      let attachmentServerRelativeUrl: string | null = null;
      let attachmentImageUrl: string | null = null;
      try {
        const attachmentsApi =
          `${siteUrl}/_api/web/lists/getbytitle('Archived-Messages')/items(${latestItem.Id})/AttachmentFiles`;
        const attResp: SPHttpClientResponse = await this.props.context.spHttpClient.get(
          attachmentsApi,
          SPHttpClient.configurations.v1
        );
        if (attResp.ok) {
          const attData = await attResp.json();
          if (attData.value && attData.value.length > 0) {
            for (const file of attData.value) {
              const srvUrl: string = file.ServerRelativeUrl || '';
              const lower = srvUrl.toLowerCase();
              const isImage = lower.endsWith('.jpg') || lower.endsWith('.jpeg') ||
                lower.endsWith('.png') || lower.endsWith('.gif');
              const isDoc = lower.endsWith('.pdf') || lower.endsWith('.doc') ||
                lower.endsWith('.docx') || lower.endsWith('.ppt') ||
                lower.endsWith('.pptx') || lower.endsWith('.xls') ||
                lower.endsWith('.xlsx');

              if (!attachmentUrl && isDoc) {
                attachmentServerRelativeUrl = srvUrl;
                attachmentUrl = `${window.location.protocol}//${window.location.host}${srvUrl}`;
              }

              if (!attachmentImageUrl && isImage) {
                attachmentImageUrl = `${window.location.protocol}//${window.location.host}${srvUrl}`;
              }
            }
          }
        }
      } catch {
        // ignore
      }

      if (!featuredImageUrl && attachmentImageUrl) {
        featuredImageUrl = attachmentImageUrl;
      }

      this.setState({
        latestMessage: {
          ...latestItem,
          FeaturedImageUrl: featuredImageUrl,
          VideoUrl: videoUrl,
          AttachmentUrl: attachmentUrl,
          AttachmentServerRelativeUrl: attachmentServerRelativeUrl
        }
      });
    } catch {
      this.setState({ latestMessage: null });
    } finally {
      this.setState({ loading: false });
    }
  };

  private handlePlayClick = (): void => {
    const videoUrl: string | undefined = this.state.latestMessage?.VideoUrl;
    if (!videoUrl || videoUrl.trim() === '') return;
    window.open(videoUrl, '_blank', 'noopener,noreferrer');
  };

  private handleReadMore = (): void => {
    this.setState({ showDetailView: true });
  };

  private handleBackFromDetail = (): void => {
    this.setState({ showDetailView: false });
  };

  private getDefaultThumbnail = (): string =>
    'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 450"%3E%3Crect width="800" height="450" fill="%23f0f0f0"/%3E%3Ctext x="400" y="200" text-anchor="middle" fill="%23999999" font-size="24" font-family="Arial"%3ELatest News%3C/text%3E%3C/svg%3E';

  private getDefaultBackground = (): string =>
    'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080"%3E%3Cdefs%3E%3ClinearGradient id="bg" x1="0%25" y1="0%25" x2="100%25" y2="100%25"%3E%3Cstop offset="0%25" style="stop-color:%234a90e2;stop-opacity:1"/%3E%3Cstop offset="100%25" style="stop-color:%237b68ee;stop-opacity:1"/%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width="1920" height="1080" fill="url(%23bg)"/%3E%3C/svg%3E';

  private formatDate = (dateString: string | null | undefined): string | null => {
    if (!dateString) return null;
    const d = new Date(dateString);
    if (isNaN(d.getTime())) return null;
    return `Published ${d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
  };

  private stripHtmlTags = (html: string): string => {
    if (!html) return '';
    const div = document.createElement('div');
    div.innerHTML = html;
    return div.textContent || div.innerText || '';
  };

  private truncateText = (text: string, maxLength: number = 280): string => {
    if (!text) return '';
    return text.length <= maxLength ? text : text.substring(0, maxLength) + '...';
  };

  public render(): React.ReactElement<IVideoBannerProps> {
    const { showDetailView, latestMessage, loading, currentSlideIndex, userName } = this.state;
    const { backgroundImages } = this.props;

    if (showDetailView) {
      return (
        <DetailedView
          messageData={latestMessage}
          onBack={this.handleBackFromDetail}
          context={this.props.context}
        />
      );
    }

    const displayTitle = latestMessage?.Title || 'Latest News';
    const displayMessage = latestMessage
      ? this.truncateText(this.stripHtmlTags(latestMessage.Content))
      : 'Stay updated with our latest announcements and news.';
    const displayDate = latestMessage
      ? (this.formatDate(latestMessage.PublishedDate) || this.formatDate(latestMessage.Created))
      : null;

    let displayBackground = this.getDefaultBackground();
    if (backgroundImages && backgroundImages.length > 0) {
      displayBackground = backgroundImages[currentSlideIndex];
    }

    const rightSideImage = latestMessage?.FeaturedImageUrl || this.getDefaultThumbnail();
    const videoUrl: string | undefined = latestMessage?.VideoUrl;
    const hasVideo = !!(videoUrl && videoUrl.trim() !== '');

    const backgroundStyle = {
      backgroundImage: `linear-gradient(rgba(0,0,0,0.4),rgba(0,0,0,0.4)),url(${displayBackground})`,
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      transition: 'background-image 0.8s ease-in-out'
    };

    if (loading) {
      return (
        <div className={styles.loadingSpinner}>
          <div className={styles.spinner}></div>
          Loading latest news...
        </div>
      );
    }

    return (
      <div className={styles.fullWidthContainer}>
        <div className={styles.welcomeStrip}>
          Welcome, {userName}
        </div>

        <div className={styles.videoBanner}>
          <div className={styles.thumbnailWrapper} style={backgroundStyle}>
            <div className={styles.overlay}></div>
            
            <div className={styles.contentWrapper}>
              <div className={styles.textContent}>
                <h2 className={styles.title}>{displayTitle}</h2>
                {displayDate && (
                  <p className={styles.publishedDate}>{displayDate}</p>
                )}
                <p className={styles.message}>{displayMessage}</p>
                <button
                  className={styles.readMoreButton}
                  onClick={this.handleReadMore}
                >
                  READ MORE
                </button>
              </div>

              <div className={styles.videoSection}>
                <div className={styles.thumbnailContainer}>
                  <img
                    className={styles.thumbnail}
                    src={rightSideImage}
                    alt={displayTitle}
                    onError={(e): void => {
                      (e.target as HTMLImageElement).src = this.getDefaultThumbnail();
                    }}
                  />
                  {hasVideo && (
                    <button
                      className={styles.playButton}
                      onClick={this.handlePlayClick}
                      aria-label="Play video"
                    >
                      <div className={styles.playIcon}></div>
                    </button>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
