import * as React from 'react';
import styles from './WelcomeBannerPart.module.scss';
import type { IWelcomeBannerPartProps } from './IWelcomeBannerPartProps';

export interface IWelcomeBannerPartState {
  isExpanded: boolean;
}

export default class WelcomeBannerPart extends React.Component<IWelcomeBannerPartProps, IWelcomeBannerPartState> {
  constructor(props: IWelcomeBannerPartProps) {
    super(props);
    this.state = {
      isExpanded: false
    };
  }

  private handlePlayVideo = (): void => {
    const { videoUrl } = this.props;
    if (videoUrl) {
      // Open video in new window or implement modal player
      window.open(videoUrl, '_blank');
    } else {
      console.log('No video URL configured');
    }
  };

  private handleReadMore = (): void => {
    this.setState({ isExpanded: !this.state.isExpanded });
  };

  public render(): React.ReactElement<IWelcomeBannerPartProps> {
    const { isExpanded } = this.state;
    const {
      messageTitle = 'Message from CEO',
      ceoName = 'CEO Name',
      ceoTitle = 'Chief Executive Officer',
      ceoMessage = 'Add your CEO message here',
      ceoExpandedMessage = 'Add extended message here',
      backgroundImageUrl,
      ceoImageUrl,
      showVideo = true,
      readMoreText = 'Read More',
      readLessText = 'Read Less'
    } = this.props;

    // Use placeholder images if not configured
    const defaultBackground = 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 400"%3E%3Cdefs%3E%3ClinearGradient id="bg" x1="0%25" y1="0%25" x2="100%25" y2="100%25"%3E%3Cstop offset="0%25" style="stop-color:%233498db;stop-opacity:1" /%3E%3Cstop offset="100%25" style="stop-color:%239b59b6;stop-opacity:1" /%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width="1200" height="400" fill="url(%23bg)" /%3E%3C/svg%3E';
    const defaultCeoImage = 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 200"%3E%3Crect width="300" height="200" fill="%23f39c12"/%3E%3Ccircle cx="150" cy="80" r="40" fill="%23fff"/%3E%3Cellipse cx="150" cy="150" rx="60" ry="40" fill="%23fff"/%3E%3Ctext x="150" y="180" text-anchor="middle" fill="%23333" font-family="Arial" font-size="14"%3ECEO Photo%3C/text%3E%3C/svg%3E';

    const backgroundStyle = {
      backgroundImage: `url(${backgroundImageUrl || defaultBackground})`
    };

    return (
      <div className={styles.welcomeBannerPart} style={backgroundStyle}>
        <div className={styles.overlay}>
          <div className={styles.contentCard}>
            <div className={styles.textSection}>
              <h2 className={styles.title}>{messageTitle}</h2>
              <div className={styles.messageContent}>
                <p className={styles.message}>"{ceoMessage}"</p>
                {isExpanded && ceoExpandedMessage && (
                  <p className={styles.expandedMessage}>"{ceoExpandedMessage}"</p>
                )}
                {/* Duplicate message as shown in the design */}
                <p className={styles.message}>"{ceoMessage}"</p>
              </div>
              <button 
                className={styles.readMoreButton}
                onClick={this.handleReadMore}
                aria-expanded={isExpanded}
              >
                {isExpanded ? readLessText : readMoreText}
              </button>
            </div>
            
            {showVideo && (
              <div className={styles.videoSection}>
                <div className={styles.videoThumbnail}>
                  <img 
                    src={ceoImageUrl || defaultCeoImage} 
                    alt={`${ceoName}, ${ceoTitle}`}
                    className={styles.ceoImage}
                  />
                  <button 
                    className={styles.playButton}
                    onClick={this.handlePlayVideo}
                    aria-label="Play video message from CEO"
                    title={`Play video message from ${ceoName}`}
                  >
                    <svg 
                      width="24" 
                      height="24" 
                      viewBox="0 0 24 24" 
                      fill="currentColor"
                      className={styles.playIcon}
                    >
                      <path d="M8 5v14l11-7z"/>
                    </svg>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }
}