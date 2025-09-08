import * as React from 'react';
import styles from './VideoBanner.module.scss' ;
import { IVideoBannerProps } from './IVideoBannerProps';

export interface IVideoBannerState {
  isPlaying: boolean;
  showModal: boolean;
}

export default class VideoBanner extends React.Component<IVideoBannerProps, IVideoBannerState> {
  private videoRef: React.RefObject<HTMLVideoElement>;
  
  constructor(props: IVideoBannerProps) {
    super(props);
    this.state = {
      isPlaying: false,
      showModal: false
    };
    this.videoRef = React.createRef();
  }

  private handlePlayClick = (): void => {
    const { videoUrl, showInModal } = this.props;
    
    if (!videoUrl) {
      alert('No video URL configured. Please configure the video URL in the web part settings.');
      return;
    }

    if (showInModal) {
      this.setState({ showModal: true, isPlaying: true });
    } else {
      // Handle inline video play or redirect
      if (this.isYouTubeUrl(videoUrl) || this.isStreamUrl(videoUrl)) {
        window.open(videoUrl, '_blank');
      } else {
        this.setState({ isPlaying: true });
      }
    }
  }

  private handleCloseModal = (): void => {
    this.setState({ showModal: false, isPlaying: false });
  }


  private isYouTubeUrl = (url: string): boolean => {
  return url.indexOf('youtube.com') !== -1 || url.indexOf('youtu.be') !== -1;
}

private isStreamUrl = (url: string): boolean => {
  return url.indexOf('microsoftstream.com') !== -1 || url.indexOf('stream.microsoft.com') !== -1;
}


  private getEmbedUrl = (url: string): string => {
    // Convert YouTube URLs to embed format
    if (this.isYouTubeUrl(url)) {
      const videoId = this.extractYouTubeId(url);
      return `https://www.youtube.com/embed/${videoId}?autoplay=1`;
    }
    return url;
  }

  private extractYouTubeId = (url: string): string => {
    const match = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([^&\n?#]+)/);
    return match ? match[1] : '';
  }

  private getDefaultThumbnail = (): string => {
    // Return a default thumbnail if none provided
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 450"%3E%3Crect width="800" height="450" fill="%23f39c12"/%3E%3Ctext x="400" y="225" text-anchor="middle" fill="white" font-size="24" font-family="Arial"%3EVideo Thumbnail%3C/text%3E%3C/svg%3E';
  }

  private getDefaultBackground = (): string => {
    // Return a default background gradient if none provided
    return 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080"%3E%3Cdefs%3E%3ClinearGradient id="bg" x1="0%25" y1="0%25" x2="100%25" y2="100%25"%3E%3Cstop offset="0%25" style="stop-color:%235dade2;stop-opacity:1" /%3E%3Cstop offset="100%25" style="stop-color:%2385c1e9;stop-opacity:1" /%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width="1920" height="1080" fill="url(%23bg)" /%3E%3C/svg%3E';
  }

  public render(): React.ReactElement<IVideoBannerProps> {
    const { 
      title, 
      message, 
      videoUrl, 
      thumbnailUrl, 
      backgroundImageUrl 
    } = this.props;
    const { isPlaying, showModal } = this.state;

    const backgroundStyle = {
      backgroundImage: `url(${backgroundImageUrl || this.getDefaultBackground()})`
    };

    return (
      <div className={styles.videoBanner} style={backgroundStyle}>
        <div className={styles.overlay}>
          <div className={styles.contentWrapper}>
            <div className={styles.textContent}>
              <h2 className={styles.title}>{title}</h2>
              <p className={styles.message}>{message}</p>
              <p className={styles.message}>{message}</p>
              {/* <button 
                className={styles.readMoreButton}
                onClick={this.handleReadMore}
              >
                {buttonText}
              </button> */}
            </div>
            
            <div className={styles.videoSection}>
              {!isPlaying || showModal ? (
                <div className={styles.thumbnailContainer}>
                  <img 
                    src={thumbnailUrl || this.getDefaultThumbnail()} 
                    alt="Video thumbnail"
                    className={styles.thumbnail}
                  />
                  <button 
                    className={styles.playButton}
                    onClick={this.handlePlayClick}
                    aria-label="Play video"
                  >
                    <svg 
                      className={styles.playIcon} 
                      viewBox="0 0 80 80" 
                      fill="none" 
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <circle cx="40" cy="40" r="40" fill="white" fillOpacity="0.95"/>
                      <path d="M32 28L52 40L32 52V28Z" fill="#333333"/>
                    </svg>
                  </button>
                </div>
              ) : (
                <video 
                  ref={this.videoRef}
                  className={styles.videoPlayer}
                  controls
                  autoPlay={this.props.autoPlay}
                  src={videoUrl}
                >
                  Your browser does not support the video tag.
                </video>
              )}
            </div>
          </div>
        </div>

        {/* Video Modal */}
        {showModal && (
          <div className={styles.modal}>
            <div className={styles.modalContent}>
              <button 
                className={styles.closeButton}
                onClick={this.handleCloseModal}
                aria-label="Close video"
              >
                ×
              </button>
              {this.isYouTubeUrl(videoUrl) || this.isStreamUrl(videoUrl) ? (
                <iframe 
                  className={styles.videoFrame}
                  src={this.getEmbedUrl(videoUrl)}
                  allowFullScreen
                  allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
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
        )}
      </div>
    );
  }
}