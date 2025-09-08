import * as React from 'react';
import { useState, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CompanyNews.module.scss';
import { ICompanyNewsProps, INewsItem } from './ICompanyNewsProps';

interface ICompanyNewsState {
  currentIndex: number;
}

const CompanyNews: React.FC<ICompanyNewsProps> = (props) => {
  const [state, setState] = useState<ICompanyNewsState>({
    currentIndex: 0
  });

  // Auto-scroll effect
  useEffect(() => {
    if (props.autoScroll && props.newsItems.length > 0) {
      const interval = setInterval(() => {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        handleNext();
      }, props.autoScrollInterval);

      return () => clearInterval(interval);
    }
  }, [props.autoScroll, props.autoScrollInterval, state.currentIndex]);

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      const options: Intl.DateTimeFormatOptions = {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      };
      return date.toLocaleDateString('en-US', options);
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    } catch (error) {
      return dateString;
    }
  };

  const handleShare = (item: INewsItem): void => {
    if ((window as any).SP && (window as any).SP.UI && (window as any).SP.UI.ModalDialog) {
      try {
        const shareUrl = item.shareUrl || window.location.href;
        const options = {
          url: shareUrl,
          title: 'Share',
          allowMaximize: false,
          showClose: true,
          width: 600,
          height: 500
        };
        (window as any).SP.UI.ModalDialog.showModalDialog(options);
      } catch (error) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        fallbackShare(item);
      }
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      fallbackShare(item);
    }
  };

  const fallbackShare = (item: INewsItem): void => {
    const shareUrl = item.shareUrl || window.location.href;
    const shareData = {
      title: item.title,
      text: `Check out this news: ${item.title}`,
      url: shareUrl
    };

    if (navigator.share) {
      navigator.share(shareData).catch(() => {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        copyToClipboard(shareUrl);
      });
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      copyToClipboard(shareUrl);
    }
  };

  const copyToClipboard = (url: string): void => {
    if (navigator.clipboard) {
      void navigator.clipboard.writeText(url).then(() => {
        alert('Link copied to clipboard!');
      });
    }
  };

  const handleReadMore = (item: INewsItem): void => {
    if (item.readMoreUrl && item.readMoreUrl !== '#') {
      window.open(item.readMoreUrl, '_blank');
    }
  };

  const getDefaultImage = (index: number): string => {
    const colors = ['#4A90E2', '#E24A90', '#90E24A', '#E2904A'];
    const color = colors[index % colors.length];
    
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 250'%3E%3Crect width='400' height='250' fill='${encodeURIComponent(color)}'/%3E%3Ctext x='200' y='125' text-anchor='middle' fill='white' font-size='24' font-family='Arial'%3ENews Image%3C/text%3E%3C/svg%3E`;
  };

  // Carousel navigation - Auto scroll only, no manual arrows
  const handleNext = (): void => {
    const maxIndex = Math.max(0, props.newsItems.length - (props.itemsToShow || 4));
    const nextIndex = state.currentIndex >= maxIndex ? 0 : state.currentIndex + 1;
    setState({ currentIndex: nextIndex });
  };

  const handleDotClick = (index: number): void => {
    setState({ currentIndex: index * (props.itemsToShow || 4) });
  };

  const visibleItems = props.newsItems.slice(state.currentIndex, state.currentIndex + (props.itemsToShow || 4));
  if (visibleItems.length < (props.itemsToShow || 4) && props.newsItems.length > 0) {
    const remaining = (props.itemsToShow || 4) - visibleItems.length;
    visibleItems.push(...props.newsItems.slice(0, remaining));
  }

  return (
    <div className={styles.companyNews}>
      {/* Header - No Buttons */}
      <h2 className={styles.sectionTitle}>{props.title}</h2>
      
      <div className={styles.newsContainer}>
        <div className={styles.carousel}>
          {/* ✅ NO ARROW BUTTONS - Removed completely */}
          
          <div className={styles.newsGrid}>
            {visibleItems.map((item, index) => (
              <div key={item.id || index} className={styles.newsCard}>
                <div className={styles.cardImage}>
                  <img 
                    src={item.imageUrl || getDefaultImage(index)} 
                    alt={item.title}
                    onError={(e) => {
                      (e.target as HTMLImageElement).src = getDefaultImage(index);
                    }}
                  />
                </div>
                
                <div className={styles.cardContent}>
                  <div className={styles.cardHeader}>
                    <div className={styles.authorInfo}>
                      <span className={styles.authorName}>{item.author}</span>
                      <span className={styles.dateSeparator}> - </span>
                      <span className={styles.date}>{formatDate(item.date)}</span>
                    </div>
                    
                    {/* Only Share Button */}
                    <button 
                      className={styles.shareButton}
                      onClick={() => handleShare(item)}
                      title="Share"
                    >
                      <Icon iconName="Share" />
                      <span>Share</span>
                    </button>
                  </div>
                  
                  <h3 className={styles.cardTitle}>{item.title}</h3>
                  
                  <button 
                    className={styles.readMoreLink}
                    onClick={() => handleReadMore(item)}
                  >
                    Read More
                  </button>
                </div>
              </div>
            ))}
          </div>
          
          {/* ✅ DOTS ONLY Navigation */}
          {props.showDots && props.newsItems.length > (props.itemsToShow || 4) && (
            <div className={styles.pagination}>
              <div className={styles.dots}>
                {[...Array(Math.ceil(props.newsItems.length / (props.itemsToShow || 4)))].map((_, index) => (
                  <button
                    key={index}
                    className={`${styles.dot} ${Math.floor(state.currentIndex / (props.itemsToShow || 4)) === index ? styles.active : ''}`}
                    onClick={() => handleDotClick(index)}
                  />
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default CompanyNews;
