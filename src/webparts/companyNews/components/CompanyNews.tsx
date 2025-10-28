import * as React from 'react';
import { useState, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CompanyNews.module.scss';
import { ICompanyNewsProps, INewsItem } from './ICompanyNewsProps';

interface ICompanyNewsState {
  currentPage: number;
}

const CompanyNews: React.FC<ICompanyNewsProps> = (props) => {
  const [state, setState] = useState<ICompanyNewsState>({
    currentPage: 0
  });

  // Sort news items by date (latest first)
  const sortedNewsItems = React.useMemo(() => {
    return [...props.newsItems].sort((a, b) => {
      const dateA = new Date(a.date).getTime();
      const dateB = new Date(b.date).getTime();
      return dateB - dateA;
    });
  }, [props.newsItems]);

  const itemsToShow = props.itemsToShow || 4;
  const totalPages = Math.ceil(sortedNewsItems.length / itemsToShow);

  useEffect(() => {
    if (props.autoScroll && totalPages > 1) {
      const interval = setInterval(() => {
        setState(prev => ({
          currentPage: (prev.currentPage + 1) % totalPages
        }));
      }, props.autoScrollInterval);

      return () => clearInterval(interval);
    }
  }, [props.autoScroll, props.autoScrollInterval, totalPages]);

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      const options: Intl.DateTimeFormatOptions = {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      };
      return date.toLocaleDateString('en-US', options);
    } catch (error) {
      return dateString;
    }
  };

  const handleShare = (item: INewsItem): void => {
    const shareUrl = item.readMoreUrl || item.shareUrl || window.location.href;
    
    if ((window as any).SP && (window as any).SP.UI && (window as any).SP.UI.ModalDialog) {
      try {
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
        fallbackShare(item, shareUrl);
      }
    } else {
      fallbackShare(item, shareUrl);
    }
  };

  const fallbackShare = (item: INewsItem, shareUrl: string): void => {
    const shareData = {
      title: item.title,
      text: `Check out this news: ${item.title}`,
      url: shareUrl
    };

    if (navigator.share) {
      navigator.share(shareData).catch(() => {
        copyToClipboard(shareUrl, item.title);
      });
    } else {
      copyToClipboard(shareUrl, item.title);
    }
  };

  const copyToClipboard = (url: string, title: string): void => {
    if (navigator.clipboard) {
      void navigator.clipboard.writeText(url).then(() => {
        alert(`Link copied to clipboard!\n\n"${title}"\n${url}`);
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

  const handleDotClick = (pageIndex: number): void => {
    setState({ currentPage: pageIndex });
  };

  // Calculate visible items
  const startIndex = state.currentPage * itemsToShow;
  const endIndex = Math.min(startIndex + itemsToShow, sortedNewsItems.length);
  const visibleItems = sortedNewsItems.slice(startIndex, endIndex);

  // Debug log
  console.log('Page:', state.currentPage, 'Start:', startIndex, 'End:', endIndex, 'Showing:', visibleItems.length, 'items');

  return (
    <div className={styles.companyNews}>
      <h2 className={styles.sectionTitle}>{props.title}</h2>
      
      <div className={styles.newsContainer}>
        <div className={styles.carousel}>
          <div className={styles.newsGrid}>
            {visibleItems.map((item, index) => (
              <div key={`${item.id}-${startIndex + index}`} className={styles.newsCard}>
                <div className={styles.cardImage}>
                  <img 
                    src={item.imageUrl || getDefaultImage(startIndex + index)} 
                    alt={item.title}
                    onError={(e) => {
                      (e.target as HTMLImageElement).src = getDefaultImage(startIndex + index);
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
          
          {props.showDots && totalPages > 1 && (
            <div className={styles.pagination}>
              <div className={styles.dots}>
                {[...Array(totalPages)].map((_, pageIndex) => (
                  <button
                    key={pageIndex}
                    className={`${styles.dot} ${state.currentPage === pageIndex ? styles.active : ''}`}
                    onClick={() => handleDotClick(pageIndex)}
                    aria-label={`Page ${pageIndex + 1}`}
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
