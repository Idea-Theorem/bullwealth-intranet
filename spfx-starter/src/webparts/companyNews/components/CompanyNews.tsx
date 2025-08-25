import * as React from 'react';
import styles from './CompanyNews.module.scss';
import { ICompanyNewsProps, INewsItem } from './ICompanyNewsProps';

export interface ICompanyNewsState {
  currentIndex: number;
  isAnimating: boolean;
}

export default class CompanyNews extends React.Component<ICompanyNewsProps, ICompanyNewsState> {
  private autoScrollInterval: any;
  
  constructor(props: ICompanyNewsProps) {
    super(props);
    this.state = {
      currentIndex: 0,
      isAnimating: false
    };
  }

  public componentDidMount(): void {
    // Auto-scroll if enabled
    if (this.props.autoScroll) {
      this.startAutoScroll();
    }
  }

  public componentDidUpdate(prevProps: ICompanyNewsProps): void {
    // Handle auto-scroll changes
    if (prevProps.autoScroll !== this.props.autoScroll) {
      if (this.props.autoScroll) {
        this.startAutoScroll();
      } else {
        this.stopAutoScroll();
      }
    }
    
    // Handle interval changes
    if (prevProps.autoScrollInterval !== this.props.autoScrollInterval && this.props.autoScroll) {
      this.stopAutoScroll();
      this.startAutoScroll();
    }
  }

  public componentWillUnmount(): void {
    this.stopAutoScroll();
  }

  private startAutoScroll = (): void => {
    this.autoScrollInterval = setInterval(() => {
      this.handleNext();
    }, this.props.autoScrollInterval);
  }

  private stopAutoScroll = (): void => {
    if (this.autoScrollInterval) {
      clearInterval(this.autoScrollInterval);
    }
  }

  private handleNext = (): void => {
    const { newsItems, itemsToShow } = this.props;
    const { currentIndex } = this.state;
    
    if (!this.state.isAnimating) {
      const maxIndex = Math.max(0, newsItems.length - itemsToShow);
      const nextIndex = currentIndex >= maxIndex ? 0 : currentIndex + 1;
      
      this.setState({ 
        currentIndex: nextIndex,
        isAnimating: true 
      }, () => {
        setTimeout(() => {
          this.setState({ isAnimating: false });
        }, 300);
      });
    }
  }

  // private handlePrevious = (): void => {
  //   const { newsItems, itemsToShow } = this.props;
  //   const { currentIndex } = this.state;
    
  //   if (!this.state.isAnimating) {
  //     const maxIndex = Math.max(0, newsItems.length - itemsToShow);
  //     const prevIndex = currentIndex <= 0 ? maxIndex : currentIndex - 1;
      
  //     this.setState({ 
  //       currentIndex: prevIndex,
  //       isAnimating: true 
  //     }, () => {
  //       setTimeout(() => {
  //         this.setState({ isAnimating: false });
  //       }, 300);
  //     });
  //   }
  // }

  private handleDotClick = (index: number): void => {
    this.stopAutoScroll();
    this.setState({ currentIndex: index });
    this.startAutoScroll();
  }

  private handleShare = (item: INewsItem): void => {
    if (item.shareUrl) {
      window.open(item.shareUrl, '_blank');
    } else {
      // Copy link to clipboard
      const url = window.location.href;
      void navigator.clipboard.writeText(url).then(() => {
        alert('Link copied to clipboard!');
      }).catch((err) => {
        console.error('Failed to copy link:', err);
        alert('Failed to copy link to clipboard');
      });
    }
  }

  private handleReadMore = (item: INewsItem): void => {
    if (item.readMoreUrl) {
      window.open(item.readMoreUrl, '_blank');
    } else {
      console.log('Read more:', item.title);
      // You can implement a modal or navigation here
    }
  }

  private getDefaultImage = (index: number): string => {
    // Return different placeholder based on index
    const colors = ['#4A90E2', '#E24A90', '#90E24A', '#E2904A'];
    const color = colors[index % colors.length];
    
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 250'%3E%3Crect width='400' height='250' fill='${encodeURIComponent(color)}'/%3E%3Ctext x='200' y='125' text-anchor='middle' fill='white' font-size='24' font-family='Arial'%3ENews Image%3C/text%3E%3C/svg%3E`;
  }

  public render(): React.ReactElement<ICompanyNewsProps> {
    const { title, newsItems, itemsToShow } = this.props;
    const { currentIndex } = this.state;
    
    // Calculate visible items
    const visibleItems = newsItems.slice(currentIndex, currentIndex + itemsToShow);
    
    // If we don't have enough items at the end, wrap around
    if (visibleItems.length < itemsToShow) {
      const remaining = itemsToShow - visibleItems.length;
      visibleItems.push(...newsItems.slice(0, remaining));
    }

    // Calculate dots
    const totalDots = Math.max(1, Math.ceil(newsItems.length / itemsToShow));
    const activeDot = Math.floor(currentIndex / itemsToShow);

    return (
      <div className={styles.companyNews}>
        <h2 className={styles.sectionTitle}>{title}</h2>
        
        <div className={styles.newsContainer}>
          <div className={styles.newsGrid}>
            {visibleItems.map((item, index) => (
              <div key={`${item.id}-${index}`} className={styles.newsCard}>
                <div className={styles.cardImage}>
                  <img 
                    src={item.imageUrl || this.getDefaultImage(parseInt(item.id))} 
                    alt={item.title}
                  />
                </div>
                
                <div className={styles.cardContent}>
                  <div className={styles.cardHeader}>
                    <div className={styles.authorInfo}>
                      <span className={styles.authorName}>{item.author}</span>-
                      <span className={styles.date}>{item.date}</span>
                    </div>
                    <button 
                      className={styles.shareButton}
                      onClick={() => this.handleShare(item)}
                      aria-label="Share"
                    >
                      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                        <path d="M12 5.5C13.1046 5.5 14 4.60457 14 3.5C14 2.39543 13.1046 1.5 12 1.5C10.8954 1.5 10 2.39543 10 3.5C10 3.70873 10.0315 3.91019 10.0896 4.09896L6.41043 6.34896C6.01417 5.94119 5.46629 5.68571 4.85714 5.68571C3.69391 5.68571 2.75 6.62963 2.75 7.79286C2.75 8.95608 3.69391 9.9 4.85714 9.9C5.46629 9.9 6.01417 9.64452 6.41043 9.23675L10.0896 11.4868C10.0315 11.6755 10 11.877 10 12.0857C10 13.1903 10.8954 14.0857 12 14.0857C13.1046 14.0857 14 13.1903 14 12.0857C14 10.9812 13.1046 10.0857 12 10.0857C11.3908 10.0857 10.843 10.3412 10.4468 10.749L6.76758 8.49896C6.82568 8.31019 6.85714 8.10873 6.85714 7.9C6.85714 7.69127 6.82568 7.48981 6.76758 7.30104L10.4468 5.05104C10.843 5.45881 11.3908 5.71429 12 5.71429V5.5Z" 
                          fill="currentColor"/>
                      </svg>
                      <span>Share</span>
                    </button>
                  </div>
                  
                  <h3 className={styles.cardTitle}>{item.title}</h3>
                  
                  <button 
                    className={styles.readMoreLink}
                    onClick={() => this.handleReadMore(item)}
                  >
                    Read More
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className={styles.pagination}>
          <div className={styles.dots}>
            {[...Array(totalDots)].map((_: any, index: number) => (
              <button
                key={index}
                className={`${styles.dot} ${index === activeDot ? styles.active : ''}`}
                onClick={() => this.handleDotClick(index * itemsToShow)}
                aria-label={`Go to page ${index + 1}`}
              />
            ))}
          </div>
        </div>
      </div>
    );
  }
}