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

  const [shareItem, setShareItem] = useState<INewsItem | null>(null);
  const [showSharePanel, setShowSharePanel] = useState<boolean>(false);

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
    setShareItem(item);
    setShowSharePanel(true);
  };

  const getEncodedUrl = (item: INewsItem): string => {
    const shareUrl = item.readMoreUrl || item.shareUrl || window.location.href;
    const absoluteUrl = shareUrl.startsWith('http')
      ? shareUrl
      : `${window.location.protocol}//${window.location.host}${shareUrl}`;
    return absoluteUrl.replace(/ /g, '%20');
  };

  const copyToClipboard = (url: string, title: string): void => {
    if (navigator.clipboard) {
      void navigator.clipboard.writeText(url).then(() => {
        alert(`Link copied to clipboard!\n\n"${title}"\n${url}`);
      });
    }
  };

  // Opens Outlook desktop app via mailto with TinyURL shortened link
  const openOutlookAppShare = async (item: INewsItem, encodedUrl: string): Promise<void> => {
    const displayName = item.title;
    const subject = encodeURIComponent(`Sharing: ${displayName}`);

    let shareUrl = encodedUrl;

    try {
      const res = await fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(encodedUrl)}`);
      if (res.ok) {
        const short = await res.text();
        if (short.startsWith('https://tinyurl.com')) {
          shareUrl = short.trim();
        }
      }
    } catch (e) { /* fallback to full URL */ }

    const body = encodeURIComponent(
      `Hi,\r\n\r\nPlease find the link below:\r\n\r\n${displayName}\r\n${shareUrl}\r\n\r\nBest regards,`
    );

    window.location.href = `mailto:?subject=${subject}&body=${body}`;
    setShowSharePanel(false);
  };

  // Opens Outlook Web (OWA) with proper HTML clickable hyperlink
  // const openOutlookWebShare = (item: INewsItem, encodedUrl: string): void => {
  //   const displayName = item.title;
  //   const subject = encodeURIComponent(`Sharing: ${displayName}`);
  //   const htmlBody = encodeURIComponent(
  //     `<p>Hi,</p>` +
  //     `<p>Please find the link below:</p>` +
  //     `<p><b>${displayName}</b></p>` +
  //     `<p><a href="${encodedUrl}" style="color:#0078d4;text-decoration:underline;">Click here to view</a></p>` +
  //     `<p>Kindly let me know if you face any access issues.</p>` +
  //     `<p>Best regards,</p>`
  //   );
  //   const owaUrl = `https://outlook.office365.com/mail/deeplink/compose?subject=${subject}&body=${htmlBody}&ishtml=true`;
  //   window.open(owaUrl, '_blank', 'noopener,noreferrer');
  //   setShowSharePanel(false);
  // };

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

  const startIndex = state.currentPage * itemsToShow;
  const endIndex = Math.min(startIndex + itemsToShow, sortedNewsItems.length);
  const visibleItems = sortedNewsItems.slice(startIndex, endIndex);

  const renderSharePanel = (): JSX.Element | null => {
    if (!showSharePanel || !shareItem) return null;

    const encodedUrl = getEncodedUrl(shareItem);

    return (
      <div style={{
        position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
        backgroundColor: 'rgba(0,0,0,0.4)', zIndex: 9999,
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      }} onClick={() => setShowSharePanel(false)}>
        <div style={{
          background: '#fff', borderRadius: '12px', padding: '24px',
          width: '380px', boxShadow: '0 8px 32px rgba(0,0,0,0.18)'
        }} onClick={(e) => e.stopPropagation()}>

          {/* Header */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, fontSize: '16px' }}>Share</span>
            <button onClick={() => setShowSharePanel(false)}
              style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '18px', color: '#605e5c' }}>✕</button>
          </div>

          <p style={{ fontSize: '13px', color: '#605e5c', marginBottom: '20px', wordBreak: 'break-all' }}>
            {shareItem.title}
          </p>

          {/* Share icons */}
          <div style={{ display: 'flex', gap: '16px', justifyContent: 'center', marginBottom: '24px', flexWrap: 'wrap' }}>

            {/* Outlook Desktop */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }}
              onClick={() => { void openOutlookAppShare(shareItem, encodedUrl); }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#0078d4', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="OutlookLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Outlook<br />App</span>
            </div>

            {/* Outlook Web */}
            {/* <div style={{ textAlign: 'center', cursor: 'pointer' }}
              onClick={() => openOutlookWebShare(shareItem, encodedUrl)}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#005a9e', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="OutlookLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Outlook<br />Web</span>
            </div> */}

            {/* Teams */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://teams.microsoft.com/share?href=${encodeURIComponent(encodedUrl)}&msgText=${encodeURIComponent(`Check out: ${shareItem.title}`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#6264a7', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="TeamsLogo" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Teams</span>
            </div>

            {/* WhatsApp */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://wa.me/?text=${encodeURIComponent(`${shareItem.title}\n${encodedUrl}`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#25d366', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="Chat" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>WhatsApp</span>
            </div>

            {/* Gmail */}
            <div style={{ textAlign: 'center', cursor: 'pointer' }} onClick={() => {
              const url = `https://mail.google.com/mail/?view=cm&su=${encodeURIComponent(`Sharing: ${shareItem.title}`)}&body=${encodeURIComponent(`Hi,\n\nPlease find the link below:\n${shareItem.title}\n${encodedUrl}\n\nBest regards`)}`;
              window.open(url, '_blank', 'noopener,noreferrer');
              setShowSharePanel(false);
            }}>
              <div style={{ width: 48, height: 48, borderRadius: '50%', background: '#ea4335', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 6px' }}>
                <Icon iconName="Mail" style={{ fontSize: 24, color: '#fff' }} />
              </div>
              <span style={{ fontSize: '11px' }}>Gmail</span>
            </div>
          </div>

          {/* Copy Link */}
          <div style={{
            display: 'flex', alignItems: 'center', gap: '10px',
            border: '1px solid #edebe9', borderRadius: '6px', padding: '10px 12px'
          }}>
            <span style={{ flex: 1, fontSize: '12px', color: '#605e5c', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {encodedUrl}
            </span>
            <button onClick={() => { copyToClipboard(encodedUrl, shareItem.title); setShowSharePanel(false); }}
              style={{ background: '#0078d4', color: '#fff', border: 'none', borderRadius: '4px', padding: '6px 14px', cursor: 'pointer', fontSize: '13px', fontWeight: 600 }}>
              Copy
            </button>
          </div>
        </div>
      </div>
    );
  };

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

      {renderSharePanel()}
    </div>
  );
};

export default CompanyNews;