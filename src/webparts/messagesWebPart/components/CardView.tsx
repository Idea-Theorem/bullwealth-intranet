import * as React from 'react';
import { 
  Text, 
  DefaultButton
} from '@fluentui/react';
import { IGroupedMessages } from './IMessagesProps';
import styles from './Messages.module.scss';

export interface ICardViewProps {
  title: string;
  groupedMessages: IGroupedMessages;
  onReadMore: (messageId: number) => void;
  columnsPerRow?: number;
}

const CardView: React.FC<ICardViewProps> = ({ 
  title, 
  groupedMessages, 
  onReadMore, 
  columnsPerRow = 4 
}) => {

  const getImageUrl = (item: any): string | null => {
    if (item.FeaturedImage && item.FeaturedImage.Url) {
      return item.FeaturedImage.Url;
    }
    return null;
  };

  const getDefaultGradient = (): string => {
    return '#4CAF50';
  };

  // Enhanced year sorting
  const getSortedYears = (): string[] => {
    return Object.keys(groupedMessages).sort((a, b) => parseInt(b) - parseInt(a));
  };

  // Fixed: Actually use this function in the render
  const getYearLabel = (year: string): string => {
    const currentYear = new Date().getFullYear();
    const yearNum = parseInt(year);
    
    if (yearNum === currentYear) {
      return `Newsletters and Message from CEO from ${year}`;
    } else {
      return `Newsletters and Message from CEO from Past ${year}`;
    }
  };

  const sortedYears = getSortedYears();

  console.log('🎯 CardView render:', {
    totalYears: sortedYears.length,
    years: sortedYears,
    totalMessages: Object.values(groupedMessages).flat().length
  });

  return (
    <div className={styles.cardView}>
      <Text variant="xxLarge" className={styles.mainTitle}>
        {title}
      </Text>

      {sortedYears.map((year) => (
        <div key={year} className={styles.yearSection}>
          {/* Fixed: Actually use getYearLabel function here */}
           <Text variant="xLarge" className={styles.yearTitle}>
            {getYearLabel(year)}
          </Text>
          
          <div 
            className={styles.cardsContainer}
            style={{
              gridTemplateColumns: `repeat(${columnsPerRow}, 1fr)`
            }}
          >
            {groupedMessages[year].map((message) => {
              const imageUrl = getImageUrl(message);
              const backgroundColor = getDefaultGradient();
              
              return (
                <div 
                  key={message.Id}
                  className={styles.messageCard}
                >
                  <div className={styles.cardContent}>
                    <div 
                      className={styles.cardImage}
                      style={{
                        ...(imageUrl ? {
                          backgroundImage: `url("${imageUrl}")`,
                          backgroundSize: 'cover',
                          backgroundPosition: 'center',
                          backgroundRepeat: 'no-repeat'
                        } : {
                          background: `linear-gradient(135deg, ${backgroundColor}, ${backgroundColor}dd)`
                        })
                      }}
                    >
                      {!imageUrl && (
                        <div className={styles.imagePlaceholder}>
                          <Text variant="large" className={styles.imageText}>
                            Newsletter
                          </Text>
                        </div>
                      )}
                    </div>

                    <div className={styles.cardDetails}>
                      <Text variant="mediumPlus" className={styles.cardTitle}>
                        {message.Title}
                      </Text>
                      
                      <div className={styles.cardFooter}>
                        <DefaultButton
                          text="Read More"
                          onClick={() => onReadMore(message.Id)}
                          className={styles.readMoreButton}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
};

export default CardView;
