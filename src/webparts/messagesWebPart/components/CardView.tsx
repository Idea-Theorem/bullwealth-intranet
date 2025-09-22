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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const getImageUrl = (item: any): string | null => {
    console.log(`\n--- CardView getImageUrl for Item ${item.Id} ---`);
    console.log('Item FeaturedImage:', item.FeaturedImage);
    
    if (item.FeaturedImage && item.FeaturedImage.Url) {
      console.log('Found image URL:', item.FeaturedImage.Url);
      return item.FeaturedImage.Url;
    }
    
    console.log('No image URL found');
    return null;
  };

  const getDefaultGradient = (): string => {
    // Single default gradient since no categories
    return '#4CAF50';
  };

  // Sort years in descending order
  const sortedYears = Object.keys(groupedMessages).sort((a, b) => parseInt(b) - parseInt(a));

  console.log('CardView render - sortedYears:', sortedYears);
  console.log('CardView render - groupedMessages:', groupedMessages);

  return (
    <div className={styles.cardView}>
      <Text variant="xxLarge" className={styles.mainTitle}>
        {title}
      </Text>

      {sortedYears.map((year) => (
        <div key={year} className={styles.yearSection}>
          {/* <Text variant="xLarge" className={styles.yearTitle}>
            Messages from {year}
          </Text> */}
          
          <div 
            className={styles.cardsContainer}
            style={{
              gridTemplateColumns: `repeat(${columnsPerRow}, 1fr)`
            }}
          >
            {groupedMessages[year].map((message) => {
              const imageUrl = getImageUrl(message);
              const backgroundColor = getDefaultGradient();
              
              console.log(`\n--- Rendering Card ${message.Id} ---`);
              console.log('Title:', message.Title);
              console.log('Image URL:', imageUrl);
              console.log('Background color:', backgroundColor);
              
              return (
                <div 
                  key={message.Id}
                  className={styles.messageCard}
                >
                  <div className={styles.cardContent}>
                    {/* Card Image - No Category Badge */}
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
                            Message
                          </Text>
                        </div>
                      )}
                      
                      {/* Debug info - remove after testing */}
                      {imageUrl && (
                        <div style={{
                          position: 'absolute',
                          bottom: '5px',
                          left: '5px',
                          background: 'rgba(0,0,0,0.7)',
                          color: 'white',
                          fontSize: '10px',
                          padding: '2px 4px',
                          borderRadius: '2px'
                        }}>
                          IMG: ✓
                        </div>
                      )}
                    </div>

                    {/* Card Details */}
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
