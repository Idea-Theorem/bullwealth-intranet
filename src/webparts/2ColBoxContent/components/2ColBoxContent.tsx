import * as React from 'react';
import styles from './2ColBoxContent.module.scss';
import type { I2ColBoxContentProps } from './I2ColBoxContentProps';
import { ContactCard } from './ContactCard';

export default class TwoColBoxContent extends React.Component<I2ColBoxContentProps, {}> {
  
  private handleEmailClick = (email: string): void => {
    if (email) {
      window.location.href = `mailto:${email}`;
    }
  };

  private handlePhoneClick = (phone: string): void => {
    if (phone) {
      window.location.href = `tel:${phone}`;
    }
  };

  public render(): React.ReactElement<I2ColBoxContentProps> {
    const {
      leftCard,
      rightCard,
      columnLayout,
      containerBackgroundColor,
      cardSpacing,
      hasTeamsContext
    } = this.props;

    const containerStyle: React.CSSProperties = {
      backgroundColor: containerBackgroundColor || 'transparent',
      gap: `${cardSpacing || 20}px`
    };

    const isReversed = columnLayout === 'right-left';

    return (
      <section className={`${styles.twoColBoxContent} ${hasTeamsContext ? styles.teams : ''}`}>
        <div 
          className={`${styles.cardsContainer} ${isReversed ? styles.reversed : ''}`} 
          style={containerStyle}
        >
          <div className={styles.cardWrapper}>
            <ContactCard
              {...leftCard}
              onEmailClick={this.handleEmailClick}
              onPhoneClick={this.handlePhoneClick}
            />
          </div>
          <div className={styles.cardWrapper}>
            <ContactCard
              {...rightCard}
              onEmailClick={this.handleEmailClick}
              onPhoneClick={this.handlePhoneClick}
            />
          </div>
        </div>
      </section>
    );
  }
}
