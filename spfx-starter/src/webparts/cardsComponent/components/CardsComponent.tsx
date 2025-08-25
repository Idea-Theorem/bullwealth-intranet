import * as React from 'react';
import styles from './CardsComponent.module.scss';
import { ICardsComponentProps } from './ICardsComponentProps';
import { escape } from '@microsoft/sp-lodash-subset';

const CardsComponent: React.FC<ICardsComponentProps> = (props) => {
  const cards = props.cards || [];

  return (
    <div className={styles.CanvasSection}>
    <div className={styles.cardsComponent}>
      <h2 className={styles.title}>{escape(props.title || '6 Core Values')}</h2>
      <p className={styles.intro}>{escape(props.intro || 'At Bullwealth, our core values guide everything we do. They define who we are as a company and how we serve our clients.')}</p>
      <div className={styles.grid}>
        {cards.map((card, index) => (
          <div key={`card-${index}`} className={styles.card}>
            <div className={styles.iconContainer}>
              <img src={card.iconUrl} alt={escape(card.cardTitle)} className={styles.icon} />
            </div>
            <h3 className={styles.cardTitle}>{escape(card.cardTitle)}</h3>
            <p className={styles.cardDescription}>{escape(card.cardDescription)}</p>
          </div>
        ))}
      </div>
    </div>
    </div>
  );
};

export default CardsComponent;