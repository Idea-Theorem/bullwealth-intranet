import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './CarouselCards.module.scss';
import { ICarouselCardsProps } from './ICarouselCardsProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';

const CarouselCards: React.FC<ICarouselCardsProps> = (props) => {
  const [currentSlide, setCurrentSlide] = useState<number>(0);
  const [isAutoPlaying, setIsAutoPlaying] = useState<boolean>(true);

  const visibleCards = props.cards ? props.cards.filter(card => card.isVisible !== false) : [];

  useEffect(() => {
    let interval: number;
    
    if (isAutoPlaying && visibleCards.length > 0) {
      interval = window.setInterval(() => {
        setCurrentSlide((prev) => (prev + 1) % visibleCards.length);
      }, 5000);
    }

    return () => {
      if (interval) {
        window.clearInterval(interval);
      }
    };
  }, [isAutoPlaying, visibleCards.length]);

  const goToSlide = (index: number): void => {
    setCurrentSlide(index);
    setIsAutoPlaying(false);
    setTimeout(() => setIsAutoPlaying(true), 10000);
  };

  const nextSlide = (): void => {
    setCurrentSlide((prev) => (prev + 1) % visibleCards.length);
    setIsAutoPlaying(false);
    setTimeout(() => setIsAutoPlaying(true), 10000);
  };

  const prevSlide = (): void => {
    setCurrentSlide((prev) => (prev - 1 + visibleCards.length) % visibleCards.length);
    setIsAutoPlaying(false);
    setTimeout(() => setIsAutoPlaying(true), 10000);
  };

  if (!visibleCards || visibleCards.length === 0) {
    return (
      <div className={styles.carouselCards}>
        <div style={{ textAlign: 'center', padding: '40px', color: '#605e5c' }}>
          <Icon iconName="DocumentSet" style={{ fontSize: '48px', marginBottom: '16px' }} />
          <p>No carousel cards configured or all cards are hidden. Please add cards from the web part properties.</p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.carouselCards}>
      <div className={styles.headerSection}>
        <h1 className={styles.mainTitle}>{props.title || 'Corporate Values - Definition and Behaviours'}</h1>
        <p className={styles.subtitle}>
          {props.subtitle || 'At Bullwealth, our core values guide everything we do. They define who we are as a company and how we serve our clients.'}
        </p>
      </div>

      <div className={styles.carouselContainer}>
        {/* ✅ LEFT ARROW BUTTON */}
        {visibleCards.length > 1 && (
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            className={styles.navButton}
            onClick={prevSlide}
            ariaLabel="Previous slide"
            title="Previous"
          />
        )}

        <div className={styles.cardWrapper}>
          {visibleCards.map((card, index) => (
            <div
              key={card.id}
              className={`${styles.card} ${index === currentSlide ? styles.active : ''}`}
              style={{
                opacity: index === currentSlide ? 1 : 0,
                visibility: index === currentSlide ? 'visible' : 'hidden',
                transition: 'opacity 0.5s ease-in-out'
              }}
            >
              <div 
                className={styles.iconContainer}
                style={{ backgroundColor: card.iconColor }}
              >
                {card.iconType === 'upload' && card.icon ? (
                  <img 
                    src={card.icon} 
                    alt={card.title} 
                    className={styles.uploadedIcon}
                  />
                ) : (
                  <Icon iconName={card.icon || 'Lightbulb'} className={styles.cardIcon} />
                )}
              </div>

              <h2 className={styles.cardTitle}>{card.title}</h2>
              <p className={styles.cardDescription}>{card.description}</p>

              {card.bulletPoints && card.bulletPoints.length > 0 && (
                <ul className={styles.bulletList}>
                  {card.bulletPoints.map((point, idx) => (
                    <li key={idx} className={styles.bulletItem}>
                      {point}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          ))}
        </div>

        {/* ✅ RIGHT ARROW BUTTON */}
        {visibleCards.length > 1 && (
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            className={styles.navButton}
            onClick={nextSlide}
            ariaLabel="Next slide"
            title="Next"
          />
        )}
      </div>

      {/* Dots Navigation */}
      {visibleCards.length > 1 && (
        <div className={styles.dotsContainer}>
          {visibleCards.map((card, index) => (
            <button
              key={card.id}
              className={`${styles.dot} ${index === currentSlide ? styles.activeDot : ''}`}
              onClick={() => goToSlide(index)}
              aria-label={`Go to slide ${index + 1}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default CarouselCards;
