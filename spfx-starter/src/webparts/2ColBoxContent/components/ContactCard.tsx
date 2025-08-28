import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './2ColBoxContent.module.scss';
import { IContactCard } from './I2ColBoxContentProps';

export interface IContactCardProps extends IContactCard {
  onEmailClick: (email: string) => void;
  onPhoneClick: (phone: string) => void;
}

export const ContactCard: React.FC<IContactCardProps> = (props) => {
  const {
    title,
    subtitle,
    name,
    email,
    phone,
    emailButtonText,
    phoneButtonText,
    showEmailButton,
    showPhoneButton,
    cardBackgroundColor,
    titleColor,
    subtitleColor,
    nameColor,
    contactColor,
    emailButtonColor,
    phoneButtonColor,
    onEmailClick,
    onPhoneClick
  } = props;

  const cardStyle: React.CSSProperties = {
    backgroundColor: cardBackgroundColor || '#ffffff'
  };

  const titleStyle: React.CSSProperties = {
    color: titleColor || '#323130'
  };

  const subtitleStyle: React.CSSProperties = {
    color: subtitleColor || '#605e5c'
  };

  const nameStyle: React.CSSProperties = {
    color: nameColor || '#323130'
  };

  const contactStyle: React.CSSProperties = {
    color: contactColor || '#0078d4'
  };

  const emailButtonStyle: React.CSSProperties = {
    backgroundColor: emailButtonColor || '#5cb85c'
  };

  const phoneButtonStyle: React.CSSProperties = {
    backgroundColor: phoneButtonColor || '#5bc0de'
  };

  return (
    <div className={styles.contactCard} style={cardStyle}>
      <div className={styles.cardHeader}>
        <div className={styles.headerContent}>
          <div className={styles.titleSection}>
            <h3 className={styles.cardTitle} style={titleStyle}>
              {title || 'Technical Support (BullWealth)'}
            </h3>
            <p className={styles.cardSubtitle} style={subtitleStyle}>
              {subtitle || 'Contact information for technical assistance'}
            </p>
          </div>
          <div className={styles.headerButtons}>
            {showEmailButton && (
              <button
                className={styles.headerButton}
                style={emailButtonStyle}
                onClick={() => onEmailClick(email)}
                type="button"
              >
                <Icon iconName="Mail" className={styles.buttonIcon} />
                {emailButtonText || 'Email'}
              </button>
            )}
            {showPhoneButton && (
              <button
                className={styles.headerButton}
                style={phoneButtonStyle}
                onClick={() => onPhoneClick(phone)}
                type="button"
              >
                <Icon iconName="Phone" className={styles.buttonIcon} />
                {phoneButtonText || 'Phone'}
              </button>
            )}
          </div>
        </div>
      </div>

      <div className={styles.cardBody}>
        <div className={styles.contactInfo}>
          <div className={styles.nameSection}>
            <span className={styles.nameLabel}>Name: </span>
            <span className={styles.nameValue} style={nameStyle}>
              {name || 'Jolari HD'}
            </span>
          </div>

          <div className={styles.contactDetails}>
            <div className={styles.contactItem}>
              <Icon iconName="Mail" className={styles.contactIcon} />
              <a href={`mailto:${email}`} className={styles.contactLink} style={contactStyle}>
                {email || 'joralad@company.com'}
              </a>
            </div>
            <div className={styles.contactItem}>
              <Icon iconName="Phone" className={styles.contactIcon} />
              <a href={`tel:${phone}`} className={styles.contactLink} style={contactStyle}>
                {phone || '555-123-4567'}
              </a>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
