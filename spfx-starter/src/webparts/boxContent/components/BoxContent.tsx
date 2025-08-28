import * as React from 'react';
import styles from './BoxContent.module.scss';
import type { IBoxContentProps } from './IBoxContentProps';
import { Icon } from '@fluentui/react/lib/Icon';

export default class BoxContent extends React.Component<IBoxContentProps, {}> {
  
  private handleButtonClick = (): void => {
    const { buttonUrl } = this.props;
    if (buttonUrl) {
      // Check if it's an external URL or internal SharePoint link
      if (buttonUrl.startsWith('http://') || buttonUrl.startsWith('https://')) {
        window.open(buttonUrl, '_blank', 'noopener,noreferrer');
      } else {
        window.location.href = buttonUrl;
      }
    }
  };

  public render(): React.ReactElement<IBoxContentProps> {
    const {
      title,
      description,
      duration,
      buttonText,
      buttonIcon,
      backgroundColor,
      titleColor,
      descriptionColor,
      buttonColor,
      showDuration,
      hasTeamsContext
    } = this.props;

    const containerStyle: React.CSSProperties = {
      backgroundColor: backgroundColor || '#ffffff'
    };

    const titleStyle: React.CSSProperties = {
      color: titleColor || '#323130'
    };

    const descriptionStyle: React.CSSProperties = {
      color: descriptionColor || '#605e5c'
    };

    const buttonStyle: React.CSSProperties = {
      backgroundColor: buttonColor || '#5cb85c',
      borderColor: buttonColor || '#5cb85c'
    };

    return (
      <section className={`${styles.boxContent} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.contentBox} style={containerStyle}>
          <div className={styles.header}>
            <div className={styles.titleSection}>
              <h2 className={styles.title} style={titleStyle}>
                {title || 'HR Platform Introduction'}
              </h2>
              <p className={styles.description} style={descriptionStyle}>
                {description || 'Get started with our HR platform. HR Platform Introductory Call Recording'}
              </p>
            </div>
            <div className={styles.actionSection}>
              <button 
                className={styles.watchButton}
                style={buttonStyle}
                onClick={this.handleButtonClick}
                type="button"
              >
                {buttonIcon && (
                  <Icon 
                    iconName={buttonIcon} 
                    className={styles.buttonIcon}
                  />
                )}
                {buttonText || 'Watch'}
              </button>
            </div>
          </div>
          
          {showDuration && duration && (
            <div className={styles.footer}>
              <span className={styles.duration}>
                {duration}
              </span>
            </div>
          )}
        </div>
      </section>
    );
  }
}
