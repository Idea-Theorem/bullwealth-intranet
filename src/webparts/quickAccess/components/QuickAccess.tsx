import * as React from 'react';
import styles from './QuickAccess.module.scss';
import { IQuickAccessProps } from './IQuickAccessProps';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { escape } from '@microsoft/sp-lodash-subset';

const QuickAccess: React.FC<IQuickAccessProps> = (props) => {
  const quickLinks = props.quickLinks || [];
  const supports = props.supports || [];

  const firstSupport = supports[0] || { name: '', email: '', phone: '' };
  const secondSupport = supports[1] || { name: '', email: '', phone: '' };

  return (
    <div className={styles.CanvasSection}>
    <div className={styles.quickAccess}>
      <h2 className={styles.title}>{escape(props.title || 'Quick Access')}</h2>
      <div className={styles.container}>
        <div className={styles.left}>
          {quickLinks.map((link, index) => (
            <a key={`link-${index}`} href={link.url || '#'} className={styles.tile} target="_blank" rel="noopener noreferrer">
              <img src={link.iconUrl} alt={escape(link.title)} className={styles.icon} />
              <div className={styles.tileTitle}>{escape(link.title)}</div>
            </a>
          ))}
        </div>
        <div className={styles.right}>
          <div className={styles.support}>
            <div className={styles.supportTitle}>{escape(firstSupport.name)}</div>
            {firstSupport.email && (
              <div className={styles.supportItem1}>
                <Icon iconName="Mail" className={styles.supportIcon} />
                {escape(firstSupport.email)}
              </div>
            )}
            {firstSupport.phone && (
              <div className={styles.supportItem}>
                <Icon iconName="Phone" className={styles.supportIcon} />
                {escape(firstSupport.phone)}
              </div>
            )}
          </div>
          <div className={styles.support}>
            <div className={styles.supportTitle}>{escape(secondSupport.name)}</div>
            {secondSupport.email && (
              <div className={styles.supportItem1}>
                <Icon iconName="Mail" className={styles.supportIcon} />
                {escape(secondSupport.email)}
              </div>
            )}
            {secondSupport.phone && (
              <div className={styles.supportItem}>
                <Icon iconName="Phone" className={styles.supportIcon} />
                {escape(secondSupport.phone)}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
    </div>
  );
};

export default QuickAccess;