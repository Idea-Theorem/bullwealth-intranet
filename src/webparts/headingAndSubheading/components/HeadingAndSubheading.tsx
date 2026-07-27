import * as React from 'react';
import styles from './HeadingAndSubheading.module.scss';
import type { IHeadingAndSubheadingProps } from './IHeadingAndSubheadingProps';

export default class HeadingAndSubheading extends React.Component<IHeadingAndSubheadingProps, {}> {
  public render(): React.ReactElement<IHeadingAndSubheadingProps> {
    const {
      heading,
      subheading,
      hasTeamsContext
      //userDisplayName
    } = this.props;

    return (
      <section className={`${styles.headingAndSubheading} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          <div className={styles.content}>
            <h1 className={styles.heading}>{heading}</h1>
            <p className={styles.subheading}>
              {subheading}
            </p>
            {/* {userDisplayName && (
              <div className={styles.userInfo}>
                Welcome, {userDisplayName}!
              </div>
            )} */}
          </div>
        </div>
      </section>
    );
  }
}
