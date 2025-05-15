import * as React from 'react';
import styles from './SpFxPageProperties.module.scss';
import type { ISpFxPagePropertiesProps } from './ISpFxPagePropertiesProps';

export default class SpFxPageProperties extends React.Component<ISpFxPagePropertiesProps> {
  public render(): React.ReactElement<ISpFxPagePropertiesProps> {
    const {
      hasTeamsContext,
      pageProperties,
      selectedPageProperties
    } = this.props;

    return (
      <section className={`${styles.spFxPageProperties} ${hasTeamsContext ? styles.teams : ''}`}>
          <h3>Selected Page Properties:</h3>
          {selectedPageProperties && selectedPageProperties.length > 0 ? (
            <div className={styles.links}>
              {selectedPageProperties.map(column => {
                const value = pageProperties[column];

                let displayValue = value || '';
                // console.log(column);
                // console.log(typeof displayValue);
                if (typeof displayValue === 'object') {
                  displayValue = JSON.stringify(value);
                }

                return (
                  <div key={column} className={styles.links}>
                    <strong>{column}:</strong> {displayValue}
                  </div>
                );
              })}
            </div>
          ) : (
            <div>No properties selected. Please configure the web part to select page properties to display.</div>
          )}
      </section>
    );
  }
}
