import * as React from 'react';
import styles from './SpFxPageProperties.module.scss';
import type { ISpFxPagePropertiesProps } from './ISpFxPagePropertiesProps';
import { escape } from '@microsoft/sp-lodash-subset';

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
              {selectedPageProperties.map(columnId => {
                const value = pageProperties[columnId];
                return value ? (
                  <div key={columnId} className={styles.links}>
                    <strong>{columnId}:</strong> {escape(String(value))}
                  </div>
                ) : null;
              })}
            </div>
          ) : (
            <div>No properties selected. Please configure the web part to select page properties to display.</div>
          )}
      </section>
    );
  }
}
