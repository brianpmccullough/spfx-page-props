import * as React from 'react';
import styles from './SpFxPageProperties.module.scss';
import type { ISpFxPagePropertiesProps } from './ISpFxPagePropertiesProps';
import PageProperty from './PageProperty';

export default class SpFxPageProperties extends React.Component<ISpFxPagePropertiesProps> {

  public render(): React.ReactElement<ISpFxPagePropertiesProps> {
    const { pageProperties, selectedPageProperties } = this.props;
    
    const propertiesToDisplay = pageProperties.filter(property =>
      selectedPageProperties.indexOf(property.internalName) !== -1
    );

    console.log(propertiesToDisplay);

    return (
      <div className={styles.spFxPageProperties}>
        {propertiesToDisplay && propertiesToDisplay.length > 0 ? (
          <div>
            {propertiesToDisplay.map((property) => (
              <PageProperty key={property.internalName} {...property} />
            ))}
          </div>
        ) : (
          <div>No properties selected. Please configure the web part to select page properties to display.</div>
        )}
      </div>
    );
  }
}
