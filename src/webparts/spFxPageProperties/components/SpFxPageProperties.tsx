import * as React from 'react';
import styles from './SpFxPageProperties.module.scss';
import type { ISpFxPagePropertiesProps } from './ISpFxPagePropertiesProps';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import PageProperty from './PageProperty';

export default class SpFxPageProperties extends React.Component<ISpFxPagePropertiesProps> {

  public render(): React.ReactElement<ISpFxPagePropertiesProps> {
    const { pageProperties } = this.props;
    
    return (
      <div className={styles.spFxPageProperties}>
        <WebPartTitle 
          title={this.props.title} 
          updateProperty={this.props.updateTitle} 
          displayMode={this.props.displayMode}
          placeholder='Web Part Title'
          />
        {pageProperties && pageProperties.length > 0 ? (
          <div>
            {pageProperties.map((property) => (
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
