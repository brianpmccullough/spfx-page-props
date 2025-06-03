import * as React from 'react';
import { createTheme, ThemeProvider } from '@fluentui/react';
import styles from './PageProperties.module.scss';
import type { IPagePropertiesProps } from './IPagePropertiesProps';
import PageProperty from './PageProperty';
import { IListColumnWithValue } from '../models/IListSiteColumn';

export default class PageProperties extends React.Component<IPagePropertiesProps> {

  public render(): React.ReactElement<IPagePropertiesProps> {
    const { pageProperties, selectedPageProperties, theme } = this.props;
    const fluentTheme = createTheme({
      palette: theme.palette,
      fonts: theme.fonts ?? undefined,
      semanticColors: theme.semanticColors ?? undefined,
      isInverted: theme.isInverted
    });

    const propertiesToDisplay: IListColumnWithValue[] = selectedPageProperties
      .map(selectedProperty => pageProperties.find(p => p.internalName.toLowerCase() === selectedProperty.toLowerCase()))
      .filter((property): property is IListColumnWithValue => property !== undefined);

    console.log(propertiesToDisplay);

    return (
      // https://github.com/microsoft/fluentui/blob/master/packages/react/src/utilities/ThemeProvider/README.md#other-call-outs
      // The `applyTo` prop is set to 'none' to avoid applying the background to the element.
      <ThemeProvider theme={fluentTheme} applyTo='none'>
      <div className={styles.pageProperties}>
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
      </ThemeProvider>
    );
  }
}
