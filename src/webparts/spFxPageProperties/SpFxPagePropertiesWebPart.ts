import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'SpFxPagePropertiesWebPartStrings';
import PageProperties from './components/PageProperties';
import { IPagePropertiesProps } from './components/IPagePropertiesProps';
import { PagePropertiesService } from './services/PagePropertiesService';
import { IPagePropertiesService } from './services/IPagePropertiesService';
import { IListColumn, IListColumnWithValue } from './models/IListSiteColumn';


export interface ISpFxPagePropertiesWebPartProps {
  title: string;
  selectedPageProperties: string[];
}

export default class SpFxPagePropertiesWebPart extends BaseClientSideWebPart<ISpFxPagePropertiesWebPartProps> {
  private _theme: IReadonlyTheme | undefined;
  private _pagePropertiesService: IPagePropertiesService;
  private _pageProperties: IListColumnWithValue[] = [];
  private _listColumns: IListColumn[] = [];

  public render(): void {

    const element: React.ReactElement<IPagePropertiesProps> = React.createElement(PageProperties, {
      theme: this._theme || {},
      title: this.properties.title,
      displayMode: this.displayMode,
      updateTitle: (value: string) => {
        this.properties.title = value;
      },
      pageProperties: this._pageProperties,
      selectedPageProperties: this.properties.selectedPageProperties || [],
    });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = themeProvider.tryGetTheme();

    this._pagePropertiesService = new PagePropertiesService(this.context);
    try {
      this._listColumns = await this._pagePropertiesService.getListColumns();
      this._pageProperties = await this._pagePropertiesService.getPageProperties(this._listColumns);
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error fetching page properties:', error);
    }

    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._theme = currentTheme;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
      this.domElement.style.setProperty('--neutralLighterAlt', currentTheme.palette?.neutralLighterAlt || null);
      this.domElement.style.setProperty('--neutralPrimary', currentTheme.palette?.neutralPrimary || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Called when the property pane is opened to ensure options are loaded.
   */
  protected async onPropertyPaneConfigurationStart(): Promise<void> {}

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedPageProperties') {
      this.properties.selectedPageProperties = newValue;
      this.render();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.PropertyPaneTitleLabel,
                }),
                PropertyFieldMultiSelect('selectedPageProperties', {
                  key: 'selectedPageProperties',
                  label: 'Select properties to display',
                  options: this._listColumns.map((column: IListColumn) => {
                    return {
                      key: column.internalName,
                      text: column.title,
                    } as IPropertyPaneDropdownOption;
                  }),
                  selectedKeys: this.properties.selectedPageProperties,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
