import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpFxPagePropertiesWebPartStrings';
import SpFxPageProperties from './components/SpFxPageProperties';
import { ISpFxPagePropertiesProps } from './components/ISpFxPagePropertiesProps';
import { PagePropertiesService } from './services/PagePropertiesService';
import { IPagePropertiesService } from './services/IPagePropertiesService';
import { IListColumn } from './models/IListSiteColumn';

export interface ISpFxPagePropertiesWebPartProps {
  description: string;
  selectedPageProperties: string[];
}

export default class SpFxPagePropertiesWebPart extends BaseClientSideWebPart<ISpFxPagePropertiesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _pagePropertiesService: IPagePropertiesService;
  private _pageProperties: Record<string, unknown> = {};
  private _listColumns: IListColumn[] = [];

  public render(): void {

    const element: React.ReactElement<ISpFxPagePropertiesProps> = React.createElement(
      SpFxPageProperties,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        pageProperties: this._pageProperties,
        selectedPageProperties: this.properties.selectedPageProperties || []
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._pagePropertiesService = new PagePropertiesService(this.context);
    try {
      this._listColumns = await this._pagePropertiesService.getListColumns();
      this._pageProperties = await this._pagePropertiesService.getPageProperties();
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

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    
  }

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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldMultiSelect('selectedPageProperties', {
                  key: 'selectedPageProperties',
                  label: 'Select properties to display',
                  options: this._listColumns.map((column: IListColumn) => {
                    return {
                      key: column.internalName,
                      text: column.title
                    } as IPropertyPaneDropdownOption;
                  }),
                  selectedKeys: this.properties.selectedPageProperties
                  
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
