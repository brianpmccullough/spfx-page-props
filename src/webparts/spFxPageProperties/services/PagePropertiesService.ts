import { SPFx, spfi, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/content-types';
import { IPagePropertiesService } from './IPagePropertiesService';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IListColumn } from '../models/IListSiteColumn';

/**
 * Service to retrieve the current page's properties using @pnp/sp and SPFx context.
 *
 * This service is designed to be instantiated in the WebPart or Customizer and passed to React components
 * via props, promoting testability and separation of SharePoint-specific logic from UI components.
 *
 * The service uses only instance fields (_sp, _pageId, _listId) initialized from the provided context.
 * It does not check for _pageContext in getPageProperties, and does not use any SharePoint-specific logic in React components.
 */
export class PagePropertiesService implements IPagePropertiesService {
  private _sp: SPFI;
  private _pageId: number | undefined;
  private _listId: string | undefined;

  /**
   * Initializes the service and sets up @pnp/sp with the provided SPFx context.
   *
   * @param context The SPFx BaseComponentContext (WebPart or Customizer context).
   * The context is used to initialize the PnPjs SPFI instance and extract the current page's list and item IDs.
   */
  constructor(context: BaseComponentContext) {
    this._sp = spfi().using(SPFx(context));
    this._pageId = context.pageContext?.listItem?.id;
    this._listId = context.pageContext?.list?.id.toString();
  }

  /**
   * Gets the properties of the current page from the associated list.
   *
   * @returns Promise resolving to a dictionary of page properties.
   * @throws Error if the service is not initialized (i.e., if _sp, _pageId, or _listId are not set).
   *
   * @example
   *   const props = await pagePropertiesService.getPageProperties();
   *   console.log(props.Title);
   */
  public async getPageProperties(
    listColumns: IListColumn[] = []
  ): Promise<Record<string, unknown>> {
    if (this._sp && this._pageId && this._listId) {
      const lookupColumnTypes = ['User', 'UserMulti', 'Lookup', 'LookupMulti'];

      const fieldNames = listColumns
        .map(function (column) {
          if (column.fieldType === 'User' || column.fieldType === 'UserMulti') {
            return [
              `${column.internalName}/Title`,
              `${column.internalName}/Id`,
              `${column.internalName}/UserName`,
            ];
          }
          return [column.internalName];
        })
        .reduce(function (acc, val) {
          return acc.concat(val);
        }, []);

      const expandFields = listColumns
        .filter((column) => {
          return lookupColumnTypes.indexOf(column.fieldType) > -1;
        })
        .map((column) => column.internalName);

      const item = await this._sp.web.lists
        .getById(this._listId)
        .items.getById(this._pageId)
        .select(fieldNames.join(','))
        .expand(expandFields.join(','))();
      return item;
    }
    throw new Error('Service not initialized.');
  }

  /**
   * Gets all site columns in all content types in the current list, excluding the 'Folder' content type.
   *
   * @returns Promise resolving to an array of IListSiteColumn objects for each field.
   * @throws Error if the service is not initialized (i.e., if _sp or _listId are not set).
   */
  public async getListColumns(): Promise<IListColumn[]> {
    if (!this._sp || !this._listId) {
      throw new Error('Service not initialized.');
    }

    const contentTypes = await this._sp.web.lists
      .getById(this._listId)
      .contentTypes()
      .then((types) => types.filter((ct) => ct.Name !== 'Folder'));

    // Use a Set to avoid duplicate field IDs
    const fieldIdSet = new Set<string>();
    const columns: IListColumn[] = [];

    for (const ct of contentTypes) {
      // Get fields for each content type
      const fields = await this._sp.web.lists
        .getById(this._listId)
        .contentTypes.getById(ct.Id.StringValue)
        .fields.select(
          'Id',
          'Title',
          'InternalName',
          'TypeAsString',
          'Hidden',
          'Group'
        )();
      for (const field of fields) {
        if (!fieldIdSet.has(field.Id)) {
          if (
            (field.Hidden ||
              field.Group === '_Hidden' ||
              field.Title === 'Document Modified By' ||
              field.Title === 'Document Created By') &&
            field.InternalName !== 'Description' &&
            field.InternalName !== 'Modified' &&
            field.InternalName !== 'Created' &&
            field.InternalName !== 'Author' &&
            field.InternalName !== 'Editor' &&
            field.InternalName !== 'FileRef'
          ) {
            continue;
          }

          columns.push({
            id: field.Id,
            title: field.Title,
            internalName: field.InternalName,
            fieldType: field.TypeAsString,
            hidden: field.Hidden,
            group: field.Group,
          });
          fieldIdSet.add(field.Id);
        }
      }
    }
    return columns;
  }
}
