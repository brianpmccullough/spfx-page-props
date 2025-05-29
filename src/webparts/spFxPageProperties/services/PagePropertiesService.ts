import { SPFx, spfi, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/content-types';
import '@pnp/sp/fields';
import { IPagePropertiesService } from './IPagePropertiesService';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IListColumn, IListColumnWithValue, LookupFieldValue, TaxonomyFieldValue, UserFieldValue } from '../models/IListSiteColumn';
import { IFieldInfo } from '@pnp/sp/fields';

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
   */
  public async getPageProperties(listColumns: IListColumn[] = []): Promise<IListColumnWithValue[]> {
    if (this._sp && this._pageId && this._listId && listColumns.length > 0) {
      const lookupColumnTypes = ['User', 'UserMulti', 'Lookup', 'LookupMulti'];

      const fieldNames = listColumns
        .map(function (column) {
          if (column.fieldType === 'User' || column.fieldType === 'UserMulti') {
            return [`${column.internalName}/Title`, `${column.internalName}/Id`, `${column.internalName}/UserName`, `${column.internalName}/Name`];
          }
          if (column.fieldType === 'Lookup' || column.fieldType === 'LookupMulti') {
            return [`${column.internalName}/Title`, `${column.internalName}/Id`];
          }
          if (column.fieldType === 'TaxonomyFieldType' || column.fieldType === 'TaxonomyFieldTypeMulti') {
            return [`${column.internalName}`, 'TaxCatchAll/ID', 'TaxCatchAll/Term'];
          }
          
          return [column.internalName];
        })
        .reduce(function (allFieldNames, fieldNameArray) {
          return allFieldNames.concat(fieldNameArray);
        }, []);

      const expandFields = listColumns
        .filter((column) => {
          return lookupColumnTypes.indexOf(column.fieldType) > -1;
        })
        .map((column) => column.internalName);
      
      if (fieldNames.indexOf('TaxCatchAll/ID') > -1) {
        expandFields.push('TaxCatchAll');
      }


      const item = await this._sp.web.lists
        .getById(this._listId)
        .items.getById(this._pageId)
        .select(fieldNames.join(','))
        .expand(expandFields.join(','))();


      const values: IListColumnWithValue[] = listColumns.map((column) => {
        const value = item[column.internalName];
        let formattedValue: string | number | boolean | Date | null | UserFieldValue | UserFieldValue[] | LookupFieldValue | LookupFieldValue[] | TaxonomyFieldValue | TaxonomyFieldValue[] = null;
        if (value !== null && value !== undefined) {
          if (column.fieldType === 'User' || column.fieldType === 'UserMulti') {
            const extractInitials = (name: string): string => {
              const parts = name.split(' ');
              return parts.length > 1
                ? parts[0].charAt(0).toUpperCase() + parts[1].charAt(0).toUpperCase()
                : name.charAt(0).toUpperCase();
            };

            if (Array.isArray(value)) {
              formattedValue = value.map((user) => ({
                id: user.Id,
                name: user.Title,
                initials: extractInitials(user.Title),
                userName: user.UserName,
                profilePictureUrl: `/_layouts/15/userphoto.aspx?size=L&accountname=${user.UserName}`
              }));
            } else {
              formattedValue = {
                id: value.Id,
                name: value.Title,
                initials: extractInitials(value.Title),
                userName: value.UserName,
                profilePictureUrl: `/_layouts/15/userphoto.aspx?size=L&accountname=${value.UserName}`
              };
            }
          } else if (column.fieldType === 'Lookup' || column.fieldType === 'LookupMulti') {
            if (Array.isArray(value)) {
              formattedValue = value.map((lookup) => ({
                id: lookup.Id,
                title: lookup.Title,
              }));
            } else {
              formattedValue = {
                id: value.Id,
                title: value.Title,
              };
            }
          } else if (column.fieldType === 'TaxonomyFieldType' || column.fieldType === 'TaxonomyFieldTypeMulti') {
            if (Array.isArray(value)) {
              formattedValue = value.map((term) => {
                return {
                  id: term.TermGuid,
                  term: term.Label,
                };
              });
            } else {
            
              const taxonomyCatchAllField = item.TaxCatchAll as {
                ID: string;
                Term: string;
                TermGuid: string;
                'odata.id': string
              }[];
              const term = taxonomyCatchAllField.find((t) => t.ID === value.WssId);
              formattedValue = term
                ? {
                    id: value.TermGuid,
                    term: term.Term,
                  }
                : null;
            }
          } else if (column.fieldType === 'DateTime') {
            formattedValue = new Date(value);
          } else if (column.fieldType === 'Boolean') {
            formattedValue = value && (value === '1' || value === true);
          } else {
            formattedValue = value;
          }
        }

        return {
          ...column,
          value: formattedValue,
        };
      });

      return values;
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

    // Get all content types except 'Folder', and include their fields
    const contentTypes = await this._sp.web.lists
      .getById(this._listId)
      .contentTypes
      .expand('Fields')()
      .then((types) => types.filter((ct) => ct.Name !== 'Folder'));

    // Use a Set to avoid duplicate field IDs
    const fieldIdSet = new Set<string>();
    const columns: IListColumn[] = [];

    for (const ct of contentTypes) {
      //TODO: log a bug for this in pnp/sp?
      // eslint-disable-next-line dot-notation, @typescript-eslint/no-explicit-any
      const fields = (ct as any)['Fields'] as IFieldInfo[];

      for (const field of fields) {
        if (!fieldIdSet.has(field.Id)) {
          if (
            (field.Hidden 
              || field.Group === '_Hidden'
              || field.InternalName.startsWith('_')
              || field.Title === 'Document Modified By' 
              || field.Title === 'Document Created By'
            ) &&
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
