import { ISPLists } from "../common/SPEntities";
export declare enum LibsOrderBy {
    Id = 1,
    Title = 2,
}
/**
 * Options used to sort and filter
 */
export interface ILibsOptions {
    orderBy?: LibsOrderBy;
    baseTemplate?: number;
    includeHidden?: boolean;
    filter?: string;
}
export interface ISPService {
    /**
     * Get the lists from SharePoint
     * @param options Options used to order and filter during the API query
     */
    getLibs(options?: ILibsOptions): Promise<ISPLists>;
    getListItems?(filterText: string, listId: string, internalColumnName: string, keyInternalColumnName?: string, webUrl?: string): Promise<any[]>;
}
