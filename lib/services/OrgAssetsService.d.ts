import { FileBrowserService } from "./FileBrowserService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILibrary, FilesQueryResult } from "./FileBrowserService.types";
import { ExtensionContext } from "@microsoft/sp-extension-base";
export declare class OrgAssetsService extends FileBrowserService {
    private _orgAssetsLibraryServerRelativeSiteUrl;
    constructor(context: ExtensionContext | WebPartContext, itemsToDownloadCount?: number);
    getListItems: (libraryName: string, folderPath: string, acceptedFilesExtensions?: string[], nextPageQueryStringParams?: string) => Promise<FilesQueryResult>;
    getSiteMediaLibraries: (includePageLibraries?: boolean) => Promise<ILibrary[]>;
    private _parseOrgAssetsLibraryItem;
}
