import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { PrincipalType, IPeoplePickerUserItem } from "../PeoplePicker";
/**
 * Service implementation to search people in SharePoint
 */
export default class SPPeopleSearchService {
    private context;
    private cachedPersonas;
    private cachedLocalUsers;
    /**
     * Service constructor
     */
    constructor(context: WebPartContext | ExtensionContext);
    /**
     * Generate the user photo link using SharePoint user photo endpoint.
     *
     * @param value
     */
    generateUserPhotoLink(value: string): string;
    /**
     * Generate sum of principal types
     *
     * PrincipalType controls the type of entities that are returned in the results.
     * Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
     * These values can be combined (example: 13 is security + SP groups + users)
     *
     * @param principalTypes
     */
    getSumOfPrincipalTypes(principalTypes: PrincipalType[]): number;
    /**
     * Retrieve the specified group
     *
     * @param groupName
     * @param siteUrl
     */
    getGroupId(groupName: string, siteUrl?: string): Promise<number | null>;
    /**
     * Search person by its email or login name
     */
    searchPersonByEmailOrLogin(email: string, principalTypes: PrincipalType[], siteUrl?: string, groupId?: number, ensureUser?: boolean): Promise<IPeoplePickerUserItem>;
    /**
     * Search All Users from the SharePoint People database
     */
    searchPeople(query: string, maximumSuggestions: number, principalTypes: PrincipalType[], siteUrl?: string, groupId?: number, ensureUser?: boolean): Promise<IPeoplePickerUserItem[]>;
    /**
     * Local site search
     */
    private localSearch(siteUrl, query, principalTypes, showHiddenInUI, groupName, exactMatch?);
    /**
     * Tenant search
     */
    private searchTenant(siteUrl, query, maximumSuggestions, principalTypes, ensureUser, groupId);
    /**
     * Retrieves the local user ID
     *
     * @param userId
     */
    private ensureUser(userId);
    /**
     * Generates Initials from a full name
     */
    private getFullNameInitials(fullName);
    /**
     * Gets the user photo url
     */
    private getUserPhotoUrl(userEmail, siteUrl);
    /**
     * Returns fake people results for the Mock mode
     */
    private searchPeopleFromMock(query);
}
