import { Client, GraphRequest } from "@microsoft/microsoft-graph-client";
export type EntraConfig = {
    tenantId: string | null;
    clientId: string | null;
    clientSecret: string | null;
    appResourceId: string | null;
};
export type EntraGroup = {
    id: string;
    displayName: string;
};
export type ApiLink = {
    path: string;
    docLink: string;
};
export declare class MicrosoftGraph {
    static _instances: {
        [groupId: number]: _MicrosoftGraph;
    };
    static _groupConfig: {
        [groupId: number]: EntraConfig | null;
    };
    static getMSEntraConfigForGroup: (groupId: number) => Promise<EntraConfig | null>;
    static reloadInstanceForGroup(groupId: number): void;
    static getUserId(groupId: number, userEmail: string): Promise<string | null>;
    static isUserAuthorizedForUpSignOn(groupId: number, userId: string): Promise<boolean>;
    static getGroupsForUser(groupId: number, userId: string): Promise<EntraGroup[]>;
    static getAllUsersAssignedToUpSignOn(groupId: number, withoutConfigRefresh: boolean): Promise<string[]>;
    static _getInstance(groupId: number, withoutConfigRefresh: boolean): Promise<_MicrosoftGraph | null>;
    static _hasConfigChanged(groupId: number, currentConfig: EntraConfig | null): boolean;
    static listNeededAPIs(): ApiLink[];
}
declare class _MicrosoftGraph {
    msGraph: Client;
    appResourceId: string;
    /**
     *
     * @param tenantId - The Microsoft Entra tenant (directory) ID.
     * @param clientId - The client (application) ID of an App Registration in the tenant.
     * @param clientSecret - A client secret that was generated for the App Registration.
     * @param appResourceId - Identifier of the ressource (UpSignOn) in the graph that users need to have access to in order to be authorized to use an UpSignOn licence
     */
    constructor(tenantId: string, clientId: string, clientSecret: string, appResourceId: string);
    /**
     * Graph paginates its collections (100 items by default) and gives the next page in "@odata.nextLink".
     * This follows every page and returns all the items.
     */
    _getAllPages<T>(firstPageRequest: GraphRequest): Promise<T[]>;
    /**
     * Gets the id of the first user to match that email address and who has been assigned the role for using UpSignOn
     *
     * @param email
     * @returns the id if such a user exists, null otherwise
     */
    _getUserIdFromEmail(email: string): Promise<string | null>;
    isUserAuthorizedForUpSignOn(userId: string): Promise<boolean>;
    getAllUsersAssignedToUpSignOn(): Promise<string[]>;
    /**
     * Returns all groups (and associated groups) that this user belongs to
     * To be used for sharing to teams ?
     * This would suppose a user can only shared to teams to which it belongs ?
     * @param email
     * @returns
     */
    getGroupsForUser(userId: string): Promise<EntraGroup[]>;
    /**
     * Returns all members of a group
     * @returns
     */
    listGroupMembers(groupId: string): Promise<{
        id: string;
        displayName: string;
    }[]>;
    checkGroupMembers(groupIds: string[]): Promise<{
        id: string;
        displayName: string;
        members: {
            "@odata.type": string;
            id: string;
            displayName: string;
            mail: string | null;
        }[];
    }[]>;
}
export {};
//# sourceMappingURL=index.d.ts.map