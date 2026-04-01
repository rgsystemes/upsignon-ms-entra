"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MicrosoftGraph = void 0;
const identity_1 = require("@azure/identity");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const azureTokenCredentials_1 = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
class MicrosoftGraph {
    static _instances = {};
    static _groupConfig = {};
    // NB: do not try catch requests to Microsoft Graph to avoid deactivating users due to throttling limit
    // THROTTLING LIMITS: https://learn.microsoft.com/en-us/graph/throttling-limits
    static getMSEntraConfigForGroup;
    static reloadInstanceForGroup(groupId) {
        delete MicrosoftGraph._instances[groupId];
    }
    static async getUserId(groupId, userEmail) {
        const graph = await MicrosoftGraph._getInstance(groupId, true);
        if (graph) {
            const userId = await graph._getUserIdFromEmail(userEmail);
            return userId;
        }
        return null;
    }
    static async isUserAuthorizedForUpSignOn(groupId, userId) {
        const graph = await MicrosoftGraph._getInstance(groupId, false);
        if (graph) {
            const isAuthorized = await graph.isUserAuthorizedForUpSignOn(userId);
            return isAuthorized;
        }
        return false;
    }
    static async getGroupsForUser(groupId, userId) {
        const graph = await MicrosoftGraph._getInstance(groupId, false);
        if (graph) {
            const groups = await graph.getGroupsForUser(userId);
            return groups;
        }
        return [];
    }
    static async getAllUsersAssignedToUpSignOn(groupId, withoutConfigRefresh) {
        const graph = await MicrosoftGraph._getInstance(groupId, withoutConfigRefresh);
        if (graph) {
            const allUsers = await graph.getAllUsersAssignedToUpSignOn();
            return allUsers;
        }
        return [];
    }
    static async _getInstance(groupId, withoutConfigRefresh) {
        if (!withoutConfigRefresh && MicrosoftGraph._instances[groupId]) {
            return MicrosoftGraph._instances[groupId];
        }
        const entraConfig = await MicrosoftGraph.getMSEntraConfigForGroup(groupId);
        if (!MicrosoftGraph._instances[groupId] || MicrosoftGraph._hasConfigChanged(groupId, entraConfig)) {
            if (entraConfig?.tenantId && entraConfig.clientId && entraConfig.clientSecret && entraConfig.appResourceId) {
                MicrosoftGraph._instances[groupId] = new _MicrosoftGraph(entraConfig.tenantId, entraConfig.clientId, entraConfig.clientSecret, entraConfig.appResourceId);
            }
            else {
                delete MicrosoftGraph._instances[groupId];
            }
            MicrosoftGraph._groupConfig[groupId] = entraConfig;
        }
        return MicrosoftGraph._instances[groupId] || null;
    }
    static _hasConfigChanged(groupId, currentConfig) {
        const cachedConfig = MicrosoftGraph._groupConfig[groupId];
        if (currentConfig == null && cachedConfig == null)
            return false;
        if (currentConfig?.tenantId != cachedConfig?.tenantId ||
            currentConfig?.clientId != cachedConfig?.clientId ||
            currentConfig?.clientSecret != cachedConfig?.clientSecret ||
            currentConfig?.appResourceId != cachedConfig?.appResourceId) {
            return true;
        }
        return false;
    }
    static listNeededAPIs() {
        return [
            {
                path: "/users",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/groups",
                docLink: "https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/users/{id | userPrincipalName}/appRoleAssignments",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list-approleassignments?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/servicePrincipals/{id}/appRoleAssignedTo",
                docLink: "https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignedto?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/groups/{id}/members/microsoft.graph.user",
                docLink: "https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/users/{id}/memberOf/microsoft.graph.group",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http",
            },
        ];
    }
}
exports.MicrosoftGraph = MicrosoftGraph;
class _MicrosoftGraph {
    msGraph;
    appResourceId;
    /**
     *
     * @param tenantId - The Microsoft Entra tenant (directory) ID.
     * @param clientId - The client (application) ID of an App Registration in the tenant.
     * @param clientSecret - A client secret that was generated for the App Registration.
     * @param appResourceId - Identifier of the ressource (UpSignOn) in the graph that users need to have access to in order to be authorized to use an UpSignOn licence
     */
    constructor(tenantId, clientId, clientSecret, appResourceId) {
        const credential = new identity_1.ClientSecretCredential(tenantId, clientId, clientSecret);
        const authProvider = new azureTokenCredentials_1.TokenCredentialAuthenticationProvider(credential, {
            // The client credentials flow requires that you request the
            // /.default scope, and pre-configure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            scopes: ["https://graph.microsoft.com/.default"],
        });
        const clientOptions = {
            authProvider,
        };
        this.msGraph = microsoft_graph_client_1.Client.initWithMiddleware(clientOptions);
        this.appResourceId = appResourceId;
    }
    /**
     * Gets the id of the first user to match that email address and who has been assigned the role for using UpSignOn
     *
     * @param email
     * @returns the id if such a user exists, null otherwise
     */
    async _getUserIdFromEmail(email) {
        if (!email.match(/^[\w-\.+]+@([\w-]+\.)+[\w-]{2,4}$/)) {
            throw "Email is malformed";
        }
        const users = await this.msGraph
            // PERMISSION = User.Read.All OR Directory.Read.All
            .api("/users")
            .header("ConsistencyLevel", "eventual")
            .filter(`mail eq '${email}' or userPrincipalName eq '${email}' or otherMails/any(oe:oe eq '${email}')`)
            .select(["id"])
            .get();
        const userId = users.value[0]?.id;
        return userId;
    }
    async isUserAuthorizedForUpSignOn(userId) {
        // PERMISSION = Directory.Read.All
        // const appRoleAssignments = await this.msGraph
        //   .api(`/users/${userId}/appRoleAssignments`) // this also works if the user is a direct member of a group assigned to UpSignOn
        //   .header("ConsistencyLevel", "eventual")
        //   .filter(`resourceId eq ${this.appResourceId}`)
        //   .get();
        // return appRoleAssignments.value.filter((as: any) => !as.deletedDateTime).length > 0;
        // ALTERNATIVE METHOD
        const allAuthorizedUserIds = await this.getAllUsersAssignedToUpSignOn();
        return allAuthorizedUserIds.indexOf(userId) >= 0;
    }
    async getAllUsersAssignedToUpSignOn() {
        const allPrincipalsRes = await this.msGraph
            // PERMISSION = Application.Read.All OR Directory.Read.All
            // https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignedto?view=graph-rest-1.0&tabs=http
            .api(`/servicePrincipals/${this.appResourceId}/appRoleAssignedTo`)
            .header("ConsistencyLevel", "eventual")
            .select(["principalType", "principalId"])
            .get();
        let allUsersId = allPrincipalsRes.value
            .filter((u) => u.principalType === "User")
            .map((u) => u.principalId);
        const allGroups = allPrincipalsRes.value.filter((u) => u.principalType === "Group");
        for (let i = 0; i < allGroups.length; i++) {
            const g = allGroups[i];
            const allGroupUsersRes = await this.listGroupMembers(g.principalId);
            allUsersId = [...allUsersId, ...allGroupUsersRes.map((u) => u.id)];
        }
        return allUsersId;
    }
    /**
     * Returns all groups (and associated groups) that this user belongs to
     * To be used for sharing to teams ?
     * This would suppose a user can only shared to teams to which it belongs ?
     * @param email
     * @returns
     */
    async getGroupsForUser(userId) {
        const groups = await this.msGraph
            // Get groups, directory roles, and administrative units that the user is a transitive member of.
            // PERMISSION = Directory.Read.All OR GroupMember.Read.All OR Directory.Read.All
            // https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http
            // .api(`/users/${userId}/memberOf`) // pour tout avoir
            // .api(`/users/${userId}/memberOf/microsoft.graph.administrativeUnit`) // pour avoir tous les administrativeUnit
            .api(`/users/${userId}/transitiveMemberOf/microsoft.graph.group`) // pour avoir tous les groupes
            .header("ConsistencyLevel", "eventual")
            .select(["id", "displayName"])
            .get();
        return groups.value;
    }
    /**
     * Returns all members of a group
     * @returns
     */
    async listGroupMembers(groupId) {
        // Get a list of the group's transitive members. A group can have users, organizational contacts, devices, service principals and other groups as members. This operation is not transitive.
        // PERMISSION = GroupMember.Read.All OR Group.Read.All OR Directory.Read.All
        // https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0&tabs=http
        const groupMembers = await this.msGraph
            .api(`/groups/${groupId}/transitiveMembers/microsoft.graph.user/`)
            .header("ConsistencyLevel", "eventual")
            .select(["id", "mail", "displayName"])
            .get();
        return groupMembers.value;
    }
    async checkGroupMembers(groupIds) {
        // PERMISSION = GroupMember.Read.All OR Group.Read.All
        const allGroups = await this.msGraph
            .api(`/groups`)
            .header("ConsistencyLevel", "eventual")
            .filter(`id in ('${groupIds.join("', '")}')`)
            .expand("members($select=id, displayName, mail)")
            .select(["id", "displayName"])
            .get();
        // beware, that mail could be empty although the user may have another email
        return allGroups.value;
        // When sharing to a group, there should be a check that verifies new users in that group and removed users from that group to adapt sharing
    }
}
