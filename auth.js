const msal = require('@azure/msal-node');

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL Node configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/configuration.md
 */
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.AAD_ENDPOINT + '/' + process.env.TENANT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        redirectUri : process.env.OAUTH_REDIRECT_URI
    }
};

/**
 * With client credentials flows permissions need to be granted in the portal by a tenant administrator.
 * The scope is always in the format '<resource>/.default'. For more, visit:
 * https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
 */

//get requests worked at this scope, but maybe it needs to be changed?
const tokenRequest = {
    scopes: [process.env.GRAPH_ENDPOINT + '/.default'],
};

//URI endpoint to make the findMeetingTimes call
const apiConfigMeetingTimes = {
    uri: process.env.GRAPH_ENDPOINT + '/v1.0/users/cletendre-umassd@devexch.umassp.edu/findMeetingTimes',
};

//URI endpoint to make the free/busy call
const apiConfigGetFreeBusy = {
    uri: process.env.GRAPH_ENDPOINT + '/v1.0/users/cletendre-umassd@devexch.umassp.edu/calendar/getSchedule'
}

const apiConfigCreateMeeting = {
    uri: process.env.GRAPH_ENDPOINT + '/v1.0/users/cletendre-umassd@devexch.umassp.edu/calendar/events'
}

//request to get a delegated acccess token
//asks the user to sign in
//example: https://dzone.com/articles/getting-access-token-for-microsoft-graph-using-oau
//this is a client credential token, we may need a user credential token which needs a user to login

const apiConfigSignInClient = {
    uri: process.env.AAD_ENDPOINT + '/umasspdev.onmicrosoft.com/oauth2/token'
}

//sign in user and get permissions
const apiConfigSignInUser = {
    uri: 'https://login.microsoftonline.com/umasspdev.onmicrosoft.com/oauth2/token'
    //uri: 'https://login.microsoftonline.com/umasspdev.onmicrosoft.com/oauth2/v2.0/authorize?client_id=fb6fd679-50bd-4e50-87f8-d08b0245b577&response_type=code&redirect_uri=http%3A%2F%2Flocalhost%2Fauth%2Fcallback%2F&response_mode=query&scope=calendars.read%20calendars.read.shared%20calendars.readwrite%20calendars.readwrite.shared&state=12345'
}

const apiConfigGetUserToken = {
    uri: process.env.AAD_ENDPOINT + '/umasspdev.onmicrosoft.com/oauth2/v2.0/token'
}

/**
 * Initialize a confidential client application. For more info, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md
 */
const cca = new msal.ConfidentialClientApplication(msalConfig);

/**
 * Acquires token with client credentials.
 * @param {object} tokenRequest
 */
async function getToken(tokenRequest) {
    return await cca.acquireTokenByClientCredential(tokenRequest);
}

//exporting data to other files
module.exports = {
    apiConfigMeetingTimes: apiConfigMeetingTimes,
    apiConfigGetFreeBusy : apiConfigGetFreeBusy,
    apiConfigSignInClient : apiConfigSignInClient,
    apiConfigSignInUser : apiConfigSignInUser,
    apiConfigGetUserToken : apiConfigGetUserToken,
    apiConfigCreateMeeting : apiConfigCreateMeeting,
    tokenRequest: tokenRequest,
    getToken: getToken
};