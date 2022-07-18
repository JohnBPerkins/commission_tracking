const msal = require('@azure/msal-node');
const config = require('../config.json');

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: config.AADEndpoint + config.tenantId,
        clientSecret: config.clientSecret,
    }
};

const apiConfig = {
    uri: config.graphEndpoint + 'v1.0/users',
};

const tokenRequest = {
    scopes: [config.graphEndpoint + '.default'],
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getToken() {
    try {
        const authResponse = await cca.acquireTokenByClientCredential(tokenRequest);
        return authResponse.accessToken;
    } catch (error) {
        console.log(error);
    }
}

module.exports = getToken;