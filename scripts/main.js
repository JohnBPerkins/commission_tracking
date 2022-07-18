import { downloadFile } from './download.mjs';

// Create require
import { createRequire } from "module";
const require = createRequire(import.meta.url);

const config = require('../config.json');
const excel = require('exceljs');

const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { ClientSecretCredential } = require("@azure/identity");
require('isomorphic-fetch');
const d = new Date();

const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});

const monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
];

const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider
});

const options = {
    authProvider,
};

//console.log(downloadFile())

async function getDrives(siteId) {
    let drives = await client.api(`/sites/${siteId}/drives`)
        .get();
    return drives.value[0].id;
}

async function getFolders(siteId) {
    let children = await client.api(`/sites/${siteId}/drive/root/children`)
        .get();
    return children
}

function checkExists(items, name) {
    for (let i = 0; i < items.value.length; i++) {
        if (items.value[i].name == name) {
            return items.value[i].id;
        }
    }
}

async function getItem(siteId, itemId, name) {
    let items = await client.api(`/sites/${siteId}/drive/items/${itemId}/children`).get();
    for (let i = 0; i < items.value.length; i++) {
        if (items.value[i].name == name)
            return items.value[i].id;
    }
}

async function main() {
    let site = await client.api('/sites/root')
        .get();
    let folders = await getFolders(site.id);
    let driveId = await getDrives(site.id)
    let folderId = '';

    // Get root folder
    for (var i = 0; i < folders.value.length; i++) {
        if (folders.value[i].name === 'Company Open Access Files') {
            folderId = folders.value[i].id;
            break;
        }
    }

    // Check if main folder exists
    let items = await client.api(`/drives/${driveId}/items/${folderId}/children`).get();
    let reportsId = checkExists(items, 'Commission Reports');
    // If main folder doesn't exist
    if (reportsId == null) {
        try {
            const folder = {
                name: 'Commission Reports',
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            var response = await client.api(`/drives/${driveId}/items/${folderId}/children`)
                .post(folder);
            reportsId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Check if year folder exists
    items = await client.api(`/drives/${driveId}/items/${reportsId}/children`).get();
    let year = d.getFullYear();
    let yearId = checkExists(items, year);
    // If year folder doesn't exist
    if (yearId == null) {
        try {
            const folder = {
                name: year.toString(),
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            var response = await client.api(`/drives/${driveId}/items/${reportsId}/children`)
                .post(folder);
            yearId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Check if month folder exists
    items = await client.api(`/drives/${driveId}/items/${yearId}/children`).get();
    let month = d.getMonth();
    let monthId = checkExists(items, monthNames[month]);
    if (monthId == null) {
        try {
            const folder = {
                name: monthNames[month],
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            var response = await client.api(`/drives/${driveId}/items/${yearId}/children`)
                .post(folder);
            monthId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Checks if the template file exists in the folder
    let templateId = await getItem(site.id, reportsId, 'template.xlsx');
    if (templateId == null) {
        console.log('Error: No template file in Commission Reports')
        return;
    }

    // Download template
    var result = await client.api(`/drive/items/${templateId}?select=id,@microsoft.graph.downloadUrl`).get();

    await downloadFile(result['@microsoft.graph.downloadUrl'], './temp/test.xlsx')

    var workbook = new excel.Workbook();
    workbook.xlsx.readFile('./temp/test.xlsx');
}

main();