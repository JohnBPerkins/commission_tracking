import { downloadFile } from './download.mjs';

// Create require
import { createRequire } from "module";
const require = createRequire(import.meta.url);

const config = require('../config.json');
const Excel = require('exceljs');

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

var employees = {};
var researcherCommission;

const client = Client.initWithMiddleware({
    debugLogging: false,
    authProvider
});

async function getItem(siteId, itemId, name) {
    let items = await client.api(`/sites/${siteId}/drive/items/${itemId}/children`).get();
    for (let i = 0; i < items.value.length; i++) {
        if (items.value[i].name == name)
            return items.value[i].id;
    }
}

async function readConfig() {
    const config = new Excel.Workbook();
    await config.xlsx.readFile('./temp/config.xlsx');

    const configSheet = config.getWorksheet('Main');

    configSheet.eachRow(function (row, rowNumber) {
        if (rowNumber != 1) {
            let contents = row.values;
            contents.shift();

            if (!employees[contents[0]]) {
                if (contents[2] == null)
                    contents[2] = 0;
                employees[contents[0]] = [{ [contents[4]]: contents[1] }, contents[2], contents[3], contents[5].text];
                if (contents[4] == 'Account Manager' && contents[6] != null)
                    employees[contents[0]].push(contents[6], contents[7], contents[8]);
            } else {
                employees[contents[0]][0][contents[4]] = contents[1];
                if (contents[4] == 'Account Manager' && contents[6] != null)
                    employees[contents[0]].push(contents[6], contents[7], contents[8]);
            }
        } else
            researcherCommission = row.getCell(13).value;
    });
}

async function main() {
    let siteId = await client.api('/sites?search=metrics').get();
    siteId = siteId.value[0].id;
    let driveId = await client.api(`/sites/${siteId}/drives`).get();
    driveId = driveId.value[0].id;

    // Check if main folder exists
    let root = await client.api(`/drives/${driveId}/root`).get();
    let folderId = await getItem(siteId, root.id, 'Commission Reports');
    // If not
    if (folderId == null) {
        try {
            const folder = {
                name: 'Commission Reports',
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            let response = await client.api(`/drives/${driveId}/items/${root.id}/children`)
                .post(folder);
            folderId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Check if year folder exists
    let year = d.getFullYear();
    let yearId = getItem(siteId, folderId, year);
    // If not
    if (yearId == null) {
        try {
            const folder = {
                name: year.toString(),
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            let response = await client.api(`/drives/${driveId}/items/${folderId}/children`)
                .post(folder);
            yearId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Checks if the template file exists in the folder
    let templateId = await getItem(siteId, folderId, 'template.xlsx');
    if (templateId == null) {
        console.log('Error: No template file found in the Commission Reports folder');
        return;
    }

    // Checks if the config file exists in the folder
    let configId = await getItem(siteId, folderId, 'config.xlsx');
    if (configId == null) {
        console.log('Error: No config file found in the Commission Reports folder');
        return;
    }

    // Checks if the billing submission form exists in the folder
    let billingId = await getItem(siteId, folderId, 'billing_submission_form.xlsx');
    if (billingId == null) {
        console.log('Error: No billing submission form found in the Commission Reports folder');
        return;
    }

    // Download template
    try {
        console.log('Downloading template file...')
        let template = await client.api(`sites/${siteId}/drive/items/${templateId}?select=id,@microsoft.graph.downloadUrl`).get();
        await downloadFile(template['@microsoft.graph.downloadUrl'], './temp/template.xlsx')
    } catch (error) {
        console.log(error);
        return;
    }

    // Download config
    try {
        console.log('Downloading config file...');
        let config = await client.api(`sites/${siteId}/drive/items/${configId}?select=id,@microsoft.graph.downloadUrl`).get();
        await downloadFile(config['@microsoft.graph.downloadUrl'], './temp/config.xlsx');

        await readConfig();
        console.log('Reading in config settings...');
    } catch (error) {
        console.log(error);
        return;
    }

    // Download billing submission form
    try {
        console.log('Downloading billing submission form...');
        let billing = await client.api(`sites/${siteId}/drive/items/${billingId}?select=id,@microsoft.graph.downloadUrl`).get();
        await downloadFile(billing['@microsoft.graph.downloadUrl'], './temp/billing.xlsx');
    } catch (error) {
        console.log(error);
        return;
    }

    // Load the billing submission form workbook
    const billing = new Excel.Workbook();
    await billing.xlsx.readFile('./temp/billing.xlsx');
    const billingForm = billing.getWorksheet('Billing');

    // Create a new template for each person in the employee dictionary
    for (let key in employees) {
        const template = new Excel.Workbook();
        await template.xlsx.readFile('./temp/template.xlsx');
        const templateCommissionSheet = template.getWorksheet('Commission');
        const templateBillingSheet = template.getWorksheet('Billing');
    }
}

main();