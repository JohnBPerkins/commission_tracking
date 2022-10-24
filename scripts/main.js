import { downloadFile } from './download.mjs';

// Create require
import { createRequire } from "module";
const require = createRequire(import.meta.url);

const config = require('../config.json');
const Excel = require('exceljs');

const fs = require('fs');

const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { ClientSecretCredential } = require("@azure/identity");
require('isomorphic-fetch');
const d = new Date();
var year = d.getFullYear();

const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});

const monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
];

var employees = {};
var ops = [];
var researcherCommission;

const client = Client.initWithMiddleware({
    debugLogging: false,
    authProvider
});

async function getItem(siteId, itemId, name) {
    try {
        let items = await client.api(`/sites/${siteId}/drive/items/${itemId}/children`).get();
        for (let i = 0; i < items.value.length; i++) {
            if (items.value[i].name == name)
                return items.value[i].id;
        }
    } catch (error) {
        console.log(error);
        console.log('test');
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

            if (contents[4] == 'Operations Manager')
                ops.push(contents[0]);
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
    fs.access('./temp', (error) => {
        if (error)
            fs.mkdirSync('./temp');
    });

    let siteId = await client.api('/sites?search=metrics').get();
    siteId = siteId.value[0].id;
    let driveId = await client.api(`/sites/${siteId}/drives`).get();
    driveId = driveId.value[0].id;

    // Check if main folder exists
    let root = await client.api(`/drives/${driveId}/root`).get();
    let mainFolderId = await getItem(siteId, root.id, 'Commission Reports');
    // If not
    if (mainFolderId == null) {
        try {
            const folder = {
                name: 'Commission Reports',
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            let response = await client.api(`/drives/${driveId}/items/${root.id}/children`)
                .post(folder);
            mainFolderId = response.id;
        } catch (error) {
            console.log(error);
        }
    }

    // Checks if the template file exists in the folder
    let templateId = await getItem(siteId, mainFolderId, 'template.xlsx');
    if (templateId == null) {
        console.log('Error: No template file found in the Commission Reports folder');
        return;
    }

    // Checks if the config file exists in the folder
    let configId = await getItem(siteId, mainFolderId, 'config.xlsx');
    if (configId == null) {
        console.log('Error: No config file found in the Commission Reports folder');
        return;
    }

    // Checks if the billing submission form exists in the folder
    let billingId = await getItem(siteId, mainFolderId, 'billing_submission_form.xlsx');
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

    // Check if year folder exists
    let yearFolderId = await getItem(siteId, mainFolderId, year);
    // If not
    if (yearFolderId == null) {
        try {
            const folder = {
                name: year.toString(),
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            };

            let response = await client.api(`/drives/${driveId}/items/${mainFolderId}/children`)
                .post(folder);
            yearFolderId = response.id;
        } catch (error) {
            console.log(error)
        }
    }

    // Create a new template for each person in the employee dictionary

    console.log('Initializing employee files...');
    for (let key in employees) {
        // Check that their folder exists
        let employeeFolderId = await getItem(siteId, yearFolderId, key);
        if (employeeFolderId == null) {
            try {
                const folder = {
                    name: key,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'fail'
                };

                let response = await client.api(`/drives/${driveId}/items/${yearFolderId}/children`)
                    .post(folder);
                employeeFolderId = response.id;
            } catch (error) {
                console.log(error);
            }
        }

        // Check if there is already an existing file
        let fileName = key.replaceAll(' ', '_') + '_Report.xlsx';
        let employeeFileId = await getItem(siteId, employeeFolderId, fileName);

        const template = new Excel.Workbook();
        await template.xlsx.readFile('./temp/template.xlsx');

        if (employeeFileId == null) {
            // Initialize files
            employees[key].push(template);
            let length = employees[key].length;
            let billingSheet = employees[key][length - 1].getWorksheet('Billing');
            let worksheet;

            if (employees[key][0].hasOwnProperty('Account Manager')) {
                billingSheet.spliceColumns(8, 2);
                employees[key][length - 1].removeWorksheet('ResearcherCommission');
                worksheet = employees[key][length - 1].getWorksheet('AMCommission');
                if (key != 'Alan Carty') {
                    worksheet.getCell('E5').value = employees[key][5];
                    worksheet.getCell('E6').value = employees[key][6];
                    worksheet.getCell('B7').value = employees[key][4];
                }
            } else if (employees[key][0].hasOwnProperty('Researcher')) {
                employees[key][length - 1].removeWorksheet('AMCommission');
                worksheet = employees[key][length - 1].getWorksheet('ResearcherCommission');
                worksheet.getCell('E5').value = researcherCommission;
                worksheet.getCell('E6').value = employees[key][0]['Researcher'];
                worksheet.getCell('E7').value = employees[key][1];
            } else {
                billingSheet.spliceColumns(8, 2);
                employees[key][length - 1].removeWorksheet('ResearcherCommission');
                worksheet = employees[key][length - 1].getWorksheet('AMCommission');
                worksheet.getCell('E5').value = employees[key][0]['Operations Manager'];
                worksheet.getCell('E6').value = 0;
            }
            worksheet.state = 'visible';
        } else {
            // Download file and read it
            let file = new Excel.Workbook();
            let newFile = new Excel.Workbook();
            let fileURL = await client.api(`sites/${siteId}/drive/items/${employeeFileId}?select=id,@microsoft.graph.downloadUrl`).get();
            await downloadFile(fileURL['@microsoft.graph.downloadUrl'], './temp/file.xlsx');
            // Wait one second
            await delay(1000);
            await file.xlsx.readFile('./temp/file.xlsx');

            let templateBilling = template.getWorksheet('Billing');
            let billingCopy = newFile.addWorksheet('Billing');
            billingCopy.model = Object.assign(templateBilling.model, {
                mergeCells: templateBilling.model.merges
            });

            console.log(`Reading ${key}...`);
            if (employees[key][0].hasOwnProperty('Account Manager') || employees[key][0].hasOwnProperty('Operations Manager')) {
                let original = file.getWorksheet('AMCommission');
                let copy = newFile.addWorksheet('AMCommission');
                copy.model = Object.assign(original.model, {
                    mergeCells: original.model.merges
                });
                billingCopy.spliceColumns(8, 2);
            } else {
                let original = file.getWorksheet('ResearcherCommission');
                let copy = newFile.addWorksheet('ResearcherCommission');
                copy.model = Object.assign(original.model, {
                    mergeCells: original.model.merges
                });
            }
            employees[key].push(newFile);
        }
    }

    // Load the billing submission form workbook
    const billing = new Excel.Workbook();
    await billing.xlsx.readFile('./temp/billing.xlsx');
    const billingSheet = billing.getWorksheet('Billing');
    // Remove unnecessary columns
    // This should be done using tables but is too buggy to use
    billingSheet.spliceColumns(3, 2);
    billingSheet.spliceColumns(4, 1);
    billingSheet.spliceColumns(6, 8);
    billingSheet.spliceColumns(10, 1);
    billingSheet.spliceColumns(16, 7);

    console.log('Filling spreadsheets...')
    billingSheet.eachRow(function (row, rowNumber) {
        // Process all of the billing files
        if (rowNumber != 1) {
            let contents = row.values;
            contents.shift();

            addRow(contents[10], contents, 'Account Manager');
            if (contents[11]) // If recruiter2
                addRow(contents[11], contents, 'Account Manager');
            if (contents[12] != 'None') // If researcher
                addRow(contents[12], contents, 'Researcher');

            // Ops Managers
            for (let manager in ops) {
                addRow(ops[manager], contents, 'Operations Manager');
            }
        }
    });

    console.log('Uploading files...');
    for (let key in employees) {
        try {
            let length = employees[key].length;
            let fileName = key.replaceAll(' ', '_') + '_Report.xlsx';
            let employeeFolderId = await getItem(siteId, yearFolderId, key);
            await employees[key][length - 1].xlsx.writeFile(`./temp/${fileName}`);
            await client.api(`sites/${siteId}/drive/items/${employeeFolderId}:/${fileName}:/content`).put(fs.readFileSync(`./temp/${fileName}`, (error, data) => {
                if (error)
                    console.log(error);
            }));
        } catch (error) {
            console.log(`${JSON.parse(error.body).message}, error with ${key}`)
        }
    }

    console.log('Done!');
}

main();

function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

function calculateSplitInvoice(contents) {
    if (contents[11]) {
        if (contents[11] == 'Split - Top Echelon Office')
            return contents[8] / contents[9] * .47;
        else
            if (contents[11] == 'Account Manager')
                return contents[8] / contents[9] * .5;
            else
                return contents[8] / contents[9]
    } else
        return contents[8] / contents[9];
}

function addRow(employee, contents, role) {

    if (employees[employee]) {
        let date = contents[5].toISOString().split('T')[0];
        date = date.split('-');

        if (date[0] == year) {
            let length = employees[employee].length;
            let billingSheet = employees[employee][length - 1].getWorksheet('Billing');

            let row = [`${date[1]}/${date[2]}/${date[0]}`];
            if (contents[15]) // If Invoice number
                row.push(contents[15]);
            else
                row.push(null);
            row.push(contents[3], contents[4], parseFloat(contents[8]) / contents[9]);

            // Logic to determine splits
            let splitFee, invoice;


            switch (role) {
                case 'Account Manager':
                    if (employees[employee][0]['Researcher'])
                        splitFee = researcherCommission;
                    else
                        splitFee = employees[employee][0]['Account Manager'];
                    invoice = calculateSplitInvoice(contents, role)
                    row.push(invoice * 1);
                    row.push(invoice * splitFee);
                    if (employees[employee][0]['Researcher']) {
                        row.push(0);
                        row.push(0);
                    }
                    break;
                case 'Operations Manager':
                    splitFee = employees[employee][0]['Operations Manager'];
                    invoice = calculateSplitInvoice(contents, role)
                    row.push(invoice * 1);
                    row.push(invoice * splitFee);
                    break;
                case 'Researcher':
                    splitFee = employees[employee][0]['Researcher'];
                    invoice = calculateSplitInvoice(contents, role)
                    row.push(invoice * 1);
                    row.push(0);
                    if (contents[13] == 'Yes')
                        row.push(invoice * employees[employee][0]['Researcher']);
                    else
                        row.push(0);

                    if (contents[4].includes('1/') || !contents[4].includes('/'))
                        row.push(250)
                    else
                        row.push(0)
                    break;
                default:
                    console.log('Error: Invalid role');
            }

            if (contents[16]) {
                row.push(null);
                row.push(contents[16]);
            }

            billingSheet.addRow(row);
        }
    }
}
