import { createRequire } from "module";
const require = createRequire(import.meta.url);
const Excel = require('exceljs');

async function main() {
    const billing = new Excel.Workbook();
    await billing.xlsx.readFile('./temp/billing.xlsx');
    const billingSheet = billing.getWorksheet('Billing');

    // Remove unnecessary columns
    // This should be done using tables but is too buggy to use
    billingSheet.spliceColumns(3, 2);
    billingSheet.spliceColumns(4, 1);
    billingSheet.spliceColumns(6, 9);
    billingSheet.spliceColumns(11, 1);
    billingSheet.spliceColumns(16, 7);

    billingSheet.eachRow(function (row, rowNumber) {
        let content = row.values
        content.shift()
        if (content[15])
            console.log(content[15])
        else
            console.log('No invoice number')
    })

    billing.xlsx.writeFile('./temp/parsed_billing.xlsx');

    //if (researcher1 && researcher2) {
    // totalfee/2 * commission rate of each researcher
    //}

    /** three types of splits
     * recruiter1 and recruiter2 is internal: divide total fee by 2 and give the recruiters @ their rate
     * recruiter1 and recruiter2 is external: recruiter1 gets half the fee @ their rate
     * recruiter1 and recruiter2 is external @te: recruiter1 gets 47% of the fee @ their rate
    */
}

main();