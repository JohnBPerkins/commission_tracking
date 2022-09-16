This program automatically generates several billing and commission tracking files for 
individual employees to view their monthly progress.

For setup:

1. Use the config.xlsx file in the Commission Reports folder within the SharePoint Metrics site to add/remove employees.
2. Copy entries from Meghan's billing submission form into the billing_submission_form.xlsx within the SP site.
  - If there are multiple invoices for one submission, copy and paste the row for however many invoices there will be. Change the 'Start date' column for each entry to reflect the invoice date, enter the invoice number, and update the candidate name to include what number invoice it is, e.g. John Doe (5/6).
3. Any billing changes that need to be made should be done in the billing_submission_form.xlsx file and not in the individual employee files otherwise they will be overwritten. However, the employees commission tracking pages will not be overwritten and can be edited freely.

To run: 

1. Download and extract the zip file.
2. Edit the config file to include the tenantId, clientId, and clientSecret.
3. Run the run.bat file.
