# Email-Excelerator
###A Microsoft Excel based bulk email generator and sender

This excel based bulk email generator comes handy when you wish to generate and send an personalized email with attachments to a bunch of people.
However,
* Microsoft Office Mail Merge is insufficient as it doesn't help with attachments 
* You don't have any computer programming experience 
* Cannot install any software on your computer (admin / corporate restriction)

Prerequisites
* Microsoft Office 2016 or above (with minor change in reference this could be run for lower versions)
* Microsoft Outlook must have active connection

The Workbook has 3 main worksheets - EmailTemplate, EmailContentRecords and LOG. Don't delete or change the name of these worksheets.

Start with EmailTemplate, fill in Subject and Body mandatorily. Optionally you can indicate if anyone needs to be CCed or BCCed and whether attachments are available.
In case attachments are available and to be included, provide the root path. 

In EmailContentRecords, first row is assumed to have header. Subsequent rows are assumed to have data for mails. First column Recipient should have email id of recipient and is expected to have values in all valid (filled) rows. For CC and BCC columns with CopyTo and BCopyTo header must be available. For file attachments a column FileName must be present.

In case the Subject and Body have tokens that needs to be substituted with specific values per mail, tokens must be present as header (exact). While generating emails, tool will try to replace tokens with values from each valid data row.

##### Complete the setup and hit the 'Generate & Send' button.
