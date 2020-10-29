# Email-Excelerator
### A Microsoft Excel based bulk email generator and sender

This excel based bulk email generator comes handy when you wish to generate and send bulk personalized emails with attachments.
However,
* Microsoft Office Mail Merge is insufficient as it doesn't help with attachments 
* You don't have any computer programming experience 
* Cannot install any software on your computer (admin / corporate restriction)

Prerequisites
* Microsoft Office 2010 (may also work for 2013 but not tested)
* Microsoft Outlook must have active connection

The Workbook has 3 main worksheets - EmailTemplate, EmailContentRecords and LOG. Don't delete or change the name of these worksheets.

Start with worksheet EmailTemplate, in the provided form, fill in Subject and Body mandatorily. Email content type can be chosen between HTML and Text (HTML chosen by default). When chosent HTML, tool will expect HTML content (only sets the MIME type). Optionally you can indicate if anyone needs to be CCed or BCCed and whether attachments are available. In case attachments are available and to be included, provide the root path. 

Move to worksheet EmailContentRecords. First row should have headers, subsequent rows are assumed to have data for mails [one row per email]. First column with header 'Recipient' is a must. In the data rows it should have email id of recipients and is expected to have values in all valid (filled) data rows. In case you have selected CC and / or BCC in the form, columns with headers 'CopyTo' and 'BCopyTo' must be available with data in data rows. Similarly for file attachments a column 'FileName' must be present. In case the Subject and Body have tokens that needs to be substituted with specific values per email, each token should have placeholders $token$ in the form and in here, a column in the header should have 'token' as an entry and data in data rows. While generating emails, tool will try to replace $tokens$ with values from each valid data row.

##### Complete the setup and hit the 'Generate & Send' button.
