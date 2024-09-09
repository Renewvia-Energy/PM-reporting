# Preventative Maintenance Reporting
Renewvia built our own preventative maintenance reporting tool in house. Every week, each of our site agents is required to submit a Google Form containing pictures and status updates about their mini-grids. Then, we use Google Apps Script to send a summary of all of the issues reported by site agents to Renewvia's O&M staff. Thanks to Google's free form automation tools, this solution is entirely free.

If you would like to hire Renewvia to implement a custom setup for you and your company's needs, please reach out at [mailto:info@renewvia.com](info@renewvia.com). This repository contains all the code we use and links to an example Google Form and corresponding Google Sheet so you can build it yourself.

## The PM Form
An example of our PM form can be found at [https://forms.gle/GKedmUH1rCoEw44U7](https://forms.gle/GKedmUH1rCoEw44U7).

After copying the form for yourself, in the "Responses" tab, click "Link to Sheets."

## The Responses Workbook
After you click "Link to Sheets," Google will generate a workbook in Sheets. Add at least two spreadsheets titled "Recent Delinquent Reports" and "Summary." Copy the content from the respective sheets in [https://docs.google.com/spreadsheets/d/1XlBYHHowUJ1oxT4PP82qm7UEG6sP-toCDGH4OPNttoQ/edit?usp=sharing](this example workbook) to complete those sheets.

In the Recent Delinquent Reports worksheet, each row is a unique project site. After each site agent has completed at least one submission of the PM report form, the rows should match the options for the Site Name question in section 1 of the [form](https://forms.gle/GKedmUH1rCoEw44U7). The columns correspond to weeks, with each cell in the table reading "X" if and only if that site did not have a submission that week. You can use this table to track which sites' agents are completing their weekly reports.

The Summary tab is more involved. Again, each row after Row 3 is a unique project site. The first three columns, Site Name, Country, and Email, need to be manually filled once by you. Each column after column H corresponds to a different question on the form. Rows 1, 2, and 3 are populated as follows:
1. is autofilled with the questions from Form Responses. In column I in the example, Row 1 is "Are all security lights functioning and firmly attached to their poles or roofs?"
2. contains the form response to that column's question that would require its inclusion in the report summary email. In column I in the example, Row 2 is "No," because a "No" response to the question "Are all security lights functioning and firmly attached to their poles or roofs?" implies that the O&M team needs to do something to fix that problem.
3. contains the text that should be printed in the report summary email to indicate a problem. In column I in the example, Row 3 is "Problem with security lights" because, if the site agent responds "No" to the question "Are all security lights functioning and firmly attached to their poles or roofs?", then the summary report should indicate a problem with the security lights. If the next column is a link to an image, then that line of the summary report will link to that image. Note that, while it is not required, you may choose to customize this linking behavior with at-codes. For example, column O, row 3 reads "Damage to @(fence) or @(gate)." When a site agent responds "Yes" to "Is any section of the fence or gate damaged?" they are expected to attach two images, one of the fence and one of the gate. That line of the summary report will be formatted like `<li>Damage to <a rel='external' target='_blank' href=${href-1}>fence</a> or <a rel='external' target='_blank' href=${href-2}>gate</a>` where `${href-1}` is a link to the image linked in column P and `${href-2}` is a link to the image linked in column Q.
Each cell in the table corresponds to that row's site's agent's response to that column's question.

## Google Apps Script
Once you have completed creating the Responses workbook in Google Sheets, create a new project in Google Apps Script. Copy the script from `Code.gs` in this repository into the project. This script is responsible for:
- Compiling and sending a weekly summary report
- Sending reminders to site agents who have not yet completed their weekly report

Note that you will need to make the following changes:
1. In Line 3, replace `WORKBOOK_ID_HERE` with the ID of your Responses workbook from the previous section. The workbook ID can be found in the URL of the workbook. In this readme, we've been using [https://docs.google.com/spreadsheets/d/1XlBYHHowUJ1oxT4PP82qm7UEG6sP-toCDGH4OPNttoQ/edit?usp=sharing](this example workbook) with URL https://docs.google.com/spreadsheets/d/1XlBYHHowUJ1oxT4PP82qm7UEG6sP-toCDGH4OPNttoQ/edit?usp=sharing. For this workbook, the ID is `1XlBYHHowUJ1oxT4PP82qm7UEG6sP-toCDGH4OPNttoQ`, so Line 3 would read `const WORKBOOK_ID = '1XlBYHHowUJ1oxT4PP82qm7UEG6sP-toCDGH4OPNttoQ'`.
2. In Line 5, replace `email1@domain.com,email2@domain.com,email3@domain.com` with a comma-separated list of the emails of everyone who should receive a copy of the weekly summary reports.

Note that the content of the emails sent by the script are hardcoded. Feel free to customize these as you see fit, e.g., replacing "Renewvia" with your company's name.

After configuring the script to your liking, navigate to the Triggers tab and add two triggers:
1. A time-based trigger on the `main` function to run every week on a given day. This will send a weekly summary report on that day.
2. A time-based trigger on the `reminder` function to run every week a day or two before the first day. This will send a reminder to every site agent who has not yet filled out a report that week.

I encourage you to test the script using your own email for the `REPORT_EMAILS` to confirm you have configured the script to your liking before configuring the triggers.

## Conclusion
Once you have completed the following steps, send the form link to your site agents for them to complete on a weekly basis. You may find it helpful to train them by walking them through their first submission.

Depending on your email provider, you may notice that the emails sent by the script get sent to your Spam or Junk folder.

You may find some artifacts in the code we use to complete additional tasks such as diesel reporting and summarizing which site agents have "too many" delinquencies. We also found it useful to modify the script to support additional reporting. For example, we have our site agents complete weekly PM reports, separate monthly PM reports, and triweekly diesel reports.