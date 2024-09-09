const ANCHOR_CODE = /@\((.+?)\)/
const VALUE_CODE = /#/g
const WORKBOOK_ID = 'WORKBOOK_ID_HERE'
const MAX_DELINQUENCIES_PER_MONTH = 1
const REPORT_EMAILS = 'email1@domain.com,email2@domain.com,email3@domain.com'

function main() {
  let report = '<h1>Weekly PM Summary Report</h1>\n<p>Text in <span style="color: blue;">blue</span> are hyperlinks. Click them to see pictures of the problems submitted by site agents.</p>'
  report+= '\n' + getPMSummary()
  
  // Send the email
  MailApp.sendEmail({
    to: REPORT_EMAILS,
    subject: 'Weekly PM Report Summary',
    htmlBody: report
  })
  Logger.log('Sent weekly PM report summary')
}

function reminder() {
  const sheet = SpreadsheetApp.openById(WORKBOOK_ID).getSheetByName('Summary')
  let summarySheet = sheet.getDataRange().getValues()
  const headers = summarySheet.shift()  // Delete header row
  summarySheet.shift()  // Delete second row
  summarySheet.shift() // Delete third row

  for (const siteRow of summarySheet) {
    if (!getResponse(siteRow, headers, 'Submission within Last 5 Days')) {
      const delinquentEmail = getResponse(siteRow, headers, 'Email')
      if (delinquentEmail) {
        MailApp.sendEmail({ 
          to: delinquentEmail,
          subject: 'Reminder: Please submit your weekly PM report by Monday',
          body: `Dear ${getResponse(siteRow, headers, 'Site Agent').trim()},\n\nThis is an automated message from Renewvia. Please do not reply to this email.\n\nOur records indicate that you have not yet submitted your Weekly Preventative Maintenance Report for this week. Preventative maintenance reports are important because they help us solve problems before they happen, increasing uptime and keeping our customers happy. Unless you have received explicit permission from core Renewvia staff, you may not be eligible to receive your full paycheck if you fail to submit your report before 5 AM UTC, Monday. If you believe you are receiving this email in error, please reach out in your site's Renewvia WhatsApp group.`
        })
        Logger.log(`Sent report reminder to ${getResponse(siteRow, headers, 'Site Agent').trim()}`)
      }
    }
  }
}

function paycheckSummary() {
  const sheet = SpreadsheetApp.openById(WORKBOOK_ID).getSheetByName('Recent Delinquent Reports')
  let delSheet = sheet.getDataRange().getValues()
  const weeks = delSheet.shift()  // Delete header row
  const today = new Date()

  // Get columns corresponding to this month
  let firstColThisMonth = 1
  while (weeks[firstColThisMonth].getYear() != today.getYear() || weeks[firstColThisMonth].getMonth() != today.getMonth()) { firstColThisMonth++ }

  // For each site
  let problems = []
  for (const siteRow of delSheet) {

    // Get the number of delinquencies this month
    let delinquencies = 0
    for (let w=firstColThisMonth; w<=siteRow.length; w++) {
      if (siteRow[w]) {
        delinquencies++
      }
    }

    // If there were too many delinquencies, add this site to the list of problems
    if (delinquencies>MAX_DELINQUENCIES_PER_MONTH) {
      problems.push(`${siteRow[0]}'s site agent was delinquent on ${delinquencies} of their weekly PM reports this month.`)
    }
  }

  // Send the list of problems
  if (problems) {
    let problemsList = '<ul>\n'
    for (const problem of problems) {
      problemsList+= `\t<li>${problem}</li>\n`
    }
    problemsList+= '</ul>'

    MailApp.sendEmail({
      to: REPORT_EMAILS,
      subject: `Monthly summary of weekly PM report delinquencies`,
      htmlBody: `<p>This month, the site agents of the following sites had more than ${MAX_DELINQUENCIES_PER_MONTH} delinquent PM reports:</p>\n${problemsList}`
    })
    Logger.log('Sent list of delinquent reporters to staff')
  }
}

function getPMSummary() {
  const sheet = SpreadsheetApp.openById(WORKBOOK_ID).getSheetByName('Summary')
  let summarySheet = sheet.getDataRange().getValues()
  const headers = summarySheet.shift()  // Delete header row
  const wrongAnswers = summarySheet.shift()  // Delete second row
  const reportBullets = summarySheet.shift() // Delete third row
  summarySheet.sort((a, b) => a[1].localeCompare(b[1])) // Sort by country

  let report = ''
  for (const siteRow of summarySheet) {
    report+= `\n<h2>${getResponse(siteRow, headers, 'Country')}, ${getResponse(siteRow, headers, 'Site Name')}`
    if (getResponse(siteRow, headers, 'Site Agent') == '#N/A') { // If the site agent has never completed a weekly PM report
      report+= '</h2>\n<p style="color: red">Site agent has never completed a weekly PM report.</p>'
    } else {  // If the site agent has completed a weekly PM report sometime in the past
      report+= `, ${getResponse(siteRow, headers, 'Site Agent')}</h2>`
      let problems = []

      // Submission within last week
      if (!getResponse(siteRow, headers, 'Submission within Last Week')) {
        problems.push(`Last submission was on ${getResponse(siteRow, headers, 'Most Recent Submission')}, more than one week ago.`)

        const delinquentEmail = getResponse(siteRow, headers, 'Email')
        if (delinquentEmail) {
          MailApp.sendEmail({ 
            to: delinquentEmail,
            subject: 'You missed your weekly PM report',
            body: `Dear ${getResponse(siteRow, headers, 'Site Agent').trim()},\n\nThis is an automated message from Renewvia. Please do not reply to this email.\n\nOur records indicate that you failed to submit your Weekly Preventative Maintenance Report last week. Preventative maintenance reports are important because they help us solve problems before they happen, increasing uptime and keeping our customers happy. Unless you have received explicit permission from core Renewvia staff, you may not be eligible to receive your full paycheck as a result of your delinquency. If you believe you are receiving this email in error, please reach out in your site's Renewvia WhatsApp group.`
          })
          Logger.log(`Sent delinquent report notification email to ${getResponse(siteRow, headers, 'Site Agent').trim()}`)
        }
      }

      // Get first non-header column
      let c = 0
      while (!wrongAnswers[c]) { c++ }
      
      // For each question
      while (c<siteRow.length) {
        // If the agent's answer matches the "wrong answer" regex
        if (RegExp(wrongAnswers[c]).test(siteRow[c])) {
          // Get the error text for the report
          let reportBullet = reportBullets[c]
          
          // If the error text is regular, push it to the problems list as an anchor with the linked picture
          if (!ANCHOR_CODE.test(reportBullet) && !VALUE_CODE.test(reportBullet)) {
            problems.push(getAnchor(siteRow[c+1], reportBullet))

          // If the error text has a special code, format it before adding it
          // Note that anchor code formatting should be done before value code formatting to avoid the user adding anchor codes
          } else {
            let response = siteRow[c]
            while (ANCHOR_CODE.test(reportBullet)) {
              const reMatch = reportBullet.match(ANCHOR_CODE)
              reportBullet = reportBullet.replace(ANCHOR_CODE, getAnchor(siteRow[++c], reMatch[0].substring(2, reMatch[0].length-1)))
            }
            reportBullet = reportBullet.replaceAll(VALUE_CODE, response)
            problems.push(reportBullet)
          }
        }

        // Advance to the next column with an answer check
        c++
        while (!wrongAnswers[c] && c<siteRow.length) { c++ }
      }

      // Append problems to report
      if (problems.length == 0) {
        report+= '\n<p>No issues reported.</p>'
      } else {
        report+= `\n<ul>\n<li>${problems.join('</li>\n<li>')}</li>\n</ul>`
      }
    }
  }

  return report
}

function getDieselSummary() {
  let report = '<h1>Diesel Generator Summary</h1>'

  const sheet = SpreadsheetApp.openById(DIESEL_SUMMARY_ID).getSheetByName('Summary')
  let summarySheet = sheet.getDataRange().getValues()
  const headers = summarySheet.shift()  // Delete header row
  for (const siteRow of summarySheet) {
    report+= `\n<h2>${getResponse(siteRow, headers, 'Site Name')}</h2>`
    if (siteRow[2] == '#N/A') {
      report+= '\n<p style="color: red">Site agent has never completed a diesel generator update report.</p>'
    } else {
      report+= `\n<ul>`
      report+= `\n<li>Last update was on ${formatDate(getResponse(siteRow, headers, 'Most Recent Submission'))}</li>`
      report+= `\n<li>Diesel Level: ${getAnchor(getResponse(siteRow, headers, 'Most Recent Photo of Diesel Tank'), `${getResponse(siteRow, headers, 'Most Recent Diesel Level (L)')} L`)}\</li>`
      report+= `\n<li>Run Hours: ${getAnchor(getResponse(siteRow, headers, 'Most Recent Photo of Run Hours'), `${getResponse(siteRow, headers, 'Most Recent Run Hours')}`)}</li>`
      report+= '\n</ul>'
    }
  }

  return report
}

function getAnchor(href, label) {
  return href ? `<a rel='external' target='_blank' href='${href}'>${label}</a>` : label
}

function getResponse(siteRow, headers, header) {
  const ind = headers.indexOf(header)
  if (ind == -1) {
    Logger.log(`${header} does not exist in the header row.`)
    throw new Error(`${header} does not exist in the header row.`)
  }
  return siteRow[ind]
}

function formatDate(dateStr) {
  let date = new Date(dateStr)
  date = `${date.getDate()} ${date.toLocaleString('default', { month: 'long' })} ${date.getFullYear()}`
}