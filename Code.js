// Google Script Code to generate invoices and mail them on demand as a PDF.  Meant to be called 
// monthly to generate invoice for previous month.  Automatically adds new invoices as new sheets
// by their ID.  Automatically hides invoices marked paid on Summary page.

// Change this to be your spreadsheet with two sheets: Summary and 1
var SHEETS_URL = "https://docs.google.com/spreadsheets/d/1oOfJGtvfI1zU_pJ6h_C7ceFho7a6zxEpXTLofygm0WI/edit"
// Summary sheet should have two sections: Settings and Invoices
//    Settings are arranged as key name in column A and value in B, rows 2-5
//    There are three expected keys to fill in: EMAIL_TO, SUBJECT_PREFIX, PDF_NAME_PREFIX
//    Everything after that is the Invoices section.  I created three headers in row 6: Invoice, Date, Paid
// Sheet 1 is your first Invoice, used as a template.  I used this template:
//    https://docs.google.com/spreadsheets/d/1PgO74GPaU6UsTU7WED2geQJstV1ASGxeFQqvNFAhFIg/copy
// 

var settings = getSheetSettings()

function getSheetSettings() {
  var ss = SpreadsheetApp.openByUrl(SHEETS_URL)
  var sheet = ss.getSheetByName('Summary')
  var data = sheet.getRange('A2:B8').getValues()
  var ret = {}
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      ret[data[i][0]] = data[i][1]
    }
  }
  ret.lastUpdated = new Date()
  console.info(ret)
  return ret
}

function firstDayInPreviousMonth(yourDate) {
  return new Date(yourDate.getFullYear(), yourDate.getMonth() - 1, 1)
}

function lastDayInPreviousMonth(yourDate) {
  return new Date(yourDate.getFullYear(), yourDate.getMonth(), 0)
}

function monthAndYear(givenDate) {
  const month = givenDate.toLocaleString('en-us', { month: 'long' })
  const year = givenDate.getFullYear()
  return month + " " + year
}

function sendLastInvoice(previousMonthDate, ss, sheets, paidInvoices) {
  // Hide all sheets (but the last one, which isn't in sheets), otherwise they appear in PDF
  for (const sheet of sheets) {
    if (!sheet.isSheetHidden()) {
      sheet.hideSheet()
    }
  }
  const monYr = monthAndYear(previousMonthDate)

  var message = {
    to: settings.EMAIL_TO,
    subject: settings.SUBJECT_PREFIX + " - " + monYr,
    body: ["Hi,", "Please find invoice " + sheets.length + " for " + monYr + " attached.", "Thank you!"].join('\n\n'),
    attachments: [ss.getAs(MimeType.PDF).setName(settings.PDF_NAME_PREFIX + sheets.length)]
  }
  if ('EMAIL_CC' in settings) {
    message.cc = settings.EMAIL_CC
  }
  MailApp.sendEmail(message)
  // Unhide any not paid
  for (var i = 0; i < sheets.length; i++) {
    if (paidInvoices.indexOf(i) == -1) {
      sheets[i].showSheet()
    }
  }
}

function generateNewInvoice() {
  var ss = SpreadsheetApp.openByUrl(SHEETS_URL)
  var sheets = ss.getSheets()

  if ('MAX_INVOICES' in settings && sheets.length > settings.MAX_INVOICES) {
    console.log('Already generated ' + settings.MAX_INVOICES + ' invoices...skipping this invoice')
    return
  } else {
    console.log('Generating invoice ' + sheets.length)
  }

  // Create new Sheet for invoice and update fields
  ss.setActiveSheet(sheets[sheets.length-1]) // Duplicate the last sheet as our template
  var newSheet = ss.duplicateActiveSheet()
  newSheet.showSheet()
  var invoiceId = sheets.length
  newSheet.setName(invoiceId)
  var idRange = newSheet.getRange('E3') // Update this to match your invoice template
  idRange.setValue(invoiceId)
  var todayRange = newSheet.getRange('E4') // Update this to match your invoice template
  var today = new Date()
  todayRange.setValue(today)
  var startDateRange = newSheet.getRange('C21') // Update this to match your invoice template
  startDateRange.setValue(firstDayInPreviousMonth(today))
  var endDateRange = newSheet.getRange('D21') // Update this to match your invoice template
  endDateRange.setValue(lastDayInPreviousMonth(today))

  // Update summary sheet
  var summarySheet = ss.getSheetByName('Summary')
  summarySheet.appendRow([invoiceId, today, false])
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build()
  summarySheet.getRange(summarySheet.getLastRow(), 3).setDataValidation(checkbox).setValue("FALSE")

  // Calculate which invoices are paid so they can stay hidden after email is sent
  var sheetData = summarySheet.getDataRange().getValues()
  var paidInvoices = []
  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][2] !== false) {
      paidInvoices.push(sheetData[i][0])
    }
  }

  // Send it
  sendLastInvoice(lastDayInPreviousMonth(today), ss, sheets, paidInvoices)
}
