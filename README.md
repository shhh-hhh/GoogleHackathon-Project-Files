# GoogleHackathon-Project-Files
GoogleHackathon 2024 Project File - Team GoogleNoGoo

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoice')
  .addItem('Send Invoice', 'sendInvoice')
  .addToUi();
  // Create Receipt menu
  ui.createMenu('Receipt')
    .addItem('Send Receipt', 'sendReceipt')
    .addToUi();
}

function sendInvoice() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Details');
  if (!sheet) {
    Logger.log('Sheet "Invoice Details" not found.');
    SpreadsheetApp.getUi().alert('Sheet "Invoice Details" not found.');
    return;
  }
  Logger.log('Sheet found, proceeding with invoice processing.');
  const templateID = '1JH10ceiX7lPlk0io5Fa6Zr9I0mlc_TmmK9XjQLseFy0';
  const invoicedata = sheet.getDataRange().getValues();
  for (let i = 1; i < invoicedata.length; i++){
    let status = invoicedata[i][8];
    let invoice_sent = invoicedata[i][9]
    if (status == 'Unpaid' && invoice_sent == 'No'){
      let IssueDate = invoicedata[i][6];
      let InvoiceNumber = invoicedata[i][0];
      let ClientName = invoicedata[i][1];
      let ClientEmail = invoicedata[i][2];
      let ClientAddress = invoicedata[i][3];
      let Service = invoicedata[i][4];
      let Amount = invoicedata[i][5];
      let DueDate = invoicedata[i][7];

      let copyID = DriveApp.getFileById(templateID).makeCopy().getId();
      let doc = DocumentApp.openById(copyID);
      let body = doc.getBody();

      body.replaceText('{{IssueDate}}', IssueDate);
      body.replaceText('{{InvoiceNumber}}', InvoiceNumber);
      body.replaceText('{{ClientName}}', ClientName);
      body.replaceText('{{ClientEmail}}', ClientEmail);
      body.replaceText('{{ClientAddress}}', ClientAddress);
      body.replaceText('{{Service}}', Service);
      body.replaceText('{{Amount}}', Amount);
      body.replaceText('{{DueDate}}', DueDate);

      doc.saveAndClose();

      let pdfname = `Invoice_${InvoiceNumber}_${ClientName}.pdf`;
      let pdfBlob = DriveApp.getFileById(copyID).getAs("application/pdf").setName(pdfname);

      GmailApp.sendEmail(ClientEmail, "Invoice " + InvoiceNumber,
        `Dear ${ClientName},\n\nPlease find your invoice details below:\n\nService: ${Service}\nAmount: $${Amount}\nIssue Date: ${IssueDate}\nDue Date: ${DueDate}\n\nPlease find your invoice attached as a PDF.\n\nThank you for your business!\n\nBest Regards,\nYour Company, \n\n***This is a system generated email. Please do not reply to this address***`,
        {
          attachments: [pdfBlob],
          name: 'Automatic Invoice Sender'
        });

      sheet.getRange(i + 1, 10).setValue("Yes");

      let folderID = '1TtpiW3vqhz7nogs7NmrpYtM3Gk2Kbnwd';
      let folder = DriveApp.getFolderById(folderID);
      folder.createFile(pdfBlob);

      DriveApp.getFileById(copyID).setTrashed(true);

    }
  }
}



//receipt generator

// Set the column indices
const STATUS_COLUMN = 9; 
const REMINDER_SENT_COLUMN = 10; 
const RECEIPT_NUMBER_COLUMN = 11; 
const RECEIPT_SENT_COLUMN = 12; 

// Function to run when the sheet is edited
function onEdit(e) {
  if (!e) {
    Logger.log("Event object is undefined. Ensure this function is not run manually.");
    return;
  }

  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Check if the edited cell is in the "Status" column
  if (range.getColumn() == STATUS_COLUMN) {
    // Get the row of the edited cell
    var row = range.getRow();

    // Check if the status is "Paid"
    if (range.getValue() == "Paid") {
      var receiptNumberCell = sheet.getRange(row, RECEIPT_NUMBER_COLUMN);
      var receiptSentCell = sheet.getRange(row, RECEIPT_SENT_COLUMN);

      // Check if the receipt number cell is empty
      if (!receiptNumberCell.getValue()) {
        // Generate a unique receipt number
        var receiptNumber = generateReceiptNumber(sheet);

        // Set the receipt number and receipt sent status
        receiptNumberCell.setValue(receiptNumber);
        receiptSentCell.setValue("No");

        Logger.log("Receipt number generated and set: " + receiptNumber);
      } else {
        Logger.log("Receipt number already exists: " + receiptNumberCell.getValue());
      }
    }
  }
}

// Function to generate a unique receipt number
function generateReceiptNumber(sheet) {
  if (!sheet) {
    Logger.log("Sheet is undefined in generateReceiptNumber.");
    return "ERROR";
  }

  Logger.log("generateReceiptNumber called with sheet: " + sheet.getName());
  var lastRow = sheet.getLastRow();
  Logger.log("Last row in the sheet: " + lastRow);

  var receiptNumbers = sheet.getRange(2, RECEIPT_NUMBER_COLUMN, lastRow - 1, 1).getValues();
  var maxNumber = 0;

  // Find the maximum receipt number in the existing data
  for (var i = 0; i < receiptNumbers.length; i++) {
    var receiptNumber = receiptNumbers[i][0];
    if (receiptNumber) {
      var num = parseInt(receiptNumber.replace('RCP', ''), 10);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // Generate the new receipt number
  if (maxNumber === 0) {
    return 'RCP001';
  } else {
    return 'RCP' + ('000' + (maxNumber + 1)).slice(-3);
  }
}

function sendReceipt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Details');
  if (!sheet) {
    Logger.log('Sheet "Invoice Details" not found.');
    SpreadsheetApp.getUi().alert('Sheet "Invoice Details" not found.');
    return;
  }
  Logger.log('Sheet found, proceeding with receipt processing.');

  const templateID = '1rwSshf9rj6FvFQ8tESXUFtk4qgf_OZ5OnBaDMuzd3nA'; // Receipt template ID
  const receiptData = sheet.getDataRange().getValues();

  for (let i = 1; i < receiptData.length; i++) {
    let receiptSent = receiptData[i][11]; // Column L (index 12)
    let paid_status = receiptData[i][8];
    if (receiptSent == 'No' && paid_status == 'Paid') { // If customer paid but haven't send receipt
      let currentDate = new Date();
      let formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      let receiptNumber = receiptData[i][10]; // Column K (index 11)
      let clientName = receiptData[i][1];
      let clientEmail = receiptData[i][2];
      let clientAddress = receiptData[i][3];
      let service = receiptData[i][4];
      let amount = receiptData[i][5];

      Logger.log('Processing row ' + (i + 1));
      Logger.log('Receipt Number: ' + receiptNumber);
      Logger.log('Client Email: ' + clientEmail);

      try {
        let copyID = DriveApp.getFileById(templateID).makeCopy().getId();
        Logger.log('Template copied, new ID: ' + copyID);

        let doc = DocumentApp.openById(copyID);
        let body = doc.getBody();

        body.replaceText('{{Date}}', formattedDate);
        body.replaceText('{{ReceiptNumber}}', receiptNumber);
        body.replaceText('{{ClientName}}', clientName);
        body.replaceText('{{ClientEmail}}', clientEmail);
        body.replaceText('{{ClientAddress}}', clientAddress);
        body.replaceText('{{Service}}', service);
        body.replaceText('{{Amount}}', amount);

        doc.saveAndClose();

        let pdfName = `Receipt_${receiptNumber}_${clientName}.pdf`;
        let pdfBlob = DriveApp.getFileById(copyID).getAs("application/pdf").setName(pdfName);
        Logger.log('PDF Blob created, name: ' + pdfName);

        GmailApp.sendEmail(clientEmail, "Receipt " + receiptNumber,
          `Dear ${clientName},\n\nI hope this message finds you well.\n\nThank you very much for your prompt payment.\n\nI have attached the receipt for your records.\n\nIf you have any questions or need further assistance, please don't hesitate to reach out.\n\n\nPlease find your receipt attached as a PDF.\n\nThank you for your business!\n\n\n\nBest Regards,\nYour Company,\n\n***This is a system generated email. Please do not reply to this address***`,
          {
            attachments: [pdfBlob],
            name: 'Automatic Receipt Sender'
          });

        Logger.log('Email sent successfully to ' + clientEmail);

        sheet.getRange(i + 1, 12).setValue("Yes");

        let folderID = '1PRwC8TxHdiVmcBGxXU16v9RT7m6f48pz'; // Corrected folder ID
        let folder = DriveApp.getFolderById(folderID);
        folder.createFile(pdfBlob);
        Logger.log('PDF saved to folder.');

        DriveApp.getFileById(copyID).setTrashed(true);
        Logger.log('Temporary file moved to trash.');

      } catch (error) {
        Logger.log('Error processing row ' + (i + 1) + ': ' + error.message);
      }
    }
  }
}
