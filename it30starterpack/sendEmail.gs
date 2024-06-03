function sendEmail() {
  const sheetName = 'Sheet 1'
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const colEmail = "B"
  const colId = "C"
  const colName = "D"
  const colRole = "H"
  const lastRow = sheet.getLastRow()

  for (let i = 2; i <= lastRow; i++){
    const email = sheet.getRange(colEmail + i).getValue().toString()
    const id = sheet.getRange(colId + i).getValue().toString()
    const name = sheet.getRange(colName + i).getValue().toString()
    const role = sheet.getRange(colRole + i).getValue().toString()

    const subject = 'Your Email Subject'
    const message = `to ${name} (${id}) on role ${role}`

    GmailApp.sendEmail(email, subject, message);
    console.log("Success ID:" + id)
  }
}
