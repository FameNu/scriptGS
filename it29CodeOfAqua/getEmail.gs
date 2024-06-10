function copyAndEditPaste() {
  const sheetName = 'All'
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  const lastRow = sheet.getLastRow()

  const colEmail = 'F'
  const colAddMail = 'J'

  for (let i = 18; i <= lastRow; i++) {
    const email = sheet.getRange(colEmail + i).getValue().toString()
    const msMail = email.replace("@mail.kmutt.ac.th", '@kmutt.ac.th')

    sheet.getRange(colAddMail + i).setValue(email)
    sheet.getRange(colEmail + i).setValue(msMail)
  }
}

function personalEmail() {
  const mainSheetName = 'All'
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetName)
  const mainSheetLastRow = mainSheet.getLastRow()

  const mainCol = {
    stdId: 'A',
    pnMail: 'K'
  }
  
  const targetId = '1c56dNCAgTDcG_y-0ur7mdleXKo8zdzMo2_CmtbMRjfI'
  const targetSheetName = 'Form Responses 1'
  const targetSheet = SpreadsheetApp.openById(targetId).getSheetByName(targetSheetName)
  const targetLastRow = targetSheet.getLastRow()

  const targetCol = {
    stdId: 'D',
    pnMail: 'C'
  }

  for (let i = 2; i <= targetLastRow; i++) {
    const stdId = targetSheet.getRange(targetCol.stdId + i).getValue().toString()

    for (let j = 18; j <= mainSheetLastRow; j++) {
      const mainStdId = mainSheet.getRange(mainCol.stdId + j).getValue().toString()

      if (stdId.trim().includes(mainStdId.trim()) || mainStdId.trim().includes(stdId.trim())) {
        const pnEmail = targetSheet.getRange(targetCol.pnMail + i).getValue().toString()
        mainSheet.getRange(mainCol.pnMail + j).setValue(pnEmail)
        break
      }
    }
  }
}
