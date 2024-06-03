function sheetToDocs() {
  const colors = ["แดง", "เขียว", "น้ำเงิน", "เหลือง", "ส้ม", "ชมพู"]
  const colID = "A"
  const colName = "C"
  const colColor = "D"

  const docsID = "target-docs-id"
  const docs = DocumentApp.openById(docsID)
  const body = docs.getBody()

  for (let c = 0; c < colors.length; c++) {
    let color = colors[c]
    console.log(color)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(color);
    const lastRow = sheet.getLastRow()

    let hd = body.appendParagraph("สี" + color)
    hd.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    hd.setFontFamily("Kanit")
    hd.setSpacingBefore(0)

    for (let i = 2; i <= lastRow; i++){
      let id = sheet.getRange(colID + i).getValue()
      let name = sheet.getRange(colName + i).getValue()
      
      body.appendParagraph((i - 1) + '. ' + id + ' - ' + name).setFontSize(14).setFontFamily("Kanit").setSpacingBefore(1)
      console.log(id, name)
    }
    body.appendPageBreak()
  }
}
