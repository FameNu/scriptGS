function toTableDocs() { // run time 18s
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("questionToDocs")
  const lastRow = sheet.getLastRow() // row สุดท้ายของ sheet

  const colCheck = 'A'
  const colID = 'E'
  const colName = 'F'
  const colNickName = 'G'
  const colEmail = 'H' // Microsoft account (outlook)
  const colRole = 'K'
  const colSpeaker = 'L'

  const role = ['Front end', 'Back end', 'Web Design', 'DevOps', 'Technical', 'Art & Design']

  const docsId = 'Target-Google-Docs-ID' // DocsTable
  const docs = DocumentApp.openById(docsId)
  const body = docs.getBody()
  
  for (let getRole of role) {
    const textTopic = `ลงทะเบียนผู้เข้าร่วมสำหรับนักศึกษาชั้นปีที่ 1\nโครงการ Hello World of Goose`
    const topic = body.appendParagraph(textTopic)
    topic.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    topic.setFontFamily("Sarabun")
    topic.setSpacingBefore(0)
    topic.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    topic.setBold(true)
    topic.setFontSize(18)

    const dateAndLocate = `จัดขึ้นในวันที่ 8 พฤศจิกายน พ.ศ. 2566\nณ อาคาร Learning Exchange (LX)`
    const subTopic = body.appendParagraph(dateAndLocate)
    subTopic.setFontFamily("Sarabun")
    subTopic.setSpacingBefore(0)
    subTopic.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    subTopic.setBold(true)
    subTopic.setFontSize(14)

    const hd = body.appendParagraph('ฝ่าย ' + getRole)
    hd.setHeading(DocumentApp.ParagraphHeading.HEADING2)
    hd.setFontFamily("Sarabun")
    hd.setSpacingBefore(1)
    hd.setSpacingAfter(10)
    hd.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    hd.setBold(true)
    hd.setFontSize(14)

    // how to set width each cell
    // setWidth(point) by 1 inch = 2.54 cm = 72 point
    let table = body.appendTable()
    let row = table.appendTableRow().setFontSize(11).setBold(true)
    row.appendTableCell("ลำดับ").setWidth(40).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    row.appendTableCell("รหัสนักศึกษา").setWidth(80).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    row.appendTableCell("ชื่อ-นามสกุล").setWidth(170)
    row.appendTableCell("ชื่อเล่น").setWidth(50)
    row.appendTableCell("เซ็นชื่อ")
    let countPerson = 0
    for (let i = 2; i <= lastRow; i++) {
      const thisRole = sheet.getRange(colRole + i).getValue() === 'Speaker' ? sheet.getRange(colSpeaker + i).getValue() : sheet.getRange(colRole + i).getValue()
      if (getRole === thisRole) {
        const id = sheet.getRange(colID + i).getValue()
        const name = sheet.getRange(colName + i).getValue()
        const nickName = sheet.getRange(colNickName + i).getValue()

        row = table.appendTableRow().setBold(false)
        row.appendTableCell(`${++countPerson}`).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
        row.appendTableCell(`${id}`).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
        row.appendTableCell(name).setWidth(170)
        row.appendTableCell(nickName)
        row.appendTableCell(" ")
      }
    }
    body.appendPageBreak()
  }
}
