function listToTableDocs() {
  const colors = ["แดง", "เขียว", "น้ำเงิน", "เหลือง", "ส้ม", "ชมพู"]
  const groupColor = {
    "แดง": "RubyTangle", 
    "เขียว": "JadyExplorer", 
    "น้ำเงิน": "WaveBlue", 
    "เหลือง": "LemonPuff", 
    "ส้ม": "SunnyFin", 
    "ชมพู": "RosyLotl"
  }
  const colID = 0
  const colName = 1
  const colNickName = 2

  const docsID = "target-docs-id" // ใบเซ็นชื่อ IT29 6/9/23
  const docs = DocumentApp.openById(docsID)
  const body = docs.getBody()

  for (let c = 0; c < colors.length; c++) {
    let color = colors[c]
    console.log(color)
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(color);
    let data = sheet.getDataRange().getValues()

    let textTopic = `ลงทะเบียนผู้เข้าร่วมสำหรับนักศึกษาชั้นปีที่ 1\nโครงการ IT29 and The Code Of Aquatia`
    let topic = body.appendParagraph(textTopic)
    topic.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    topic.setFontFamily("Sarabun")
    topic.setSpacingBefore(0)
    topic.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    topic.setBold(true)
    topic.setFontSize(18)

    let dateAndLocate = `จัดขึ้นในวันที่ 6 กันยายน พ.ศ. 2566\nณ อาคารเรียนรวม 2 (CB2)`
    let subTopic = body.appendParagraph(dateAndLocate)
    subTopic.setFontFamily("Sarabun")
    subTopic.setSpacingBefore(0)
    subTopic.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    subTopic.setBold(true)
    subTopic.setFontSize(14)
 
    let hd = body.appendParagraph("สี" + color + " (" + groupColor[color] + ")")
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

    for (let i = 1; i < data.length; i++){
      let id = data[i][colID]
      let name = data[i][colName]
      let nickName = data[i][colNickName]

      row = table.appendTableRow().setBold(false)
      row.appendTableCell(`${i}`).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      row.appendTableCell(`${id}`).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      row.appendTableCell(name).setWidth(170)
      row.appendTableCell(nickName)
      row.appendTableCell(" ")
    } 
    body.appendPageBreak()
  }
}
