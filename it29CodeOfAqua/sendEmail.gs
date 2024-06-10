function runAll() {
  sendEmailPassAll()
  snedEmailNotPassAll()
}

function sendEmailPassAll() {
  const sheetName = 'passAll'
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)

  const colEmail = "B"
  const colId = "C"
  const colName = "D"
  const colRole = "H"
  const colDepartment = 'I'
  const lastRow = sheet.getLastRow()

  let count = 0;
  for (let i = 2; i <= lastRow; i++){
    const email = sheet.getRange(colEmail + i).getValue().toString()
    const id = sheet.getRange(colId + i).getValue().toString()
    const name = sheet.getRange(colName + i).getValue().toString()
    const role = sheet.getRange(colRole + i).getValue().toString() == 'Speaker' ?
      `${sheet.getRange(colRole + i).getValue().toString()} ประจำวิชา ${sheet.getRange(colDepartment + i).getValue().toString()}` :
      sheet.getRange(colRole + i).getValue().toString()
    

    const subject = 'IT#30 Starter Pack | Welcome to our team!!'
    const message = 
    `\tเรียน ${name} ${id}

\tทางทีมงานโครงการ IT#30 Starter Pack มีความยินดีเป็นอย่างยิ่งที่ท่านให้ความสนใจในการเป็นส่วนหนึ่งของทีมงานโครงการ IT#30 Starter Pack

\tขอแสดงความยินดีด้วย คุณได้ผ่านการคัดเลือกและได้ทำงานกับเราในฝ่าย ${role}
โดยสามารถติดตามรายละเอียดเพิ่มเติมและการนัดหมายต่าง ๆ ของท่านผ่านแพลตฟอร์ม Discord เซิร์ฟเวอร์ "IT#30 Starter Pack: Staff"
และกดยืนยันสิทธิ์ภายในวันที่ 10 มิถุนายน 2567 เวลา 16:00 น. ทาง link ที่แนบมากับ Email ฉบับนี้

\tการนัดหมายล่วงหน้า ในวันที่ 10 มิถุนายน 2567 เวลา 20.00 น. จะเป็นการประชุมครั้งแรกของทีมงานในโครงการ IT#30 Starter Pack เราหวังว่าเราจะได้พบกันในวันเวลาดังกล่าว

\tยืนยันสิทธิ์: https://forms.gle/EVD9KkbAh3HduEp76

\tDiscord: https://discord.gg/gZMATPgDfY

“We are excited to start working with everyone soon!”
---------------------------------------------------------------------
เฟม IT#28 ประธานโครงการ IT#30 Starter Pack
Email ติดต่อกลับ: phuwamet.panj@mail.kmutt.ac.th`

    GmailApp.sendEmail(email, subject, message);
    // console.log(subject)
    // console.log('------------')
    // console.log(message)
    console.log("<Pass> Success ID:" + id)
    count++
    // console.log('<----------------------------------------->')
  }
  console.log(count)
}

function snedEmailNotPassAll() {
  const sheetName = 'notPass'
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)

  const colEmail = "B"
  const colId = "C"
  const colName = "D"
  const colRole = "H"
  const colDepartment = 'I'
  const lastRow = sheet.getLastRow()

  let count = 0
  for (let i = 2; i <= lastRow; i++){
    const email = sheet.getRange(colEmail + i).getValue().toString()
    const id = sheet.getRange(colId + i).getValue().toString()
    const name = sheet.getRange(colName + i).getValue().toString()
    const role = sheet.getRange(colRole + i).getValue().toString() == 'Speaker' ?
      `${sheet.getRange(colRole + i).getValue().toString()} ประจำวิชา ${sheet.getRange(colDepartment + i).getValue().toString()}` :
      sheet.getRange(colRole + i).getValue().toString()
    

    const subject = 'IT#30 Starter Pack'
    const message = 
    `\tเรียน ${name} ${id}

\tก่อนอื่นเลยพวกเราขอขอบคุณผู้สมัครทุกท่านที่สนใจเข้าร่วมเป็นส่วนหนึ่งของ IT#30 Starter Pack หลังจากที่ได้อ่านคำตอบและทำความรู้จักกับทุกคนมากขึ้น พวกเรารู้สึกประทับใจมากแต่เนื่องจากมีผู้สมัครจำนวนมากและตำแหน่งที่มีจำกัด ทีมงานเสียใจที่ไม่สามารถตอบรับคุณเข้ามาเป็นทีมงานในฝ่าย ${role} ในขณะนี้ได้

\tอย่างไรก็ตาม ทางทีมงานขอให้ท่านไม่ย่อท้อและยินดีต้อนรับท่านเป็นอย่างมากในการสมัครเข้าร่วมโครงการในครั้งต่อไป หากท่านสนใจทางเรายินดีที่จะให้คำแนะนำและข้อมูลเพิ่มเติมเกี่ยวกับการสมัครและการเตรียมตัวสำหรับโครงการในอนาคต

\tทางเราขอขอบคุณท่านสำหรับความสนใจและความพยายามที่ได้แสดงออกมาในการสมัครเข้าร่วมโครงการครั้งนี้ และหวังเป็นอย่างยิ่งว่าจะได้พบกับท่านอีกในโอกาสต่อไป

ด้วยความเคารพ
---------------------------------------------------------------------
เฟม IT#28 ประธานโครงการ IT#30 Starter Pack
Email ติดต่อกลับ: phuwamet.panj@mail.kmutt.ac.th`

    GmailApp.sendEmail(email, subject, message);
    // console.log(subject)
    // console.log('------------')
    // console.log(message)
    console.log("<Not Pass> Success ID:" + id)
    count++
    // console.log('<----------------------------------------->')
  }
  console.log(count)
}
