function sendEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sendEmail')
  
  const lastRow = sheet.getLastRow()

  const pasteCol = {
    id: 'A',
    name: 'B',
    email: 'C',
    gmail: 'D'
  }

  const subject = `ขอเชิญชวนนักศึกษาสาขาเทคโนโลยีสารสนเทศชั้นปีที่ 1 เข้าร่วมโครงการ “IT#30 Starter Pack”`

  const imageId = 'image-id-on-google-drive'
  const image = DriveApp.getFileById(imageId)

  for (let i = 2; i <= lastRow; i++) {
    const id = sheet.getRange(pasteCol.id + i).getValue().toString()
    const name = sheet.getRange(pasteCol.name + i).getValue().toString()
    const email = sheet.getRange(pasteCol.email + i).getValue().toString()
    const gmail = sheet.getRange(pasteCol.gmail + i).getValue().toString()

    const message = `ถึง ${name} รหัสนักศึกษา ${id}

\tเนื่องจากการประชาสัมพันธ์ของโครงการ IT#30 Starter Pack ที่ผ่านมาไม่ทั่วถึงอย่างที่ควรทางคณะดำเนินงานจึงได้มีการส่ง Email ฉบับนี้เพื่อเชิญชวนนักศึกษาสาขาเทคโนโลยีสารสนเทศชั้นปีที่ 1 เข้าสู่โครงการ “IT#30 Starter Pack”

\tเนื่องด้วยนักศึกษาสาขาวิชาเทคโนโลยีสารสนเทศชั้นปี 2 และ 3 ได้มีการจัดโครงการปรับพื้นฐานสำหรับนักศึกษาสาขาเทคโนโลยีสารสนเทศชั้นปีที่ 1 เพื่อเตรียมความพร้อมก่อนเปิดภาคการศึกษา (เป็นกิจกรรมสมัครใจ ไม่บังคับเข้าร่วม)

โดยในโครงการจะประกอบด้วยการเรียนการสอนและกิจกรรมต่าง ๆ เพื่อมอบความสุข และเนื้อหาสาระให้กับน้อง ๆ

โดยวิชาที่จัดสอนมีดังนี้:
- Web Technology
- IT Fundamental
- Computational

กิจกรรมจะจัดในวันที่ 23 - 26 กรกฏาคม 2567 โดยเป็นการจัดกิจกรรมในรูปแบบ On-Site ณ ชั้น 11 อาคารการเรียนรู้พหุวิทยาการ (Learning Exchange:LX) มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี

หากนักศึกษาชั้นปีที่ 1 มีความสนใจสามารถลงทะเบียนใน Link ของ Google Form หรือ Scan QR Code ที่แนบมากับ Email ฉบับนี้ ภายในวันที่ 15 กรกฏาคม 2567 เวลา 23.59 น.
และจะประกาศผู้ที่มีสิทธิ์เข้าร่วม ภายในวันที่ 17 กรกฏาคม 2567 (แจ้งทาง Email)

สามารถติดตามข่าวสารเพิ่มเติม หรือหากมีข้อสงสัยสามารถติดต่อได้ที่ Instragram: sit.it.starterpack หรือคลิก Link Instragram ที่แนบมากับ Email ฉบับนี้

Google Form: https://forms.gle/1VJKYn5ZUaCv7upS7

Instragram: https://www.instagram.com/sit.it.starterpack/

เฟม IT#28
ประธานโครงการ IT#30 Starter Pack`
// const message = `<p>ถึง ${name} รหัสนักศึกษา ${id}</p>
//       <p>เนื่องจากการประชาสัมพันธ์ของโครงการ IT#30 Starter Pack ที่ผ่านมาไม่ทั่วถึงอย่างที่ควรทางคณะดำเนินงานจึงได้มีการส่ง Email ฉบับนี้ เพื่อเชิญชวนนักศึกษาสาขาเทคโนโลยีสารสนเทศชั้นปีที่ 1 เข้าสู่โครงการ “IT#30 Starter Pack”</p>
//       <p>เนื่องด้วยนักศึกษาสาขาวิชาเทคโนโลยีสารสนเทศชั้นปี 2 และ 3 ได้มีการจัดโครงการปรับพื้นฐานสำหรับนักศึกษาสาขาเทคโนโลยีสารสนเทศชั้นปีที่ 1 เพื่อเตรียมความพร้อมก่อนเปิดภาคการศึกษา<br>(เป็นกิจกรรมสมัครใจ ไม่บังคับเข้าร่วม)</p>
//       <p>โดยในโครงการจะประกอบด้วยการเรียนการสอนและกิจกรรมต่าง ๆ เพื่อมอบความสุข และเนื้อหาสาระให้กับน้อง ๆ</p>
//       <div>
//         <p>โดยวิชาที่จัดสอนมีดังนี้</p>
//         <ul>
//           <li>Web Technology</li>
//           <li>IT Fundamental</li>
//           <li>Computational</li>
//         </ul>
//       </div>
//       <p>กิจกรรมจะจัดในวันที่ 23 - 26 กรกฏาคม 2567 โดยเป็นการจัดกิจกรรมในรูปแบบ On-Site ณ ชั้น 11 อาคารเรียนรู้พหุวิทยาการ (Learning Exchange:LX) มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี</p>
//       <p>หากนักศึกษาชั้นปีที่ 1 มีความสนใจสามารถลงทะเบียนใน Link ของ Google Form หรือ Scan QR Code ที่แนบมากับ Email ฉบับนี้ ภายในวันที่ 15 กรกฏาคม 2567 เวลา 23.59 น.<br>และจะประกาศผู้ที่มีสิทธิ์เข้าร่วม ภายในวันที่ 17 กรกฏาคม 2567 (แจ้งทาง Email)</p>
//       <p>สามารถติดตามข่าวสารเพิ่มเติม หรือหากมีข้อสงสัยสามารถติดต่อได้ที่ Instragram: sit.it.starterpack หรือคลิก Link Instragram ที่แนบมากับ Email ฉบับนี้</p>
//       <p>Google Form: <a href="https://forms.gle/1VJKYn5ZUaCv7upS7">https://forms.gle/1VJKYn5ZUaCv7upS7</a></p>
//       <p>Instragram: <a href="https://www.instagram.com/sit.it.starterpack/">https://www.instagram.com/sit.it.starterpack/</a></p>
//       <p>เฟม IT#28<br>ประธานโครงการ IT#30 Starter Pack</p>`


    // MailApp.sendEmail({
    //   to: email,
    //   subject: subject,
    //   body: message,
    //   attachments: [image]
    // })
    // MailApp.sendEmail({
    //   to: gmail,
    //   subject: subject,
    //   body: message,
    //   attachments: [image]
    // })
    MailApp.sendEmail({
      to: 'phuwamet.panj@kmutt.ac.th',
      subject: subject,
      body: message,
      // htmlBody: message,
      attachments: [image]
    })
    break
  }
}
