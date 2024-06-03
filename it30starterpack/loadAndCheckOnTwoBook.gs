function loadinterview() {
  const mainSheetName = 'SortResponses'
  const mainBook = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = mainBook.getSheetByName(mainSheetName)
  const mainSheetLastRow = mainSheet.getLastRow()
  const passSheetName = 'passToInterview'
  const notPassSheetName = 'notPassToInterview'

  const colApplyIdMain = 'A'
  const colEmail = 'B'
  const colStdId = 'D'
  const colName = 'H'
  const colNickName = 'I'
  const colTel = 'J'
  const colDiscord = 'K'

  const targetId = 'target-book-id'
  const targetBook = SpreadsheetApp.openById(targetId)

  const colPass = 'A'
  const colApplyIdTarget = 'B'
  const sheetNames = ['ITFund', 'CTP', 'WT']

  const getTargetSheetLastRow = (sheet) => {
    const range = sheet.getRange(colApplyIdTarget + '3:' + colApplyIdTarget).getValues()
    let countRow = 0
    for (let i = 0; i < range.length; i++) {
      if (range[i].toString() == '') return countRow + 2
      countRow++
    }
  }

  const extractDataFromMain = (rowMain) => {
    return [
      mainSheet.getRange(colApplyIdMain + rowMain).getValue().toString(),
      mainSheet.getRange(colEmail + rowMain).getValue().toString(),
      mainSheet.getRange(colStdId + rowMain).getValue().toString(),
      mainSheet.getRange(colName + rowMain).getValue().toString(),
      mainSheet.getRange(colNickName + rowMain).getValue().toString(),
      mainSheet.getRange(colTel + rowMain).getValue().toString(),
      mainSheet.getRange(colDiscord + rowMain).getValue().toString()
    ]
  }
  
  const count = {
    pass: 2,
    noPass: 2
  }

  const passApplyId = []
  let notPassApplyId = []

  // get all pass and primary role
  for (const sheetName of sheetNames) {
    const targetSheet = targetBook.getSheetByName(sheetName)
    const targetSheetLastRow = getTargetSheetLastRow(targetSheet)

    for (let targetRow = 3; targetRow <= targetSheetLastRow; targetRow++) {
      const passCell = targetSheet.getRange(colPass + targetRow).getValue().toString()
      const pass = passCell.toLocaleLowerCase() == 'true'

      const targetApplyId = targetSheet.getRange(colApplyIdTarget + targetRow).getValue().toString()

      if (pass && targetApplyId.includes(sheetName)) {
        passApplyId.push(targetApplyId)
      } else if (!pass && targetApplyId.includes(sheetName)) {
        notPassApplyId.push(targetApplyId)
      }
    }
  }

  const copyNotPassId = [...notPassApplyId]
  for (const notPass of copyNotPassId) {
    const getSheetNameFromId = notPass.slice(0, -2)
    const checkSheets = sheetNames.filter((sheetName) => sheetName != getSheetNameFromId)
    for (const sheetName of checkSheets) {
      const targetSheet = targetBook.getSheetByName(sheetName)
      const targetSheetLastRow = getTargetSheetLastRow(targetSheet)

      for (let targetRow = targetSheetLastRow; targetRow >= 3; targetRow--) {
        const passCell = targetSheet.getRange(colPass + targetRow).getValue().toString()
        const pass = passCell.toLocaleLowerCase() == 'true'

        const targetApplyId = targetSheet.getRange(colApplyIdTarget + targetRow).getValue().toString()
        if (pass && targetApplyId == notPass) {
          console.log('pass', targetApplyId)
          passApplyId.push(targetApplyId)
          
          notPassApplyId = notPassApplyId.filter((id) => id != targetApplyId)
          break
        }
      }
    }
  }
  
  for (const idPass of passApplyId) {
    const passToInterviewSheet = mainBook.getSheetByName(passSheetName)
    
    for (let row = 2; row <= mainSheetLastRow; row++) {
      const findApplyId = mainSheet.getRange(colApplyIdMain + row).getValue().toString()

      if (findApplyId == idPass) {
        const data = extractDataFromMain(row)
        copyDataToTargetSheet(passToInterviewSheet, count.pass++, data)

        break
      }
    }
  }

  for (const idNotPass of notPassApplyId) {
    const notPassToInterviewSheet = mainBook.getSheetByName(notPassSheetName)

    for (let row = 2; row <= mainSheetLastRow; row++) {
      const findApplyId = mainSheet.getRange(colApplyIdMain + row).getValue().toString()

      if (findApplyId == idNotPass) {
        const data = extractDataFromMain(row)
        copyDataToTargetSheet(notPassToInterviewSheet, count.noPass++, data)

        break
      }
    }
  }
}

const copyDataToTargetSheet = (targetSheet, row, data) => {
  data.forEach((value, index) => {
    targetSheet.getRange(getColLetter(index + 1) + row).setValue(value)
  })
}
