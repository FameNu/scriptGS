function moveToSheet() {
  // Config Sheet
  // This Sheet
  const devMode = false
  const sheetName = devMode ? 'Form Responses 1' : 'SortResponses'
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  const lastRow = sheet.getLastRow()

  // target sheet
  const sheetAllStaffId = 'target-sheet-id'
  const sheetSpeakerId = 'target-sheet-id'

  const colApplyId = 'B'
  const colApplYear = 'C'
  const colTargetSheetStartAns = 'D'

  const colTimeStmap = 'A'
  const colStdId = 'D'
  const colYear = 'E'
  const colRole = 'N'
  const colRoleSpeakerSecond = {
    'ITFund': 'AB',
    'CTP': 'AO',
    'WT': 'BA'
  }

  // speaker (introduce)
  const colRoleSpeakerPrimary = 'U'

  const ansRanges = {
    'Speaker': { startCol: 'Q', endCol: 'T' },
    'Art & Design': { startCol: 'BG', endCol: 'BK' },
    'PR': { startCol: 'BL', endCol: 'BP' },
    'Copyreader': { startCol: 'BQ', endCol: 'BT' },
    'HR': { startCol: 'BU', endCol: 'BY' },
    'Technical': { startCol: 'BZ', endCol: 'CD' }
  };

  const speakerAnsRange = {
    'ITFund Primary': { startCol: 'V', endCol: 'AA' },
    'ITFund Secondary': { startCol: 'AC', endCol: 'AH' },
    'CTP Primary': { startCol: 'AI', endCol: 'AN' },
    'CTP Secondary': { startCol: 'AP', endCol: 'AU' },
    'WT Primary': { startCol: 'AV', endCol: 'AZ' },
    'WT Secondary': { startCol: 'BB', endCol: 'BF' }
  }

  const speakerPrimaryCount = (priRole) => {
    let count = 0
    for (let row = 2; row <= lastRow; row++) {
      const role = sheet.getRange(colRoleSpeakerPrimary + row).getValue().toString()
      if (priRole == fullRoleToShort(role)) count++
    }
    return count
  }

  const speakerPrimaryRoleCount = {
    'ITFund': speakerPrimaryCount('ITFund'),
    'CTP': speakerPrimaryCount('CTP'),
    'WT': speakerPrimaryCount('WT')
  }

  let count = {
    'Art & Design': 0,
    'PR': 0,
    'Copyreader': 0,
    'HR': 0,
    'Technical': 0,
    'ITFund': 0,
    'CTP': 0,
    'WT': 0,
  }
  const colCount = getColIndex(colTargetSheetStartAns)

  for (let row = 2; row <= lastRow; row++) {
    const role = sheet.getRange(colRole + row).getValue().toString()
    const year = sheet.getRange(colYear + row).getValue().toString()

    const useThatRole = (Object.keys(ansRanges)).find((key) => role.toLocaleLowerCase().includes(key.toLocaleLowerCase()))
    if (!useThatRole) continue

    const mainQuestions = runCols(sheet, 1, ansRanges[useThatRole])
    const mainAnswers = runCols(sheet, row, ansRanges[useThatRole])

    if (
      useThatRole == 'Speaker'
    ) {
      const primarySpeaker = sheet.getRange(colRoleSpeakerPrimary + row).getValue().toString()
      const primarySpeakerShort = fullRoleToShort(primarySpeaker)

      const targetBook = SpreadsheetApp.openById(sheetSpeakerId)
      const targetSheet = targetBook.getSheetByName(primarySpeakerShort)
      const questionsEachDepartment = runCols(sheet, 1, speakerAnsRange[primarySpeakerShort + ' Primary'])
      const questions = mainQuestions.concat(questionsEachDepartment)

      const answerEachDepartment = runCols(sheet, row, speakerAnsRange[primarySpeakerShort + ' Primary'])
      const primaryAnswers = mainAnswers.concat(answerEachDepartment)

      for (let i = 0; i < primaryAnswers.length; i++) {
        const colLetter = getColLetter(colCount + i)

        if (count[primarySpeakerShort] == 0) targetSheet.getRange(colLetter + 2).setValue(questions[i])

        const rowPriSheet = count[primarySpeakerShort] + 3

        const applyId = primarySpeakerShort + (rowPriSheet - 2).toString().padStart(2, '0')
        if (!devMode) sheet.getRange(colTimeStmap + row).setValue(applyId)

        targetSheet.getRange(colApplyId + rowPriSheet).setValue(applyId)
        targetSheet.getRange(colApplYear + rowPriSheet).setValue(year)
        targetSheet.getRange(colLetter + rowPriSheet).setValue(primaryAnswers[i])

        const speakerSecRole = sheet.getRange(colRoleSpeakerSecond[primarySpeakerShort] + row).getValue().toString()
        const speakerSecRoleShort = fullRoleToShort(speakerSecRole)
        if ( (Object.keys(colRoleSpeakerSecond)).some((key) => key.includes(speakerSecRoleShort)) ) {
          const secTargetSheet = targetBook.getSheetByName(speakerSecRoleShort)

          const answerSecDepartment = runCols(sheet, row, speakerAnsRange[speakerSecRoleShort + ' Secondary'])
          const secondAnswers = mainAnswers.concat(answerSecDepartment)

          const rowSecSheet = speakerPrimaryRoleCount[speakerSecRoleShort] + 5

          secTargetSheet.getRange(colApplyId + rowSecSheet).setValue(applyId)
          secTargetSheet.getRange(colApplYear + rowSecSheet).setValue(year)
          secTargetSheet.getRange(colLetter + rowSecSheet).setValue(secondAnswers[i])
        }
      }
      count[primarySpeakerShort]++
    } else {
      const targetSheet = SpreadsheetApp.openById(sheetAllStaffId).getSheetByName(useThatRole)


      for (let i = 0; i < mainAnswers.length; i++) {
        const colLetter = getColLetter(colCount + i)
        const rowTarget = count[useThatRole] + 3

        if (count[useThatRole] == 0) targetSheet.getRange(colLetter + 2).setValue(mainQuestions[i])

        const applyId = fullRoleToShort(useThatRole) + (rowTarget - 2).toString().padStart(2, '0')
        if (!devMode) sheet.getRange(colTimeStmap + row).setValue(applyId)

        targetSheet.getRange(colApplyId + rowTarget).setValue(applyId)
        targetSheet.getRange(colApplYear + rowTarget).setValue(year)
        targetSheet.getRange(colLetter + rowTarget).setValue(mainAnswers[i])
      }
      count[useThatRole]++
    }
  }
}

function fullRoleToShort(role) {
  switch (role) {
    case 'Web Technology':
      return 'WT'
    case 'Information Technology Fundamental':
      return 'ITFund'
    case 'Computational Thinking and Programming':
      return 'CTP'
    case 'Copyreader':
      return 'CR'
    case 'Technical':
      return 'TNC'
    case 'Art & Design':
      return 'AD'
    default:
      return role
  }
}

const runCols = (sheet, row, { startCol, endCol }) => {
  const anses = []
  const startIndex = getColIndex(startCol)
  const lastIndex = getColIndex(endCol)
  for (let col = startIndex; col <= lastIndex; col++) {
    const colLetter = getColLetter(col)
    const cell = sheet.getRange(colLetter + row).getValue().toString()

    anses.push(cell)
  }
  return anses
}

const getColIndex = (col) => {
  const letters = col.toUpperCase().split('')
  let index = 0
  for (let i = 0; i < letters.length; i++) {
    index *= 26
    index += letters[i].charCodeAt(0) - 64;
  }
  return index
}

const getColLetter = (index) => {
  let letter = ''
  while (index > 0) {
    const modulo = (index - 1) % 26
    letter = String.fromCharCode(65 + modulo) + letter;
    index = Math.floor((index - 1) / 26)
  }
  return letter
}
