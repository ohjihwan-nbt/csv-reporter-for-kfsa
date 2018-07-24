(_=>{
  const [,,inputExcelFileName = 'input.xlsx', outputFileNameWithoutExtension = 'output'] = process.argv
  const startTime = Date.now()
  const XLSX = require('xlsx')
  const fs = require('fs')
  const buffer = fs.readFileSync(inputExcelFileName)
  const workBooks = XLSX.read(buffer, {type: 'buffer'})
  
  const {
    user_info: userInfoFromSheet, 
    raw_data: rawDataFromSheet, 
    item_info: itemInfoFromSheet 
  } = workBooks.Sheets
  
  const isValidRange = (str, startOf = 'A', endOf = 'Z') => {
    const firstCharacter = str[0]
    return firstCharacter.charCodeAt() >= startOf.charCodeAt() && firstCharacter.charCodeAt() <= endOf.charCodeAt()
  }
  
  const getUserInfoFromSheetData = userInfoFromSheet => {
    const userInfo = {}
    let currentUserId = undefined
  
    for(let key in userInfoFromSheet) {
      if(isValidRange(key, 'A', 'Z')) {
        let value = userInfoFromSheet[key].v
        const firstCharacter = key[0]

        switch (firstCharacter) {
          case 'A': 
            if(!isNaN(String(value).trim())) currentUserId = value, userInfo[currentUserId] = {}
            break
          default: 
            const currentUserInfo = userInfo[currentUserId] || {}
            const header = userInfoFromSheet[`${firstCharacter}1`].v
  
            currentUserInfo[header] = value
            break
        }
      }
    }
  
    return userInfo
  }
  
  const getItemInfoFromSheetData = itemInfoFromSheet => {
    const itemInfo = {}
    let currentItemNumber = undefined
  
    for(let key in itemInfoFromSheet) {
      if(isValidRange(key, 'A', 'C')) {
        let value = itemInfoFromSheet[key].v
        const firstCharacter = key[0]
        
        switch (firstCharacter) {
          case 'A': 
            if(!isNaN(value)) currentItemNumber = value, itemInfo[currentItemNumber] = {answer: undefined, score: undefined}
            break
          case 'B':
            if(currentItemNumber) itemInfo[currentItemNumber]['answer'] = value
            break
          case 'C': 
            if(currentItemNumber) itemInfo[currentItemNumber]['score'] = value
            break
          default: break
        }
      }
    }
  
    return itemInfo
  }
  
  const getRawDataFromSheet = rawDataFromSheet => {
    const rawData = {}
    let currentUserId = undefined
    let currentQuizNumber = undefined
  
    for(let key in rawDataFromSheet) {
      if(isValidRange(key, 'A', 'C')) {
        let value = rawDataFromSheet[key].v
        const firstCharacter = key[0]
  
        switch (firstCharacter) {
          case 'A': 
            if(!isNaN(value)) currentUserId = value
            if(!rawData[currentUserId]) rawData[currentUserId] = {}
            break
          case 'B':
            if(currentUserId) currentQuizNumber = value
            break
          case 'C': 
            if(currentUserId && value !== -1) rawData[currentUserId][currentQuizNumber] = value
            break
          default: break
        }
      }
    }
  
    return rawData
  }

  const getParticipantsList = rawDataFromSheet => {
    const participantsList = []

    for(let key in rawDataFromSheet) {
      const value = rawDataFromSheet[key].v

      if(key[0] === 'A' && !participantsList.includes(value) && !isNaN(value)) {
        participantsList.push(value)
      }
    }
    
    return participantsList
  }
  
  const isInvalidValue = value => value === ' ' || isNaN(value)
  
  const getOutputObject = (participantsList, userInfo, itemInfo, rawData) => {
    if(!(participantsList && userInfo && itemInfo && rawData)) return
  
    const lastNumber = Math.max(...Object.keys(itemInfo).map(item => Number(item)))
    const SEPERATOR = '\t'
  
    let csvHeader = 'userId'
    for(let item in itemInfo) csvHeader += `${SEPERATOR}${item}`
    csvHeader += `${SEPERATOR}점수${SEPERATOR}응시직종${SEPERATOR}접수일시${SEPERATOR}성명${SEPERATOR}휴대폰${SEPERATOR}소방본부/직종/학교${SEPERATOR}소방서/지부/학과${SEPERATOR}대상물명${SEPERATOR}이메일`
  
    let csvRow = ''
    let score
  
    for(let user in userInfo) {
      if(isInvalidValue(user)) continue
      else if(!participantsList.includes(Number(user))) continue
      
      score = 0
      csvRow += user
  
      for(let quizNumber = 1; quizNumber <= lastNumber; quizNumber++) {
        const answerInfo = itemInfo[quizNumber]
        const currentQuizAnswer = rawData[user] && rawData[user][quizNumber] || 0
        
        if(answerInfo && currentQuizAnswer === answerInfo.answer) {
          csvRow += `${SEPERATOR}O`
          score += answerInfo.score
        } else {
          csvRow += `${SEPERATOR}X`
        }
      }
  
      csvRow += `${SEPERATOR}${parseFloat(score).toFixed(1)}`
  
      for(let currentUser in userInfo[user]) {
        csvRow += `${SEPERATOR}${userInfo[user][currentUser]}`
      }
      
      csvRow += '\n'
    }
  
    return csvHeader + '\n' + csvRow
  }
  
  const userInfo = getUserInfoFromSheetData(userInfoFromSheet)
  const itemInfo = getItemInfoFromSheetData(itemInfoFromSheet)
  const rawData = getRawDataFromSheet(rawDataFromSheet)
  const participantsList = getParticipantsList(rawDataFromSheet)

  const output = getOutputObject(participantsList, userInfo, itemInfo, rawData)
  
  fs.writeFileSync(`./${outputFileNameWithoutExtension}.xls`, output, 'utf8')

  const consumedTimeInSeconds = (Date.now() - startTime) / 1000
  console.info(`${outputFileNameWithoutExtension}.xls write has completed in ${consumedTimeInSeconds} seconds.`)
})()
