(_=>{
  /* 
    엑셀파일을 받아 output 파일을 제작한다. 
  
    Input 엑셀파일 시트 정보
      1번. raw_data   : kibana에서 추출된 user_id, option_number, currentQuizNumber (총 세개 필드를 가진 리스트)
      2번. user_info  : user_id, 응시직종, 접수일시, 성명, 휴대폰, 소방본무/직종/학교, 소방서/지부/학과, 대상물명, 이메일 (총 아홉개 필드를 가진 리스트)
      3번. item_info  : 문항번호, 정답, 배점 (총 세개 필드를 가진 리스트)
  
    Output 파일
      1번. output     : user_id, 문항번호(item_info 문항번호 max값)만큼의 컬럼, user_info 시트 내 user_id를 제외한 컬럼 전체
  
    Input 파일 처리 관련
      - userId 데이터를 모두 검색하던 중, currentQuizNumber에 option_number가 -1인 데이터는 처리하지 않는다. (중복정답제출 처리하지 않음)
  */
  const XLSX = require('xlsx')
  const fs = require('fs')
  const buffer = fs.readFileSync('input.xlsx')
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
  
  // 셀 전체를 순회하는 로직이므로 순서가 보장됩니다.
  const getUserInfoFromSheetData = userInfoFromSheet => {
    const userInfo = {}
    let currentUserId = undefined
  
    for(let key in userInfoFromSheet) {
      if(isValidRange(key, 'A', 'Z')) {
        // A: userId
        // B: 응시직종
        // C: 접수일시
        // D: 성명
        // E: 휴대폰
        // F: 소방본부/직종/학교
        // G: 소방서/지부/학과
        // H: 대상물명
        // I: 이메일
        let value = userInfoFromSheet[key].v
        const firstCharacter = key[0]
        
        switch (firstCharacter) {
          case 'A': 
            if(!isNaN(value)) currentUserId = value, userInfo[currentUserId] = {}
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
    const rawData = {} // { 1: {1: 1, 2: 1, ..., 15: 1}} ===> {userId: {questionNumber: answer, ...}}
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
  
  const getOutputObject = (userInfo, itemInfo, rawData) => {
    if(!(userInfo && itemInfo && rawData)) return
  
    const lastNumber = Math.max(...Object.keys(itemInfo).map(item => Number(item)))
  
    let csvHeader = 'userId'
    for(let item in itemInfo) csvHeader += `,${item}`
    csvHeader += ',점수,응시직종,접수일시,성명,휴대폰,소방본부/직종/학교,소방서/지부/학과,대상물명,이메일'
  
    // userId별 row 생성
    let csvRow = ''
    let score // 누적 점수
  
    for(let user in userInfo) {
      score = 0
      csvRow += user
  
      // 문항별 응답 결과 작성
      for(let quizNumber = 1; quizNumber <= lastNumber; quizNumber++) {
        const answerInfo = itemInfo[quizNumber]
        const currentQuizAnswer = rawData[user][quizNumber] || 0
        
        if(answerInfo && currentQuizAnswer === answerInfo.answer) {
          csvRow += ',O'
          score += answerInfo.score
        } else {
          csvRow += ',X'
        }
      }
  
      // 점수 작성
      csvRow += `,${score}`
  
      // 응시직종 ~ 이메일까지 사용자 정보 추가
      for(let currentUser in userInfo[user]) {
        csvRow += `,${userInfo[user][currentUser]}`
      }
      
      csvRow += '\n'
    }
  
    return csvHeader + '\n' + csvRow
  }
  
  const userInfo = getUserInfoFromSheetData(userInfoFromSheet) // { 1: {응시직종: '', 접수일시: '', ..., 이메일: ''}, ... }
  const itemInfo = getItemInfoFromSheetData(itemInfoFromSheet) // { 1: {answer: 1, score: 1}, ..., 15: {answer: 1, score: 1} }
  const rawData = getRawDataFromSheet(rawDataFromSheet) // { 1: {1: 1, 2: 1, ..., 15: 1}} ===> {userId: {questionNumber: answer, ...}}
  
  // {userId: {questionNumber: isCorrect}, score: Number, 응시직종: '', 접수일시: '', ..., 이메일: ''}
  const output = getOutputObject(userInfo, itemInfo, rawData)
  
  // write file
  fs.writeFileSync('./output.csv', output, 'utf8')
})()
