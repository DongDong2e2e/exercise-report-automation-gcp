// =================================================================
// ===================== ⚙️ 전역 설정 관리 ⚙️ =====================
// =================================================================

/**
 * 프로젝트의 모든 주요 설정값을 객체 형태로 반환하는 함수입니다.
 * 모든 설정은 이 함수 내에서 관리하여 안정성을 높입니다.
 */
function getProjectConfig() {
  // 1. 스크립트 속성에서 민감한 정보를 불러옵니다.
  const scriptProperties = PropertiesService.getScriptProperties();
  
  return {
    // --- 민감 정보 (Script Properties에서 관리) ---
    GEMINI_API_KEY: scriptProperties.getProperty('GEMINI_API_KEY'),
    REPORT_RECIPIENT_EMAIL: scriptProperties.getProperty('REPORT_RECIPIENT_EMAIL'),
    USER_NAME: scriptProperties.getProperty('USER_NAME'),

    // --- 시트 및 파일 이름 설정 (여기서 직접 수정 가능) ---
    RAW_DATA_SHEET_PREFIX: '운동데이터_',
    STRUCTURED_LOG_SHEET: 'structured_log',
    MAPPING_SHEET: '운동분류',
    INBODY_SHEET: 'Inbody_data',

    // --- [추가됨] Debounce 트리거 관리를 위한 설정 ---
    DEBOUNCE_TRIGGER_HANDLER: 'processDataUpdate' // 임시 트리거가 실행할 함수 이름
  };
}

// --- 설정 끝 ---


/**
 * 🛠️ 최초 1회만 실행하여 프로젝트를 설정하는 함수
 */
function setup() {
  const config = getProjectConfig(); // 설정값 불러오기
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName(config.STRUCTURED_LOG_SHEET)) {
    const sheet = ss.insertSheet(config.STRUCTURED_LOG_SHEET);
    const header = [
      '날짜', '운동명', '세트_구분', '세트번호', '무게(kg)', '횟수/시간', '단위', 
      '볼륨(kg)', '대분류', '도구', '움직임', '주동근'
    ];
    sheet.appendRow(header);
    Logger.log(`'${config.STRUCTURED_LOG_SHEET}' 시트를 생성했습니다.`);
  }

  // 기존의 모든 트리거 삭제 (onEdit 포함)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log("기존의 모든 트리거를 삭제했습니다.");

  // [수정됨] onEdit 트리거 설정 (이제 이 트리거는 '예약'만 담당)
  ScriptApp.newTrigger('runOnEditTrigger').forSpreadsheet(ss).onEdit().create();
  Logger.log("'runOnEditTrigger'가 설정되었습니다.");
  
  ScriptApp.newTrigger('sendWeeklyReportTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Logger.log("'sendWeeklyReportTrigger'가 매주 월요일 오전 8-9시로 설정되었습니다.");
  
  ScriptApp.newTrigger('monthlyTasksTrigger').timeBased().onMonthDay(1).atHour(8).create();
  Logger.log("'monthlyTasksTrigger'가 매월 1일 오전 8-9시로 설정되었습니다.");
  
  SpreadsheetApp.getUi().alert('✅ 모든 리포트(주간/월간/분기/연간) 트리거 설정이 완료되었습니다!');
}


// --- ⏰ 트리거 실행 함수들 ---

/**
 * [수정됨] Debouncing 기법이 적용된 onEdit 트리거 함수입니다.
 * 실제 데이터 처리를 직접 하지 않고, 90초 뒤에 실행될 '임시 트리거'를 생성/갱신하는 역할만 합니다.
 */
function runOnEditTrigger(e) {
  const config = getProjectConfig();
  try {
    // 1. 수정된 시트가 '운동데이터_'로 시작하는지 확인합니다. 아니라면 즉시 종료.
    const sheetName = e.source.getActiveSheet().getName();
    if (!sheetName.startsWith(config.RAW_DATA_SHEET_PREFIX)) {
      return;
    }

    // 2. 기존에 만들어졌던 '임시 트리거'가 있다면 삭제합니다.
    //    사용자가 연속으로 타이핑할 때마다 이전 예약을 취소하는 효과를 줍니다.
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === config.DEBOUNCE_TRIGGER_HANDLER) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 3. 90초 후에 실제 데이터 처리 함수(processDataUpdate)를 딱 한 번 실행할 '임시 트리거'를 새로 생성합니다.
    ScriptApp.newTrigger(config.DEBOUNCE_TRIGGER_HANDLER)
      .timeBased()
      .after(90 * 1000) // 90초 (1.5분)
      .create();
    
    Logger.log(`'${sheetName}' 시트 수정 감지. 90초 후 데이터 동기화를 예약합니다.`);

  } catch (err) {
    Logger.log(`onEdit 트리거 예약 오류: ${err.message}`);
  }
}

/**
 * [추가됨] 임시 트리거에 의해 실제로 데이터 처리를 실행하는 함수입니다.
 * 이 함수가 호출된다는 것은 사용자의 마지막 수정 후 90초가 지났음을 의미합니다.
 */
function processDataUpdate() {
  Logger.log("예약된 데이터 업데이트를 시작합니다.");
  updateStructuredLogSheet();
}


function monthlyTasksTrigger() {
  const today = new Date();
  const month = today.getMonth() + 1;

  if (month === 1) { sendReport('year'); } 
  else if ([4, 7, 10].includes(month)) { sendReport('quarter'); } 
  else { sendReport('month'); }
}

function sendWeeklyReportTrigger() {
  sendReport('week');
}

// --- 데이터 파싱 및 동기화 함수들 ---
function updateStructuredLogSheet() { 
  const config = getProjectConfig();
  try { 
    const infoMap = getExerciseInfoMap(); 
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const targetSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(config.RAW_DATA_SHEET_PREFIX)); 
    if (targetSheets.length === 0) return; 
    const allParsedData = []; 
    targetSheets.forEach(sheet => { parseSheetData(sheet, infoMap, allParsedData); }); 
    syncDataToSheet(allParsedData); // 수정된 syncDataToSheet 함수 호출
    Logger.log("데이터 변환 및 동기화 완료."); 
  } catch (e) { 
    Logger.log(`파싱/동기화 오류: ${e.stack}`); 
  } 
}

function getExerciseInfoMap() { 
  const config = getProjectConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const mappingSheet = ss.getSheetByName(config.MAPPING_SHEET); 
  if (!mappingSheet) throw new Error(`'${config.MAPPING_SHEET}' 시트 없음.`); 
  const data = mappingSheet.getDataRange().getValues(); 
  const map = {}; 
  for (let i = 1; i < data.length; i++) { 
    const name = data[i][0]; 
    if (!name || name.startsWith('**')) continue; 
    map[name.trim()] = { 
      category: data[i][1]?.trim() || '미분류', 
      calcMultiplier: (data[i][2] == 2) ? 2 : 1, 
      tool: data[i][3]?.trim() || '', 
      movement: data[i][4]?.trim() || '', 
      target: data[i][5]?.trim() || '' 
    }; 
  } 
  return map; 
}

function parseSheetData(sheet, infoMap, allParsedData) { 
  const data = sheet.getDataRange().getValues(); 
  const datePattern = /(\d{4})[-.\s]*(\d{1,2})[-.\s]*(\d{1,2}).*/; 
  const setPattern = /^(?:(\d+)\s*세트|Warm-up)\s*(?:\((F|D)\))?:\s*([\d.]+)\s*(kg|lbs)\s*([\d.]+)\s*(?:회|reps)/i; 
  const setPatternRepsOnly = /^(?:(\d+)\s*세트|Warm-up)\s*(?:\((F|D)\))?:\s*([\d.]+)\s*(?:회|reps)/i; 
  const setPatternTime = /^(?:(\d+)\s*세트|Warm-up)\s*(?:\((F|D)\))?:\s*([\d.]+)\s*(초|분|시간|min|sec|s)/i; 
  const LBS_TO_KG = 0.453592; 
  let currentDate = null; 
  let currentExercise = null; 
  for (const row of data) { 
    const line = row[0].toString().trim(); 
    if (!line || line.includes("기록이 몸을 만든다")) continue; 
    let dateMatch = line.match(datePattern); 
    if (dateMatch) { 
      currentDate = `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`; 
      currentExercise = null; 
      continue; 
    } 
    if (!line.includes(':') && !/^\d/.test(line) && isNaN(line[0])) { 
      currentExercise = line.trim(); 
      continue; 
    } 
    if (currentDate && currentExercise) { 
      let match, setType = '본세트', setNumStr, weight = 0, repsOrTime = 0, unit = '', volume = 0; 
      if (line.includes('(F)')) setType = '실패세트'; 
      else if (line.includes('(D)')) setType = '드롭세트'; 
      else if (line.toLowerCase().startsWith('warm-up')) setType = '웜업'; 
      if (match = line.match(setPattern)) { 
        setNumStr = match[1]; 
        let rawWeight = parseFloat(match[3]); 
        weight = (match[4].toLowerCase() === 'lbs') ? rawWeight * LBS_TO_KG : rawWeight; 
        repsOrTime = parseFloat(match[5]); 
        unit = '회'; 
      } else if (match = line.match(setPatternRepsOnly)) { 
        setNumStr = match[1]; 
        repsOrTime = parseFloat(match[3]); 
        unit = '회'; 
      } else if (match = line.match(setPatternTime)) { 
        setNumStr = match[1]; 
        repsOrTime = parseFloat(match[3]); 
        unit = (match[4].toLowerCase() === '분' || match[4] === 'min') ? '분' : '초'; 
      } else { 
        continue; 
      } 
      const setNum = setType === '웜업' ? 'Warm-up' : (setNumStr || '1'); 
      const info = infoMap[currentExercise] || { category: '미분류', calcMultiplier: 1, tool: '', movement: '', target: '' }; 
      if (unit === '회') { 
        volume = weight * repsOrTime * info.calcMultiplier; 
      } 
      allParsedData.push([currentDate, currentExercise, setType, setNum, weight, repsOrTime, unit, volume, info.category, info.tool, info.movement, info.target]); 
    } 
  } 
}

/**
 * [수정됨] 데이터를 시트에 동기화하는 '스마트 업데이트' 방식 함수입니다.
 * 불필요한 데이터 삭제/쓰기 작업을 최소화하여 성능을 개선합니다.
 */
function syncDataToSheet(allData) { 
  const config = getProjectConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const logSheet = ss.getSheetByName(config.STRUCTURED_LOG_SHEET); 

  // 1. 파싱된 모든 데이터를 날짜, 운동명, 세트 순으로 정렬합니다.
  allData.sort((a, b) => { 
    if (a[0] > b[0]) return 1; if (a[0] < b[0]) return -1;
    if (a[1] > b[1]) return 1; if (a[1] < b[1]) return -1;
    const setA = isNaN(a[3]) ? 0 : parseInt(a[3]); 
    const setB = isNaN(b[3]) ? 0 : parseInt(b[3]); 
    return setA - setB; 
  }); 

  const newDataRowCount = allData.length;
  const oldDataRowCount = logSheet.getLastRow() - 1; // 헤더 제외

  // 2. 새 데이터가 있으면 시트에 덮어씁니다.
  if (newDataRowCount > 0) {
    logSheet.getRange(2, 1, newDataRowCount, allData[0].length).setValues(allData);
  }

  // 3. 만약 기존 데이터가 새 데이터보다 많았다면, 남는 부분을 깔끔하게 지웁니다.
  if (oldDataRowCount > newDataRowCount) {
    const startRowToClear = newDataRowCount + 2; // 지우기 시작할 행 번호
    const numRowsToClear = oldDataRowCount - newDataRowCount; // 지울 행의 개수
    logSheet.getRange(startRowToClear, 1, numRowsToClear, logSheet.getLastColumn()).clearContent();
    Logger.log(`${numRowsToClear}개의 오래된 데이터를 시트에서 삭제했습니다.`);
  }
}


// =================================================================
// ================= ✨ 4단계 고도화 아키텍처 적용 ✨ =================
// =================================================================

// (이하 모든 코드는 원본과 동일하게 매우 훌륭하므로 수정하지 않았습니다.)

function sendReport(reportType) {
  const config = getProjectConfig();
  try {
    Logger.log(`[${reportType}] 4단계 리포트 생성을 시작합니다.`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(config.STRUCTURED_LOG_SHEET);
    const inbodySheet = ss.getSheetByName(config.INBODY_SHEET);

    if (!logSheet || !inbodySheet) throw new Error("필수 시트를 찾을 수 없습니다.");
    
    const stats = analyzeDataForPeriod(logSheet, inbodySheet, reportType);
    
    if (stats.current.totalWorkoutDays === 0) { 
      Logger.log(`이번 ${stats.periodName} 운동 기록이 없어 리포트를 발송하지 않습니다.`);
      return; 
    }
    
    Logger.log(`[${reportType}] 1단계: 과거 데이터 컨텍스트 요약 시작`);
    const historyContext = (stats.previous.totalWorkoutDays > 0) 
      ? callGeminiAPI(createHistoryAnalysisPrompt(stats), 'text')
      : "이전 기간의 운동 기록이 없습니다.";
    Logger.log(`[${reportType}] 1단계 완료. 요약된 과거 컨텍스트: \n${historyContext}`);

    Logger.log(`[${reportType}] 2단계: 현재 데이터 심층 분석 시작`);
    const tacticalAnalysis = callGeminiAPI(createTacticalAnalysisPrompt(stats, historyContext), 'text');
    Logger.log(`[${reportType}] 2단계 완료. 도출된 심층 인사이트: \n${tacticalAnalysis}`);

    Logger.log(`[${reportType}] 3단계: 맞춤형 루틴 생성 시작`);
    const recommendedRoutine = callGeminiAPI(createRoutineGenerationPrompt(stats, tacticalAnalysis), 'text');
    Logger.log(`[${reportType}] 3단계 완료. 생성된 추천 루틴: \n${recommendedRoutine}`);

    Logger.log(`[${reportType}] 4단계: 최종 리포트 생성 시작`);
    const reportHtml = callGeminiAPI(createFinalReportPrompt(stats, reportType, tacticalAnalysis, recommendedRoutine), 'html');
    
    const subject = `💪 ${config.USER_NAME}님, ${stats.periodName} 운동 리포트 + 맞춤 루틴이 도착했습니다!`;
    MailApp.sendEmail({ to: config.REPORT_RECIPIENT_EMAIL, subject: subject, htmlBody: reportHtml });
    Logger.log(`[${reportType}] 리포트 이메일을 성공적으로 발송했습니다.`);

  } catch (e) {
    Logger.log(`[${reportType}] 리포트 생성 오류: ${e.toString()}\n${e.stack}`);
    const config = getProjectConfig();
    MailApp.sendEmail(config.REPORT_RECIPIENT_EMAIL, `🚨 [${reportType}] 운동 리포트 생성 오류`, `오류가 발생했습니다: ${e.message}\n\n${e.stack}`);
  }
}

function analyzeDataForPeriod(logSheet, inbodySheet, periodType) {
  const config = getProjectConfig();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const formatDate = (date) => Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  let startDate, endDate, prevStartDate, prevEndDate, periodName;
  let weeksInPeriod = 1;

  switch(periodType) {
    case 'week':
      const dayOfWeek = today.getDay();
      endDate = new Date(today.getTime() - (dayOfWeek + 1) * 24 * 60 * 60 * 1000);
      startDate = new Date(endDate.getTime() - 6 * 24 * 60 * 60 * 1000);
      prevEndDate = new Date(startDate.getTime() - 1);
      prevStartDate = new Date(prevEndDate.getTime() - 6 * 24 * 60 * 60 * 1000);
      periodName = '주간';
      weeksInPeriod = 1;
      break;
    case 'month':
      endDate = new Date(today.getFullYear(), today.getMonth(), 0);
      startDate = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
      prevEndDate = new Date(startDate.getTime() - 1);
      prevStartDate = new Date(prevEndDate.getFullYear(), prevEndDate.getMonth(), 1);
      periodName = `${startDate.getFullYear()}년 ${startDate.getMonth() + 1}월`;
      weeksInPeriod = 4.345;
      break;
    case 'quarter':
      const currentQuarter = Math.floor(today.getMonth() / 3);
      endDate = new Date(today.getFullYear(), currentQuarter * 3, 0);
      startDate = new Date(endDate.getFullYear(), endDate.getMonth() - 2, 1);
      prevEndDate = new Date(startDate.getTime() - 1);
      prevStartDate = new Date(prevEndDate.getFullYear(), prevEndDate.getMonth() - 2, 1);
      periodName = `${startDate.getFullYear()}년 ${Math.floor(startDate.getMonth() / 3) + 1}분기`;
      weeksInPeriod = 13;
      break;
    case 'year':
      const lastYear = today.getFullYear() - 1;
      endDate = new Date(lastYear, 11, 31);
      startDate = new Date(lastYear, 0, 1);
      prevEndDate = new Date(startDate.getTime() - 1);
      prevStartDate = new Date(prevEndDate.getFullYear(), 0, 1);
      periodName = `${lastYear}년 연간`;
      weeksInPeriod = 52;
      break;
  }
  
  const startDateStr = formatDate(startDate);
  const endDateStr = formatDate(endDate);
  const prevStartDateStr = formatDate(prevStartDate);
  const prevEndDateStr = formatDate(prevEndDate);

  const logData = logSheet.getDataRange().getValues().filter(row => row[0]);
  const inbodyData = inbodySheet.getDataRange().getValues().filter(row => row[0]);

  const header = logData[0];
  const [dateIdx, exerciseIdx, setTypeIdx, weightIdx, repsIdx, unitIdx, volumeIdx, categoryIdx] = 
    ['날짜', '운동명', '세트_구분', '무게(kg)', '횟수/시간', '단위', '볼륨(kg)', '대분류'].map(h => header.indexOf(h));

  const allTimeData = logData.slice(1).filter(r => r[setTypeIdx] !== '웜업' && r[unitIdx] === '회');
  
  const extractStatsForPeriod = (start, end) => {
    const periodData = allTimeData.filter(r => {
      const rowDateStr = formatDate(new Date(r[dateIdx]));
      return rowDateStr >= start && rowDateStr <= end;
    });

    if (periodData.length === 0) return { totalWorkoutDays: 0, totalVolume: 0, mainFocusBodyPart: '없음', topExercises: [], bestPerformance: { weight: 0 } };

    const workoutDays = [...new Set(periodData.map(r => r[dateIdx].toString()))];
    const totalVolume = periodData.reduce((sum, r) => sum + (r[volumeIdx] || 0), 0);
    const categoryVol = {}, exerciseVol = {};
    periodData.forEach(r => {
      const category = r[categoryIdx] || '미분류';
      const exercise = r[exerciseIdx];
      const volume = r[volumeIdx] || 0;
      categoryVol[category] = (categoryVol[category] || 0) + volume;
      exerciseVol[exercise] = (exerciseVol[exercise] || 0) + volume;
    });
    const mainFocusBodyPart = Object.keys(categoryVol).length ? Object.keys(categoryVol).reduce((a, b) => categoryVol[a] > categoryVol[b] ? a : b) : '없음';
    const topExercises = Object.entries(exerciseVol).sort(([, a], [, b]) => b - a).slice(0, 5).map(([name, volume]) => ({ exercise: name, volume: volume.toFixed(0) + 'kg' }));
    let periodBest = { weight: 0 };
    periodData.forEach(r => { if (r[weightIdx] > periodBest.weight) periodBest = { exercise: r[exerciseIdx], weight: r[weightIdx], reps: r[repsIdx] }; });
    
    return { totalWorkoutDays: workoutDays.length, totalVolume: totalVolume.toFixed(0), mainFocusBodyPart, topExercises, bestPerformance: periodBest };
  };

  const currentStats = extractStatsForPeriod(startDateStr, endDateStr);
  const previousStats = extractStatsForPeriod(prevStartDateStr, prevEndDateStr);

  const avgWorkoutDaysPerWeek = (currentStats.totalWorkoutDays > 0) ? Math.max(1, Math.round(currentStats.totalWorkoutDays / weeksInPeriod)) : 0;

  const previousAllData = allTimeData.filter(r => formatDate(new Date(r[dateIdx])) < startDateStr && r[exerciseIdx] === currentStats.bestPerformance.exercise);
  const previousBestWeight = previousAllData.reduce((max, r) => Math.max(max, r[weightIdx]), 0);
  let pr = { exercise: '없음', record: '' };
  if (currentStats.bestPerformance.weight > 0 && currentStats.bestPerformance.weight > previousBestWeight) {
    pr.exercise = currentStats.bestPerformance.exercise;
    pr.record = `${currentStats.bestPerformance.weight.toFixed(1)}kg x ${currentStats.bestPerformance.reps}회`;
  }

  const startInbody = inbodyData.slice(1).filter(r => formatDate(new Date(r[0])) < startDateStr).pop() || Array(6).fill('N/A');
  const endInbody = inbodyData.slice(1).filter(r => formatDate(new Date(r[0])) <= endDateStr).pop() || startInbody;
  
  const getChange = (latestVal, prevVal) => { if (!isFinite(latestVal) || !isFinite(prevVal)) return ''; const diff = parseFloat(latestVal) - parseFloat(prevVal); if (diff > 0) return ` (+${diff.toFixed(2)} ▲)`; if (diff < 0) return ` (${diff.toFixed(2)} ▼)`; return ' (변화 없음)'; };
  const formatPercent = val => (typeof val === 'number' ? (val * 100).toFixed(1) + '%' : (val || 'N/A'));

  return {
    userName: config.USER_NAME, periodName, startDate: startDateStr, endDate: endDateStr,
    current: currentStats,
    previous: previousStats,
    avgWorkoutDaysPerWeek,
    prExercise: pr.exercise, prRecord: pr.record,
    endWeight: `${endInbody[2]} kg${getChange(endInbody[2], startInbody[2])}`,
    endMuscleMass: `${endInbody[3]} kg${getChange(endInbody[3], startInbody[3])}`,
    endBodyFatPercent: `${formatPercent(endInbody[5])}${getChange(endInbody[5], startInbody[5])}`
  };
}

// --- 🤖 AI 프롬프트 생성 함수들 ---

function createHistoryAnalysisPrompt(stats) {
  return `**Persona:** 당신은 피트니스 데이터 기록 분석가 '아카이브'입니다. 당신의 임무는 과거 데이터를 객관적으로 요약하는 것입니다.
**Task:** 아래 ${stats.userName}님의 **이전 기간** 운동 데이터를 간결하게 요약해주세요. 어떤 해석이나 조언도 하지 말고, 오직 사실만을 나열하세요.
**Input Data (Previous Period):**
- 총 운동일수: ${stats.previous.totalWorkoutDays}일
- 총 볼륨: ${stats.previous.totalVolume} kg
- 주력 운동 부위: ${stats.previous.mainFocusBodyPart}
- 볼륨 상위 운동: ${JSON.stringify(stats.previous.topExercises)}
**Output:** 이전 기간의 운동 패턴은 다음과 같음: [운동일수, 총 볼륨, 주력 부위, 상위 운동을 바탕으로 한 문장의 객관적인 요약]`;
}

function createTacticalAnalysisPrompt(stats, historyContext) {
  return `**Persona:** 당신은 전문 피트니스 데이터 분석가 '옵티머스'입니다.
**Task:** '아카이브'가 요약한 과거 데이터와 아래 제공된 현재 데이터를 **비교 분석**하여, ${stats.userName}님의 성과에 대한 핵심 인사이트를 도출해주세요.
**Input Data 1: Historical Context (from 'Archive')**
${historyContext}
**Input Data 2: Current Period Data (${stats.periodName}: ${stats.startDate} ~ ${stats.endDate})**
- 총 운동일수: ${stats.current.totalWorkoutDays}일
- 총 볼륨: ${stats.current.totalVolume} kg
- 주력 운동 부위: ${stats.current.mainFocusBodyPart}
- 볼륨 상위 운동: ${JSON.stringify(stats.current.topExercises)}
- 신기록(PR) 달성: ${stats.prExercise} (${stats.prRecord})
- 인바디 변화 (이전 전체 기간 대비 현재): 체중: ${stats.endWeight}, 골격근량: ${stats.endMuscleMass}, 체지방률: ${stats.endBodyFatPercent}
**Instructions (Think step-by-step):**
1. **Compare & Contrast:** 현재와 과거 데이터를 비교하여 변화된 패턴(예: 볼륨 증가/감소, 운동일수 변화, 주력 부위 변경 등)을 찾아내세요.
2. **Synthesize:** 이 변화가 인바디 결과나 PR 달성과 어떤 연관이 있는지 종합적으로 분석하세요.
3. **Conclude:** 분석을 바탕으로 칭찬할 점, 고려할 점, 다음을 위한 구체적인 제안을 도출하세요.
**Output Format:**
### 옵티머스의 데이터 분석 노트
**1. 성장 및 변화 포인트 (Growth & Changes):**
* [예: "이전 기간 대비 총 볼륨이 2,500kg 증가했으며, 이는 주력 부위인 하체 운동의 빈도가 늘어난 덕분으로 보입니다."]
**2. 주목할 성과 (Key Achievements):**
* [PR, 인바디의 긍정적 변화 등을 과거와 비교하며 구체적으로 칭찬]
**3. 다음을 위한 전략 제안 (Strategic Suggestions):**
* [분석된 성장/정체 패턴을 기반으로 다음 기간의 목표를 구체적으로 제시. 예: "상체 볼륨이 2주 연속 정체 상태이니, 다음 주 벤치프레스 마지막 세트는 드롭세트로 진행하여 새로운 자극을 주는 것을 추천합니다."]`
}

function createRoutineGenerationPrompt(stats, tacticalAnalysis) {
  return `
    **Persona:** 당신은 선수의 과거 기록과 현재 상태를 모두 파악하고 있는 엘리트 스트렝스 코치 '스트라테고스'입니다. 당신의 임무는 다음 주를 위한 가장 효과적인 운동 루틴을 설계하는 것입니다.

    **Task:** 아래 제공된 ${stats.userName}님의 데이터 분석 결과를 바탕으로, 다음 주를 위한 **사용자의 평균 운동 빈도에 맞는 최적의 운동 루틴**을 추천해주세요. 루틴은 반드시 분석 결과에 명시된 '전략 제안'을 반영해야 합니다.

    **Input Data 1: Athlete's Current Profile**
    - 이름: ${stats.userName}
    - ✨ **평균 주당 운동일수:** ${stats.avgWorkoutDaysPerWeek}일
    - 주로 수행하는 운동(선호도): ${JSON.stringify(stats.current.topExercises.map(e => e.exercise))}
    - 최근 PR (현재 근력 수준): ${stats.prExercise} ${stats.prRecord}
    - 주력 운동 부위: ${stats.current.mainFocusBodyPart}

    **Input Data 2: Tactical Analysis (from 'Optimus')**
    ---
    ${tacticalAnalysis}
    ---

    **Instructions for Routine Generation:**
    1.  **Dynamic Split:** **'평균 주당 운동일수'(${stats.avgWorkoutDaysPerWeek}일)에 맞춰** 가장 이상적인 분할 루틴을 설계하세요. (예: 4일이면 4분할, 5일이면 5분할 등)
    2.  **Goal-Oriented:** '전략 제안'을 최우선 목표로 설정하세요. (예: 제안이 '상체 볼륨 증대'라면, 상체 운동의 비중이나 강도를 높이세요.)
    3.  **Personalized:** 선호 운동 목록을 참고하여 루틴을 구성하되, 분석 결과에서 '개선/고려할 점'으로 지적된 약점 부위를 보완할 수 있는 운동을 최소 1개 이상 포함시키세요.
    4.  **Progressive Overload:** 최근 PR 기록을 바탕으로 현실적인 무게와 횟수를 제안하세요. (예: "기존 PR 무게의 80%로 5회 5세트" 또는 "기존 무게에서 2.5kg 증량하여 도전")
    5.  **Clear Structure:** 각 Day별로 루틴을 명확하게 구분하고, 운동마다 '운동명: 무게 x 횟수, 0세트' 형식으로 제시하세요.

    **Output Format:**

    ### 스트라테고스의 추천 주간 루틴

    **목표:** [분석 결과의 '전략 제안'을 한 문장으로 요약]
    **추천 분할:** [AI가 설계한 분할법, 예: 5분할 (가슴-등-하체-어깨-팔)]

    **Day 1: [주요 부위]**
    *   ...
    **Day 2: [주요 부위]**
    *   ...
    (사용자의 평균 운동일수에 맞춰 Day 개수를 동적으로 생성)
  `;
}

function createFinalReportPrompt(stats, reportType, tacticalAnalysis, recommendedRoutine) {
  const persona = `You are a friendly and motivating personal trainer in Korea named '버니'. Your client is ${stats.userName}.`;
  
  const reportDetails = {
    week: { title: `💪 ${stats.userName}님의 주간 운동 리포트`, intro: `지난 한 주도 정말 수고 많으셨어요! 땀 흘린 만큼 어떤 변화가 있었는지 함께 살펴볼까요?` },
    month: { title: `🗓️ ${stats.userName}님, ${stats.periodName} 운동 리포트`, intro: `한 달간의 노력이 쌓여 멋진 결과를 만들었어요. ${stats.periodName}의 성과를 확인해 보세요!` },
    quarter: { title: `📈 ${stats.userName}님, ${stats.periodName} 종합 리포트`, intro: `지난 3개월의 꾸준함이 만든 놀라운 변화! 분기 리포트를 통해 장기적인 성장을 확인해 보세요.` },
    year: { title: `🎉 ${stats.userName}님, 경이로운 한 해를 돌아보며! ${stats.periodName} 연간 리포트`, intro: `1년 동안의 위대한 여정에 진심으로 박수를 보냅니다! ${stats.userName}님의 놀라운 변화를 함께 축하하고 싶어요.` }
  };
  const reportDetail = reportDetails[reportType] || reportDetails.week;

  return `**Persona:** ${persona}
**Task:** Create a comprehensive fitness report email in Korean for ${stats.userName}, formatted in HTML. You must integrate the "Tactical Analysis" and the "Recommended Routine".
**Input Data 1: Data Summary (${stats.periodName})**
- Period: ${stats.startDate} ~ ${stats.endDate}
- Total workout days: ${stats.current.totalWorkoutDays}
- Main focus: ${stats.current.mainFocusBodyPart}
- Total volume: ${stats.current.totalVolume} kg
- New PR: ${stats.prExercise} with ${stats.prRecord}
- InBody (Weight): ${stats.endWeight}
- InBody (Muscle): ${stats.endMuscleMass}
- InBody (Body Fat): ${stats.endBodyFatPercent}
**Input Data 2: Tactical Analysis (from 'Optimus')**
---
${tacticalAnalysis}
---
**Input Data 3: Recommended Routine (from 'Strategos')**
---
${recommendedRoutine}
---
**Instructions for Final HTML Report:**
1. **Title, Intro, Summary:** Use the provided details: Title: "${reportDetail.title}", Intro: "${reportDetail.intro}".
2. **"📊 버니의 성장 코멘트":** Rewrite the "Tactical Analysis" in your friendly, personal trainer tone.
3. **[NEW SECTION] "🎯 다음 주 추천 루틴":** Create a new section below the comment. Convert the "Recommended Routine" into a visually appealing HTML format. Emphasize the 'Goal' as this week's mission.
4. **Conclusion:** Write a strong, motivating closing statement.
5. **Styling:** Use basic HTML. Highlight positive changes (▲) in green (#4CAF50) and negative changes (▼) in red (#f44336). Make the routine section stand out.`;
}

function callGeminiAPI(prompt, responseType = 'html') {
  const config = getProjectConfig();
  if (config.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY' || !config.GEMINI_API_KEY) {
    throw new Error("Gemini API 키가 설정되지 않았습니다. 스크립트 상단의 GEMINI_API_KEY를 확인해주세요.");
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=${config.GEMINI_API_KEY}`;
  
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { 
      "temperature": 0.6, 
      "topK": 1, 
      "topP": 1, 
      "maxOutputTokens": 8192,
      "responseMimeType": responseType === 'html' ? "text/html" : "text/plain"
    }
  };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };
  
  let response;
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200 || (responseCode >= 400 && responseCode < 500)) {
      break;
    }
    
    Logger.log(`API 호출 실패 (시도 ${i + 1}/${maxRetries}), 응답 코드: ${responseCode}. 5초 후 재시도합니다.`);
    Utilities.sleep(5000);
  }

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error(`Gemini API 호출 실패: ${responseCode} - ${responseText}`);
  }
  
  const json = JSON.parse(responseText);

  if (!json.candidates || json.candidates[0].finishReason === 'MAX_TOKENS') {
    Logger.log(`API 응답이 토큰 제한으로 인해 잘렸을 수 있습니다. Finish Reason: ${json.candidates[0].finishReason}`);
  }

  if (!json.candidates || !json.candidates[0].content || !json.candidates[0].content.parts) {
    throw new Error(`Gemini API 응답 형식이 올바르지 않습니다: ${responseText}`);
  }
  
  let resultText = json.candidates[0].content.parts[0].text;
  
  if (responseType === 'html') {
    resultText = resultText.replace(/^```html\n/, '').replace(/\n```$/, '');
  }
  
  return resultText;
}


// =================================================================
// ================== ✨ 채팅 기능 구현부 시작 ✨ ===================
// =================================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('AI 피트니스 챗봇');
}

function processUserMessage(message) {
  try {
    const toolCalls = routeQueryToTools(message);
    const retrievedData = executeToolCalls(toolCalls);
    const finalAnswer = generateFinalResponse(message, retrievedData);
    return finalAnswer;
  } catch (e) {
    Logger.log(`챗봇 오류: ${e.stack}`);
    return `처리 중 오류가 발생했습니다: ${e.message}`;
  }
}

function routeQueryToTools(message) {
  const today = new Date().toISOString().split('T')[0];
  const prompt = `**Persona:** 당신은 사용자의 질문을 이해하고, 어떤 데이터가 필요한지 판단하는 똑똑한 '라우터' AI입니다.
**Task:** 사용자의 질문을 분석하여, 답변에 필요한 '도구(tool)'와 '파라미터(params)'를 결정하고 JSON 배열 형식으로 반환해주세요. 여러 도구가 필요할 수 있습니다.
**Available Tools:**
1. \`search_workout_logs\`: 운동 기록(운동명, 무게, 횟수, 볼륨 등)에 대한 질문에 사용.
   - \`params\`: \`{"exercise_names": ["운동명"], "date_range_start": "YYYY-MM-DD", "date_range_end": "YYYY-MM-DD", "metric": "highest_weight" | "total_volume" | null}\`
2. \`search_inbody_records\`: 인바디 기록(체중, 근육량, 체지방률)에 대한 질문에 사용.
   - \`params\`: \`{"date_range_start": "YYYY-MM-DD", "date_range_end": "YYYY-MM-DD", "metric": "latest" | "change" | null}\`
   - 'metric'이 'change'이면 시작과 끝 데이터를 모두 찾아야 함. 'latest'이면 가장 마지막 데이터만 찾음.

**Rules:**
- 날짜 관련 표현(지난주, 이번달 등)은 오늘(${today})을 기준으로 'YYYY-MM-DD' 형식으로 정확히 계산해야 합니다.
- 관련 도구가 없으면 빈 배열 \`[]\`을 반환하세요.

[예시]
- 질문: "지난주 벤치프레스 총 볼륨 알려줘"
  -> \`[{"tool": "search_workout_logs", "params": {"exercise_names": ["벤치프레스"], "date_range_start": "2025-10-19", "date_range_end": "2025-10-25", "metric": "total_volume"}}]\`
- 질문: "가장 최근 인바디 기록 뭐야?"
  -> \`[{"tool": "search_inbody_records", "params": {"date_range_start": null, "date_range_end": null, "metric": "latest"}}]\`
- 질문: "지난달에 운동 열심히 했는데, 근육량 변화는 어때?"
  -> \`[{"tool": "search_workout_logs", "params": {"exercise_names": null, "date_range_start": "2025-09-01", "date_range_end": "2025-09-30", "metric": "total_volume"}}, {"tool": "search_inbody_records", "params": {"date_range_start": "2025-09-01", "date_range_end": "2025-09-30", "metric": "change"}}]\`
- 질문: "안녕?"
  -> \`[]\`

[실제 분석 요청]
질문: "${message}"
JSON:`;
  
  const resultText = callGeminiAPI(prompt, 'text').replace(/```json\n|```/g, '').trim();
  Logger.log(`1단계 - 라우팅 결과 (JSON): ${resultText}`);
  
  try {
    const parsed = JSON.parse(resultText);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    Logger.log(`JSON 파싱 오류: ${e.message}`);
    return [];
  }
}

function executeToolCalls(toolCalls) {
  if (!toolCalls || toolCalls.length === 0) {
    return "검색할 특정 데이터가 없습니다. 일반적인 대화를 나눠주세요.";
  }
  
  const results = toolCalls.map(call => {
    let result = `[Tool: ${call.tool}에 대한 결과]\n`;
    switch (call.tool) {
      case 'search_workout_logs':
        result += findWorkoutData(call.params);
        break;
      case 'search_inbody_records':
        result += findInbodyData(call.params);
        break;
      default:
        result += "알 수 없는 도구입니다.";
    }
    return result;
  });
  
  const aggregatedResult = results.join('\n\n');
  Logger.log(`2단계 - 도구 실행 및 결과 취합:\n${aggregatedResult}`);
  return aggregatedResult;
}

function findWorkoutData(conditions) {
  const config = getProjectConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(config.STRUCTURED_LOG_SHEET);
  const allData = logSheet.getRange("A2:H" + logSheet.getLastRow()).getValues();

  const [dateIdx, exerciseIdx, , , weightIdx, repsIdx, , volumeIdx] = [0, 1, 2, 3, 4, 5, 6, 7];
  let filteredData = allData;

  if (conditions.exercise_names && conditions.exercise_names.length > 0) {
    filteredData = filteredData.filter(row => conditions.exercise_names.some(name => row[exerciseIdx].includes(name)));
  }
  if (conditions.date_range_start && conditions.date_range_end) {
    const start = new Date(conditions.date_range_start + "T00:00:00");
    const end = new Date(conditions.date_range_end + "T23:59:59");
    filteredData = filteredData.filter(row => {
      const rowDate = new Date(row[dateIdx]);
      return rowDate >= start && rowDate <= end;
    });
  }
  
  if (filteredData.length === 0) return "해당 조건의 운동 기록을 찾지 못했습니다.";

  if (conditions.metric) {
    if (conditions.metric === "highest_weight") {
      let bestSet = filteredData.reduce((best, current) => (current[weightIdx] > best[weightIdx]) ? current : best, filteredData[0]);
      return `최고 기록: ${bestSet[exerciseIdx]} ${bestSet[weightIdx]}kg x ${bestSet[repsIdx]}회 (${new Date(bestSet[dateIdx]).toLocaleDateString()})`;
    }
    if (conditions.metric === "total_volume") {
      let totalVolume = filteredData.reduce((sum, row) => sum + (row[volumeIdx] || 0), 0);
      return `총 볼륨: ${totalVolume.toFixed(0)} kg (${filteredData.length} 세트)`;
    }
  }

  const slicedData = filteredData.slice(-30);
  let summary = `검색된 기록 (${filteredData.length}개 중 최근 30개):\n`;
  summary += slicedData.map(row => `${new Date(row[dateIdx]).toLocaleDateString()}: ${row[exerciseIdx]} ${row[weightIdx]}kg x ${row[repsIdx]}회`).join('\n');
  return summary;
}

function findInbodyData(conditions) {
  const config = getProjectConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inbodySheet = ss.getSheetByName(config.INBODY_SHEET);
  const allData = inbodySheet.getRange("A2:F" + inbodySheet.getLastRow()).getValues();
  const [dateIdx, , weightIdx, muscleIdx, , fatPercentIdx] = [0, 1, 2, 3, 4, 5];

  let filteredData = allData;
  if (conditions.date_range_start && conditions.date_range_end) {
    const start = new Date(conditions.date_range_start + "T00:00:00");
    const end = new Date(conditions.date_range_end + "T23:59:59");
    filteredData = filteredData.filter(row => {
      const rowDate = new Date(row[dateIdx]);
      return rowDate >= start && rowDate <= end;
    });
  }

  if (filteredData.length === 0) return "해당 기간의 인바디 기록을 찾지 못했습니다.";

  const formatRecord = (row) => `${new Date(row[dateIdx]).toLocaleDateString()}: 체중 ${row[weightIdx]}kg, 골격근량 ${row[muscleIdx]}kg, 체지방률 ${(row[fatPercentIdx]*100).toFixed(1)}%`;

  if (conditions.metric === 'latest') {
    return `가장 최근 기록: ${formatRecord(filteredData[filteredData.length - 1])}`;
  }
  if (conditions.metric === 'change') {
    const startRecord = formatRecord(filteredData[0]);
    const endRecord = formatRecord(filteredData[filteredData.length - 1]);
    const muscleChange = filteredData[filteredData.length - 1][muscleIdx] - filteredData[0][muscleIdx];
    return `기간 내 변화:\n- 시작: ${startRecord}\n- 종료: ${endRecord}\n- 골격근량 변화: ${muscleChange.toFixed(2)}kg`;
  }

  return filteredData.map(row => formatRecord(row)).join('\n');
}

function generateFinalResponse(message, retrievedData) {
  const prompt = `**Persona:** 당신은 사용자의 운동 기록과 인바디 기록을 모두 알고 있는 친절하고 전문적인 AI 피트니스 비서 '버니'입니다. 항상 한국어로, 격려하는 말투로 답변해주세요.
**Task:** 사용자의 질문에 대해, 제공된 '검색된 데이터'를 반드시 종합적으로 참고하여 답변을 생성해주세요.
**User's Question:** "${message}"
**Retrieved Context (Data from Tools):**
---
${retrievedData}
---
**Instruction:**
- 제공된 데이터를 바탕으로 질문에 대해 상세하고 친절하게 답변해주세요.
- 여러 도구의 결과가 있다면, 두 결과를 자연스럽게 연결하여 하나의 이야기처럼 설명해주세요. (예: "지난달 운동 볼륨이 높았던 만큼, 인바디에서도 근육량이 증가한 결과가 나타났네요!")
- 기록에 없는 내용은 "기록을 찾아봤는데, 그 정보는 없었어요."라고 솔직하게 말해주세요.
**Answer (in Korean):**`;

  return callGeminiAPI(prompt, 'text');
}


function getUserInfo() {
  const config = getProjectConfig();
  return {
    name: config.USER_NAME
  };
}

// =================================================================
// =================== ✨ 테스트 전용 함수들 ✨ =====================
// =================================================================

function TEST_sendMonthlyReport() {
  sendReport('month');
}

function TEST_sendQuarterlyReport() {
  sendReport('quarter');
}

function TEST_sendYearlyReport() {
  sendReport('year');
}
