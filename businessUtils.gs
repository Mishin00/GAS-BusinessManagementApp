function writeBusiness(business) {
  const sheet = getSheet(SHEET_BUSINESSES);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const headerMap = getHeaderIndexMap(headers)
  const id = data.length > 0 ? Math.max(...data.map(row => Number(row[headerMap.id]))) + 1 : 1;
  const today = new Date();

  sheet.appendRow([
    id,
    business.title,
    'Unstarted',
    business.requestDate,
    business.dueDate,
    '', // startDate
    '', // completionDate
    0, //progressRate
    '', //flag
    business.notes ?? '',
    today
  ]);

  return id;
}

function updateBusiness(updated) {
  const sheet = getSheet(SHEET_BUSINESSES);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const headerMap = getHeaderIndexMap(headers)

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][headerMap.id]) === String(updated.id)) {
      data[i][headerMap.title] = updated.title;
      data[i][headerMap.status] = updated.status;
      data[i][headerMap.requestDate] = updated.requestDate;
      data[i][headerMap.dueDate] = updated.dueDate;
      data[i][headerMap.startDate] = updated.startDate || data[i][headerMap.startDate];
      data[i][headerMap.completionDate] = updated.completionDate || data[i][headerMap.completionDate];
      data[i][headerMap.notes] = updated.notes;
      data[i][headerMap.lastUpdated] = new Date();
      
      sheet.getRange(i + 2, 1, 1, headers.length).setValues([data[i]]);
      break;
    }
  }
}

function getBusinesses() {

  const rows = getDataRangeValues(SHEET_BUSINESSES);
  const headers = rows.shift();
  const businesses =  rows.map(row => {
    const obj = Object.fromEntries(headers.map((h, i) => {
      let value = row[i];
      // Date オブジェクトなら文字列に変換
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }
      return [h, value];
    }));
    return obj;
});

  return businesses;
}

function getBusinessDetails(id) {
  const business = getBusinesses().find(b => String(b.id) === String(id));
  const tasks = getTasks().filter(t => String(t.businessId) === String(id));
  return { business, tasks };
}

function calculateProgress(tasks, headerMap) {
  let totalWeight = 0, completedWeight = 0;
  tasks.forEach(task => {
    const weight = Number(task[headerMap.weight]);
    const [done, total] = String(task[headerMap.progress]).split('/').map(Number);
    totalWeight += weight;
    
    completedWeight += (done / total) * weight;
  });

  return totalWeight ? Math.round((completedWeight / totalWeight) * 100) : 0;
}

function determineFlag(status, progressRate, startDate, dueDate) {
  const now = new Date();
  if (status === 'Unstarted') {
    const daysLeft = Math.floor((new Date(dueDate) - now) / (1000 * 60 * 60 * 24));
    return daysLeft <= 3 ? '✕' : '';
  }

  const start = new Date(startDate);
  const due = new Date(dueDate);
  const elapsed = Math.floor((now - start) / (1000 * 60 * 60 * 24));
  const total = Math.max(1, Math.floor((due - start) / (1000 * 60 * 60 * 24)));
  const expectedRate = Math.floor((elapsed / total) * 100);
  const diff = progressRate - expectedRate;

  if (diff < -20) return '✕';
  if (diff < -10) return '△';
  if (diff <= 10) return '〇';
  return '◎';
}

function updateProgressAndFlag(businessId) {
  const businessSheet = getSheet(SHEET_BUSINESSES);
  const taskSheet = getSheet(SHEET_TASKS);

  const businessData = businessSheet.getDataRange().getValues();
  const taskData = taskSheet.getDataRange().getValues();
  const bHeaders = businessData[0];
  const tHeaders = taskData[0];
  const bHeaderMap = getHeaderIndexMap(bHeaders)
  const tHeaderMap = getHeaderIndexMap(tHeaders)

  const businessIndex = businessData.findIndex(row => String(row[0]) === String(businessId));
  if (businessIndex === -1) return;

  const tasks = taskData.filter(row => String(row[1]) === String(businessId));

  const progressRate = calculateProgress(tasks,tHeaderMap)

  const now = new Date();
  const startDate = businessData[businessIndex][bHeaderMap.startDate] || now;
  const dueDate = businessData[businessIndex][bHeaderMap.dueDate];

  let flag = determineFlag(businessData[businessIndex][bHeaderMap.status], progressRate, startDate, dueDate);

  const row = businessData[businessIndex];
  const nowStatus = row[bHeaderMap.status]
  let newStatus = nowStatus;
  let allCompleted;
  if (tasks.length){
    allCompleted = tasks.every(row => row[tHeaderMap.status] === 'Completed');
    const isInProgress = tasks.some(row => row[tHeaderMap.status] === 'InProgress');
    if (nowStatus === 'InProgress' || nowStatus === 'OnHold'){
      newStatus = allCompleted ? 'Completed': nowStatus;
    }else if (nowStatus === 'Unstarted'){
      newStatus = isInProgress ? 'InProgress': nowStatus;
      row[bHeaderMap.startDate] = isInProgress ? now: '';
    }
  }

  row[bHeaderMap.progressRate] = progressRate;
  row[bHeaderMap.flag] = flag;
  row[bHeaderMap.status] = newStatus;
  if (row[bHeaderMap.status]==="Completed" && !row[bHeaderMap.completionDate]) row[bHeaderMap.completionDate] = now;
  
  row[bHeaderMap.lastUpdated] = now;

  businessSheet.getRange(businessIndex + 1, 1, 1, row.length).setValues([row]);
}

function updateProgressAndFlagAll(){
  const businesses = getBusinesses()
  businesses.map(b=>updateProgressAndFlag(b.id))
}
