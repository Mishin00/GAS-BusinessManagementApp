function getTasks() {
  const rows = getDataRangeValues(SHEET_TASKS);
  const headers = rows.shift();
  return rows.map(row => Object.fromEntries(headers.map((h, i) => {
    let value = row[i];
    // Date オブジェクトなら文字列に変換
    if (value instanceof Date) {
      value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }
    return [h, value];

  }
  )));
}



function writeTask(task) {
  const sheet = getSheet(SHEET_TASKS);
  const data = sheet.getDataRange().getValues();
  const nextId = data.length > 1 ? Math.max(...data.slice(1).map(row => Number(row[0]))) + 1 : 1;
  const today = new Date();

  sheet.appendRow([
    nextId,
    task.businessId,
    task.title,
    'Unstarted',
    task.weight,
    '0/100',
    task.notes ?? '',
    today
  ]);

  updateProgressAndFlag(task.businessId);
  return nextId;
}

function updateTask(updated) {
  const sheet = getSheet(SHEET_TASKS);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const headerMap = getHeaderIndexMap(headers)

  let businessId;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][headerMap.id]) === String(updated.id)) {
      
      data[i][headerMap.title] = updated.title;
      data[i][headerMap.status]  = updated.status;
      data[i][headerMap.weight]  = updated.weight;
      data[i][headerMap.progress]  = updated.progress;
      data[i][headerMap.notes] = updated.notes;
      
      data[i][headerMap.lastUpdated] = new Date();
      sheet.getRange(i + 2, 1, 1, headers.length).setValues([data[i]]);
      businessId = data[i][headerMap.businessId]
      break;
    }
  }

  updateProgressAndFlag(businessId);
}
