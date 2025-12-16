/**
 * ⚠ SOURCE OF TRUTH:
 * Этот файл дублирует Code.gs из Google Apps Script.
 * Все изменения сначала тестируются в Codex,
 * затем вручную переносятся в GAS.
 */


// =========================================
// 1. КОНФИГУРАЦИЯ
// =========================================
const INCOME_SHEET = 'Мастер Классы';
const EXPENSE_SHEET = 'Расходные материалы';
const REF_SHEET = 'Справочник';
const BOOKING_SHEET = 'Календарь_Брони';

const BOOKING_HEADERS = [
  'ID',
  'Дата',
  'Время_Начала',
  'Время_Окончания',
  'Название',
  'Цена',
  'Количество',
  'Сумма',
  'Предоплата',
  'Статус',
  'Дата_Создания'
];

const BOOKING_COLS = {
  id: 1,
  date: 2,
  startTime: 3,
  endTime: 4,
  title: 5,
  price: 6,
  participants: 7,
  total: 8,
  prepayment: 9,
  status: 10,
  createdAt: 11
};

const BOOKING_STATUSES = {
  planned: 'planned',
  done: 'done',
  canceled: 'canceled'
};

// =========================================
// 2. ИНИЦИАЛИЗАЦИЯ И СВЯЗЬ
// =========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 Аналитика')
    .addItem('🟡 Венера ', 'openAppWindow')
    .addToUi();
}

function openAppWindow() {
  var template = HtmlService.createTemplateFromFile('Index');
  var html = template.evaluate()
      .setTitle('Мастерская Венера')
      .setWidth(1450).setHeight(900)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
  SpreadsheetApp.getUi().showModalDialog(html, 'Мастерская Венера');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('CRM Mobile')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if (action === 'getInitialConfig') result = getInitialConfig();
    else if (action === 'saveTransaction') result = saveTransaction(data);
    else if (action === 'getTableData') result = getTableData(data.type);
    else if (action === 'editTransaction') result = editTransaction(data);
    else if (action === 'deleteTransaction') result = deleteTransaction(data);
    else if (action === 'getAnalyticsData') result = getAnalyticsData(data.year, data.monthIdx);
    else if (action === 'getCalendarData') result = getCalendarData(data.year, data.month);
    else if (action === 'getBookingsByDate') result = getBookingsByDate(data.date);
    else if (action === 'addBooking') result = addBooking(data);
    else if (action === 'updateBooking') result = updateBooking(data.id, data);
    else if (action === 'changeBookingStatus') result = changeBookingStatus(data.id, data.status);
    else throw new Error("Unknown action: " + action);

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================
// 3. БАЗОВЫЕ МЕТОДЫ
// =========================================

function getInitialConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName(REF_SHEET);
  let incomeItems = [], expenseItems = [];
  
  if (refSheet) {
    const lastRow = refSheet.getLastRow();
    if (lastRow > 1) {
      const data = refSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      data.forEach(r => {
        if(r[0]) incomeItems.push(String(r[0]).trim());
        if(r[1]) expenseItems.push(String(r[1]).trim());
      });
    }
  }
  
  const years = new Set([new Date().getFullYear()]);
  [INCOME_SHEET, EXPENSE_SHEET].forEach(name => {
     const s = ss.getSheetByName(name);
     if(s && s.getLastRow() > 1) {
        const dates = s.getRange(2, 1, s.getLastRow()-1, 1).getValues();
        dates.forEach(r => { if(r[0] instanceof Date) years.add(r[0].getFullYear()); });
     }
  });
  return {
    incomeItems: incomeItems,
    expenseItems: expenseItems,
    years: Array.from(years).sort((a,b)=>b-a),
    months: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
  };
}

function saveTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.type === 'income' ? INCOME_SHEET : EXPENSE_SHEET;
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["Дата", "Название", "Кол-во", "Цена", "Итого"]); 
    }

    const date = new Date();
    let rowData = (data.type === 'income') 
        ? [date, data.name, data.qty, data.price, data.total] 
        : [date, data.name, data.price, data.qty, data.total];

    sheet.appendRow(rowData);
    if (data.isNew) {
      const refSheet = ss.getSheetByName(REF_SHEET);
      if (refSheet) {
        const col = data.type === 'income' ? 1 : 2;
        const lastRow = refSheet.getLastRow();
        let targetRow = lastRow + 1;
        // Простая проверка пустых ячеек
        const range = refSheet.getRange(1, col, lastRow + 5, 1).getValues();
        for(let i=1; i<range.length; i++) {
           if(!range[i][0]) { targetRow = i+1; break; }
        }
        refSheet.getRange(targetRow, col).setValue(data.name);
      }
    }
    return { success: true, message: "Добавлено" };
  } catch (e) {
    return { success: false, message: "Ошибка: " + e.toString() };
  }
}

function getTableData(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sName = (type === 'income') ? INCOME_SHEET : EXPENSE_SHEET;
  const sheet = ss.getSheetByName(sName);
  
  if(!sheet || sheet.getLastRow() < 2) return [];

  const lastRow = sheet.getLastRow();
  const limit = 50;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  const vals = sheet.getRange(startRow, 1, numRows, 5).getValues();
  let result = [];
  for(let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    if(!r[0]) continue;
    result.push({
      rowId: startRow + i,
      dateStr: Utilities.formatDate(r[0], ss.getSpreadsheetTimeZone(), "dd.MM.yyyy"),
      name: r[1],
      qty: (type === 'income') ? r[2] : r[3],
      price: (type === 'income') ? r[3] : r[2],
      total: r[4]
    });
  }
  return result;
}

function editTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sName = (data.type === 'income') ? INCOME_SHEET : EXPENSE_SHEET;
    const sheet = ss.getSheetByName(sName);
    const total = data.price * data.qty;
    let vals = (data.type === 'income') ? [[data.name, data.qty, data.price, total]] : [[data.name, data.price, data.qty, total]];
    sheet.getRange(data.rowId, 2, 1, 4).setValues(vals);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sName = (data.type === 'income') ? INCOME_SHEET : EXPENSE_SHEET;
    const sheet = ss.getSheetByName(sName);
    sheet.deleteRow(data.rowId);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =========================================
// 4. АНАЛИТИКА
// =========================================

function getAnalyticsData(year, monthIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = {
    current: { income: 0, expense: 0, profit: 0, visitors: 0, avgCheck: 0 },
    prev: { income: 0, expense: 0, profit: 0, visitors: 0 },
    growth: { income: 0, expense: 0, profit: 0, visitors: 0 },
    pulse: { labels: [], income: [], expense: [], profit: [] },
    pieIncome: { labels: [], data: [] },
    pieExpense: { labels: [], data: [] },
    attendance: { labels: [], data: [] },
    avgCheckChart: { labels: [], data: [] }
  };
  
  const y = parseInt(year);
  const m = monthIdx === 'all' ? null : parseInt(monthIdx);
  let dStart, dEnd, pStart, pEnd;
  
  if (m === null) { 
     dStart = new Date(y, 0, 1);
     dEnd = new Date(y, 11, 31, 23, 59, 59);
     pStart = new Date(y - 1, 0, 1);
     pEnd = new Date(y - 1, 11, 31, 23, 59, 59);
  } else { 
     dStart = new Date(y, m, 1);
     dEnd = new Date(y, m + 1, 0, 23, 59, 59);
     let pm = m - 1;
     let py = y; if (pm < 0) { pm = 11; py = y - 1; }
     pStart = new Date(py, pm, 1);
     pEnd = new Date(py, pm + 1, 0, 23, 59, 59);
  }

  let trendData = {};
  let incomeCats = {};
  let expenseCats = {};
  const getTrendKey = (date) => (m === null) ? date.getMonth() : date.getDate();
  const steps = (m === null) ? 12 : new Date(y, m+1, 0).getDate();
  
  for(let i = (m === null ? 0 : 1); i <= (m === null ? 11 : steps); i++) {
      trendData[i] = { inc: 0, exp: 0, vis: 0 };
  }

  function processSheet(sName, type) {
    const s = ss.getSheetByName(sName);
    if(!s || s.getLastRow() < 2) return;
    const data = s.getRange(2, 1, s.getLastRow()-1, 5).getValues();
    
    data.forEach(r => {
      const d = r[0];
      if(!(d instanceof Date)) return;
      const sum = parseFloat(r[4]) || 0;
      const qty = (type === 'income') ? (parseFloat(r[2]) || 0) : 0;
      const name = String(r[1]);

      if (d >= dStart && d <= dEnd) {
         if (type === 'income') {
            res.current.income += sum; res.current.visitors += qty;
            incomeCats[name] = (incomeCats[name] || 0) + sum;
         } else {
            res.current.expense += sum;
            expenseCats[name] = (expenseCats[name] || 0) + sum;
         }
         const k = getTrendKey(d);
         if (trendData[k]) {
            if (type === 'income') { trendData[k].inc += sum; trendData[k].vis += qty; }
            else trendData[k].exp += sum;
         }
      }
      if (d >= pStart && d <= pEnd) {
         if (type === 'income') { res.prev.income += sum; res.prev.visitors += qty; }
         else { res.prev.expense += sum; }
      }
    });
  }

  processSheet(INCOME_SHEET, 'income');
  processSheet(EXPENSE_SHEET, 'expense');
  
  res.current.profit = res.current.income - res.current.expense;
  res.prev.profit = res.prev.income - res.prev.expense;
  if (res.current.visitors > 0) res.current.avgCheck = Math.round(res.current.income / res.current.visitors);
  
  const calcGrowth = (curr, prev) => {
     if (prev === 0) return curr > 0 ? 100 : 0;
     return Math.round(((curr - prev) / prev) * 100);
  };
  
  res.growth.income = calcGrowth(res.current.income, res.prev.income);
  res.growth.expense = calcGrowth(res.current.expense, res.prev.expense);
  res.growth.profit = calcGrowth(res.current.profit, res.prev.profit);
  res.growth.visitors = calcGrowth(res.current.visitors, res.prev.visitors);

  const monthNames = ["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"];
  Object.keys(trendData).sort((a,b) => parseInt(a)-parseInt(b)).forEach(k => {
     const label = (m === null) ? monthNames[k] : k;
     res.pulse.labels.push(label);
     res.pulse.income.push(trendData[k].inc);
     res.pulse.expense.push(trendData[k].exp);
     res.pulse.profit.push(trendData[k].inc - trendData[k].exp);
     
     res.attendance.labels.push(label);
     res.attendance.data.push(trendData[k].vis);
     
     res.avgCheckChart.labels.push(label);
     const avg = trendData[k].vis > 0 ? Math.round(trendData[k].inc / trendData[k].vis) : 0;
     res.avgCheckChart.data.push(avg);
  });
  
  const sortedInc = Object.entries(incomeCats).sort((a,b) => b[1] - a[1]).slice(0, 5);
  res.pieIncome.labels = sortedInc.map(x => x[0]);
  res.pieIncome.data = sortedInc.map(x => x[1]);
  
  const sortedExp = Object.entries(expenseCats).sort((a,b) => b[1] - a[1]).slice(0, 5);
  res.pieExpense.labels = sortedExp.map(x => x[0]);
  res.pieExpense.data = sortedExp.map(x => x[1]);

  return res;
}

// =========================================
// 5. КАЛЕНДАРЬ БРОНИРОВАНИЙ (АРХИТЕКТУРА)
// =========================================

function ensureBookingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BOOKING_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(BOOKING_SHEET);
    sheet.appendRow(BOOKING_HEADERS);
    return sheet;
  }

  const headerRange = sheet.getRange(1, 1, 1, BOOKING_HEADERS.length);
  const headerValues = headerRange.getValues()[0];
  const hasHeaders = headerValues.some(value => Boolean(value));
  if (!hasHeaders) {
    headerRange.setValues([BOOKING_HEADERS]);
  }

  return sheet;
}

function isValidBookingStatus(status) {
  return Object.values(BOOKING_STATUSES).indexOf(status) !== -1;
}

function getCalendarData(year, month) {
  const sheet = ensureBookingSheet();
  const lastRow = sheet.getLastRow();
  const daysMap = {};

  if (lastRow < 2) {
    return { success: true, filters: { year: year, month: month }, days: [] };
  }

  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const targetYear = parseInt(year, 10);
  const targetMonth = parseInt(month, 10);
  const data = sheet.getRange(2, 1, lastRow - 1, BOOKING_HEADERS.length).getValues();

  data.forEach(row => {
    const dateCell = row[BOOKING_COLS.date - 1];
    const status = String(row[BOOKING_COLS.status - 1] || '').trim();

    if (!(dateCell instanceof Date)) return;

    // Используем часовой пояс таблицы, чтобы не смещать дату при агрегации
    const dateYear = parseInt(Utilities.formatDate(dateCell, timezone, 'yyyy'), 10);
    const dateMonth = parseInt(Utilities.formatDate(dateCell, timezone, 'M'), 10) - 1;
    if (dateYear !== targetYear || dateMonth !== targetMonth) return;

    const dateKey = Utilities.formatDate(dateCell, timezone, 'yyyy-MM-dd');
    if (!daysMap[dateKey]) {
      daysMap[dateKey] = {
        date: dateKey,
        totalBookings: 0,
        statusSummary: {
          planned: 0,
          done: 0,
          canceled: 0
        }
      };
    }

    daysMap[dateKey].totalBookings += 1;
    if (isValidBookingStatus(status)) {
      daysMap[dateKey].statusSummary[status] += 1;
    }
  });

  const days = Object.values(daysMap).sort((a, b) => a.date.localeCompare(b.date));

  return {
    success: true,
    filters: { year: targetYear, month: targetMonth },
    days: days
  };
}

function getBookingsByDate(date) {
  const sheet = ensureBookingSheet();
  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const [year, month, day] = String(date || '').split('-').map(Number);

  if (!year || !month || !day) {
    return { success: false, message: 'Invalid date format', date: date, bookings: [] };
  }

  const targetDate = new Date(year, month - 1, day); // Локальная дата без смещения часового пояса
  const targetKey = Utilities.formatDate(targetDate, timezone, 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { success: true, date: targetKey, bookings: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, BOOKING_HEADERS.length).getValues();
  const bookings = [];

  values.forEach(row => {
    const dateCell = row[BOOKING_COLS.date - 1];
    if (!(dateCell instanceof Date)) return;

    const rowDateKey = Utilities.formatDate(dateCell, timezone, 'yyyy-MM-dd');
    if (rowDateKey !== targetKey) return;

    const startTimeCell = row[BOOKING_COLS.startTime - 1];
    const endTimeCell = row[BOOKING_COLS.endTime - 1];

    bookings.push({
      id: row[BOOKING_COLS.id - 1],
      name: row[BOOKING_COLS.title - 1],
      startTime: (startTimeCell instanceof Date) ? Utilities.formatDate(startTimeCell, timezone, 'HH:mm') : '',
      endTime: (endTimeCell instanceof Date) ? Utilities.formatDate(endTimeCell, timezone, 'HH:mm') : '',
      qty: Number(row[BOOKING_COLS.participants - 1]) || 0,
      price: Number(row[BOOKING_COLS.price - 1]) || 0,
      total: Number(row[BOOKING_COLS.total - 1]) || 0,
      prepayment: Number(row[BOOKING_COLS.prepayment - 1]) || 0,
      status: String(row[BOOKING_COLS.status - 1] || '').trim() || BOOKING_STATUSES.planned
    });
  });

  return {
    success: true,
    date: targetKey,
    bookings: bookings
  };
}

function addBooking(data) {
  const sheet = ensureBookingSheet();
  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  const dateStr = data.date;
  const dateObj = dateStr ? new Date(dateStr) : null;
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return { success: false, message: 'Invalid booking date' };
  }

  const parseTime = (timeStr) => {
    if (!timeStr) return null;
    const parsed = new Date(`${dateStr}T${timeStr}:00`);
    return (parsed instanceof Date && !isNaN(parsed.getTime())) ? parsed : null;
  };

  const startTime = parseTime(data.startTime);
  const endTime = parseTime(data.endTime);
  const price = Number(data.price) || 0;
  const participants = Number(data.participants) || 0;
  const prepayment = Number(data.prepayment) || 0;
  const total = price * participants;

  const id = `BK-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
  const createdAt = new Date();

  const row = [];
  row[BOOKING_COLS.id - 1] = id;
  row[BOOKING_COLS.date - 1] = dateObj;
  row[BOOKING_COLS.startTime - 1] = startTime;
  row[BOOKING_COLS.endTime - 1] = endTime;
  row[BOOKING_COLS.title - 1] = data.name || '';
  row[BOOKING_COLS.price - 1] = price;
  row[BOOKING_COLS.participants - 1] = participants;
  row[BOOKING_COLS.total - 1] = total;
  row[BOOKING_COLS.prepayment - 1] = prepayment;
  row[BOOKING_COLS.status - 1] = BOOKING_STATUSES.planned;
  row[BOOKING_COLS.createdAt - 1] = createdAt;

  sheet.appendRow(row);

  return {
    success: true,
    booking: {
      id: id,
      name: data.name || '',
      startTime: startTime ? Utilities.formatDate(startTime, timezone, 'HH:mm') : '',
      endTime: endTime ? Utilities.formatDate(endTime, timezone, 'HH:mm') : '',
      qty: participants,
      price: price,
      total: total,
      prepayment: prepayment,
      status: BOOKING_STATUSES.planned
    }
  };
}

function updateBooking(id, data) {
  ensureBookingSheet();
  return {
    success: false,
    message: 'updateBooking is not implemented yet',
    id: id,
    payload: data
  };
}

function changeBookingStatus(id, status) {
  ensureBookingSheet();
  if (!isValidBookingStatus(status)) {
    return { success: false, message: 'Invalid booking status', id: id, status: status };
  }
  return {
    success: false,
    message: 'changeBookingStatus is not implemented yet',
    id: id,
    status: status
  };
}
