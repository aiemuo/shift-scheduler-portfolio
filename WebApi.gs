/**
 * Apps Script Web App API. All spreadsheet access is explicit and uses the
 * Script Properties SPREADSHEET_ID value initialized from the bound sheet.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .setTitle("シフト管理");
}

function setupWebAppConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("対象Spreadsheetを開いてから実行してください。");

  const required = [SHEET_EMP, SHEET_JOB, SHEET_CFG, SHEET_OUT, SHEET_FORM, SHEET_WARN, SHEET_CAT, SHEET_QUOTA];
  const missing = required.filter(function(name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error("必要なシートがありません: " + missing.join(", "));

  const settings = ss.getSheetByName(SHEET_CFG);
  const employeeHeader = normalizeCellValue_(ss.getSheetByName(SHEET_EMP).getRange("A1").getValue());
  const settingHeader = normalizeCellValue_(settings.getRange("A1").getValue());
  if (employeeHeader !== "名前" || settingHeader !== "開始日") {
    throw new Error("このSpreadsheetはシフト管理用の既存スキーマと一致しません。");
  }

  const properties = PropertiesService.getScriptProperties();
  const previousId = properties.getProperty("SPREADSHEET_ID");
  if (previousId && previousId !== ss.getId()) {
    throw new Error("別のSpreadsheet IDが設定済みです。安全のため自動で変更しません。");
  }
  properties.setProperty("SPREADSHEET_ID", ss.getId());
  SpreadsheetApp.getUi().alert("Webアプリ接続先を設定しました。既存シートの内容は変更していません。");
}

function apiGetAppData() {
  return executeWebApi_(false, function(ss) {
    return buildWebAppState_(ss);
  });
}

function apiGenerateShift() {
  return executeWebApi_(true, function(ss) {
    return generateShiftCore_(ss);
  });
}

function apiPrepareShiftTemplate() {
  return executeWebApi_(true, function(ss) {
    return prepareShiftTemplateCore_(ss, true);
  });
}

function apiPrepareQuotaTemplate() {
  return executeWebApi_(true, function(ss) {
    return prepareQuotaTemplateCore_(ss, true);
  });
}

function apiSaveShiftCell(payload) {
  return executeWebApi_(true, function(ss) {
    const cell = resolveWebShiftCell_(ss, payload);
    if (cell.fixed) throw webUserError_("固定中のセルです。先に固定を解除してください。");

    const value = requireString_(payload.value, "業務");
    const jobs = readSingleColumn_(ss, SHEET_JOB, 1, 2);
    if (value !== "" && value !== "休" && jobs.indexOf(value) === -1) {
      throw webUserError_("選択した業務が業務シートにありません。画面を更新してください。");
    }

    cell.range.setValue(safeSheetText_(value));
    return { employeeName: cell.employeeName, date: cell.date, value: value, fixed: false };
  });
}

function apiFixShiftCell(payload) {
  return executeWebApi_(true, function(ss) {
    const cell = resolveWebShiftCell_(ss, payload);
    assertExpectedFixedStateForWeb_(cell, payload && payload.expectedFixed);
    if (!cell.value) throw webUserError_("空欄は固定できません。先に業務または「休」を設定してください。");
    cell.range.setNote(FIXED_NOTE);
    cell.range.setBackground(FIXED_BG);
    return { employeeName: cell.employeeName, date: cell.date, value: cell.value, fixed: true };
  });
}

function apiUnfixShiftCell(payload) {
  return executeWebApi_(true, function(ss) {
    const cell = resolveWebShiftCell_(ss, payload);
    assertExpectedFixedStateForWeb_(cell, payload && payload.expectedFixed);
    if (cell.fixed) {
      cell.range.setNote("");
      cell.range.setBackground(AUTO_BG);
    }
    return { employeeName: cell.employeeName, date: cell.date, value: cell.value, fixed: false };
  });
}

function apiUnfixAllShiftCells(payload) {
  return executeWebApi_(true, function(ss) {
    const expectedCells = requireFixedCellListForWeb_(payload && payload.cells, "固定解除対象");
    const expectedKeys = new Set(expectedCells.map(function(item) {
      return fixedCellKeyForWeb_(item.employeeName, item.date);
    }));
    const currentFixed = listFixedShiftCellsForWeb_(ss);
    const currentKeys = new Set(currentFixed.map(function(item) {
      return fixedCellKeyForWeb_(item.employeeName, item.date);
    }));

    if (expectedKeys.size !== currentKeys.size) {
      throw webUserError_("固定状態が他の操作で変更されています。画面を再読み込みしてからやり直してください。");
    }
    expectedKeys.forEach(function(key) {
      if (!currentKeys.has(key)) {
        throw webUserError_("固定状態が他の操作で変更されています。画面を再読み込みしてからやり直してください。");
      }
    });

    if (!currentFixed.length) return { changed: [], count: 0 };

    const resolved = currentFixed.map(function(item) {
      const cell = resolveWebShiftCell_(ss, item);
      if (!cell.fixed) {
        throw webUserError_("固定状態が他の操作で変更されています。画面を再読み込みしてからやり直してください。");
      }
      return cell;
    });

    resolved.forEach(function(cell) {
      cell.range.setNote("");
      cell.range.setBackground(AUTO_BG);
    });

    return {
      changed: resolved.map(function(cell) {
        return { employeeName: cell.employeeName, date: cell.date, value: cell.value, fixed: false };
      }),
      count: resolved.length,
    };
  });
}

function apiApplyFixedStateTransaction(payload) {
  return executeWebApi_(true, function(ss) {
    const changes = requireFixedStateChangesForWeb_(payload && payload.changes);
    if (!changes.length) return { changed: [], count: 0 };

    const resolved = changes.map(function(change) {
      const cell = resolveWebShiftCell_(ss, change);
      if (cell.fixed !== change.expectedFixed) {
        throw webUserError_("固定状態が他の操作で変更されています。画面を再読み込みしてからやり直してください。");
      }
      if (change.targetFixed && !cell.value) {
        throw webUserError_("空欄のセルは固定できません。シフト内容を確認してください。");
      }
      return { cell: cell, targetFixed: change.targetFixed };
    });

    resolved.forEach(function(item) {
      if (item.targetFixed) {
        item.cell.range.setNote(FIXED_NOTE);
        item.cell.range.setBackground(FIXED_BG);
      } else {
        item.cell.range.setNote("");
        item.cell.range.setBackground(AUTO_BG);
      }
    });

    return {
      changed: resolved.map(function(item) {
        return {
          employeeName: item.cell.employeeName,
          date: item.cell.date,
          value: item.cell.value,
          fixed: item.targetFixed,
        };
      }),
      count: resolved.length,
    };
  });
}

function apiSaveShiftSettings(payload) {
  return executeWebApi_(true, function(ss) {
    if (!payload || typeof payload !== "object") throw webUserError_("設定内容を読み取れません。");
    const start = parseIsoDateForWeb_(payload.startDate, "開始日");
    const end = parseIsoDateForWeb_(payload.endDate, "終了日");
    if (dateKey_(start) > dateKey_(end)) throw webUserError_("開始日は終了日以前の日付にしてください。");

    if (payload.intervalDays === "" || payload.intervalDays == null) throw webUserError_("業務間隔を入力してください。");
    const interval = Number(payload.intervalDays);
    if (!Number.isInteger(interval) || interval < 0 || interval > 366) {
      throw webUserError_("業務間隔は0〜366日の整数で指定してください。");
    }

    if (!Array.isArray(payload.holidays)) throw webUserError_("会社休業日の形式が正しくありません。");
    const holidayInputs = payload.holidays;
    if (holidayInputs.length > 366) throw webUserError_("会社休業日は366件以内で指定してください。");
    const holidayByIso = {};
    holidayInputs.forEach(function(value) {
      const date = parseIsoDateForWeb_(value, "会社休業日");
      holidayByIso[formatDateForWeb_(date)] = date;
    });
    const holidays = Object.keys(holidayByIso).sort().map(function(key) { return holidayByIso[key]; });

    const sheet = ss.getSheetByName(SHEET_CFG);
    sheet.getRange("B1:B3").setValues([[start], [end], [interval]]);
    if (holidays.length) sheet.getRange(5, 2, holidays.length, 1).setValues(holidays.map(function(date) { return [date]; }));
    const previousLastRow = sheet.getLastRow();
    const oldHolidayEndRow = Math.max(previousLastRow, 4 + holidays.length);
    const staleHolidayStartRow = 5 + holidays.length;
    if (oldHolidayEndRow >= staleHolidayStartRow) {
      sheet.getRange(staleHolidayStartRow, 2, oldHolidayEndRow - staleHolidayStartRow + 1, 1).clearContent();
    }
    return { startDate: formatDateForWeb_(start), endDate: formatDateForWeb_(end), intervalDays: interval, holidayCount: holidays.length };
  });
}

function apiSaveEmployees(payload) {
  return executeWebApi_(true, function(ss) {
    const items = requireArray_(payload, "従業員");
    if (items.length > 200) throw webUserError_("従業員は200名以内で登録してください。");
    if (!items.length) throw webUserError_("従業員を1名以上登録してください。");
    const seen = new Set();
    const rows = items.map(function(item) {
      const name = requireText_(item && item.name, "従業員名", 100);
      const group = requireString_(item && item.group, "区分");
      if (group !== "午前" && group !== "午後") throw webUserError_("区分は「午前」または「午後」を指定してください。");
      const key = name.toLocaleLowerCase();
      if (seen.has(key)) throw webUserError_("従業員名が重複しています: " + name);
      seen.add(key);
      return [safeSheetText_(name), group, item.must === true];
    });
    replaceSchemaRows_(ss.getSheetByName(SHEET_EMP), 2, 1, 3, rows);
    return { employeeCount: rows.length };
  });
}

function apiSaveJobs(payload) {
  return executeWebApi_(true, function(ss) {
    const items = requireArray_(payload, "業務");
    if (items.length > 200) throw webUserError_("業務は200件以内で登録してください。");
    if (!items.length) throw webUserError_("業務を1件以上登録してください。");
    const seen = new Set();
    const rows = items.map(function(value) {
      const name = requireText_(value, "業務名", 100);
      const key = name.toLocaleLowerCase();
      if (seen.has(key)) throw webUserError_("業務名が重複しています: " + name);
      seen.add(key);
      return [safeSheetText_(name)];
    });
    replaceSchemaRows_(ss.getSheetByName(SHEET_JOB), 2, 1, 1, rows);
    return { jobCount: rows.length };
  });
}

function apiSaveCategories(payload) {
  return executeWebApi_(true, function(ss) {
    const items = requireArray_(payload, "業務分類");
    if (items.length > 200) throw webUserError_("業務分類は200件以内で登録してください。");
    const seen = new Set();
    const rows = items.map(function(item) {
      const job = requireText_(item && item.job, "業務", 100);
      const category = requireString_(item && item.category, "分類");
      if (category && category !== COUNT_CAT_PRIMARY && category !== COUNT_CAT_SECONDARY) {
        throw webUserError_("分類は「区分1」「区分2」または未設定にしてください。");
      }
      const key = job.toLocaleLowerCase();
      if (seen.has(key)) throw webUserError_("業務分類に同じ業務が重複しています: " + job);
      seen.add(key);
      return [safeSheetText_(job), category];
    });
    replaceSchemaRows_(ss.getSheetByName(SHEET_CAT), 2, 1, 2, rows);
    return { categoryCount: rows.length };
  });
}

function apiSaveQuotas(payload) {
  return executeWebApi_(true, function(ss) {
    const items = requireArray_(payload, "回数設定");
    if (items.length > 100) throw webUserError_("回数設定は100件以内で登録してください。");
    const seen = new Set();
    const rows = items.map(function(item) {
      if (!item || item.workDays === "" || item.primary === "" || item.secondary === "" || item.workDays == null || item.primary == null || item.secondary == null) {
        throw webUserError_("回数設定の出勤日数・区分1・区分2をすべて入力してください。");
      }
      const days = Number(item && item.workDays);
      const primary = Number(item && item.primary);
      const secondary = Number(item && item.secondary);
      if (!Number.isInteger(days) || days < 0 || days > 366) throw webUserError_("出勤日数は0〜366の整数で指定してください。");
      if (!Number.isInteger(primary) || primary < 0 || primary > 366 || !Number.isInteger(secondary) || secondary < 0 || secondary > 366) {
        throw webUserError_("区分1・区分2の回数は0〜366の整数で指定してください。");
      }
      if (seen.has(days)) throw webUserError_("同じ出勤日数の回数設定が重複しています: " + days);
      seen.add(days);
      return [days, primary, secondary];
    });
    replaceSchemaRows_(ss.getSheetByName(SHEET_QUOTA), 2, 1, 3, rows);
    return { quotaCount: rows.length };
  });
}

function withShiftLock_(operation) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw webUserError_("別の処理が実行中です。少し待ってから再度お試しください。");
  }
  try {
    return operation();
  } finally {
    lock.releaseLock();
  }
}

function executeWebApi_(mutating, operation) {
  try {
    return withShiftLock_(function() {
      const ss = getWebAppSpreadsheet_();
      const data = operation(ss);
      if (mutating) PropertiesService.getScriptProperties().setProperty("LAST_WEB_ACTION", new Date().toISOString());
      return { ok: true, data: data };
    });
  } catch (error) {
    Logger.log("Web API error: " + String(error && error.stack ? error.stack : error));
    return {
      ok: false,
      message: error && error.webMessage ? error.webMessage : "処理に失敗しました。入力内容とSpreadsheetの状態を確認してください。",
    };
  }
}

function getWebAppSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw webUserError_("Webアプリの接続設定がありません。Spreadsheetの「シフト」メニューから「Webアプリ接続設定」を実行してください。");
  try {
    return SpreadsheetApp.openById(id);
  } catch (error) {
    throw webUserError_("対象Spreadsheetにアクセスできません。ログイン中のGoogleアカウントに編集権限があるか確認してください。");
  }
}

function buildWebAppState_(ss) {
  const cfg = readConfig_(ss);
  const employees = readEmployeesForWeb_(ss);
  const jobs = readSingleColumn_(ss, SHEET_JOB, 1, 2);
  const categories = readCategoriesForWeb_(ss);
  const categoryJobs = new Set();
  categories.forEach(function(row) { categoryJobs.add(row.job); });
  jobs.forEach(function(job) {
    if (!categoryJobs.has(job)) categories.push({ job: job, category: "" });
  });
  const warnings = readWarningsForWeb_(ss);
  const properties = PropertiesService.getScriptProperties();
  const lastAction = properties.getProperty("LAST_WEB_ACTION") || "";
  let lastActionLabel = "Webアプリの更新記録はありません";
  if (lastAction) {
    const lastActionDate = new Date(lastAction);
    if (!isNaN(lastActionDate.getTime())) lastActionLabel = Utilities.formatDate(lastActionDate, TZ, "yyyy/MM/dd HH:mm");
  }

  return {
    settings: {
      startDate: formatDateForWeb_(cfg.startDate),
      endDate: formatDateForWeb_(cfg.endDate),
      intervalDays: cfg.intervalDays,
      holidays: readHolidaysForWeb_(ss, cfg),
    },
    employees: employees,
    jobs: jobs,
    categories: categories,
    quotas: readQuotasForWeb_(ss),
    warnings: warnings,
    formRequests: readFormRequestsForWeb_(ss),
    lastAction: lastActionLabel,
    shift: readShiftSnapshotForWeb_(ss, cfg, employees),
  };
}

function readEmployeesForWeb_(ss) {
  const sh = ss.getSheetByName(SHEET_EMP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(function(row) {
      return { name: String(row[0] || "").trim(), group: String(row[1] || "").trim(), must: parseBool_(row[2]) };
    })
    .filter(function(employee) { return employee.name !== ""; });
}

function readCategoriesForWeb_(ss) {
  const sh = ss.getSheetByName(SHEET_CAT);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .map(function(row) { return { job: String(row[0] || "").trim(), category: String(row[1] || "").trim() }; })
    .filter(function(row) { return row.job !== ""; });
}

function readQuotasForWeb_(ss) {
  const sh = ss.getSheetByName(SHEET_QUOTA);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues()
    .map(function(row) { return { workDays: String(row[0] || "").trim(), primary: String(row[1] || "").trim(), secondary: String(row[2] || "").trim() }; })
    .filter(function(row) { return row.workDays !== "" || row.primary !== "" || row.secondary !== ""; });
}

function readHolidaysForWeb_(ss, cfg) {
  const keys = Array.from(readCompanyHolidays_(ss, cfg, cfg.startDate));
  return keys.sort();
}

function readWarningsForWeb_(ss) {
  const sh = ss.getSheetByName(SHEET_WARN);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .filter(function(row) { return String(row[0] || row[1] || "").trim() !== ""; })
    .map(function(row) { return { date: String(row[0] || ""), message: String(row[1] || "") }; });
}

function readFormRequestsForWeb_(ss) {
  const sh = ss.getSheetByName(SHEET_FORM);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  const lastRow = sh.getLastRow();
  const lastColumn = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(value) { return String(value).trim(); });
  const timestampIndex = header.indexOf("タイムスタンプ");
  const nameIndex = header.indexOf("氏名");
  const vacationIndex = header.indexOf("休暇希望日");
  if (nameIndex < 0 || vacationIndex < 0) return [];

  const rowCount = Math.min(100, lastRow - 1);
  const startRow = lastRow - rowCount + 1;
  const readColumn = function(columnIndex) {
    return sh.getRange(startRow, columnIndex + 1, rowCount, 1).getDisplayValues()
      .map(function(row) { return String(row[0] || ""); });
  };
  const timestamps = timestampIndex < 0 ? [] : readColumn(timestampIndex);
  const names = readColumn(nameIndex);
  const vacationDates = readColumn(vacationIndex);
  const requests = [];
  for (let index = rowCount - 1; index >= 0; index--) {
    const row = {
      timestamp: timestampIndex < 0 ? "" : timestamps[index],
      name: names[index],
      vacationDates: vacationDates[index],
    };
    if (row.name || row.vacationDates) requests.push(row);
  }
  return requests;
}

function readShiftSnapshotForWeb_(ss, cfg, employees) {
  const dates = enumerateDates_(cfg.startDate, cfg.endDate).map(function(date) {
    return { iso: formatDateForWeb_(date), label: Utilities.formatDate(date, TZ, "M/d") };
  });
  const rows = employees.map(function(employee) {
    return { name: employee.name, group: employee.group, cells: dates.map(function() { return { value: "", fixed: false }; }) };
  });
  const sh = ss.getSheetByName(SHEET_OUT);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 2) return { dates: dates, rows: rows };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowNames = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) { return normalizeCellValue_(row[0]); });
  const dateColumn = mapShiftDateColumnsForWeb_(headers.slice(1), cfg, 2);
  const employeeRow = new Map();
  rowNames.forEach(function(name, index) { employeeRow.set(name, index + 2); });
  const valueMatrix = sh.getRange(2, 2, lastRow - 1, lastCol - 1).getValues();
  const noteMatrix = sh.getRange(2, 2, lastRow - 1, lastCol - 1).getNotes();
  const backgroundMatrix = sh.getRange(2, 2, lastRow - 1, lastCol - 1).getBackgrounds();

  rows.forEach(function(row, employeeIndex) {
    const actualRow = employeeRow.get(row.name);
    if (!actualRow) return;
    const dataRowIndex = actualRow - 2;
    row.cells.forEach(function(cell, dateIndex) {
      const actualCol = dateColumn.get(dates[dateIndex].iso);
      if (!actualCol) return;
      const dataColIndex = actualCol - 2;
      cell.value = normalizeCellValue_(valueMatrix[dataRowIndex][dataColIndex]);
      cell.fixed = isFixedMark_(noteMatrix[dataRowIndex][dataColIndex], backgroundMatrix[dataRowIndex][dataColIndex]);
    });
  });
  return { dates: dates, rows: rows };
}

function resolveWebShiftCell_(ss, payload) {
  if (!payload || typeof payload !== "object") throw webUserError_("対象セルを読み取れません。");
  const employeeName = requireText_(payload.employeeName, "従業員名", 100);
  const date = parseIsoDateForWeb_(payload.date, "日付");
  const cfg = readConfig_(ss);
  const dateKey = dateKey_(date);
  const startKey = dateKey_(cfg.startDate);
  const endKey = dateKey_(cfg.endDate);
  if (dateKey < startKey || dateKey > endKey) throw webUserError_("選択した日付は現在のシフト期間外です。");

  const sh = ss.getSheetByName(SHEET_OUT);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 2) throw webUserError_("シフト表がありません。先にひな形を作成してください。");
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const names = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) { return normalizeCellValue_(row[0]); });
  const rowIndex = names.indexOf(employeeName);
  if (rowIndex < 0) throw webUserError_("従業員がシフト表にありません。ひな形を更新してください。");

  const headers = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  const dateColumns = mapShiftDateColumnsForWeb_(headers, cfg, 2);
  const actualCol = dateColumns.get(formatDateForWeb_(date));
  if (!actualCol) throw webUserError_("日付がシフト表にありません。ひな形を更新してください。");

  const range = sh.getRange(rowIndex + 2, actualCol);
  const note = range.getNote();
  const background = range.getBackground();
  return {
    range: range,
    employeeName: employeeName,
    date: formatDateForWeb_(date),
    value: normalizeCellValue_(range.getValue()),
    fixed: isFixedMark_(note, background),
  };
}

function assertExpectedFixedStateForWeb_(cell, expectedFixed) {
  if (typeof expectedFixed !== "boolean") return;
  if (cell.fixed !== expectedFixed) {
    throw webUserError_("固定状態が他の操作で変更されています。画面を再読み込みしてからやり直してください。");
  }
}

function fixedCellKeyForWeb_(employeeName, date) {
  return employeeName + "\u0000" + date;
}

function requireFixedCellListForWeb_(value, label) {
  if (!Array.isArray(value)) throw webUserError_(label + "の形式が正しくありません。");
  if (value.length > 5000) throw webUserError_(label + "が多すぎます。画面を再読み込みしてからやり直してください。");
  const seen = new Set();
  return value.map(function(item) {
    if (!item || typeof item !== "object") throw webUserError_(label + "の形式が正しくありません。");
    const employeeName = requireText_(item.employeeName, "従業員名", 100);
    const date = formatDateForWeb_(parseIsoDateForWeb_(item.date, "日付"));
    const key = fixedCellKeyForWeb_(employeeName, date);
    if (seen.has(key)) throw webUserError_(label + "に同じセルが重複しています。");
    seen.add(key);
    if (item.expectedFixed !== true) throw webUserError_(label + "の期待状態が正しくありません。");
    return { employeeName: employeeName, date: date, expectedFixed: true };
  });
}

function requireFixedStateChangesForWeb_(value) {
  if (!Array.isArray(value)) throw webUserError_("固定履歴の形式が正しくありません。");
  if (value.length > 5000) throw webUserError_("固定履歴の対象が多すぎます。画面を再読み込みしてからやり直してください。");
  const seen = new Set();
  return value.map(function(item) {
    if (!item || typeof item !== "object") throw webUserError_("固定履歴の形式が正しくありません。");
    const employeeName = requireText_(item.employeeName, "従業員名", 100);
    const date = formatDateForWeb_(parseIsoDateForWeb_(item.date, "日付"));
    if (typeof item.expectedFixed !== "boolean" || typeof item.targetFixed !== "boolean") {
      throw webUserError_("固定履歴の状態が正しくありません。");
    }
    const key = fixedCellKeyForWeb_(employeeName, date);
    if (seen.has(key)) throw webUserError_("固定履歴に同じセルが重複しています。");
    seen.add(key);
    return {
      employeeName: employeeName,
      date: date,
      expectedFixed: item.expectedFixed,
      targetFixed: item.targetFixed,
    };
  });
}

function listFixedShiftCellsForWeb_(ss) {
  const cfg = readConfig_(ss);
  const employees = readEmployeesForWeb_(ss);
  const snapshot = readShiftSnapshotForWeb_(ss, cfg, employees);
  const fixed = [];
  snapshot.rows.forEach(function(row) {
    row.cells.forEach(function(cell, dateIndex) {
      if (cell.fixed) fixed.push({ employeeName: row.name, date: snapshot.dates[dateIndex].iso });
    });
  });
  return fixed;
}

function mapShiftDateColumnsForWeb_(headers, cfg, firstColumn) {
  const expectedDates = enumerateDates_(cfg.startDate, cfg.endDate);
  const columns = new Map();
  let dateIndex = 0;
  headers.forEach(function(value, headerIndex) {
    const label = value instanceof Date ? formatMD_(value) : String(value || "").trim();
    while (dateIndex < expectedDates.length && formatMD_(expectedDates[dateIndex]) !== label) dateIndex++;
    if (dateIndex < expectedDates.length) {
      columns.set(formatDateForWeb_(expectedDates[dateIndex]), firstColumn + headerIndex);
      dateIndex++;
    }
  });
  return columns;
}

function replaceSchemaRows_(sheet, startRow, startCol, width, rows) {
  if (!sheet) throw webUserError_("保存先のシートがありません。");
  const lastRow = sheet.getLastRow();
  if (rows.length) sheet.getRange(startRow, startCol, rows.length, width).setValues(rows);
  const staleStartRow = startRow + rows.length;
  if (lastRow >= staleStartRow) sheet.getRange(staleStartRow, startCol, lastRow - staleStartRow + 1, width).clearContent();
}

function safeSheetText_(value) {
  const text = String(value == null ? "" : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function requireArray_(value, label) {
  if (!Array.isArray(value)) throw webUserError_(label + "の形式が正しくありません。");
  return value;
}

function requireString_(value, label) {
  if (typeof value !== "string") throw webUserError_(label + "の形式が正しくありません。");
  return value.trim();
}

function requireText_(value, label, maxLength) {
  const text = requireString_(value, label);
  if (!text || text.length > maxLength) throw webUserError_(label + "は1〜" + maxLength + "文字で入力してください。");
  return text;
}

function parseIsoDateForWeb_(value, label) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw webUserError_(label + "をYYYY-MM-DD形式で入力してください。");
  }
  const parts = value.split("-").map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    throw webUserError_(label + "の日付が不正です。");
  }
  return date;
}

function formatDateForWeb_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM-dd");
}

function webUserError_(message) {
  const error = new Error(message);
  error.webMessage = message;
  return error;
}
