const APP = Object.freeze({
  sheets: {
    Masters: ['Type', 'Value'],
    Lessons: ['ID', 'Date', 'Period', 'Subject', 'Grade', 'Class', 'Room', 'Topic', 'NumberMode', 'LessonNo', 'Memo', 'StartTime', 'EndTime', 'UpdatedAt'],
    Records: ['RecordID', 'LessonID', 'Date', 'Plan', 'Result', 'Homework', 'Reflection', 'UpdatedAt'],
    Settings: ['Key', 'Value'],
    Logs: ['Timestamp', 'Action', 'Detail'],
    BasicTimetable: ['Weekday', 'Period', 'Subject', 'Grade', 'Class', 'Room', 'Topic', 'NumberMode', 'StartTime', 'EndTime']
  },
  lessonModes: ['auto', 'manual', 'none'],
  masterTypes: ['Subject', 'Grade', 'Class', 'Room', 'Topic'],
  backupVersion: 1,
  timeZone: 'Asia/Tokyo'
});

function dispatchAction_(action, payload) {
  switch (action) {
    case 'setup': return setupSpreadsheet();
    case 'bootstrap': return getBootstrapData_(payload.date);
    case 'list': return listLessons_(payload);
    case 'week': return getWeek_(payload.date);
    case 'save': return saveLesson_(payload);
    case 'duplicate': return duplicateLesson_(payload);
    case 'delete': return deleteLesson_(payload.ID);
    case 'reindex': return withLock_(function () { return reindexLessons_(); });
    case 'move': return moveOrSwapLesson_(payload);
    case 'saveRecord': return saveRecord_(payload);
    case 'listRecords': return listRecords_(payload);
    case 'saveBasic': return saveBasic_(payload);
    case 'deleteBasic': return deleteBasic_(payload);
    case 'applyBasic': return applyBasicTimetable_(payload.date);
    case 'saveMaster': return saveMaster_(payload);
    case 'deleteMaster': return deleteMaster_(payload);
    case 'exportJson': return exportJson_();
    case 'exportCsv': return exportCsv_(payload);
    case 'restoreJson': return restoreJson_(payload);
    case 'importJson': return restoreJson_(payload.data || payload);
    default: throw new Error('未対応のactionです: ' + action);
  }
}

function setupSpreadsheet() {
  const ss = BrowserTables.getActiveSpreadsheet();
  Object.keys(APP.sheets).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = APP.sheets[name];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#dbeafe');
      sheet.autoResizeColumns(1, headers.length);
    } else {
      const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (current.join('|') !== headers.join('|')) {
        throw new Error(name + ' シートのヘッダーが仕様と異なります。既存データを確認してください。');
      }
    }
  });
  log_('SETUP', 'Spreadsheet schema initialized');
  return { sheets: Object.keys(APP.sheets) };
}

function getBootstrapData_(date) {
  ensureSchema_();
  const target = date || formatDate_(new Date());
  const week = getWeek_(target);
  return {
    date: target,
    lessons: listLessons_({}),
    week: week,
    masters: getMasters_(),
    settings: rowsToObjects_('Settings'),
    basicTimetable: rowsToObjects_('BasicTimetable')
  };
}

function listLessons_(filters) {
  ensureSchema_();
  filters = filters || {};
  return rowsToObjects_('Lessons')
    .filter(function (lesson) {
      return (!filters.date || lesson.Date === filters.date) &&
        (!filters.from || lesson.Date >= filters.from) &&
        (!filters.to || lesson.Date <= filters.to);
    })
    .sort(compareLessons_);
}

function getWeek_(date) {
  const base = parseDate_(date || formatDate_(new Date()));
  const day = base.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  const monday = addDays_(base, mondayOffset);
  const friday = addDays_(monday, 4);
  const from = formatDate_(monday);
  const to = formatDate_(friday);
  applyBasicTimetable_(from, true);
  return {
    from: from,
    to: to,
    days: [0, 1, 2, 3, 4].map(function (offset) {
      const value = formatDate_(addDays_(monday, offset));
      return { date: value, lessons: listLessons_({ date: value }) };
    })
  };
}

function saveBasic_(input) {
  ensureSchema_();
  const weekday = Number(input.Weekday);
  const period = Number(input.Period);
  if (!Number.isInteger(weekday) || weekday < 1 || weekday > 5) throw new Error('曜日は月曜から金曜で指定してください。');
  if (!Number.isInteger(period) || period < 1 || period > 12) throw new Error('時限は1〜12の整数で指定してください。');
  if (!String(input.Subject || '').trim()) throw new Error('教科は必須です。');
  const row = {
    Weekday: weekday,
    Period: period,
    Subject: String(input.Subject).trim(),
    Grade: String(input.Grade || '').trim(),
    Class: String(input.Class || '').trim(),
    Room: String(input.Room || '').trim(),
    Topic: String(input.Topic || '').trim(),
    NumberMode: ['auto', 'none'].includes(input.NumberMode) ? input.NumberMode : 'auto',
    StartTime: String(input.StartTime || ''),
    EndTime: String(input.EndTime || '')
  };
  return withLock_(function () {
    const sheet = getSheet_('BasicTimetable');
    const existing = rowsToObjectsWithRows_('BasicTimetable').find(function (item) {
      return Number(item.data.Weekday) === weekday && Number(item.data.Period) === period;
    });
    const values = APP.sheets.BasicTimetable.map(function (header) { return serializeCell_(header, row[header]); });
    if (existing) sheet.getRange(existing.row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    log_('SAVE_BASIC', weekday + '-' + period);
    return rowsToObjects_('BasicTimetable');
  });
}

function deleteBasic_(input) {
  ensureSchema_();
  return withLock_(function () {
    const sheet = getSheet_('BasicTimetable');
    const rows = rowsToObjectsWithRows_('BasicTimetable').filter(function (item) {
      return Number(item.data.Weekday) === Number(input.Weekday) && Number(item.data.Period) === Number(input.Period);
    }).map(function (item) { return item.row; }).sort(function (a, b) { return b - a; });
    rows.forEach(function (row) { sheet.deleteRow(row); });
    log_('DELETE_BASIC', input.Weekday + '-' + input.Period);
    return rowsToObjects_('BasicTimetable');
  });
}

function applyBasicTimetable_(date, silent) {
  ensureSchema_();
  return withLock_(function () {
    const templates = rowsToObjects_('BasicTimetable');
    if (!templates.length) return { created: 0 };
    const base = parseDate_(date || formatDate_(new Date()));
    const day = base.getDay();
    const monday = addDays_(base, day === 0 ? -6 : 1 - day);
    const from = formatDate_(monday);
    const to = formatDate_(addDays_(monday, 4));
    const lessons = listLessons_({ from: from, to: to });
    let created = 0;
    templates.forEach(function (template) {
      const lessonDate = formatDate_(addDays_(monday, Number(template.Weekday) - 1));
      const occupied = lessons.some(function (item) {
        return item.Date === lessonDate && Number(item.Period) === Number(template.Period);
      });
      if (occupied) return;
      const lesson = normalizeLesson_({
        Date: lessonDate, Period: template.Period, Subject: template.Subject,
        Grade: template.Grade || '-', Class: template.Class || '-', Room: template.Room,
        Topic: template.Topic, NumberMode: template.NumberMode || 'auto',
        StartTime: template.StartTime, EndTime: template.EndTime
      });
      appendObject_('Lessons', lesson);
      lessons.push(lesson);
      created += 1;
    });
    if (created) {
      reindexLessons_();
      if (!silent) log_('APPLY_BASIC', from + ': ' + created);
    }
    return { created: created, from: from, to: to };
  });
}

function saveLesson_(input) {
  ensureSchema_();
  const lesson = normalizeLesson_(input);
  return withLock_(function () {
    const lessons = rowsToObjects_('Lessons');
    const existing = lessons.find(function (item) { return item.ID === lesson.ID; });
    const oldGroup = existing ? lessonGroupKey_(existing) : '';
    upsertObject_('Lessons', 'ID', lesson);
    if (existing && existing.Date !== lesson.Date) updateRecordDates_(lesson.ID, lesson.Date);
    reindexLessons_([oldGroup, lessonGroupKey_(lesson)].filter(Boolean));
    log_(existing ? 'UPDATE_LESSON' : 'CREATE_LESSON', lesson.ID);
    return rowsToObjects_('Lessons').find(function (item) { return item.ID === lesson.ID; });
  });
}

function deleteLesson_(id) {
  if (!id) throw new Error('削除対象のIDが必要です。');
  ensureSchema_();
  return withLock_(function () {
    const lessons = rowsToObjects_('Lessons');
    const target = lessons.find(function (item) { return item.ID === id; });
    if (!target) throw new Error('授業が見つかりません。');
    deleteRowsByValue_('Lessons', 'ID', id);
    deleteRowsByValue_('Records', 'LessonID', id);
    reindexLessons_([lessonGroupKey_(target)]);
    log_('DELETE_LESSON', id);
    return { ID: id };
  });
}

function duplicateLesson_(input) {
  if (!input.ID) throw new Error('複製元のIDが必要です。');
  ensureSchema_();
  return withLock_(function () {
    const source = rowsToObjects_('Lessons').find(function (item) { return item.ID === input.ID; });
    if (!source) throw new Error('複製元の授業が見つかりません。');
    const copy = Object.assign({}, source, {
      ID: '',
      Date: input.Date === undefined ? source.Date : input.Date,
      Period: input.Period === undefined ? source.Period : input.Period,
      StartTime: input.StartTime === undefined ? source.StartTime : input.StartTime,
      EndTime: input.EndTime === undefined ? source.EndTime : input.EndTime
    });
    // 手動番号をそのまま複製すると採番列を意図せず飛ばすため、自動採番へ戻す。
    if (copy.NumberMode === 'manual' && input.keepManualNumber !== true) {
      copy.NumberMode = 'auto';
      copy.LessonNo = '';
    }
    const lesson = normalizeLesson_(copy);
    appendObject_('Lessons', lesson);
    reindexLessons_([lessonGroupKey_(lesson)]);
    log_('DUPLICATE_LESSON', input.ID + ' -> ' + lesson.ID);
    return rowsToObjects_('Lessons').find(function (item) { return item.ID === lesson.ID; });
  });
}

function moveOrSwapLesson_(payload) {
  if (!payload.ID || !payload.Date || !payload.Period) {
    throw new Error('ID、移動先Date、Periodが必要です。');
  }
  ensureSchema_();
  const targetDate = String(payload.Date).trim();
  const targetPeriod = Number(payload.Period);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('DateはYYYY-MM-DD形式で指定してください。');
  parseDate_(targetDate);
  if (!Number.isInteger(targetPeriod) || targetPeriod < 1 || targetPeriod > 12) throw new Error('Periodは1〜12の整数で指定してください。');
  return withLock_(function () {
    const lessons = rowsToObjects_('Lessons');
    const source = lessons.find(function (item) { return item.ID === payload.ID; });
    if (!source) throw new Error('移動元の授業が見つかりません。');
    const destination = lessons.find(function (item) {
      return item.ID !== source.ID && item.Date === targetDate && Number(item.Period) === targetPeriod;
    });
    const original = { Date: source.Date, Period: source.Period, StartTime: source.StartTime, EndTime: source.EndTime };
    source.Date = targetDate;
    source.Period = targetPeriod;
    if (payload.StartTime !== undefined) source.StartTime = payload.StartTime;
    if (payload.EndTime !== undefined) source.EndTime = payload.EndTime;
    source.UpdatedAt = new Date().toISOString();
    upsertObject_('Lessons', 'ID', source);
    updateRecordDates_(source.ID, source.Date);
    if (destination) {
      destination.Date = original.Date;
      destination.Period = original.Period;
      destination.StartTime = original.StartTime;
      destination.EndTime = original.EndTime;
      destination.UpdatedAt = new Date().toISOString();
      upsertObject_('Lessons', 'ID', destination);
      updateRecordDates_(destination.ID, destination.Date);
    }
    reindexLessons_();
    log_(destination ? 'SWAP_LESSON' : 'MOVE_LESSON', source.ID + (destination ? ' <-> ' + destination.ID : ''));
    return { moved: source.ID, swapped: destination ? destination.ID : null };
  });
}

function reindexLessons_(groupKeys) {
  ensureSchema_();
  const sheet = getSheet_('Lessons');
  const headers = APP.sheets.Lessons;
  const lessons = rowsToObjects_('Lessons');
  const selected = groupKeys && groupKeys.length ? new Set(groupKeys) : null;
  const reindexed = computeReindexedLessons_(lessons, selected);
  if (reindexed.length) {
    const values = reindexed.map(function (row) {
      return headers.map(function (header) { return serializeCell_(header, row[header]); });
    });
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  log_('REINDEX', selected ? Array.from(selected).join(', ') : 'all');
  return reindexed;
}

/** Spreadsheet非依存。ローカルテスト可能な再採番本体。 */
function computeReindexedLessons_(lessons, selectedGroups) {
  const result = lessons.map(function (item) { return Object.assign({}, item); });
  const groups = {};
  result.forEach(function (lesson) {
    const key = lessonGroupKey_(lesson);
    if (selectedGroups && !selectedGroups.has(key)) return;
    if (!groups[key]) groups[key] = [];
    groups[key].push(lesson);
  });
  Object.keys(groups).forEach(function (key) {
    let nextNumber = 1;
    groups[key].sort(compareLessons_).forEach(function (lesson) {
      if (lesson.NumberMode === 'none') {
        lesson.LessonNo = '';
        return;
      }
      if (lesson.NumberMode === 'manual') {
        const manual = Number(lesson.LessonNo);
        if (!Number.isInteger(manual) || manual < 1) throw new Error('手動授業番号は1以上の整数で指定してください。');
        nextNumber = Math.max(nextNumber, manual + 1);
        return;
      }
      lesson.LessonNo = nextNumber;
      nextNumber += 1;
    });
  });
  return result;
}

function saveRecord_(input) {
  ensureSchema_();
  if (!input.LessonID) throw new Error('LessonIDが必要です。');
  return withLock_(function () {
    const lesson = rowsToObjects_('Lessons').find(function (item) { return item.ID === input.LessonID; });
    if (!lesson) throw new Error('対象授業が見つかりません。');
    const records = rowsToObjects_('Records');
    const existing = input.RecordID
      ? records.find(function (item) { return item.RecordID === input.RecordID; })
      : records.find(function (item) { return item.LessonID === input.LessonID; });
    if (input.RecordID && !existing) throw new Error('編集対象の授業記録が見つかりません。');
    if (existing && existing.LessonID !== input.LessonID) throw new Error('授業記録と授業の対応が一致しません。');
    const record = {
      RecordID: existing ? existing.RecordID : Utilities.getUuid(),
      LessonID: input.LessonID,
      Date: lesson.Date,
      Plan: String(input.Plan || ''),
      Result: String(input.Result || ''),
      Homework: String(input.Homework || ''),
      Reflection: String(input.Reflection || ''),
      UpdatedAt: new Date().toISOString()
    };
    upsertObject_('Records', 'RecordID', record);
    log_('SAVE_RECORD', record.RecordID);
    return record;
  });
}

function listRecords_(filters) {
  ensureSchema_();
  filters = filters || {};
  return rowsToObjects_('Records').filter(function (record) {
    const text = [record.Plan, record.Result, record.Homework, record.Reflection].join(' ').toLowerCase();
    return (!filters.LessonID || record.LessonID === filters.LessonID) &&
      (!filters.query || text.indexOf(String(filters.query).toLowerCase()) >= 0);
  });
}

function getMasters_() {
  const grouped = { Subject: [], Grade: [], Class: [], Room: [], Topic: [] };
  rowsToObjects_('Masters').forEach(function (item) {
    if (!grouped[item.Type]) grouped[item.Type] = [];
    grouped[item.Type].push(item.Value);
  });
  return grouped;
}

function saveMaster_(input) {
  ensureSchema_();
  const master = normalizeMaster_(input);
  return withLock_(function () {
    const exists = rowsToObjects_('Masters').some(function (item) {
      return item.Type === master.Type && item.Value.toLocaleLowerCase() === master.Value.toLocaleLowerCase();
    });
    if (exists) throw new Error('同じマスタ項目が既に登録されています。');
    appendObject_('Masters', master);
    log_('SAVE_MASTER', master.Type + ': ' + master.Value);
    return getMasters_();
  });
}

function deleteMaster_(input) {
  ensureSchema_();
  const master = normalizeMaster_(input);
  return withLock_(function () {
    const sheet = getSheet_('Masters');
    const rows = rowsToObjectsWithRows_('Masters')
      .filter(function (item) { return item.data.Type === master.Type && item.data.Value === master.Value; })
      .map(function (item) { return item.row; })
      .sort(function (a, b) { return b - a; });
    if (!rows.length) throw new Error('削除対象のマスタ項目が見つかりません。');
    rows.forEach(function (row) { sheet.deleteRow(row); });
    log_('DELETE_MASTER', master.Type + ': ' + master.Value);
    return getMasters_();
  });
}

function exportJson_() {
  ensureSchema_();
  const data = {};
  Object.keys(APP.sheets).forEach(function (name) { data[name] = rowsToObjects_(name); });
  return { exportedAt: new Date().toISOString(), version: APP.backupVersion, data: data };
}

function exportCsv_(input) {
  ensureSchema_();
  const name = String((input && (input.sheet || input.name)) || 'Lessons');
  if (!Object.prototype.hasOwnProperty.call(APP.sheets, name)) throw new Error('CSV出力対象のシート名が不正です。');
  const headers = APP.sheets[name];
  const lines = [headers].concat(rowsToObjects_(name).map(function (row) {
    return headers.map(function (header) { return row[header]; });
  })).map(function (row) { return row.map(csvEscape_).join(','); });
  return {
    filename: name + '-' + formatDate_(new Date()) + '.csv',
    mimeType: 'text/csv;charset=utf-8',
    content: '\uFEFF' + lines.join('\r\n')
  };
}

function restoreJson_(input) {
  ensureSchema_();
  const backup = typeof input === 'string' ? JSON.parse(input) : (input.backup || input);
  const validated = validateBackup_(backup);
  return withLock_(function () {
    Object.keys(APP.sheets).forEach(function (name) {
      const sheet = getSheet_(name);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, APP.sheets[name].length).clearContent();
      const rows = validated[name];
      if (rows.length) {
        const values = rows.map(function (row) {
          return APP.sheets[name].map(function (header) { return serializeCell_(header, row[header]); });
        });
        sheet.getRange(2, 1, values.length, APP.sheets[name].length).setValues(values);
      }
    });
    reindexLessons_();
    log_('RESTORE_JSON', 'version=' + backup.version + ', lessons=' + validated.Lessons.length);
    return { restored: true, version: backup.version, counts: Object.keys(validated).reduce(function (out, name) {
      out[name] = validated[name].length;
      return out;
    }, {}) };
  });
}

/** Spreadsheet非依存。復元前に全データを検査し、不完全な復元を防ぐ。 */
function validateBackup_(backup) {
  if (!backup || typeof backup !== 'object' || Array.isArray(backup)) throw new Error('バックアップJSONが不正です。');
  if (Number(backup.version) !== APP.backupVersion) throw new Error('対応していないバックアップバージョンです。');
  if (!backup.data || typeof backup.data !== 'object' || Array.isArray(backup.data)) throw new Error('バックアップにdataがありません。');
  const result = {};
  Object.keys(APP.sheets).forEach(function (name) {
    const rows = backup.data[name];
    if (!Array.isArray(rows)) throw new Error(name + ' のデータがありません。');
    result[name] = rows.map(function (row, index) {
      if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error(name + ' の' + (index + 1) + '行目が不正です。');
      const normalized = {};
      APP.sheets[name].forEach(function (header) {
        if (!Object.prototype.hasOwnProperty.call(row, header)) throw new Error(name + ' の' + (index + 1) + '行目に ' + header + ' がありません。');
        normalized[header] = row[header] === null || row[header] === undefined ? '' : row[header];
      });
      return normalized;
    });
  });
  validateRestoredRows_(result);
  return result;
}

function validateRestoredRows_(data) {
  const lessonIds = {};
  data.Lessons.forEach(function (row, index) {
    const lesson = normalizeLesson_(row);
    if (!row.ID || lessonIds[row.ID]) throw new Error('Lessons のIDが空または重複しています（' + (index + 1) + '行目）。');
    lessonIds[row.ID] = lesson.Date;
  });
  const recordIds = {};
  data.Records.forEach(function (row, index) {
    if (!row.RecordID || recordIds[row.RecordID]) throw new Error('Records のRecordIDが空または重複しています（' + (index + 1) + '行目）。');
    if (!lessonIds[row.LessonID]) throw new Error('Records のLessonIDに対応する授業がありません（' + (index + 1) + '行目）。');
    row.Date = lessonIds[row.LessonID];
    recordIds[row.RecordID] = true;
  });
  const masters = {};
  data.Masters.forEach(function (row, index) {
    const master = normalizeMaster_(row);
    const key = master.Type + '|' + master.Value.toLocaleLowerCase();
    if (masters[key]) throw new Error('Masters に重複があります（' + (index + 1) + '行目）。');
    masters[key] = true;
  });
  data.BasicTimetable.forEach(function (row, index) {
    const weekday = Number(row.Weekday);
    const period = Number(row.Period);
    if (!Number.isInteger(weekday) || weekday < 1 || weekday > 5 || !Number.isInteger(period) || period < 1 || period > 12 ||
      !String(row.Subject || '').trim() || ['auto', 'none'].indexOf(String(row.NumberMode || 'auto')) < 0) {
      throw new Error('BasicTimetable の' + (index + 1) + '行目が不正です。');
    }
  });
}

function normalizeMaster_(input) {
  const type = String((input && input.Type) || '').trim();
  const value = String((input && input.Value) || '').trim();
  if (APP.masterTypes.indexOf(type) < 0) throw new Error('Typeは ' + APP.masterTypes.join('、') + ' のいずれかです。');
  if (!value) throw new Error('Valueは必須です。');
  if (value.length > 100) throw new Error('Valueは100文字以内で指定してください。');
  if(/[\u0000-\u001F\u007F]/.test(value)) throw new Error('Valueに制御文字は使用できません。');
  if (value.charAt(0) === '=') throw new Error('Valueを「=」から始めることはできません。');
  return { Type: type, Value: value };
}

function csvEscape_(value) {
  let text = value === null || value === undefined ? '' : String(value);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return '"' + text.replace(/"/g, '""') + '"';
}

function normalizeLesson_(input) {
  const date = String(input.Date || '').trim();
  const period = Number(input.Period);
  const mode = String(input.NumberMode || 'auto');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('DateはYYYY-MM-DD形式で指定してください。');
  parseDate_(date);
  if (!Number.isInteger(period) || period < 1 || period > 12) throw new Error('Periodは1〜12の整数で指定してください。');
  if (!APP.lessonModes.includes(mode)) throw new Error('NumberModeはauto、manual、noneのいずれかです。');
  if (!String(input.Subject || '').trim()) throw new Error('教科は必須です。');
  if (!String(input.Grade || '').trim()) throw new Error('学年は必須です。');
  if (!String(input.Class || '').trim()) throw new Error('クラスは必須です。');
  if (mode === 'manual' && (!Number.isInteger(Number(input.LessonNo)) || Number(input.LessonNo) < 1)) {
    throw new Error('手動授業番号は1以上の整数で指定してください。');
  }
  return {
    ID: input.ID || Utilities.getUuid(),
    Date: date,
    Period: period,
    Subject: String(input.Subject).trim(),
    Grade: String(input.Grade).trim(),
    Class: String(input.Class).trim(),
    Room: String(input.Room || '').trim(),
    Topic: String(input.Topic || '').trim(),
    NumberMode: mode,
    LessonNo: mode === 'manual' ? Number(input.LessonNo) : (mode === 'none' ? '' : Number(input.LessonNo) || ''),
    Memo: String(input.Memo || ''),
    StartTime: String(input.StartTime || ''),
    EndTime: String(input.EndTime || ''),
    UpdatedAt: new Date().toISOString()
  };
}

function ensureSchema_() {
  const ss = BrowserTables.getActiveSpreadsheet();
  const missing = Object.keys(APP.sheets).filter(function (name) { return !ss.getSheetByName(name); });
  if (missing.length) setupSpreadsheet();
}

function getSheet_(name) {
  const sheet = BrowserTables.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(name + ' シートがありません。setupSpreadsheetを実行してください。');
  return sheet;
}

function rowsToObjects_(name) {
  return rowsToObjectsWithRows_(name).map(function (item) { return item.data; });
}

function rowsToObjectsWithRows_(name) {
  const sheet = getSheet_(name);
  const lastRow = sheet.getLastRow();
  const headers = APP.sheets[name];
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues().map(function (row, index) {
    const data = {};
    headers.forEach(function (header, column) { data[header] = normalizeCell_(header, row[column]); });
    return { row: index + 2, data: data };
  });
}

function appendObject_(name, object) {
  const headers = APP.sheets[name];
  getSheet_(name).appendRow(headers.map(function (header) { return serializeCell_(header, object[header]); }));
}

function upsertObject_(name, key, object) {
  const sheet = getSheet_(name);
  const headers = APP.sheets[name];
  const existing = rowsToObjectsWithRows_(name).find(function (item) { return item.data[key] === object[key]; });
  const values = headers.map(function (header) { return serializeCell_(header, object[header]); });
  if (existing) sheet.getRange(existing.row, 1, 1, headers.length).setValues([values]);
  else sheet.appendRow(values);
}

function deleteRowsByValue_(name, key, value) {
  const sheet = getSheet_(name);
  rowsToObjectsWithRows_(name)
    .filter(function (item) { return item.data[key] === value; })
    .map(function (item) { return item.row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (row) { sheet.deleteRow(row); });
}

function updateRecordDates_(lessonId, date) {
  rowsToObjectsWithRows_('Records')
    .filter(function (item) { return item.data.LessonID === lessonId && item.data.Date !== date; })
    .forEach(function (item) {
      item.data.Date = date;
      item.data.UpdatedAt = new Date().toISOString();
      const values = APP.sheets.Records.map(function (header) { return serializeCell_(header, item.data[header]); });
      getSheet_('Records').getRange(item.row, 1, 1, values.length).setValues([values]);
    });
}

function lessonGroupKey_(lesson) {
  return JSON.stringify([String(lesson.Subject), String(lesson.Grade), String(lesson.Class)]);
}

function compareLessons_(a, b) {
  return String(a.Date).localeCompare(String(b.Date)) ||
    Number(a.Period) - Number(b.Period) ||
    String(a.StartTime || '').localeCompare(String(b.StartTime || '')) ||
    String(a.ID).localeCompare(String(b.ID));
}

function normalizeCell_(header, value) {
  if (header === 'Date' && value instanceof Date) return formatDate_(value);
  if (header === 'Period' || header === 'LessonNo') return value === '' ? '' : Number(value);
  return value === null || value === undefined ? '' : String(value);
}

function serializeCell_(header,value) { return value == null ? "" : value; }

function parseDate_(value) {
  const parts = String(value).split('-').map(Number);
  if (parts.length !== 3) throw new Error('日付形式が正しくありません。');
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    throw new Error('存在しない日付です。');
  }
  return date;
}

function formatDate_(date) {
  return Utilities.formatDate(date, APP.timeZone, 'yyyy-MM-dd');
}

function addDays_(date, count) {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  copy.setDate(copy.getDate() + count);
  return copy;
}

function withLock_(callback) { return callback(); }

function log_(action, detail) {
  const sheet = BrowserTables.getActiveSpreadsheet().getSheetByName('Logs');
  if (sheet) sheet.appendRow([new Date().toISOString(), action, detail || '']);
}


// Local table storage preserves the existing timetable and backup schema.
const STORE_KEY='learning-apps-jp:timetable:v1';
let localTables=Object.create(null);
const Utilities={getUuid:()=>crypto.randomUUID(),formatDate:d=>[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-')};
class LocalSheet {
  constructor(name){this.name=name;}
  get rows(){return localTables[this.name];}
  getLastRow(){let n=this.rows.length;while(n&&this.rows[n-1].every(v=>v===''))n--;return n;}
  getRange(row,col,height,width){const sheet=this;return {
    getValues(){return Array.from({length:height},(_,r)=>Array.from({length:width},(_,c)=>sheet.rows[row-1+r]?.[col-1+c]??''));},
    setValues(values){for(let r=0;r<height;r++){const target=sheet.rows[row-1+r]??(sheet.rows[row-1+r]=[]);for(let c=0;c<width;c++)target[col-1+c]=values[r][c];}return this;},
    clearContent(){for(let r=0;r<height;r++)for(let c=0;c<width;c++)if(sheet.rows[row-1+r])sheet.rows[row-1+r][col-1+c]='';return this;},
    setFontWeight(){return this;},setBackground(){return this;}
  };}
  appendRow(values){this.rows.splice(this.getLastRow(),this.rows.length-this.getLastRow(),values);if(this.name==='Logs'&&this.rows.length>501)this.rows.splice(1,this.rows.length-501);}
  deleteRow(row){this.rows.splice(row-1,1);}
  setFrozenRows(){} autoResizeColumns(){}
}
const BrowserTables={getActiveSpreadsheet:()=>({getSheetByName:name=>Object.hasOwn(localTables,name)?new LocalSheet(name):null,insertSheet:name=>{localTables[name]=[];return new LocalSheet(name);}})};
function readLocalTables(){
  const raw=localStorage.getItem(STORE_KEY);
  if(raw===null){localTables=Object.create(null);return;}
  let parsed;try{parsed=JSON.parse(raw);}catch{throw new Error('保存データが壊れています。上書きせず処理を停止しました。');}
  if(!parsed||parsed.version!==1||!parsed.tables)throw new Error('保存データの形式を確認できません。');
  for(const [name,headers] of Object.entries(APP.sheets)){
    const rows=parsed.tables[name];
    if(!Array.isArray(rows)||!rows.length||!Array.isArray(rows[0])||rows[0].join('|')!==headers.join('|')||rows.some(r=>!Array.isArray(r)))throw new Error('保存データの表が不正です：'+name);
  }
  localTables=parsed.tables;
}
let localQueue=Promise.resolve();
function localCall(action,payload={}){
  const run=()=>{
    readLocalTables();const before=structuredClone(localTables);
    try{
      const result=dispatchAction_(action,payload);
      const serialized=JSON.stringify({version:1,tables:localTables});
      try{localStorage.setItem(STORE_KEY,serialized);}catch{throw new Error('端末に保存できませんでした。保存容量やブラウザーの設定を確認してください。変更は取り消しました。');}
      return structuredClone(result);
    }catch(error){localTables=before;throw error;}
  };
  const pending=localQueue.then(()=>globalThis.navigator?.locks?navigator.locks.request(STORE_KEY,run):run());
  localQueue=pending.catch(()=>{});return pending;
}
