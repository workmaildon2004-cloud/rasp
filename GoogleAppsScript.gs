const PLANNING_HEADERS = {
  TeacherWorkload: ["Факультет", "Преподаватель", "Предмет", "Выбранные виды занятий", "Вид занятия", "Часы", "Семестр"],
  ServiceTeachers: ["Факультет", "Преподаватель", "Должность", "Дни работы"],
  Schedule: ["ID", "Факультет", "Группа", "ID группы", "Код группы", "Количество студентов", "Преподаватель", "Предмет", "Тип занятия", "День недели", "Начало", "Длительность, мин", "Пар подряд", "Аудитория", "Цвет", "Еженедельно", "IDs групп", "Данные групп"]
};
const DIRECTORY_SPREADSHEET_ID = "1eEUxo_gzLf51HghBK1UQK0JV5Q_yVTSZ68h7fPkXANc";

function doGet(e) {
  const name = e?.parameter?.sheet;
  if (!name) {
    return json({ error: "Укажите имя листа в URL, например ?sheet=TeacherWorkload" });
  }
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = isPlanningSheet(name) ? getOrCreatePlanningSheet(spreadsheet, name) : spreadsheet.getSheetByName(name);

  if (!sheet) {
    return json({ error: "Лист не найден: " + name });
  }

  if (isPlanningSheet(name)) migrateLegacyPlanningSheet(sheet, name);
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const data = rows
    .filter(row => row.some(cell => cell !== ""))
    .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));

  if (name === "Schedule") {
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    data.forEach(record => {
      if (record["Начало"] instanceof Date) {
        record["Начало"] = Utilities.formatDate(record["Начало"], timezone, "HH:mm");
      }
      if (record["Длительность, мин"] instanceof Date) {
        const hours = Number(Utilities.formatDate(record["Длительность, мин"], timezone, "H"));
        const minutes = Number(Utilities.formatDate(record["Длительность, мин"], timezone, "m"));
        record["Длительность, мин"] = hours * 60 + minutes;
      }
    });
  }

  return json(data);
}

function testPlanningSheets() {
  ["TeacherWorkload", "ServiceTeachers", "Schedule"].forEach(name => {
    const records = JSON.parse(doGet({ parameter: { sheet: name } }).getContent());
    if (!Array.isArray(records)) throw new Error(records.error || `Не удалось проверить лист ${name}`);
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    const headers = sheet.getRange(1, 1, 1, PLANNING_HEADERS[name].length).getValues()[0];
    if (JSON.stringify(headers) !== JSON.stringify(PLANNING_HEADERS[name])) {
      throw new Error(`${name}: заголовки таблицы не совпадают`);
    }
    Logger.log(`${name}: OK (${records.length} строк, табличный формат)`);
  });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action === "deleteDirectoryGroup") return deleteDirectoryGroup(body);
  if (body.action === "deleteDirectoryRecord") return deleteDirectoryRecord(body);
  if (body.action === "updateDirectoryRecord") return updateDirectoryRecord(body);
  if (body.action === "appendDirectoryRecord") return appendDirectoryRecord(body);
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = isPlanningSheet(body.sheet) ? getOrCreatePlanningSheet(spreadsheet, body.sheet) : spreadsheet.getSheetByName(body.sheet);

  if (!sheet) return json({ error: "Лист не найден: " + body.sheet });

  if (isPlanningSheet(body.sheet)) {
    migrateLegacyPlanningSheet(sheet, body.sheet);
    const rows = planningSnapshotToRows(body.sheet, body.data);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, PLANNING_HEADERS[body.sheet].length).setValues([PLANNING_HEADERS[body.sheet]]);
    if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    return json({ ok: true, rows: rows.length });
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(header => body.data[header] ?? ""));
  return json({ ok: true });
}

function getDirectorySheetAndColumns(body) {
  const sheetsByType = {
    teachers: { computer: "Subjects", electrical: "EI.subject" },
    subjects: { computer: "Subjects", electrical: "EI.subject" },
    rooms: { computer: "Rooms", electrical: "EI.rooms" },
    groups: { computer: "Groups", electrical: "EI.groups" }
  };
  const sheetName = sheetsByType[body.type]?.[body.faculty];
  const sheet = sheetName ? SpreadsheetApp.openById(DIRECTORY_SPREADSHEET_ID).getSheetByName(sheetName) : null;
  if (!sheet || !sheet.getLastColumn()) return { sheet, columns: null };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  return {
    sheet,
    columns: {
      teacher: findDirectoryColumn(headers, /оқытушы|преподав|фио|teacher/iu),
      role: findDirectoryColumn(headers, /дәреж|лауазым|должност|зван|position|rank/iu),
      subject: findDirectoryColumn(headers, /пән|предмет|дисциплин|subject/iu),
      roomNumber: findDirectoryColumn(headers, /кабинет|аудитор|room/iu),
      roomType: findDirectoryColumn(headers, /түрі|тип|type/iu),
      groupCode: findDirectoryColumn(headers, /шифр|код|code/iu),
      groupName: findDirectoryColumn(headers, /топ|групп|group|атауы|назван|name/iu),
      studentCount: findDirectoryColumn(headers, /барлығы|всего|колич|числен|обуча|студент|student/iu)
    }
  };
}

function directoryRecordMatches(row, columns, type, data) {
  const normalize = value => String(value ?? "").trim().toLocaleLowerCase();
  if (type === "teachers") return normalize(row[columns.teacher]) === normalize(data.teacher);
  if (type === "rooms") return normalize(row[columns.roomNumber]) === normalize(data.roomNumber);
  if (type === "subjects") {
    return normalize(row[columns.teacher]) === normalize(data.teacher) &&
      normalize(row[columns.subject]) === normalize(data.subject);
  }
  return normalize(row[columns.groupCode]) === normalize(data.groupCode) &&
    normalize(row[columns.groupName]) === normalize(data.groupName);
}

function deleteDirectoryRecord(body) {
  if (!["teachers", "rooms", "groups", "subjects"].includes(body.type) ||
    !["computer", "electrical"].includes(body.faculty)) return json({ ok: false, error: "invalid_request" });
  const { sheet, columns } = getDirectorySheetAndColumns(body);
  if (!sheet || !columns) return json({ ok: false, error: "not_found" });
  const data = body.data && typeof body.data === "object" ? body.data : {};
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  const matches = rows.map((row, index) => directoryRecordMatches(row, columns, body.type, data) ? index + 2 : 0).filter(Boolean);
  if (matches.length === 0) return json({ ok: false, error: "not_found" });
  if (matches.length > 1) return json({ ok: false, error: "ambiguous" });
  sheet.deleteRow(matches[0]);
  return json({ ok: true });
}

function updateDirectoryRecord(body) {
  if (!["teachers", "rooms", "groups", "subjects"].includes(body.type) ||
    !["computer", "electrical"].includes(body.faculty)) return json({ ok: false, error: "invalid_request" });
  const { sheet, columns } = getDirectorySheetAndColumns(body);
  if (!sheet || !columns) return json({ ok: false, error: "not_found" });
  const original = body.original && typeof body.original === "object" ? body.original : {};
  const data = body.data && typeof body.data === "object" ? body.data : {};
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  const matches = rows.map((row, index) => directoryRecordMatches(row, columns, body.type, original) ? index + 2 : 0).filter(Boolean);
  if (matches.length === 0) return json({ ok: false, error: "not_found" });
  if (matches.length > 1) return json({ ok: false, error: "ambiguous" });
  const values = {
    teacher: String(data.teacher ?? "").trim(),
    role: String(data.role ?? "").trim(),
    subject: String(data.subject ?? "").trim(),
    roomNumber: String(data.roomNumber ?? "").trim(),
    roomType: String(data.roomType ?? "").trim(),
    groupCode: String(data.groupCode ?? "").trim(),
    groupName: String(data.groupName ?? "").trim(),
    studentCount: Number(data.studentCount)
  };
  if (body.type === "groups" && (!values.groupCode || !values.groupName || !Number.isInteger(values.studentCount) || values.studentCount < 0 || values.studentCount > 9999 || values.groupCode.length > 80 || values.groupName.length > 120)) {
    return json({ ok: false, error: "invalid_record" });
  }
  if (body.type === "teachers" && (!values.teacher || values.teacher.length > 160 || values.role.length > 120)) return json({ ok: false, error: "invalid_record" });
  if (body.type === "rooms" && (!values.roomNumber || !values.roomType || values.roomNumber.length > 60 || !["Лекция", "Лаборатория и практика"].includes(values.roomType))) return json({ ok: false, error: "invalid_record" });
  if (body.type === "subjects" && (!values.teacher || !values.subject || values.teacher.length > 160 || values.subject.length > 160 || values.role.length > 120)) return json({ ok: false, error: "invalid_record" });
  const row = matches[0];
  const updates = body.type === "teachers" ? { teacher: values.teacher, role: values.role }
    : body.type === "rooms" ? { roomNumber: values.roomNumber, roomType: values.roomType }
      : body.type === "subjects" ? { teacher: values.teacher, subject: values.subject, role: values.role }
        : { groupCode: values.groupCode, groupName: values.groupName, studentCount: values.studentCount };
  Object.entries(updates).forEach(([key, value]) => {
    if (columns[key] >= 0) sheet.getRange(row, columns[key] + 1).setValue(value);
  });
  return json({ ok: true });
}

function deleteDirectoryGroup(body) {
  const allowedFaculties = ["computer", "electrical"];
  const groupCode = String(body.groupCode ?? "").trim();
  const groupName = String(body.groupName ?? "").trim();
  if (!allowedFaculties.includes(body.faculty) || !groupCode || !groupName ||
    groupCode.length > 80 || groupName.length > 120) {
    return json({ ok: false, error: "invalid_request" });
  }

  const sheetName = body.faculty === "computer" ? "Groups" : "EI.groups";
  const sheet = SpreadsheetApp.openById(DIRECTORY_SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return json({ ok: false, error: "not_found" });
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const codeColumn = findDirectoryColumn(headers, /шифр|код|code/iu);
  const nameColumn = findDirectoryColumn(headers, /топтар|групп|group name|назван/iu);
  if (codeColumn < 0 || nameColumn < 0) return json({ ok: false, error: "missing_columns" });

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const normalizeValue = value => String(value ?? "").trim().toLocaleLowerCase();
  const matches = [];
  rows.forEach((row, index) => {
    if (normalizeValue(row[codeColumn]) === normalizeValue(groupCode) &&
      normalizeValue(row[nameColumn]) === normalizeValue(groupName)) {
      matches.push(index + 2);
    }
  });
  if (!matches.length) return json({ ok: false, error: "not_found" });
  if (matches.length > 1) return json({ ok: false, error: "ambiguous" });
  sheet.deleteRow(matches[0]);
  return json({ ok: true });
}

function appendDirectoryRecord(body) {
  body = body && typeof body === "object" ? body : {};
  const allowedTypes = ["teachers", "rooms", "groups", "subjects"];
  const allowedFaculties = ["computer", "electrical"];
  if (!allowedTypes.includes(body.type) || !allowedFaculties.includes(body.faculty)) {
    return json({ ok: false, error: "invalid_request" });
  }

  const sheetsByType = {
    teachers: { computer: "Subjects", electrical: "EI.subject" },
    subjects: { computer: "Subjects", electrical: "EI.subject" },
    rooms: { computer: "Rooms", electrical: "EI.rooms" },
    groups: { computer: "Groups", electrical: "EI.groups" }
  };
  const spreadsheet = SpreadsheetApp.openById(DIRECTORY_SPREADSHEET_ID);
  const sheetName = sheetsByType[body.type][body.faculty];
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return json({ ok: false, error: "sheet_not_found" });

  let lastColumn = sheet.getLastColumn();
  if (!lastColumn) return json({ ok: false, error: "missing_headers" });
  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const columns = {
    teacher: findDirectoryColumn(headers, /оқытушы|преподав|фио|teacher/iu),
    role: findDirectoryColumn(headers, /дәреж|лауазым|должност|зван|position|rank/iu),
    subject: findDirectoryColumn(headers, /пән|предмет|дисциплин|subject/iu),
    roomNumber: findDirectoryColumn(headers, /кабинет|аудитор|room/iu),
    roomType: findDirectoryColumn(headers, /түрі|тип|type/iu),
    groupCode: findDirectoryColumn(headers, /шифр|код|code/iu),
    groupName: findDirectoryColumn(headers, /топ|групп|group|атауы|назван|name/iu),
    studentCount: findDirectoryColumn(headers, /барлығы|всего|колич|числен|обуча|студент|student/iu)
  };
  const data = body.data && typeof body.data === "object" ? body.data : {};
  const value = key => String(data[key] ?? "").trim();
  const values = {};
  let requiredColumns = [];

  if (body.type === "teachers" || body.type === "subjects") {
    values.teacher = value("teacher");
    values.role = value("role");
    values.subject = value("subject");
    requiredColumns = ["teacher", ...(body.type === "subjects" ? ["subject"] : [])];
    if (!values.teacher || (body.type === "subjects" && !values.subject) ||
      values.teacher.length > 160 || values.role.length > 120 || values.subject.length > 160) {
      return json({ ok: false, error: "invalid_record" });
    }
  } else if (body.type === "rooms") {
    values.roomNumber = value("roomNumber");
    values.roomType = value("roomType");
    requiredColumns = ["roomNumber", "roomType"];
    if (!values.roomNumber || !values.roomType || values.roomNumber.length > 60) {
      return json({ ok: false, error: "invalid_record" });
    }
    if (!["Лекция", "Лаборатория и практика"].includes(values.roomType)) {
      return json({ ok: false, error: "invalid_record" });
    }
  } else {
    values.groupCode = value("groupCode");
    values.groupName = value("groupName");
    values.studentCount = Number(data.studentCount);
    requiredColumns = ["groupCode", "groupName", "studentCount"];
    if (!values.groupCode || !values.groupName || !Number.isInteger(values.studentCount) ||
      values.studentCount < 0 || values.studentCount > 9999 ||
      values.groupCode.length > 80 || values.groupName.length > 120) {
      return json({ ok: false, error: "invalid_record" });
    }
  }

  if (body.type !== "groups" && requiredColumns.some(key => columns[key] < 0)) {
    return json({ ok: false, error: "missing_columns" });
  }
  if ((body.type === "teachers" || body.type === "subjects") && columns.subject < 0) {
    return json({ ok: false, error: "missing_columns" });
  }
  if (body.type === "groups") {
    const newHeaders = {
      groupCode: "Код группы",
      groupName: "Название группы",
      studentCount: "Количество студентов"
    };
    const missing = requiredColumns.filter(key => columns[key] < 0);
    if (missing.length) {
      const startColumn = lastColumn + 1;
      const headersToAdd = missing.map(key => newHeaders[key]);
      sheet.getRange(1, startColumn, 1, headersToAdd.length).setValues([headersToAdd]);
      headers = headers.concat(headersToAdd);
      missing.forEach((key, index) => {
        columns[key] = startColumn - 1 + index;
      });
      lastColumn += headersToAdd.length;
    }
  }

  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
  const normalizeValue = item => String(item ?? "").trim().toLocaleLowerCase();
  const duplicate = existingRows.some(row => {
    if (body.type === "teachers") {
      return normalizeValue(row[columns.teacher]) === normalizeValue(values.teacher);
    }
    if (body.type === "subjects") {
      return normalizeValue(row[columns.teacher]) === normalizeValue(values.teacher) &&
        normalizeValue(row[columns.subject]) === normalizeValue(values.subject);
    }
    if (body.type === "rooms") {
      return normalizeValue(row[columns.roomNumber]) === normalizeValue(values.roomNumber);
    }
    return normalizeValue(row[columns.groupCode]) === normalizeValue(values.groupCode) ||
      normalizeValue(row[columns.groupName]) === normalizeValue(values.groupName);
  });
  if (duplicate) return json({ ok: false, error: "duplicate" });

  const columnValues = new Map(Object.entries(columns)
    .filter(([key, index]) => index >= 0 && Object.prototype.hasOwnProperty.call(values, key))
    .map(([key, index]) => [index, values[key]]));
  sheet.appendRow(headers.map((_, index) => columnValues.has(index) ? columnValues.get(index) : ""));
  return json({ ok: true });
}

function findDirectoryColumn(headers, matcher) {
  return headers.findIndex(header => matcher.test(String(header).toLocaleLowerCase()));
}

function isPlanningSheet(name) {
  return Object.prototype.hasOwnProperty.call(PLANNING_HEADERS, name);
}

function getOrCreatePlanningSheet(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PLANNING_HEADERS[name].length).setValues([PLANNING_HEADERS[name]]);
  }
  return sheet;
}

function migrateLegacyPlanningSheet(sheet, name) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  if (headers.includes("payload")) {
    let snapshot = name === "TeacherWorkload" ? {} : [];
    for (let index = data.length - 1; index > 0; index--) {
      const payloadIndex = headers.indexOf("payload");
      if (typeof data[index][payloadIndex] !== "string" || !data[index][payloadIndex]) continue;
      snapshot = JSON.parse(data[index][payloadIndex]);
      break;
    }
    const rows = planningSnapshotToRows(name, snapshot);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, PLANNING_HEADERS[name].length).setValues([PLANNING_HEADERS[name]]);
    if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    return;
  }

  if (name === "TeacherWorkload" &&
    JSON.stringify(headers) === JSON.stringify(PLANNING_HEADERS.TeacherWorkload.slice(0, 6))) {
    sheet.getRange(1, 7).setValue("Семестр");
    if (data.length > 1) {
      sheet.getRange(2, 7, data.length - 1, 1)
        .setValues(data.slice(1).map(row => [row.some(value => value !== "") ? "fall" : ""]));
    }
    return;
  }

  if (name === "Schedule" && JSON.stringify(headers) === JSON.stringify(PLANNING_HEADERS.Schedule.slice(0, 16))) {
    sheet.getRange(1, 1, 1, PLANNING_HEADERS.Schedule.length).setValues([PLANNING_HEADERS.Schedule]);
    return;
  }

  if (JSON.stringify(headers.slice(0, PLANNING_HEADERS[name].length)) !== JSON.stringify(PLANNING_HEADERS[name])) {
    throw new Error(`${name}: неизвестный формат заголовков; лист не изменён`);
  }
}

function planningSnapshotToRows(name, snapshot) {
  if (name === "Schedule") {
    if (!Array.isArray(snapshot)) throw new Error("Schedule: ожидается список занятий");
    const days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница"];
    return snapshot.map(item => [
      item.id || "",
      planningFacultyLabel(item.faculty),
      item.group || "",
      item.groupId || "",
      item.groupCode || "",
      Number(item.groupSize) || 0,
      item.teacher || "",
      item.subject || "",
      item.kind || "",
      days[Number(item.day)] || "",
      item.start || "",
      Number(item.duration) || 0,
      Number(item.lessonCount) || 1,
      item.room || "",
      item.color || "",
      item.weekly !== false,
      JSON.stringify(Array.isArray(item.groupIds) ? item.groupIds : item.groupId ? [item.groupId] : []),
      JSON.stringify(Array.isArray(item.groupRefs) ? item.groupRefs : [])
    ]).filter(row => row[0] && row[7]);
  }

  if (name === "ServiceTeachers") {
    if (!Array.isArray(snapshot)) throw new Error("ServiceTeachers: ожидается список преподавателей");
    return snapshot.map(item => [
      planningFacultyLabel(item.faculty),
      item.teacher || "",
      item.role || "",
      Array.isArray(item.workDays) ? item.workDays.map(day => ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница"][day]).filter(Boolean).join(", ") : ""
    ]).filter(row => row[1]);
  }

  if (!snapshot || typeof snapshot !== "object" || Array.isArray(snapshot)) {
    throw new Error("TeacherWorkload: ожидается объект учебной нагрузки");
  }
  const subjects = new Map();
  Object.entries(snapshot).forEach(([key, value]) => {
    let parts;
    try {
      parts = JSON.parse(key);
    } catch (error) {
      throw new Error(`TeacherWorkload: некорректный ключ нагрузки ${key}`, { cause: error });
    }
    if (!Array.isArray(parts) || (parts.length !== 4 && parts.length !== 5)) return;
    const [faculty, teacher, subject, rawType] = parts;
    const semester = parts.length === 5 ? parts[4] : "fall";
    if (semester !== "fall" && semester !== "spring") {
      throw new Error(`TeacherWorkload: некорректный семестр в ключе нагрузки ${key}`);
    }
    const subjectKey = JSON.stringify([faculty, teacher, subject, semester]);
    if (!subjects.has(subjectKey)) subjects.set(subjectKey, { faculty, teacher, subject, semester, selectedTypes: [], hours: new Map() });
    const record = subjects.get(subjectKey);
    if (rawType === "types" && Array.isArray(value)) record.selectedTypes = value.map(String);
    else if (rawType === "type" && typeof value === "string" && !record.selectedTypes.length) record.selectedTypes = [value];
    else if (typeof value === "number" || (typeof value === "string" && value.trim() !== "")) record.hours.set(String(rawType), value);
  });

  const rows = [];
  subjects.forEach(record => {
    const types = [...new Set([...record.selectedTypes, ...record.hours.keys()])];
    types.forEach(type => rows.push([
      planningFacultyLabel(record.faculty),
      record.teacher,
      record.subject,
      record.selectedTypes.join(", "),
      type,
      record.hours.get(type) ?? "",
      record.semester
    ]));
  });
  return rows;
}

function planningFacultyLabel(faculty) {
  return ({ computer: "Компьютерная инженерия", electrical: "Электрическая инженерия" })[faculty] || faculty || "";
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
