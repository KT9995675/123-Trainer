/**
 * Рендеринг главной страницы веб-приложения
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Тренажер "ВСЕМ ПЯТЬ!"')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Колонка K = почта; лог с колонки L (12) */
var LOG_START_COL = 12;
var USER_DATA_COLS = 11;
/** Лист «Задачи»: E=Тема, F=Показы, G=Решено, H=Сложность (формула в таблице) */
var TASK_COL_TOPIC = 5;
var TASK_COL_SHOWS = 6;
var TASK_COL_SOLVED = 7;
var MIN_PASSWORD_LEN = 6;
var RESET_WINDOW_MS = 3 * 60 * 60 * 1000;
var MAX_RESETS_IN_WINDOW = 2;
var MAX_STEPS = 5;
var MIN_TOPIC_ATTEMPTS = 2;
var SESSION_SIZE = 10;

/**
 * Один раз запустить вручную из редактора Apps Script (▶ Выполнить)
 * для выдачи разрешения MailApp.sendEmail. Письмо придёт на вашу почту.
 */
function authorizeMailOnce() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error("Не удалось определить вашу почту. Запустите функцию, будучи залогиненным владельцем таблицы.");
  }
  MailApp.sendEmail(email, 'Тренажер — тест разрешения на почту', 'MailApp авторизован. Восстановление пароля будет работать.');
}

/**
 * Авторизация: логин = ID, телефон или почта
 */
function loginUser(login, pass) {
  try {
    const userRow = _findUserByLogin(login);
    if (!userRow) {
      return { success: false, message: "Пользователь не найден. Проверьте логин." };
    }

    const data = userRow.data;
    if (data[2].toString().trim() !== pass.toString().trim()) {
      return { success: false, message: "Неверный пароль. Попробуйте еще раз." };
    }

    const userId = data[0].toString();
    const mustChangePassword = _isPasswordTemp(userId);

    return {
      success: true,
      mustChangePassword: mustChangePassword,
      user: {
        id: userId,
        name: data[1].toString(),
        rowNum: userRow.rowNum
      }
    };
  } catch (e) {
    return { success: false, message: "Ошибка авторизации на сервере: " + e.message };
  }
}

/**
 * Регистрация нового ученика
 */
function registerUser(name, phone, email, password, schoolClass, city, school) {
  try {
    const cleanName = (name || "").toString().trim();
    const cleanPhone = _normalizePhone(phone);
    const cleanEmail = _normalizeEmail(email);
    const cleanPass = (password || "").toString().trim();
    const cleanClass = (schoolClass || "").toString().trim();
    const cleanCity = (city || "").toString().trim();
    const cleanSchool = (school || "").toString().trim();

    if (!cleanName) return { success: false, message: "Укажите имя." };
    if (!cleanPhone) return { success: false, message: "Укажите телефон." };
    if (!_isValidEmail(cleanEmail)) return { success: false, message: "Укажите корректную почту." };
    if (cleanPass.length < MIN_PASSWORD_LEN) {
      return { success: false, message: "Пароль должен быть не короче " + MIN_PASSWORD_LEN + " символов." };
    }

    if (_isPhoneTaken(cleanPhone)) {
      return { success: false, message: "Пользователь с таким телефоном уже зарегистрирован." };
    }
    if (_isEmailTaken(cleanEmail)) {
      return { success: false, message: "Пользователь с такой почтой уже зарегистрирован." };
    }

    const newId = _generateNextUserId();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    sheet.appendRow([
      newId,
      cleanName,
      cleanPass,
      cleanPhone,
      cleanClass,
      cleanCity,
      cleanSchool,
      0,
      0,
      0,
      cleanEmail
    ]);

    return {
      success: true,
      userId: newId.toString(),
      message: "Регистрация успешна! Запомните ваш ID: " + newId
    };
  } catch (e) {
    return { success: false, message: "Ошибка регистрации: " + e.message };
  }
}

/**
 * Восстановление пароля: 4 цифры на почту, не более 2 раз за 3 часа
 */
function requestPasswordReset(identifier) {
  try {
    const userRow = _findUserByPhoneOrEmail(identifier);
    if (!userRow) {
      return { success: false, message: "Пользователь с таким телефоном или почтой не найден." };
    }

    const userId = userRow.data[0].toString();
    const email = _normalizeEmail(userRow.data[10]);
    if (!email) {
      return { success: false, message: "У учетной записи не указана почта. Обратитесь к администратору." };
    }

    const limit = _canRequestPasswordReset(userId);
    if (!limit.allowed) {
      return {
        success: false,
        message: "Слишком много запросов. Повторите через " + limit.waitMinutes + " мин."
      };
    }

    const tempPassword = _generateTempPassword();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    sheet.getRange(userRow.rowNum, 3).setValue(tempPassword);
    _setPasswordTempFlag(userId);
    _recordResetAttempt(userId);

    MailApp.sendEmail({
      to: email,
      subject: 'Тренажер "ВСЕМ ПЯТЬ!" — временный пароль',
      body: "Здравствуйте, " + userRow.data[1].toString() + "!\n\n" +
        "Ваш временный пароль: " + tempPassword + "\n\n" +
        "Войдите с этим паролем — система сразу попросит придумать новый постоянный.\n\n" +
        "Если вы не запрашивали восстановление, сообщите администратору."
    });

    return {
      success: true,
      message: "Зайдите на почту " + _maskEmail(email) + ". В письме — временный пароль из 4 цифр. Введите его при входе."
    };
  } catch (e) {
    return { success: false, message: "Ошибка восстановления пароля: " + e.message };
  }
}

/**
 * Установка постоянного пароля после входа с временным
 */
function setPermanentPassword(userId, tempPassword, newPassword, confirmPassword) {
  try {
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    if (!_isPasswordTemp(userId)) {
      return { success: false, message: "Смена пароля не требуется." };
    }

    const currentPass = userRow.data[2].toString().trim();
    if (currentPass !== (tempPassword || "").toString().trim()) {
      return { success: false, message: "Неверный текущий (временный) пароль." };
    }

    const cleanNew = (newPassword || "").toString().trim();
    const cleanConfirm = (confirmPassword || "").toString().trim();
    if (cleanNew.length < MIN_PASSWORD_LEN) {
      return { success: false, message: "Новый пароль — не короче " + MIN_PASSWORD_LEN + " символов." };
    }
    if (cleanNew !== cleanConfirm) {
      return { success: false, message: "Пароли не совпадают." };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    sheet.getRange(userRow.rowNum, 3).setValue(cleanNew);
    _clearPasswordTempFlag(userId);

    return { success: true, message: "Пароль успешно изменён." };
  } catch (e) {
    return { success: false, message: "Ошибка смены пароля: " + e.message };
  }
}

/**
 * Данные для экрана старта после авторизации
 * @param {string} userId
 */
function getTrainerHome(userId) {
  try {
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    const name = userRow.data[1] ? userRow.data[1].toString().trim() : "Ученик";
    let level = parseInt(userRow.data[7], 10);
    if (isNaN(level)) level = 0;

    const attempts = _parseUserLogHistory(userRow.rowNum);
    const solvedIds = _getUserSolvedTaskIds(attempts);
    const topicStats = _buildUserTopicStats(attempts);
    const steps = _buildStepSummaries(level, solvedIds);

    let greeting;
    if (level >= MAX_STEPS) {
      greeting = name + ", все ступени освоены, поздравляем!";
    } else {
      greeting = name + ", ты готовишься взять ступень " + (level + 1) + ". Успехов!";
    }

    return {
      success: true,
      name: name,
      studentLevel: level,
      allCompleted: level >= MAX_STEPS,
      greeting: greeting,
      topics: topicStats.topics,
      topTopics: topicStats.topTopics,
      weakTopics: topicStats.weakTopics,
      steps: steps
    };
  } catch (e) {
    return { success: false, message: "Ошибка загрузки старта: " + e.message };
  }
}

/**
 * Получение задач для новой сессии (только чтение уровня, без записи в лог)
 * @param {string} userId
 * @param {number} selectedStep — выбранная ступень 1..5
 */
function getTaskForUser(userId, selectedStep) {
  try {
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    let level = parseInt(userRow.data[7], 10);
    if (isNaN(level)) level = 0;

    let step = parseInt(selectedStep, 10);
    if (isNaN(step) || step < 1 || step > MAX_STEPS) {
      return { success: false, message: "Неверная ступень." };
    }

    const maxAvailable = level >= MAX_STEPS ? MAX_STEPS : level + 1;
    if (step > maxAvailable) {
      return { success: false, message: "Эта ступень пока недоступна." };
    }

    const attempts = _parseUserLogHistory(userRow.rowNum);
    const solvedIds = _getUserSolvedTaskIds(attempts);
    const queueData = _buildSessionQueue(level, step, solvedIds);

    return {
      success: true,
      allCompleted: false,
      sessionTasks: queueData.sessionTasks,
      queueIds: queueData.queueIds,
      sessionIndex: 0,
      playingStep: step,
      currentRunStats: _buildRunStats(level, 1, 0, 0, step)
    };
  } catch (e) {
    return { success: false, message: "Ошибка получения задачи: " + e.message };
  }
}

/**
 * Запись завершённой сессии в лог одним блоком (победа или поражение)
 * @param {number} selectedStep — ступень, на которой шла сессия
 */
function finalizeSession(userId, rowNum, queueIds, pairs, runStatus, selectedStep) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    let row = parseInt(rowNum, 10);

    if (isNaN(row) || row < 2) {
      const userRow = _getUserRowData(userId);
      if (!userRow) return { success: false, message: "Пользователь не найден." };
      row = userRow.rowNum;
    }

    const userData = sheet.getRange(row, 8, 1, 3).getValues()[0];
    let level = parseInt(userData[0], 10);
    if (isNaN(level)) level = 0;

    let step = parseInt(selectedStep, 10);
    if (isNaN(step) || step < 1 || step > MAX_STEPS) {
      return { success: false, message: "Неверная ступень сессии." };
    }

    let correctCount = 0;
    let errorCount = 0;
    for (let i = 0; i < pairs.length; i++) {
      if (pairs[i].isCorrect) {
        correctCount++;
      } else {
        errorCount++;
      }
    }

    if (runStatus === "win" || runStatus === "all_clear" || runStatus === "practice_win") {
      if (correctCount < 7) {
        return { success: false, message: "Недостаточно верных ответов для победы." };
      }
    } else if (runStatus === "fail") {
      if (errorCount < 3) {
        return { success: false, message: "Недостаточно ошибок для поражения." };
      }
    } else {
      return { success: false, message: "Неверный статус сессии." };
    }

    const markerCol = _getNextLogColumn(sheet, row);
    const logRow = ["##", queueIds.join(",")];
    for (let i = 0; i < pairs.length; i++) {
      logRow.push(pairs[i].taskId.toString());
      logRow.push(pairs[i].isCorrect ? "Да" : "Нет");
    }

    sheet.getRange(row, markerCol, 1, logRow.length).setValues([logRow]);

    let newLevel = level;
    const isAssault = step === level + 1;
    if ((runStatus === "win" || runStatus === "all_clear") && isAssault) {
      newLevel = level + 1;
    }

    sheet.getRange(row, 8, 1, 3).setValues([[newLevel, 0, 0]]);

    _updateTaskStats(pairs);

    let resultStatus = runStatus;
    if ((runStatus === "win" || runStatus === "all_clear") && newLevel >= MAX_STEPS) {
      resultStatus = "all_clear";
    }

    return {
      success: true,
      runStatus: resultStatus,
      currentRunStats: _buildRunStats(newLevel, 1, 0, 0, step)
    };
  } catch (e) {
    return { success: false, message: "Ошибка сохранения сессии: " + e.message };
  }
}

/**
 * Обновление счётчиков Показы (F) и Решено (G) на листе «Задачи»
 * Сложность (H) — формула в ячейках таблицы.
 * @param {Object[]} pairs - [{taskId, isCorrect}, ...]
 */
function _updateTaskStats(pairs) {
  if (!pairs || pairs.length === 0) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Задачи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, TASK_COL_SOLVED).getValues();
  const rowById = {};

  for (let i = 0; i < data.length; i++) {
    const id = data[i][0].toString().trim();
    if (id) rowById[id] = i;
  }

  const increments = {};
  for (let i = 0; i < pairs.length; i++) {
    const taskId = pairs[i].taskId.toString().trim();
    if (!taskId) continue;
    if (!increments[taskId]) {
      increments[taskId] = { shows: 0, solved: 0 };
    }
    increments[taskId].shows++;
    if (pairs[i].isCorrect) {
      increments[taskId].solved++;
    }
  }

  for (const taskId in increments) {
    const rowIdx = rowById[taskId];
    if (rowIdx === undefined) continue;

    const inc = increments[taskId];
    let shows = parseInt(data[rowIdx][TASK_COL_SHOWS - 1], 10) || 0;
    let solved = parseInt(data[rowIdx][TASK_COL_SOLVED - 1], 10) || 0;

    shows += inc.shows;
    solved += inc.solved;

    sheet.getRange(rowIdx + 2, TASK_COL_SHOWS, 1, 2).setValues([[shows, solved]]);
  }
}

/**
 * Следующая свободная колонка лога (L…)
 */
function _getNextLogColumn(sheet, row) {
  const lastColumn = sheet.getLastColumn();
  let nextCol = LOG_START_COL;

  if (lastColumn >= LOG_START_COL) {
    const width = lastColumn - LOG_START_COL + 1;
    const logValues = sheet.getRange(row, LOG_START_COL, 1, width).getValues()[0];
    let lastFilledCol = LOG_START_COL - 1;
    for (let i = 0; i < logValues.length; i++) {
      if (logValues[i] !== "" && logValues[i] !== null) {
        lastFilledCol = LOG_START_COL + i;
      }
    }
    nextCol = lastFilledCol + 1;
  }

  return nextCol;
}

/**
 * Генерация очереди из 10 задач (только в памяти, без записи в таблицу)
 * @param {number} studentLevel — H ученика
 * @param {number} selectedStep — выбранная ступень
 * @param {Object} solvedIds — map taskId → true
 */
function _buildSessionQueue(studentLevel, selectedStep, solvedIds) {
  const allTasks = _getTasksByLevel(selectedStep);

  if (allTasks.length < 7) {
    throw new Error("В базе данных меньше 7 задач для Ступени " + selectedStep);
  }

  const solved = solvedIds || {};
  const isTaken = selectedStep <= studentLevel;
  let ordered;

  if (!isTaken) {
    ordered = _shuffleArray(allTasks.slice());
  } else {
    const unsolved = [];
    const done = [];
    for (let i = 0; i < allTasks.length; i++) {
      const task = allTasks[i];
      if (solved[task.id]) {
        done.push(task);
      } else {
        unsolved.push(task);
      }
    }
    ordered = _shuffleArray(unsolved).concat(_shuffleArray(done));
  }

  const picked = ordered.slice(0, Math.min(SESSION_SIZE, ordered.length));
  const queueIds = picked.map(function(t) { return t.id.toString(); });

  return {
    queueIds: queueIds,
    sessionTasks: picked.map(function(t) {
      return {
        id: t.id,
        imageUrl: t.imageUrl,
        hint: t.hint,
        answer: t.answer
      };
    })
  };
}

/**
 * Объект статистики текущей сессии для фронтенда
 */
function _buildRunStats(level, taskNumber, correctCount, errorCount, playingStep) {
  return {
    studentLevel: level,
    playingStep: playingStep,
    taskNumber: taskNumber,
    errors: errorCount,
    correctInSession: correctCount,
    remainingCorrect: Math.max(0, 7 - correctCount)
  };
}

/**
 * Все попытки ученика из горизонтального лога (L…)
 * @param {number} rowNum
 * @returns {Object[]} [{taskId, isCorrect}, ...]
 */
function _parseUserLogHistory(rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < LOG_START_COL) return [];

  const width = lastColumn - LOG_START_COL + 1;
  const values = sheet.getRange(rowNum, LOG_START_COL, 1, width).getValues()[0];
  const attempts = [];
  let i = 0;

  while (i < values.length) {
    const cell = values[i];
    if (cell === "" || cell === null) {
      i++;
      continue;
    }
    if (cell.toString().trim() !== "##") {
      i++;
      continue;
    }

    i++;
    if (i < values.length) i++;

    while (i + 1 < values.length) {
      const marker = values[i];
      if (marker === "" || marker === null) {
        i++;
        continue;
      }
      if (marker.toString().trim() === "##") break;

      const taskId = values[i].toString().trim();
      const result = values[i + 1];
      i += 2;
      if (!taskId) continue;

      attempts.push({
        taskId: taskId,
        isCorrect: result && result.toString().trim().toLowerCase() === "да"
      });
    }
  }

  return attempts;
}

/**
 * ID задач, которые ученик хотя бы раз решил верно
 */
function _getUserSolvedTaskIds(attempts) {
  const solved = {};
  for (let i = 0; i < attempts.length; i++) {
    if (attempts[i].isCorrect) {
      solved[attempts[i].taskId] = true;
    }
  }
  return solved;
}

/**
 * Статистика по темам ученика
 */
function _buildUserTopicStats(attempts) {
  const metaMap = _getTaskMetaMap();
  const byTopic = {};

  for (let i = 0; i < attempts.length; i++) {
    const taskId = attempts[i].taskId;
    const meta = metaMap[taskId];
    const topic = meta && meta.topic ? meta.topic : "Без темы";
    if (!byTopic[topic]) {
      byTopic[topic] = { attempts: 0, correct: 0 };
    }
    byTopic[topic].attempts++;
    if (attempts[i].isCorrect) {
      byTopic[topic].correct++;
    }
  }

  const topics = [];
  for (const topicName in byTopic) {
    const item = byTopic[topicName];
    topics.push({
      name: topicName,
      attempts: item.attempts,
      correct: item.correct,
      percent: item.attempts > 0 ? Math.round(item.correct / item.attempts * 100) : 0
    });
  }
  topics.sort(function(a, b) { return a.name.localeCompare(b.name, "ru"); });

  const ranked = topics.filter(function(t) { return t.attempts >= MIN_TOPIC_ATTEMPTS; });
  const topTopics = ranked.slice().sort(function(a, b) {
    return b.percent - a.percent || b.attempts - a.attempts;
  }).slice(0, 5);
  const weakTopics = ranked.slice().sort(function(a, b) {
    return a.percent - b.percent || b.attempts - a.attempts;
  }).slice(0, 5);

  return { topics: topics, topTopics: topTopics, weakTopics: weakTopics };
}

/**
 * Сводка по ступеням: доступность и нерешённые задачи
 */
function _buildStepSummaries(studentLevel, solvedIds) {
  const steps = [];
  const maxAvailable = studentLevel >= MAX_STEPS ? MAX_STEPS : studentLevel + 1;

  for (let step = 1; step <= MAX_STEPS; step++) {
    const tasks = _getTasksByLevel(step);
    let unsolvedCount = 0;
    for (let i = 0; i < tasks.length; i++) {
      if (!solvedIds[tasks[i].id]) unsolvedCount++;
    }

    steps.push({
      step: step,
      taken: step <= studentLevel,
      available: step <= maxAvailable,
      totalCount: tasks.length,
      unsolvedCount: unsolvedCount
    });
  }

  return steps;
}

/**
 * Карта метаданных задач: тема
 */
function _getTaskMetaMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Задачи");
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow <= 1) return map;

  const data = sheet.getRange(2, 1, lastRow - 1, TASK_COL_TOPIC).getValues();
  for (let i = 0; i < data.length; i++) {
    const taskId = data[i][0].toString().trim();
    if (!taskId) continue;
    map[taskId] = {
      topic: data[i][4] ? data[i][4].toString().trim() : ""
    };
  }
  return map;
}

/**
 * Перемешивание массива (Fisher–Yates)
 */
function _shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

/**
 * Поиск пользователя по ID
 */
function _getUserRowData(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, USER_DATA_COLS).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === userId.toString().trim()) {
      return { rowNum: i + 2, data: data[i] };
    }
  }
  return null;
}

/**
 * Поиск по ID, телефону или почте (для входа)
 */
function _findUserByLogin(login) {
  const key = (login || "").toString().trim();
  if (!key) return null;

  const normPhone = _normalizePhone(key);
  const normEmail = _normalizeEmail(key);
  const all = _getAllUserRows();

  for (let i = 0; i < all.length; i++) {
    const row = all[i];
    const data = row.data;
    if (data[0].toString().trim() === key) return row;
    if (normPhone && _normalizePhone(data[3]) === normPhone) return row;
    if (normEmail && _normalizeEmail(data[10]) === normEmail) return row;
  }
  return null;
}

/**
 * Поиск по телефону или почте (для восстановления)
 */
function _findUserByPhoneOrEmail(identifier) {
  const normPhone = _normalizePhone(identifier);
  const normEmail = _normalizeEmail(identifier);
  if (!normPhone && !normEmail) return null;

  const all = _getAllUserRows();
  for (let i = 0; i < all.length; i++) {
    const data = all[i].data;
    if (normPhone && _normalizePhone(data[3]) === normPhone) return all[i];
    if (normEmail && _normalizeEmail(data[10]) === normEmail) return all[i];
  }
  return null;
}

function _getAllUserRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, USER_DATA_COLS).getValues();
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({ rowNum: i + 2, data: data[i] });
  }
  return rows;
}

function _isPhoneTaken(phone) {
  const norm = _normalizePhone(phone);
  const all = _getAllUserRows();
  for (let i = 0; i < all.length; i++) {
    if (_normalizePhone(all[i].data[3]) === norm) return true;
  }
  return false;
}

function _isEmailTaken(email) {
  const norm = _normalizeEmail(email);
  const all = _getAllUserRows();
  for (let i = 0; i < all.length; i++) {
    if (_normalizeEmail(all[i].data[10]) === norm) return true;
  }
  return false;
}

function _generateNextUserId() {
  const all = _getAllUserRows();
  let maxId = 1000;
  for (let i = 0; i < all.length; i++) {
    const id = parseInt(all[i].data[0], 10);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

function _generateTempPassword() {
  return Math.floor(1000 + Math.random() * 9000).toString();
}

function _normalizePhone(phone) {
  const digits = (phone || "").toString().replace(/\D/g, "");
  if (!digits) return "";
  if (digits.length === 11 && digits.charAt(0) === "8") {
    return "7" + digits.substring(1);
  }
  return digits;
}

function _normalizeEmail(email) {
  return (email || "").toString().trim().toLowerCase();
}

function _isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function _maskEmail(email) {
  const parts = email.split("@");
  if (parts.length !== 2) return email;
  const name = parts[0];
  const masked = name.length <= 2 ? "**" : name.substring(0, 2) + "***";
  return masked + "@" + parts[1];
}

function _pwdTempKey(userId) {
  return "pwdTemp_" + userId.toString().trim();
}

function _setPasswordTempFlag(userId) {
  PropertiesService.getScriptProperties().setProperty(_pwdTempKey(userId), "1");
}

function _clearPasswordTempFlag(userId) {
  PropertiesService.getScriptProperties().deleteProperty(_pwdTempKey(userId));
}

function _isPasswordTemp(userId) {
  return PropertiesService.getScriptProperties().getProperty(_pwdTempKey(userId)) === "1";
}

function _resetAttemptsKey(userId) {
  return "reset_" + userId.toString().trim();
}

function _getResetAttempts(userId) {
  const raw = PropertiesService.getScriptProperties().getProperty(_resetAttemptsKey(userId));
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    const now = Date.now();
    return arr.filter(function(t) { return now - t < RESET_WINDOW_MS; });
  } catch (e) {
    return [];
  }
}

function _recordResetAttempt(userId) {
  const attempts = _getResetAttempts(userId);
  attempts.push(Date.now());
  PropertiesService.getScriptProperties().setProperty(
    _resetAttemptsKey(userId),
    JSON.stringify(attempts)
  );
}

function _canRequestPasswordReset(userId) {
  const now = Date.now();
  const attempts = _getResetAttempts(userId).filter(function(t) { return now - t < RESET_WINDOW_MS; });
  if (attempts.length < MAX_RESETS_IN_WINDOW) {
    return { allowed: true };
  }
  const sorted = attempts.slice().sort(function(a, b) { return a - b; });
  const unlockAt = sorted[MAX_RESETS_IN_WINDOW - 1] + RESET_WINDOW_MS;
  if (now < unlockAt) {
    return { allowed: false, waitMinutes: Math.ceil((unlockAt - now) / 60000) };
  }
  PropertiesService.getScriptProperties().deleteProperty(_resetAttemptsKey(userId));
  return { allowed: true };
}

/**
 * ВНУТРЕННИЙ СБОР ЗАДАЧ УРОВНЯ (По первой цифре ID задачи)
 */
function _getTasksByLevel(level) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Задачи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(1, 1, lastRow, TASK_COL_TOPIC).getValues();
  const tasks = [];
  const levelDigit = level.toString();

  for (let i = 1; i < data.length; i++) {
    const taskId = data[i][0].toString().trim();
    if (!taskId) continue;

    if (taskId.charAt(0) === levelDigit) {
      tasks.push({
        id: taskId,
        imageUrl: data[i][1],
        hint: data[i][2],
        answer: data[i][3],
        topic: data[i][4] ? data[i][4].toString() : ""
      });
    }
  }
  return tasks;
}

/**
 * Служебная: собрать ссылки на файлы из папки Google Drive (с именами)
 * Запускать вручную из редактора Apps Script (▶ Выполнить).
 * Результат: журнал выполнения + новый Google Doc «Ссылки на файлы» в корне Диска.
 */
function getFileLinks() {
  try {
    const FOLDER_ID = "1sbuWOiX7UiQr8P1M44lNPP4kdkW7839U";
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    const items = [];

    while (files.hasNext()) {
      const file = files.next();
      items.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }

    if (items.length === 0) {
      Logger.log("В папке нет файлов.");
      return { success: false, message: "В папке нет файлов." };
    }

    items.sort(function(a, b) {
      return a.name.localeCompare(b.name, "ru");
    });

    const lines = items.map(function(item) {
      return item.name + " | " + item.url;
    });

    Logger.log(lines.join("\n"));

    const doc = DocumentApp.create("Ссылки на файлы");
    const body = doc.getBody();
    body.appendParagraph("Папка: " + folder.getName());
    body.appendParagraph("Файлов: " + items.length);
    body.appendParagraph("");
    lines.forEach(function(line) {
      body.appendParagraph(line);
    });
    doc.saveAndClose();
    Logger.log("Документ создан: " + doc.getUrl());

    return {
      success: true,
      count: items.length,
      docUrl: doc.getUrl(),
      items: items
    };
  } catch (e) {
    Logger.log("Ошибка getFileLinks: " + e.message);
    return { success: false, message: e.message };
  }
}
