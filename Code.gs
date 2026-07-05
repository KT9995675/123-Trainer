/**
 * Рендеринг главной страницы веб-приложения
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Математический Тренажер')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Авторизация пользователя по Логину (ID) и Паролю
 * Структура: A=ID, B=Имя, C=Пароль
 */
function loginUser(login, pass) {
  try {
    const userRow = _getUserRowData(login);
    if (!userRow) {
      return { success: false, message: "Пользователь с таким логином не найден." };
    }

    const data = userRow.data;
    if (data[2].toString().trim() !== pass.toString().trim()) {
      return { success: false, message: "Неверный пароль. Попробуйте еще раз." };
    }

    return {
      success: true,
      user: { id: data[0].toString(), name: data[1].toString() }
    };
  } catch (e) {
    return { success: false, message: "Ошибка авторизации на сервере: " + e.message };
  }
}

/**
 * Получение текущей задачи для ученика
 * @param {string} userId
 * @param {boolean} startNewSession — true при первом запросе после входа (новая сессия)
 */
function getTaskForUser(userId, startNewSession) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    const row = userRow.rowNum;
    let level = parseInt(userRow.data[7]); // Колонка H (8) - Текущая ступень
    if (isNaN(level)) level = 0;

    if (startNewSession) {
      _resetSessionCounters(sheet, row);
      userRow.data[8] = 0;
      userRow.data[9] = 0;
    }

    let верно = parseInt(userRow.data[8]) || 0;   // Колонка I (9)
    let ошибки = parseInt(userRow.data[9]) || 0; // Колонка J (10)

    // Если все 5 ступеней уже пройдены (ученик получил 5-ю ступень)
    if (level >= 5) {
      return { success: true, allCompleted: true };
    }

    // Если школьник зашел с накопленным критическим числом ошибок (страховка)
    if (ошибки >= 3) {
      _resetSessionCounters(sheet, row);
      верно = 0;
      ошибки = 0;
    }

    const session = _parseCurrentSessionLog(sheet, row);
    let completedPairsCount = session.completedPairsCount;
    let activeTaskId = session.activeTaskId;
    const usedTaskIds = session.usedTaskIds;

    // Текущий порядковый номер попытки в этой сессии
    const taskNumber = completedPairsCount + 1;

    let currentTaskId = activeTaskId;

    // Если активной задачи нет в логе, генерируем новую
    if (!currentTaskId) {
      const targetLevel = level + 1;
      const allTasks = _getTasksByLevel(targetLevel);

      const availableTasks = allTasks.filter(function(t) {
        return usedTaskIds.indexOf(t.id.toString().trim()) === -1;
      });

      if (availableTasks.length === 0) {
        return { success: false, message: "В базе данных нет доступных задач для Ступени " + targetLevel };
      }

      const randomTask = availableTasks[Math.floor(Math.random() * availableTasks.length)];
      currentTaskId = randomTask.id;

      const nextCellCol = session.sessionStartCol + (completedPairsCount * 2);
      sheet.getRange(row, nextCellCol).setValue(currentTaskId);
    }

    const targetLevel = level + 1;
    const allTasks = _getTasksByLevel(targetLevel);
    const taskObj = allTasks.find(function(t) { return t.id.toString() === currentTaskId.toString(); });

    if (!taskObj) {
      return { success: false, message: "Задача с ID " + currentTaskId + " отсутствует во вкладке 'Задачи'." };
    }

    return {
      success: true,
      allCompleted: false,
      task: {
        id: taskObj.id,
        imageUrl: taskObj.imageUrl,
        hint: taskObj.hint
      },
      currentRunStats: _buildRunStats(level, taskNumber, верно, ошибки)
    };
  } catch (e) {
    return { success: false, message: "Ошибка получения задачи: " + e.message };
  }
}

/**
 * Проверка ответа и заполнение вердикта в логе
 */
function submitAnswer(userId, taskId, answer) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    const row = userRow.rowNum;
    let level = parseInt(userRow.data[7]);
    if (isNaN(level)) level = 0;

    let верно = parseInt(userRow.data[8]) || 0;
    let ошибки = parseInt(userRow.data[9]) || 0;

    const targetLevel = level + 1;
    const allTasks = _getTasksByLevel(targetLevel);
    const taskObj = allTasks.find(function(t) { return t.id.toString() === taskId.toString(); });
    if (!taskObj) return { success: false, message: "Задача не найдена в базе данных." };

    const correctAnswer = taskObj.answer.toString().trim();
    const userAnswer = answer.toString().trim();
    const isCorrect = userAnswer.toLowerCase() === correctAnswer.toLowerCase();

    if (isCorrect) {
      верно += 1;
    } else {
      ошибки += 1;
    }

    const session = _parseCurrentSessionLog(sheet, row);
    let verdictCol = -1;

    for (let i = 0; i < session.sessionSlice.length; i += 2) {
      const tId = session.sessionSlice[i] ? session.sessionSlice[i].toString().trim() : "";
      const verdict = session.sessionSlice[i + 1] ? session.sessionSlice[i + 1].toString().trim() : "";

      if (tId === taskId.toString().trim() && verdict === "") {
        verdictCol = session.sessionStartCol + i + 1;
        break;
      }
    }

    if (verdictCol === -1) {
      const nextCol = session.sessionStartCol + (session.completedPairsCount * 2) + 1;
      verdictCol = nextCol;
    }

    sheet.getRange(row, verdictCol).setValue(isCorrect ? "Да" : "Нет");

    const completedPairsCount = session.completedPairsCount + 1;
    let runStatus = "continue";
    let responseStats = null;

    if (ошибки >= 3) {
      runStatus = "fail";
      responseStats = _buildRunStats(level, completedPairsCount, верно, ошибки);
      _resetSessionCounters(sheet, row);
      верно = 0;
      ошибки = 0;
    } else if (верно >= 7) {
      responseStats = _buildRunStats(level, completedPairsCount, 7, ошибки);
      level += 1;
      _resetSessionCounters(sheet, row);
      верно = 0;
      ошибки = 0;

      if (level >= 5) {
        runStatus = "all_clear";
      } else {
        runStatus = "win";
      }
    }

    sheet.getRange(row, 8).setValue(level);
    sheet.getRange(row, 9).setValue(верно);
    sheet.getRange(row, 10).setValue(ошибки);

    const taskNumber = (runStatus === "continue") ? (completedPairsCount + 1) : 1;
    const updatedStats = responseStats || _buildRunStats(level, taskNumber, верно, ошибки);

    return {
      success: true,
      isCorrect: isCorrect,
      correctAnswer: correctAnswer,
      runStatus: runStatus,
      currentRunStats: updatedStats,
      stats: updatedStats
    };

  } catch (e) {
    return { success: false, message: "Ошибка верификации ответа на бэкенде: " + e.message };
  }
}

/**
 * Сброс счётчиков сессии (I, J) без стирания лога; маркер ## отделяет сессии
 */
function _resetSessionCounters(sheet, row) {
  sheet.getRange(row, 9).setValue(0);
  sheet.getRange(row, 10).setValue(0);

  const lastColumn = sheet.getLastColumn();
  let nextCol = 11;

  if (lastColumn >= 11) {
    const logValues = sheet.getRange(row, 11, 1, lastColumn - 10).getValues()[0];
    let lastFilledCol = 10;
    for (let i = 0; i < logValues.length; i++) {
      if (logValues[i] !== "" && logValues[i] !== null) {
        lastFilledCol = 11 + i;
      }
    }
    nextCol = lastFilledCol + 1;
  }

  sheet.getRange(row, nextCol).setValue("##");
}

/**
 * Разбор лога текущей сессии (после последнего маркера ##)
 */
function _parseCurrentSessionLog(sheet, row) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 11) {
    return {
      sessionStartCol: 11,
      sessionSlice: [],
      completedPairsCount: 0,
      activeTaskId: null,
      usedTaskIds: []
    };
  }

  const fullLog = sheet.getRange(row, 11, 1, lastColumn - 10).getValues()[0];
  let sessionStartIndex = 0;

  for (let i = 0; i < fullLog.length; i++) {
    if (fullLog[i].toString().trim() === "##") {
      sessionStartIndex = i + 1;
    }
  }

  const sessionStartCol = 11 + sessionStartIndex;
  const sessionSlice = fullLog.slice(sessionStartIndex);
  const usedTaskIds = [];
  let completedPairsCount = 0;
  let activeTaskId = null;

  for (let i = 0; i < sessionSlice.length; i += 2) {
    const tId = sessionSlice[i] ? sessionSlice[i].toString().trim() : "";
    const verdict = sessionSlice[i + 1] ? sessionSlice[i + 1].toString().trim() : "";

    if (tId === "" || tId === "##") break;

    usedTaskIds.push(tId);

    if (verdict !== "") {
      completedPairsCount++;
    } else {
      activeTaskId = tId;
      break;
    }
  }

  return {
    sessionStartCol: sessionStartCol,
    sessionSlice: sessionSlice,
    completedPairsCount: completedPairsCount,
    activeTaskId: activeTaskId,
    usedTaskIds: usedTaskIds
  };
}

/**
 * Объект статистики текущей сессии для фронтенда
 */
function _buildRunStats(level, taskNumber, correctCount, errorCount) {
  return {
    studentLevel: level,
    taskNumber: taskNumber,
    errors: errorCount,
    correctInSession: correctCount,
    remainingCorrect: Math.max(0, 7 - correctCount)
  };
}

/**
 * ВНУТРЕННИЙ ПОИСК ПОЛЬЗОВАТЕЛЯ (Считывает первые 10 колонок A-J)
 */
function _getUserRowData(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(1, 1, lastRow, 10).getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === userId.toString().trim()) {
      return { rowNum: i + 1, data: data[i] };
    }
  }
  return null;
}

/**
 * ВНУТРЕННИЙ СБОР ЗАДАЧ УРОВНЯ (По первой цифре ID задачи)
 */
function _getTasksByLevel(level) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Задачи");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(1, 1, lastRow, 4).getValues();
  const tasks = [];
  const levelDigit = level.toString();

  for (let i = 1; i < data.length; i++) {
    const taskId = data[i][0].toString().trim();
    if (!taskId) continue;

    // Первая цифра ID = номер ступени (101 → 1, 205 → 2)
    if (taskId.charAt(0) === levelDigit) {
      tasks.push({
        id: taskId,
        imageUrl: data[i][1],
        hint: data[i][2],
        answer: data[i][3]
      });
    }
  }
  return tasks;
}
