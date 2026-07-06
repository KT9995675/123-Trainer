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
      user: {
        id: data[0].toString(),
        name: data[1].toString(),
        rowNum: userRow.rowNum
      }
    };
  } catch (e) {
    return { success: false, message: "Ошибка авторизации на сервере: " + e.message };
  }
}

/**
 * Получение задач для новой сессии (только чтение уровня, без записи в лог)
 * @param {string} userId
 * @param {boolean} startNewSession — оставлен для совместимости с клиентом
 */
function getTaskForUser(userId, startNewSession) {
  try {
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    let level = parseInt(userRow.data[7]);
    if (isNaN(level)) level = 0;

    if (level >= 5) {
      return { success: true, allCompleted: true };
    }

    const queueData = _buildSessionQueue(level);

    return {
      success: true,
      allCompleted: false,
      sessionTasks: queueData.sessionTasks,
      queueIds: queueData.queueIds,
      sessionIndex: 0,
      currentRunStats: _buildRunStats(level, 1, 0, 0)
    };
  } catch (e) {
    return { success: false, message: "Ошибка получения задачи: " + e.message };
  }
}

/**
 * Запись завершённой сессии в лог одним блоком (победа или поражение)
 * @param {string} userId
 * @param {number} rowNum — номер строки ученика (с loginUser)
 * @param {string[]} queueIds — очередь из 10 ID задач
 * @param {Object[]} pairs — [{taskId, isCorrect}, ...]
 * @param {string} runStatus — "win" | "fail" | "all_clear"
 */
function finalizeSession(userId, rowNum, queueIds, pairs, runStatus) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    let row = parseInt(rowNum, 10);

    if (isNaN(row) || row < 2) {
      const userRow = _getUserRowData(userId);
      if (!userRow) return { success: false, message: "Пользователь не найден." };
      row = userRow.rowNum;
    }

    const userData = sheet.getRange(row, 8, 1, 3).getValues()[0];
    let level = parseInt(userData[0]);
    if (isNaN(level)) level = 0;

    let correctCount = 0;
    let errorCount = 0;
    for (let i = 0; i < pairs.length; i++) {
      if (pairs[i].isCorrect) {
        correctCount++;
      } else {
        errorCount++;
      }
    }

    if (runStatus === "win" || runStatus === "all_clear") {
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
    if (runStatus === "win" || runStatus === "all_clear") {
      newLevel = level + 1;
    }

    sheet.getRange(row, 8, 1, 3).setValues([[newLevel, 0, 0]]);

    let resultStatus = runStatus;
    if ((runStatus === "win" || runStatus === "all_clear") && newLevel >= 5) {
      resultStatus = "all_clear";
    }

    return {
      success: true,
      runStatus: resultStatus,
      currentRunStats: _buildRunStats(newLevel, 1, 0, 0)
    };
  } catch (e) {
    return { success: false, message: "Ошибка сохранения сессии: " + e.message };
  }
}

/**
 * Следующая свободная колонка лога (K…)
 */
function _getNextLogColumn(sheet, row) {
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

  return nextCol;
}

/**
 * Генерация очереди из 10 задач (только в памяти, без записи в таблицу)
 */
function _buildSessionQueue(level) {
  const SESSION_SIZE = 10;
  const targetLevel = level + 1;
  const allTasks = _getTasksByLevel(targetLevel);

  if (allTasks.length < 7) {
    throw new Error("В базе данных меньше 7 задач для Ступени " + targetLevel);
  }

  const shuffled = allTasks.slice().sort(function() { return Math.random() - 0.5; });
  const picked = shuffled.slice(0, Math.min(SESSION_SIZE, shuffled.length));
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
