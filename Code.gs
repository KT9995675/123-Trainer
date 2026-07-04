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
 */
function getTaskForUser(userId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Пользователи");
    const userRow = _getUserRowData(userId);
    if (!userRow) return { success: false, message: "Пользователь не найден." };

    const row = userRow.rowNum;
    let level = parseInt(userRow.data[7]); // Колонка H (8) - Текущая ступень
    if (isNaN(level)) level = 0;

    let верно = parseInt(userRow.data[8]) || 0;   // Колонка I (9)
    let ошибки = parseInt(userRow.data[9]) || 0; // Колонка J (10)

    // Если все 5 ступеней уже пройдены (ученик получил 5-ю ступень)
    if (level >= 5) {
      return { success: true, allCompleted: true };
    }

    // Читаем живой лог пар начиная с 11-й колонки (K)
    const lastColumn = sheet.getLastColumn();
    let logValues = [];
    if (lastColumn >= 11) {
      logValues = sheet.getRange(row, 11, 1, lastColumn - 10).getValues()[0];
    }

    let completedPairsCount = 0;
    let activeTaskId = null;
    const usedTaskIds = [];

    // Парсим лог парами: [ID задачи, Вердикт]
    for (let i = 0; i < logValues.length; i += 2) {
      let tId = logValues[i] ? logValues[i].toString().trim() : "";
      let verdict = logValues[i+1] ? logValues[i+1].toString().trim() : "";

      if (tId === "") break; // Лог пуст/закончился

      usedTaskIds.push(tId);

      if (verdict !== "") {
        completedPairsCount++;
      } else {
        // Есть ID задачи, но нет вердикта — это текущая активная задача сессии
        activeTaskId = tId;
        break;
      }
    }

    // Если школьник зашел с накопленным критическим числом ошибок (страховка)
    if (ошибки >= 3) {
      _clearSessionProgress(sheet, row);
      return getTaskForUser(userId);
    }

    // Текущий порядковый номер задачи в этой сессии
    const taskNumber = completedPairsCount + 1;

    let currentTaskId = activeTaskId;

    // Если активной задачи нет в логе (первый запуск или прошлая задача решена), генерируем новую
    if (!currentTaskId) {
      // Номер искомой ступени в базе = уровень + 1 (при уровне 0 решаем задачи 1-й ступени)
      const targetLevel = level + 1;
      const allTasks = _getTasksByLevel(targetLevel);

      // Исключаем задачи, которые уже мелькали в этой сессии
      const availableTasks = allTasks.filter(function(t) {
        return usedTaskIds.indexOf(t.id.toString().trim()) === -1;
      });

      if (availableTasks.length === 0) {
        return { success: false, message: "В базе данных нет доступных задач для Ступени " + targetLevel };
      }

      // Выбираем случайную задачу из оставшихся
      const randomTask = availableTasks[Math.floor(Math.random() * availableTasks.length)];
      currentTaskId = randomTask.id;

      // Фиксируем ID новой задачи в лог (ячейка под ID текущей пары)
      const nextCellCol = 11 + (completedPairsCount * 2);
      sheet.getRange(row, nextCellCol).setValue(currentTaskId);
    }

    // Ищем данные карточки задачи в базе
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
      currentRunStats: {
        studentLevel: level,
        taskNumber: taskNumber,
        errors: ошибки
      }
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

    // Ищем задачу для сверки ответа
    const targetLevel = level + 1;
    const allTasks = _getTasksByLevel(targetLevel);
    const taskObj = allTasks.find(function(t) { return t.id.toString() === taskId.toString(); });
    if (!taskObj) return { success: false, message: "Задача не найдена в базе данных." };

    const correctAnswer = taskObj.answer.toString().trim();
    const userAnswer = answer.toString().trim();

    // Сверка без учета регистра
    const isCorrect = userAnswer.toLowerCase() === correctAnswer.toLowerCase();

    if (isCorrect) {
      верно += 1;
    } else {
      ошибки += 1;
    }

    // Ищем в логе ячейку вердикта для текущей задачи (где ID совпал, а вердикт пуст)
    const lastColumn = sheet.getLastColumn();
    let logValues = [];
    if (lastColumn >= 11) {
      logValues = sheet.getRange(row, 11, 1, lastColumn - 10).getValues()[0];
    }

    let verdictCol = -1;
    let completedPairsCount = 0;

    for (let i = 0; i < logValues.length; i += 2) {
      let tId = logValues[i] ? logValues[i].toString().trim() : "";
      let verdict = logValues[i+1] ? logValues[i+1].toString().trim() : "";

      if (tId === taskId.toString().trim() && verdict === "") {
        verdictCol = 11 + i + 1;
        break;
      }
      if (verdict !== "") {
        completedPairsCount++;
      }
    }

    // Если по какой-то причине пустая ячейка под вердикт не найдена циклом
    if (verdictCol === -1) {
      let filledCells = 0;
      for (let i = 0; i < logValues.length; i++) {
        if (logValues[i] !== "") filledCells++;
        else break;
      }
      verdictCol = 11 + filledCells + (filledCells % 2 === 0 ? 1 : 0);
    }

    // Записываем вердикт в таблицу
    sheet.getRange(row, verdictCol).setValue(isCorrect ? "Да" : "Нет");
    completedPairsCount++;

    let runStatus = "continue"; // Статусы: continue, win, fail, all_clear

    if (ошибки >= 3) {
      runStatus = "fail";
      верно = 0;
      ошибки = 0;
      _clearSessionProgress(sheet, row);
    } else if (completedPairsCount >= 7) {
      // Успешно решили 7 задач сессии с допустимым количеством ошибок
      level += 1;
      верно = 0;
      ошибки = 0;
      _clearSessionProgress(sheet, row);

      if (level >= 5) {
        runStatus = "all_clear";
      } else {
        runStatus = "win";
      }
    }

    // Сохраняем обновленные статусы обратно в таблицу (колонки H, I, J)
    sheet.getRange(row, 8).setValue(level);
    sheet.getRange(row, 9).setValue(верно);
    sheet.getRange(row, 10).setValue(ошибки);

    const updatedStats = {
      studentLevel: level,
      taskNumber: (runStatus === "continue") ? (completedPairsCount + 1) : 1,
      errors: ошибки
    };

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
 * Вспомогательная функция: Полная очистка прогресса текущей сессии (колонки I, J и лог с K)
 */
function _clearSessionProgress(sheet, row) {
  sheet.getRange(row, 9).setValue(0);  // Обнуляем Количество верных (I)
  sheet.getRange(row, 10).setValue(0); // Обнуляем Количество ошибок (J)

  const currentLastCol = sheet.getLastColumn();
  if (currentLastCol >= 11) {
    // Чистим лог пар начиная с колонки K до самого конца заполненных данных
    sheet.getRange(row, 11, 1, currentLastCol - 10).clearContent();
  }
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
