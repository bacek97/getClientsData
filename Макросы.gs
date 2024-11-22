/** @OnlyCurrentDoc */
// https://script.google.com/home/projects/1_YbzqPPpXRj9nk3cuyYpYU4RWrr7LOZHwNH_ZSd2L1fAdYjdtBe3qO1N/edit
// https://docs.google.com/spreadsheets/d/16hsYtMQ1pzwmshBlbMwImspjZPJuu06hmOOVe7xdc5o/edit?usp=sharing


function getClientsData() {
  var ui = SpreadsheetApp.getUi();

  var token = handleAuthentication(ui);

  if (token) {
    fetchAndWriteClientsData(token);
  }
}

function handleAuthentication(ui) {
  var response = ui.alert('Вы зарегистрированы?', ui.ButtonSet.YES_NO);
  var token = '';

  if (response === ui.Button.NO) {
    registerUser(ui);
  }

  token = loginUser(ui);

  if (!token) {
    throw new Error('Не удалось авторизоваться.');
  }

  return token;
}

function registerUser(ui) {
  var response = ui.prompt('Регистрация', 'Введите имя пользователя:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var username = response.getResponseText();
    var data = { username: username };
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(data),
      muteHttpExceptions: true,
    };

    var code = UrlFetchApp.fetch('http://94.103.91.4:5000/auth/registration', options).getResponseCode();
    if (code == 400) {
      ui.alert('Пользователь уже существует', ui.ButtonSet.OK);
    } else if (code == 201) {
      ui.alert('Пользователь успешно зарегистрирован', ui.ButtonSet.OK);
    } else {
      throw new Error('Ошибка регистрации: ' + code);
    }
  }
}

function loginUser(ui) {
  var token = ''
  var response = ui.prompt('Авторизация', 'Введите имя пользователя:', ui.ButtonSet.OK);
  if (response.getSelectedButton() == ui.Button.OK) {
    var username = response.getResponseText();
    var data = { username: username };
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(data),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch('http://94.103.91.4:5000/auth/login', options);
    var code = response.getResponseCode();
    if (code == 201) {
      var responseData = JSON.parse(response.getContentText());
      if (responseData.token) {
        token = responseData.token;
      }
    } else if (code == 401) {
      ui.alert('Пользователь не найден', ui.ButtonSet.OK);
    } else {
      throw new Error('Ошибка авторизации: ' + code);
    }
  }
  return token;
}

function fetchAndWriteClientsData(token) {
  var limit = 800;  // max 1000
  var offset = 0;
  var offsetRows = 2;
  var headerPrinted = false;
  var sheet = SpreadsheetApp.getActive().getActiveSheet();

  while (true) {
    var clients = fetchClients(token, limit, offset);
    if (clients.length === 0) break;

    var statuses = fetchStatuses(token, clients.map(c => c.id));
    var mergedData = mergeUsersAndStatuses(clients, statuses);

    if (!headerPrinted) {
      writeToSheet(sheet, [Object.keys(mergedData[0])], 1, 1);
      headerPrinted = true;
    }

    writeToSheet(sheet, mergedData.map(Object.values), offsetRows, 1);

    offset += limit;
    offsetRows += clients.length;

    if (clients.length < limit) break;
  }
}

function fetchClients(token, limit, offset) {
  var options = {
    method: 'get',
    contentType: 'application/json',
    headers: { Authorization: token },
  };
  var response = UrlFetchApp.fetch(`http://94.103.91.4:5000/clients?limit=${limit}&offset=${offset}`, options);
  return JSON.parse(response.getContentText());
}

function fetchStatuses(token, userIds) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: token },
    payload: JSON.stringify({ userIds: userIds }),
  };
  var response = UrlFetchApp.fetch('http://94.103.91.4:5000/clients', options);
  return JSON.parse(response.getContentText());
}

function mergeUsersAndStatuses(users, statuses) {
  return users.map(user => {
    const status = statuses.find(s => s.id === user.id)?.status || '';
    return { ...user, status };
  });
}

function writeToSheet(sheet, data, startRow, startColumn) {
  if (data.length !== 0) {
    sheet.getRange(startRow, startColumn, data.length, data[0].length).setValues(data);
  }
}
