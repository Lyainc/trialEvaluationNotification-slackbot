function sendToSheet() {
  var timeTableSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시용평가 일정");
  var headers = getHeaders(timeTableSheet);
  var values = getValues(timeTableSheet);
  var dateColumns = getDateColumns(headers);
  var today = new Date();
  var threeDaysLater = new Date(today);
  threeDaysLater.setDate(today.getDate() + 3); // 오늘로부터 3일 후

  var [groupMessages, privateMessages] = generateMessages(values, headers, dateColumns, today, threeDaysLater);

  var groupUserIds = getUserIdsForMessages(values, headers, today, threeDaysLater);
  var privateUserIds = getUserIdsForMessages(values, headers, today, threeDaysLater);
  
  // 테스트용 코드
  // appendMessages(timeTableSheet, groupMessages);
  // appendMessages(timeTableSheet, privateMessages);

  sendDirectMessages(groupMessages, groupUserIds);
  sendDirectMessages(privateMessages, privateUserIds);
}

function sendDirectMessages(messages, userIds) {
  messages.forEach(message => {
    userIds.forEach(userId => {
      sendDirectMessageToUser(message, userId);
    });
  });
}

function sendDirectMessageToUser(message, userId) {
  var slackUrl = "https://slack.com/api/chat.postMessage";
  var token = "Bearer MY_SLACK_APP_TOKEN;
  var payload = {
    "channel": userId,
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": message
        }
      }
    ]
  };

  var options = {
    "method": "post",
    "contentType": "application/json; charset=utf-8",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true,
    'headers': {
      'Authorization': token
    }
  };

  try {
    var response = UrlFetchApp.fetch(slackUrl, options);
    Logger.log(response.getContentText()); // 응답 확인 (디버깅용)
  } catch (error) {
    Logger.log("Error occurred while sending message to user: " + error);
  }
}

function getHeaders(timeTableSheet) {
  var headerRange = timeTableSheet.getRange(2, 1, 1, timeTableSheet.getLastColumn());
  var headers = headerRange.getValues()[0];
  return headers;
}

function getValues(timeTableSheet) {
  var dataRange = timeTableSheet.getDataRange(); // 시트에서 모든 데이터 범위를 가져옵니다.
  var values = dataRange.getValues(); // 데이터 범위에서 값들을 가져옵니다.
  return values.slice(2); // 2번째 행부터 가져오도록 수정합니다. (B3부터)
}

function getDateColumns(headers) {
  return headers.map((header, index) => isDateColumn(header) ? index + 1 : undefined).filter(column => column !== undefined);
}

function isDateColumn(header) {
  return /일|날짜/.test(header);
}

function generateMessages(values, headers, dateColumns, today, threeDaysLater) {
  var groupMessages = [];
  var privateMessages = [];

  values.forEach(row => {
    var name = row[1]; // 이름은 values 배열의 첫 번째 요소로 가정

    dateColumns.forEach(columnIndex => {
      var header = headers[columnIndex]; // 인덱스 수정
      var date = new Date(row[columnIndex]);

      // privateMessages 조건: 셀프 리뷰 작성 마감일만 체크
      if (header === "셀프 리뷰 작성 마감일") {
        if (isSameDay(date, today)) {
          privateMessages.push(`\n${name}님! 셀프 리뷰 작성 마감일이 오늘입니다.\n`);
        } else if (isSameDay(date, threeDaysLater)) {
          privateMessages.push(`\n${name}님! 셀프 리뷰 작성 마감이 3일 남았습니다.\n`);
        }
      } 
      // groupMessages 조건: 모든 header에 대해 당일과 3일 뒤의 날짜를 체크
      if (isSameDay(date, threeDaysLater)) {
        groupMessages.push(`\n${name}님의 ${header}이 3일 남았습니다.\n`);
      }
      if (isSameDay(date, today)) {
        groupMessages.push(`\n오늘은 ${name}님의 ${header}입니다.\n`);
      }
    });
  });

  return [groupMessages, privateMessages];
}

function getUserIdsForMessages(values, headers, today, threeDaysLater) {
  var userIds = [];
  var teamLeadIndex = headers.indexOf("팀리드 userId");
  var chapterLeadIndex = headers.indexOf("챕터리드 userId");

  values.forEach(row => {
    var teamLeadUserId = row[teamLeadIndex];
    var chapterLeadUserId = row[chapterLeadIndex];

    headers.forEach((header, index) => {
      var date = new Date(row[index]);
      if (isSameDay(date, today) || isSameDay(date, threeDaysLater)) {
        if (teamLeadUserId && !userIds.includes(teamLeadUserId)) userIds.push(teamLeadUserId);
        if (chapterLeadUserId && !userIds.includes(chapterLeadUserId)) userIds.push(chapterLeadUserId);
      }
    });
  });
  var peopleTeamUserId = ["PEOPLE_TEAM_USER_ID"]
  userIds.push(...peopleTeamUserId);
  userIds = Array.from(new Set(userIds));
  
  return userIds;
}

// 테스트용 코드
// function appendMessages(sheet, messages) {
//   if (messages.length > 0) {
//     var lastRow = sheet.getLastRow();
//     var targetRange = sheet.getRange(lastRow + 1, 1, messages.length, 1);
//     targetRange.setValues(messages.map(message => [message]));
//   }
// }

function isSameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() && date1.getMonth() === date2.getMonth() && date1.getDate() === date2.getDate();
}