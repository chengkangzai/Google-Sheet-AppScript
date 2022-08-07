function isDivideAbleBy(input) {
  return getTotalCountDay() % input == 0;
}

function isTheDateOf(date) {
  var ss = SpreadsheetApp
    .openById("<<spreadsheetID>")
    .getSheetByName("<<sheetName>>")
    .getRange('D1')
    .getValue();
  return Utilities.formatDate(new Date(ss.toString()), "GMT+8", "dd/MM") == date;
}

function getTotalCountDay() {
  return parseFloat(SpreadsheetApp.openById("<<spreadsheetID>")
    .getSheetByName("<<sheetName>>")
    .getRange('B2')
    .getValue().toString());
}

function getFormattedCountDate() {
  var sheets = SpreadsheetApp
    .openById("<<spreadsheetID>")
    .getSheetByName("<<sheetName>>");
  //Date Value
  var date = {
    year: sheets.getRange('B4').getValue(),
    month: sheets.getRange('B3').getValue(),
    day: sheets.getRange('D3').getValue()
  };
  return date;
}

function getWordForDiff() {
  var date = getFormattedCountDate();
  year = date.year;
  month = date.month;
  day = date.day;
  var dateInword = `${year}年${month}个月又${day}天`;
  if (month == 0 && day == 0) {
    //Should be catch before it reach here !
    //10/01/2021
    dateInword = `${year}年`;
  } else if (month == 0) {
    //11/01/2021
    dateInword = `${year}年又${day}天`;
  } else if (day == 0) {
    //10/02/2021
    dateInword = `${year}年${month}个月`;
  } else if (year == 0 && month == 0) {
    //less then a month
    dateInword = `${day}天`;
  } else if (year == 0 && day == 0) {
    //exactly month
    dateInword = `${month}个月`;
  } else if (year == 0) {
    dateInword = `${month}个月又${day}天`;
  }
  return dateInword;
}

function getQuote() {
  var totalDay = getTotalCountDay();


  function getPrependWord() {
    var msg = [
      "<<prefix word>>",
    ]
    return msg[Math.floor(Math.random() * msg.length)] + getWordForDiff();
  }

  var msgArr = [
    "<<prefix word>>" + getWordForDiff() + "<<suffix word>>",
    getPrependWord() + "<<suffix word>>",
    
  ];

  return msgArr[Math.floor(Math.random() * msgArr.length)];
}

// Main
function wordMessage() {
  //the importance are by low to high ... the bottom one should be most important word
  var msg = getQuote();

  if (isDivideAbleBy(100)) {
    msg = "<<100th day message>>";
  }

  if (isTheDateOf("14/02")) {
    msg = "<<14/02 message>>";
  }

  sendToDiscord(msg);
}

function sendToDiscord(msg) {
  var payload = {
    "username": "<<bot name>>",
    "embeds": [{ "title": msg }]
  };
  Logger.log(payload);
  var discordUrl = "<<discord webhook url>>";
  var params = {
    headers: { 'Content-Type': 'application/json' },
    method: "POST",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var send = UrlFetchApp.fetch(discordUrl, params);
  if (send.getContentText() == null) console.log(send.getContentText());
}