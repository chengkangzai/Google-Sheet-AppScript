function postMessageToNotification(message) {
  var msg = message;
  //Discord
  var discordUrl = '<<Discord Webhook URL>>';
  var payload = JSON.stringify({ content: msg });

  var params = {
    headers: { 'Content-Type': 'application/json' },
    method: "POST",
    payload: payload,
    muteHttpExceptions: true
  };

  var dResponse = UrlFetchApp.fetch(discordUrl, params);

  //Telegram
  //var telegramUrl = '<<';
  //var tResponse = UrlFetchApp.fetch(telegramUrl + "?chat_id=-<<telegram chat id>>&text=" + encodeURIComponent(msg));
}

function pushMessageToTeams(person, times, roomName, itemlist, link, remarks, office365UserName) {

  //https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#adaptive-card
  //Personal Chat address for ta in teams
  var taTeamsAddress = "https://teams.microsoft.com/l/chat/0/0?users=" + office365UserName + "@<<office 365 domain>>";

  //JSON file for Webhook in Teams
  var teamsJSON = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "00bfa5",
    "summary": "<<Summary of the message>>",
    "sections": [{
      "activityTitle": "<<Activity title>>",
      "activitySubtitle": times,
      "activityImage": "https://miro.medium.com/max/300/0*9U_PkckAUtKGrb_R.png",
      "facts": [{
        "name": "<<facts name 1>>",
        "value": "<<facts value 1>>"
      }],
      "markdown": true
    }],
    "potentialAction": [{
      "@type": "ActionCard",
      "name": "Actions",
      "actions": [{
        "@type": "OpenUri",
        "name": "Talk to this person",
        "targets": [{ "os": "default", "uri": taTeamsAddress }]
      }, {
        "@type": "OpenUri",
        "name": "Open the form",
        "targets": [{ "os": "default", "uri": link }]
      }]
    }]
  };

  //Production url
  var teamsUrl = '<<Teams Webhook URL>>';
  var payload = JSON.stringify(teamsJSON);

  var config = {
    headers: { 'Content-Type': 'application/json' },
    method: "POST",
    payload: payload,
    muteHttpExceptions: true
  };

  var teamsResponse = UrlFetchApp.fetch(teamsUrl, config);
}

function searchForEmail(person) {
  //https://stackoverflow.com/questions/11727975/how-to-access-data-on-different-google-spreadsheet-through-google-apps-script
  var ss = SpreadsheetApp.openById("<<spreadsheet id for storing email>>");
  var sheets = ss.getSheetByName("<<sheet name for storing email>>");

  var whilecondition = true;
  while (whilecondition == true) {
    var lastTArow = 60;
    for (i = 3; i <= lastTArow; i++) {
      var matchedTA = sheets.getRange('B' + i).getValue();
      var matchedTAEmail = sheets.getRange('C' + i).getValue();
      if (matchedTA == person) {
        return matchedTAEmail;
      }
    }
  }
}

function checkStatus(timestampCol, startCol, endCol, roomCol, nameCol, itemRow, formType, remarksCol) {
  //Get Current Timestamp
  var time = new Date();
  var times = Utilities.formatDate(time, "GMT+08:00", "HH:mm:ss dd-MMM-yyyy");

  //Get Current Sheet
  var s = SpreadsheetApp.getActiveSheet();

  var spreadsheetID = "<<spreadsheet id>>";  //Spreadsheet ID *Remember to change it when doing backup or copy to edit
  var link = "https://docs.google.com/spreadsheets/d/" + spreadsheetID;

  //Set Variables
  var curr = s.getActiveCell();  //Find which cell the user is key in right now
  var RowNum = curr.getRow();
  var ColNum = curr.getColumn();
  var timestampCell = s.getRange((RowNum), (timestampCol));
  var timestampFlag = false;

  for (var i = startCol; i <= endCol; i++) {
    var CurrCell = s.getRange((RowNum), (i));

    //Check if Any Cell is Empty
    if (CurrCell.getValue().toString() == '' && CurrCell.getBackground() != '#bdbdbd' && CurrCell.getBackground() != '#ffffff') { //If Current Cell is Empty and isn't Grey or White(merged)
      timestampFlag = false;
      break;
    } else {
      timestampFlag = true;
    }
  }

  //After Checking Cell
  if (RowNum > itemRow) {
    if (timestampFlag == false) {
      timestampCell.clearContent();
    } else {
      timestampCell.setValue(times);
    }
  }

  Utilities.sleep(5000);

  if (timestampCell.getValue() != "") {
    var location = s.getRange((RowNum), roomCol).getValue();
    var person = s.getRange((RowNum), nameCol).getValue();
    var remarks = s.getRange((RowNum), remarksCol).getValue();

    var msg = "";
    var item = "";
    var itemlist = [];
    var itemFlag = false;

    for (var i = startCol; i <= endCol; i++) {
      var CurrCell = s.getRange((RowNum), (i)).getValue();

      //Check if Cell got X or Non-Readable
      if (CurrCell == "X" || CurrCell == "Non-Readable") {

        //Check Column Name
        if (formType == 2) {
          var chkMerge = s.getRange((itemRow), (i));
          if (chkMerge.isPartOfMerge() == true && chkMerge.getValue() == "") {
            item = s.getRange((itemRow - 1), (i)).getValue();
            itemlist.push(item);
            itemFlag = true;
          } else if (chkMerge.isPartOfMerge() == false && chkMerge.getValue() != "") {
            if (s.getRange((itemRow - 1), (i)).getValue() == "") {
              for (var a = 1; a < 4; a++) {
                var ColName = s.getRange((itemRow - 1), (i - a)).getValue();
                if (ColName != "") {
                  item = ColName + " (" + chkMerge.getValue() + ")";
                  itemlist.push(item);
                  itemFlag = true;
                }
                break;
              }
            } else {
              item = s.getRange((itemRow - 1), (i)).getValue() + " (" + chkMerge.getValue() + ")";
              itemlist.push(item);
              itemFlag = true;
            }
          } else {
            item = "Unknown Item";
            itemlist.push(item);
            itemFlag = true;
          }
        }
        else if (formType = 1) {
          item = s.getRange((itemRow), (i)).getValue();
          itemlist.push(item);
          itemFlag = true;
        }
      }
    }

    if (itemFlag == true) {
      //Message Content
      var msgArr = [
        location + " seem got some issues (" + itemlist + "). ",
      ];
      msgArr = msgArr[Math.floor(Math.random() * msgArr.length)];
      msg = msgArr + "\n" + "Reported by: " + person + " at " + times + "\n" + "Link (" + s.getName() + " | Line " + RowNum + " | Remark: \"" + remarks + "\" ): " + link;
      var office365UserName = searchForEmail(person);
      pushMessageToTeams(person, times, location, itemlist, link, remarks, office365UserName);
      postMessageToNotification(msg);
    }
  }
}

function onEdit(e) {
  //Get Current Sheet
  var s = SpreadsheetApp.getActiveSheet();

  if (s.getName() == "<<type A of Form>") {
    if (e.OldValue == e.Value) {
      checkStatus(30, 2, 28, 1, 2, 3, 2, 29);
    }
  }
  else if (s.getName() == "<<Type B of Form>>") {
    if (e.OldValue == e.Value) {
      checkStatus(17, 2, 15, 1, 2, 2, 1, 16);
    }
  }
}