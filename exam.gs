var ss = SpreadsheetApp.openById("<<Spreadsheet id>>");
var sheets = ss.getSheetByName("<<Sheet name>>");

function getInfo(cell) {
    var data = sheets.getRange(cell).getValue();
    return data;
}

function processData() {
    for (let i = 3; i <= 7; i++) {
        if (getInfo('H' + i) == "1") {
            sendWarning(formatMessage(i));
            Utilities.sleep(500);
        }
    }
}

function formatMessage(row) {
    var group = getInfo('A' + row);
    var course = getInfo('B' + row);
    var event = getInfo('C' + row);
    var platform = getInfo('D' + row);
    var date = Utilities.formatDate(getInfo('E' + row), "GMT+8", "dd/MM");
    var time = Utilities.formatDate(getInfo('F' + row), "GMT+8", "h:mm a");

    var message = [];
    message.push(`
Wild Alert !!!!  :alarm_clock:  :alarm_clock:
There is an Exam Tomorrow :3 (${date})
    Time : ${time}
    Group : ${group}
    Course : ${course}
    event : ${event}
    platform : ${platform}
GLHF
    `);

    return message[Math.floor(Math.random() * message.length)];
}

function sendWarning(msg) {
    var payload = {
        "username": "Reminder BOT",
        "embeds": [{
            "title": msg,
            "color": "14177041"
        }]
    };

    var payload = JSON.stringify(payload);
    Logger.log(payload);

    var discordUrl = "<<Discord Webhook URL>>";
    var params = {
        headers: { 'Content-Type': 'application/json' },
        method: "POST",
        payload: payload,
        muteHttpExceptions: true
    };

    var send = UrlFetchApp.fetch(discordUrl, params);
    Logger.log(send.getContentText());
}