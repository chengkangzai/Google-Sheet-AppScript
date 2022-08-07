var ss = SpreadsheetApp.openById("<<source spreadsheet id>>");
var sheets = ss.getSheetByName("<<Sheet name>>");

function getInfo(cell) {
    return sheets.getRange(cell).getValue();
}

function getRecordLength() {
    var lastRow = 40;
    var i = 3;
    for (i = i; i <= lastRow; i++) {
        if (getInfo('A' + i) == "") {
            lastRow = i;
        }
    }
    return i - 4;
}

function processData() {
    for (let i = 3; i <= getRecordLength(); i++) {
        if (getInfo('G' + i) == "1") {
            BITE_EXAM_sendWarning(formatMessage(i));
            Utilities.sleep(500);
        }
    }
}

function formatMessage(row) {
    var course = getInfo('A' + row);
    var task = getInfo('B' + row);
    var desc = getInfo('C' + row);
    var note = getInfo('D' + row);
    var when = Utilities.formatDate(getInfo('F' + row), "GMT+8", "dd/MM");

    var message = [];
    message.push(`Wild Alert !!!!  :alarm_clock:  :alarm_clock:  \nThere is one Task due Tomorrow :3 (${when})\n\n  Course : ${course} \n  Task : ${task} \n  Desc : ${desc} \n  Note : ${note} \n\n GLHF`);

    return message[Math.floor(Math.random() * message.length)];
}

function BITE_EXAM_sendWarning(msg) {
    var payload = {
        "username": "Reminder BOT",
        "embeds": [{
            "title": msg,
            "color": "14177041"
        }]
    };

    var payload = JSON.stringify(payload);

    var discordUrl = "<<Discord Webhook URL>>";
    var params = {
        headers: { 'Content-Type': 'application/json' },
        method: "POST",
        payload: payload,
        muteHttpExceptions: true
    };

    var send = UrlFetchApp.fetch(discordUrl, params);
}