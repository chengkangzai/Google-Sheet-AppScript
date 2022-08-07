function pushMessageToTeams(type, date, time, module) {
    //https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#adaptive-card
    var teamsJSON = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "00bfa5",
        "summary": type + " BOT",
        "sections": [{
            "activityTitle": type + " BOT",
            "activitySubtitle": date,
            "activityImage": "https://miro.medium.com/max/300/0*9U_PkckAUtKGrb_R.png",
            "facts": [{
                "name": "Type",
                "value": type
            }, {
                "name": "Date",
                "value": date
            }, {
                "name": "Time",
                "value": time
            }, {
                "name": "Module",
                "value": module
            }, {
                "name": "type",
                "value": type
            }],
            "markdown": true
        }]
    };

    //Devevelopment url
    var teamsUrl = '<<Teams Webhook URL>>';

    var payload = JSON.stringify(teamsJSON);

    var config = {
        headers: {
            'Content-Type': 'application/json'
        },
        method: "POST",
        payload: payload,
        muteHttpExceptions: true
    };

    var teamsResponse = UrlFetchApp.fetch(teamsUrl, config);
    Logger.log(teamsResponse.getContentText());
}

function processInformation() {
    var ss = SpreadsheetApp.openById("<<Spreadsheet ID>>");
    var sheets = ss.getSheetByName("<<Sheet Name>>");

    var lastRow = 31;
    for (var i = 1; i <= lastRow; i++) {
        Logger.log("Searching")
        var type = sheets.getRange('A' + i).getValue();
        var date = sheets.getRange('B' + i).getValue();
        date = Utilities.formatDate(date, "GMT+08:00", "dd-MMM-yyyy");
        var time = sheets.getRange('C' + i).getValue();
        time = Utilities.formatDate(time, "GMT+08:00", "HH:mm");
        newDate = date + " " + time;
        var module = sheets.getRange('D' + i).getValue();

        Utilities.sleep(5000);

        Logger.log(type, date, time, module);
        if (type !== "") {
            pushMessageToTeams(type, date, time, module);
            Utilities.sleep(5000);
        } else if (type == "") {
            lastRow = i;
        }
    }
}

function getDate() {
    var date = new Date();
    date = Utilities.formatDate(date, "GMT+08:00", "u");
    processInformation();
}