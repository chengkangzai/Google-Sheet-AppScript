var setup = {
  apu: {
    spreadsheetID: "<<SpreadSheet ID>>",

    class: {
      sheetName: "<<Sheet Name>>",
      col_roomName: "A",
      col_deviceModel: "R",
      col_deviceIP: "S"
    },

    lab: {
      sheetName: "<<Sheet Name>>",
      col_roomName: "B",
      col_deviceModel: "L",
      col_deviceIP: "M"
    }
  }
}
function doGet(e) {
  var result
  var searchByRoom = (e.parameter.searchbyroom ? e.parameter.searchbyroom : false)
  var searchByKeyword = (e.parameter.searchbykeyword ? e.parameter.searchbykeyword : false)

  try {
    result = getInfo(e.parameter.location, e.parameter.roomtype, searchByRoom, searchByKeyword)
  } catch (err) {
    result = "[ERROR] " + err
  }

  var content = JSON.stringify(result)

  return ContentService
    .createTextOutput(content)
    .setMimeType(ContentService.MimeType.JSON)
}

function getInfo(_location, _roomType, _roomName, _roomKeyword) {
  var spreadsheetID = setup[_location].spreadsheetID
  var s = setup[_location][_roomType]
  var content = {
    location: _location,
    roomType: _roomType,
    searchByKeyword: _roomKeyword,
    searchByRoom: _roomName,

    resultFound: 0,
    devices: []
  }

  var col_roomName = columnLetterToIndex(s.col_roomName)
  var col_deviceModel = columnLetterToIndex(s.col_deviceModel)
  var col_deviceIP = columnLetterToIndex(s.col_deviceIP)

  // Open the respective sheet and retrieve the entire dataset
  var sheet = SpreadsheetApp.openById(spreadsheetID).getSheetByName(s.sheetName)
  var data = sheet.getDataRange().getValues()

  // Looping through all the rows within room's name column
  // (1) To find and match room with the exact '_roomName' specified
  // (2) To find and match room which consists of '_roomKeyword'
  for (var row = 0; row < data.length; row++) {
    var oriRN = data[row][col_roomName]
    var lowRN = data[row][col_roomName].toLowerCase()
    var matchCondition = false

    if (_roomKeyword) {
      // If keyword is detected, do this
      matchCondition = (lowRN.indexOf(_roomKeyword.toLowerCase()) > -1)
    } else {
      // Otherwise, match using regex
      var regex = new RegExp("^" + _roomName + "\\b", "i")
      matchCondition = regex.test(lowRN)
    }

    // Yay, we found a match, lets store the device info somewhere
    // The loop continues til it reaches the end, therefore a collection of devices can be hoarded
    if (matchCondition) {
      var deviceModel = data[row][col_deviceModel]
      var deviceIP = data[row][col_deviceIP]

      content.resultFound = ++content.resultFound

      var deviceInfo = {
        name: oriRN,
        model: deviceModel,
        ipaddr: deviceIP
      }

      content.devices.push(deviceInfo)
    }
  }

  return content
}

function columnLetterToNumber(letter) {
  for (var p = 0, n = 0; p < letter.length; p++) {
    n = letter[p].charCodeAt() - 64 + n * 26
  }
  return n
}

function columnLetterToIndex(letter) {
  return columnLetterToNumber(letter) - 1
}