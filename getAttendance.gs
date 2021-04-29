function getAttendees() {
  const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('使用シート');
  const talkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('声かけする生徒');
  const attendanceUrl = infoSheet.getRange('B4').getValue();
  if(attendanceUrl == ""){
    Browser.msgBox("「使用シート」に今月の出欠簿を貼ってください", Browser.Buttons.OK_CANCEL);
    return;
  }
  
  const attendanceSheet = SpreadsheetApp.openByUrl(attendanceUrl).getSheetByName('出欠簿');
  const today = new Date();
  const month = today.getMonth();
  const day = today.getDate();
  const col = 18 + day * 4;
  const sheetDate = attendanceSheet.getRange(1, 17 + day * 4).getValue();

  let attendees = [];
  let counter = 10;

  if(sheetDate.getMonth() != month){
    Browser.msgBox("「使用シート」に今月の出欠簿を貼ってください", Browser.Buttons.OK_CANCEL);
    return;
  }
  
  while(1){
    let status = attendanceSheet.getRange(counter, col).getValue(); 
    let name = attendanceSheet.getRange(counter, 2).getValue() + attendanceSheet.getRange(counter, 3).getValue();
    if(name == ""){
      break;
    }
    if(status == "出席" || status == "遅刻"){     
      attendees.push(name);
    }
    counter += 1;
  }

  for(let i = 0; i < attendees.length; i ++){
    talkSheet.getRange(i + 5, 2).setValue(attendees[i]);
  }
}

//毎日0時にB列を消す
function clearCell(){
  const talkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('声かけする生徒');
  talkSheet.getRange("B5:B").clearContent();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("出席者");
  menu.addItem("本日の出席者を表示", "getAttendees");
  menu.addToUi();
}