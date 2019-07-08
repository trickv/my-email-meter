/* inspiration:
https://gmail.googleblog.com/2012/04/know-your-gmail-stats-using-gmail-meter.html
https://www.hongkiat.com/blog/apps-scripts-gmail-users/
https://script.google.com/macros/s/AKfycbzo-X2bnwDO3jOlQXmi-u5ZBOtmVcHRknaSpF-zsh2sHHw2hWU/exec
*/

function logData() {
  var d = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var row = sheet.getLastRow() + 1;
  var lastColumn = sheet.getLastColumn();
  Logger.log("seeking to row " + row);

  // write datetime  
  sheet.getRange("A" + row).setValue(d);
  
  // built-in functions to GmailApp which may be useful & simple & cheap stats over time
  sheet.getRange("C" + row).setValue(GmailApp.getInboxUnreadCount());
  sheet.getRange("D" + row).setValue(GmailApp.getPriorityInboxUnreadCount());
  sheet.getRange("E" + row).setValue(GmailApp.getSpamUnreadCount());
  sheet.getRange("F" + row).setValue(GmailApp.getStarredUnreadCount());

  if (sheet.getRange("G1").getValue() != "AUTO:") {
    Logger.log("sheet borked, bailing out");
    sheet.getRange("G" + row).setValue("sheet borked, bailing out");
    throw("sheet borked");
  }
  
  // loop through custom searches, run them, and write stats
  for (x = 8; x < 50; x++) {
    sheet.getRange("G" + row).setValue("Ex#" + x);
    var customStartTime = new Date();
    Logger.log("custom " + x);
    var destinationCell = sheet.getRange(row, x);
    //sheet.getRange
    var search = sheet.getRange(1, x).getValue();
    Logger.log("custom " + x + " search " + search);
    if (search.substring(0, 1) == "X") {
      Logger.log("X");
      destinationCell.setValue("X");
    } else if (search == "") {
      break;
    } else if (x >= lastColumn) {
      throw("hit lastColumn, this is bad");
    } else {
      destinationCell.setValue(getSearchCount(search));
    }
    var customSearchTimeDiff = ((new Date()) - customStartTime) / 1000;
    sheet.getRange(2, x).setValue(customSearchTimeDiff);
  }
  
  // if we made it this far write OK to G{row}
  sheet.getRange("G" + row).setValue("OK");
  
  // calculate total execution time and write
  var y = new Date();
  var diff = (y - d) / 1000; // time difference is in ms, and I think in seconds for this
  sheet.getRange("B" + row).setValue(diff);
  Logger.log("Time taken: " + diff)
}

function getSearchCount(searchString) {
  var count = 0;
  while (true) {
    additional = GmailApp.search(searchString, count, 500).length
    Logger.log("search:" + searchString + ":loop " + count + " additional " + additional)
    count += additional;
    if (additional == 0) {
      break
    }
  }
  return count;
}

function doHudRow(hudRowNumber, lastRow, description, dataColumn, hours, data, hud) {
  var hoursAgoRow = lastRow - (6*hours); // HACK runs every 10 mins at the moment but should detect properly
  var cell = hud.getRange("B" + hudRowNumber);

  if (hoursAgoRow < 5) {
    cell.setValue("not enough data");
    return;
  }
  var change = data.getRange(dataColumn + lastRow).getValue() - data.getRange(dataColumn + hoursAgoRow).getValue()
  var color = "#00ff00";
  if (change > 0) {
    change = "+" + change;
    color = "#ff0000";
  }
  hud.getRange("A" + hudRowNumber).setValue(description + " change last " + hours + "h");
  cell.setValue(change);
  cell.setBackground(color);
  cell.setFontSize(40);
}
  
function updateHud() {
  var d = new Date();
  var hud = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HUD");
  hud.clear();
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");

  var lastRow = data.getLastRow();
  
  // if the last row is in progress it won't say OK in G, so assume the previous row is.
  // if the previous row isn't...well...this'll all blow up at some point.
  if (data.getRange("G" + lastRow).getValue() != "OK") {
    lastRow--;
  }
  var dataUpdated = data.getRange("A" + lastRow).getValue();
  
  displayRow = 1;
  doHudRow(displayRow++, lastRow, "important", "I", 4, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 24, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 24*3, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 24*7, data, hud);
  displayRow++;
  doHudRow(displayRow++, lastRow, "important unread", "K", 4, data, hud);
  doHudRow(displayRow++, lastRow, "important unread", "K", 24, data, hud);
  doHudRow(displayRow++, lastRow, "important unread", "K", 24*3, data, hud);
  doHudRow(displayRow++, lastRow, "important unread", "K", 24*7, data, hud);

  displayRow += 3;
  
  hud.getRange("A" + displayRow++).setValue("Last update:");
  var updatedCell = hud.getRange("A" + displayRow++)
  updatedCell.setValue(dataUpdated);
  updatedCell.setNumberFormat("YYYY-mm-dd HH:MM:SS");
  var ageCell =  hud.getRange("A" + displayRow++)
  ageCell.setValue("=now()-" + updatedCell.getA1Notation());
  ageCell.setNumberFormat("HH:MM:SS");
}