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
  var runInterval = 5; // HACK: this should just seek! In the mean time, assume data points are 5 minutes apart.
  var hoursAgoRow = lastRow - ((60/runInterval)*hours); // HACK: updateData logs every X mins at the moment but should detect properly
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
  } else if (change == 0) {
    color = "#cccccc";
  }
  hud.getRange("A" + hudRowNumber).setValue(description + " change last " + hours + "h");
  cell.setValue(change);
  cell.setBackground(color);
  cell.setFontSize(20);
}
  
function updateHud() {
  var d = new Date();
  var hud = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HUD");
  //hud.clear(); // FIXME: need a way to run this only when the layout has changed since last run.
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");

  var lastRow = data.getLastRow();
  
  // if the last row is in progress it won't say OK in G, so assume the previous row is.
  // if the previous row isn't...well...this'll all blow up at some point.
  if (data.getRange("G" + lastRow).getValue() != "OK") {
    lastRow--;
  }
  var dataUpdated = data.getRange("A" + lastRow).getValue();
  
  displayRow = 1;
  
  /*
  var importantCountNow =  data.getRange("I" + lastRow).getValue();
  hud.getRange("A" + displayRow).setValue("important NOW");
  hud.getRange("B" + displayRow++).setValue(importantCountNow);
  doHudRow(displayRow++, lastRow, "important", "I", 1, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 4, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 24, data, hud);
  doHudRow(displayRow++, lastRow, "important", "I", 24*7, data, hud);
  displayRow++;
  */
  
  var importantUnreadCountNow =  data.getRange("K" + lastRow).getValue();
  hud.getRange("A" + displayRow).setValue("important unread NOW");
  hud.getRange("B" + displayRow++).setValue(importantUnreadCountNow);
  doHudRow(displayRow++, lastRow, "important unread", "K", 4, data, hud);
  doHudRow(displayRow++, lastRow, "important unread", "K", 24, data, hud);
  doHudRow(displayRow++, lastRow, "important unread", "K", 24*7, data, hud);
  displayRow++;
  var starredCountNow =  data.getRange("L" + lastRow).getValue();
  hud.getRange("A" + displayRow).setValue("starred NOW");
  hud.getRange("B" + displayRow++).setValue(starredCountNow);
  doHudRow(displayRow++, lastRow, "starred", "L", 1, data, hud);
  doHudRow(displayRow++, lastRow, "starred", "L", 24*0.5, data, hud);
  doHudRow(displayRow++, lastRow, "starred", "L", 24, data, hud);
  doHudRow(displayRow++, lastRow, "starred", "L", 24*7, data, hud);

  displayRow += 3;
  
  updateHudCharts();
  
  hud.getRange("A" + displayRow++).setValue("Last update:");
  var updatedCell = hud.getRange("A" + displayRow++)
  updatedCell.setValue(dataUpdated);
  updatedCell.setNumberFormat("YYYY-mm-dd HH:MM:SS");
  var ageCell =  hud.getRange("A" + displayRow++)
  ageCell.setValue("=now()-" + updatedCell.getA1Notation());
  ageCell.setNumberFormat("HH:MM:SS");
}

function updateHudCharts() {
  var hud = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HUD");
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  charts = hud.getCharts();
  chart = charts[0];
  Logger.log(chart.getId());
  var lastRow = data.getLastRow();
  var chartHistoryTimeHours = 12;
  var rangeTextData = "K" + (lastRow-(chartHistoryTimeHours*(60/5))) + ":K" + (lastRow);
  var rangeTextA = "A" + (lastRow-(chartHistoryTimeHours*(60/5))) + ":A" + (lastRow);
  chart = chart.modify()
      .setOption("title", "")
      .clearRanges()
      .addRange(data.getRange("A1"))
      .addRange(data.getRange("I1"))
      .addRange(data.getRange(rangeTextData))
      .build();
  hud.updateChart(chart);
  Logger.log("done");
}

function onOpen() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  var menu = [
    {name: 'Update', functionName: 'updateHud'},
  ];
  doc.addMenu("Custom Script", menu);
}