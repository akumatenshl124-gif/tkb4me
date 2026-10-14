/**
 * Self App – tiện ích cho Google Sheets
 * Cách dùng: Extensions > Apps Script > dán toàn bộ file này > Save > tải lại Sheet.
 * Sẽ xuất hiện menu "Self App" trên thanh menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Self App')
    .addItem('Thêm 4 tuần vào HabitLog', 'extend4Weeks')
    .addItem('Thêm 12 tuần vào HabitLog', 'extend12Weeks')
    .addToUi();
}
function extend4Weeks()  { extendHabitLog_(4); }
function extend12Weeks() { extendHabitLog_(12); }

function extendHabitLog_(weeks) {
  var tz = Session.getScriptTimeZone();
  var ss = SpreadsheetApp.getActive();
  var habits = ss.getSheetByName('Habits').getDataRange().getValues();
  var log = ss.getSheetByName('HabitLog');
  var last = log.getRange(log.getLastRow(), 2).getValue();          // ngày cuối cùng đã có
  var existing = {};
  log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[r[0]] = true; });

  var rows = [];
  for (var d = 1; d <= weeks * 7; d++) {
    var day = new Date(last.getFullYear(), last.getMonth(), last.getDate() + d);
    var dow = (day.getDay() + 6) % 7;                                 // 0 = Thứ 2 ... 6 = Chủ nhật
    var weekStart = new Date(day.getFullYear(), day.getMonth(), day.getDate() - dow);
    var monthNo = day.getFullYear() * 100 + (day.getMonth() + 1);
    var dayKey = Utilities.formatDate(day, tz, 'yyyy-MM-dd');
    for (var i = 1; i < habits.length; i++) {
      var h = habits[i];                                              // A HabitID, B Name, E Active, F..L = T2..CN
      if (h[4] !== true || h[5 + dow] != 1) continue;
      var id = dayKey + '|' + h[0];
      if (existing[id]) continue;
      rows.push([id, day, weekStart, monthNo, h[0], h[1], false, '']);
    }
  }
  if (rows.length) {
    var start = log.getLastRow() + 1;
    log.getRange(start, 1, rows.length, 8).setValues(rows);
    log.getRange(start, 2, rows.length, 2).setNumberFormat('dd/mm/yyyy');
    log.getRange(start, 7, rows.length, 1).insertCheckboxes();
  }
  SpreadsheetApp.getUi().alert('Đã thêm ' + rows.length + ' dòng vào HabitLog.');
}
