function append_mtr_week() {

  var mtr_week_file_name = 'mtr_nieuwe_week';
  
  // find new week
  var gs_files = DriveApp.getFilesByName(mtr_week_file_name);

  // copy new week to R3-montage, if there is one
  if (gs_files.hasNext()) {
    var R3_ss = SpreadsheetApp.getActiveSpreadsheet();
    var mtr_week_file = gs_files.next();
    var mtr_week_ss = SpreadsheetApp.open(mtr_week_file);
    var mtr_week_sheet = mtr_week_ss.getSheets()[0];
    
    mtr_week_sheet.copyTo(R3_ss);
    
    var R3_mtr_week_sheet = R3_ss.getSheetByName('Kopie van mtr.tsv');
    var R3_mtr_sheet = R3_ss.getSheetByName('montage');
    
    // add 1 empty row at end of MT-schedule sheet
    R3_mtr_sheet.insertRowsAfter(R3_mtr_sheet.getMaxRows(), 1);
    
    // place "cursor" at last row in column D (A-C are hidden)
    R3_mtr_sheet.getRange('D1')
                .getNextDataCell(SpreadsheetApp.Direction.DOWN)
                .activate();
    R3_mtr_sheet.getCurrentCell()
                .offset(1, -3) // move cursor to column A of last row (empty)
                .activate();
                
    // copy new week to end of MT-schedule
    R3_mtr_week_sheet.getDataRange()
                     .copyTo(R3_mtr_sheet.getActiveRange(), 
                             SpreadsheetApp.CopyPasteType.PASTE_VALUES, 
                             false); // do not transpose

    // remove mtr-week-sheet from R3
    R3_ss.deleteSheet(R3_mtr_week_sheet);

    // "delete" mtr-week-file on Drive 
    DriveApp.getFilesByName(mtr_week_file_name).next().setTrashed(true);
  } 
}
