class Utils {
  static showInfo(message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('✅ Info', message, ui.ButtonSet.OK);
  }

  static showError(message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('💥 Error', message, ui.ButtonSet.OK);
  }

  static askConfirmation(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
  }
}