class Utils {
  static showInfo(message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('✅ Info', message, ui.ButtonSet.OK);
  }

  static showError(exception) {
    const ui = SpreadsheetApp.getUi();
    
    // Create detailed error message
    const lines = [];
    if (exception.stack) {
      lines.push(`${exception.stack}`);
    } else {
      if (exception.message) {
        lines.push(`Message: ${exception.message}`);
      }
      if (exception.fileName) {
        lines.push(`File: ${exception.fileName}`);
      }
      if (exception.lineNumber) {
        lines.push(`Line: ${exception.lineNumber}`);
      }
    }
    
    // Fallback if no detailed info available
    const errorDetails = lines.length > 0 ? lines.join('\n') : exception.toString();
    
    ui.alert('💥 Error', errorDetails, ui.ButtonSet.OK);
  }

  static askConfirmation(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
  }
}