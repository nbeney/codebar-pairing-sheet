function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Codebar 🏳️‍🌈')
    .addItem('Reset... 🗑', 'reset')
    .addItem('Format CSV 🎨', 'formatCsv')
    .addSeparator()
    .addItem('Select coach ➖', 'selectCoach')
    .addItem('Assign coach to student ➕', 'assignSelectedCoachToStudent')
    .addSeparator()
    .addItem('Show pairings ↹', 'showPairings')
    .addItem('Show numbers 🔢', 'showNumbers')
    .addSeparator()
    .addItem('Sort by name', 'sortByName')
    .addItem('Sort by role/name', 'sortByRoleName')
    .addItem('Sort by group/role/name', 'sortByGroupRoleName')
    .addSeparator()
    .addItem('Filter all', 'filterAll')
    .addItem('Filter present only 👍', 'filterPresentOnly')
    .addItem('Filter absent only 👎', 'filterAbsentOnly')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Demo/Testing 🧪')
        .addItem('Paste sample Pairing CSV data 📚', 'Demo.pasteSamplePairingCsvData')
        .addItem('Register randomly 📚', 'Demo.registerAtRandom')
        .addItem('Pair randomly 📚', 'Demo.pairAtRandom')
    )
    .addSeparator()
    .addItem('Help 🛟', 'showHelp')
    .addToUi();

    reset();
    Demo.pasteSamplePairingCsvData();
    formatCsv();
}

function reset() {
  if (!Utils.askConfirmation('❌ Warning', 'This will clear the whole sheet. Do you want to continue?')) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }

  sheet.getDataRange().clearDataValidations();
  sheet.clear();
  sheet.setFrozenRows(0);

  const bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  const message = SpreadsheetApp.newRichTextValue()
    .setText('Paste the output of Pairing CSV here then run the Format CSV 🎨 macro')
    //        0         1         2         3         4         5         6         7         8
    //        012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
    .setTextStyle(20, 31, bold)
    .setTextStyle(50, 60, bold)
    .build();
  sheet.getRange("A1").setRichTextValue(message);
  sheet.getRange("A1").setFontStyle('italic');
  sheet.getRange("A1").activate();
}

function formatCsv() {
  try {
    Format.splitCSVIntoColumns();
    Format.fillEmptyCells();
    Format.compactPronouns();
    Format.flagNewcomers();
    Format.flagCoachesAndStudents();
    Format.deleteNewAttendeesColumn();
    Format.normalizeTechnologies('Skills', SKILLS_MAP);
    Format.normalizeTechnologies('Note', SKILLS_MAP);
    Format.copySkillsForCoaches();
    Format.deleteSkillsAndRenameTutorialColumn();
    Format.insertRegisteredColumn();
    Format.insertGroupColumn();
    Format.setGroupForCoachesAndStudents();
    Format.sortAttendees();
    Format.freezeTopRow();
    Format.formatHeaderRow();
    Format.duplicateHeaders();
    Format.resizeColumnsToFit();
    Format.clipColumns();
    Format.addFilter();

    Utils.showInfo('Formatting completed successfully!');
  } catch (e) {
    Utils.showError(e.message);
  }
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 1
function selectCoach() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rowIdx = sheet.getActiveCell().getRow();
  const leftRole = sheet.getRange(rowIdx, COL_ROLE_1).getValue();
  const rightRole = sheet.getRange(rowIdx, COL_ROLE_2).getValue();

  if (leftRole !== ROLE_COACH && rightRole !== ROLE_COACH) {
    Utils.showInfo('Current row does not contain a coach.');
    return;
  }

  const colIdx = (leftRole === ROLE_COACH) ? COL_REGISTERED_1 : COL_REGISTERED_2;

  sheet.getRange(rowIdx, colIdx).offset(0, 0, 1, NUM_COLS).activate();

  PropertiesService.getScriptProperties().setProperty(PROP_COACH_ROW_INDEX, rowIdx.toString());
  PropertiesService.getScriptProperties().setProperty(PROP_COACH_COL_INDEX, colIdx.toString());
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 2
function assignSelectedCoachToStudent() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const studentRowIdx = sheet.getActiveCell().getRow();
  const leftRole = sheet.getRange(studentRowIdx, COL_ROLE_1).getValue();
  const rightRole = sheet.getRange(studentRowIdx, COL_ROLE_2).getValue();
  const coachRowIdx = PropertiesService.getScriptProperties().getProperty(PROP_COACH_ROW_INDEX);
  const coachColIdx = PropertiesService.getScriptProperties().getProperty(PROP_COACH_COL_INDEX);

  if (!coachRowIdx || !coachColIdx) {
    Utils.showInfo('Please select a coach first.');
    return;
  } else if (leftRole === ROLE_COACH) {
    Utils.showInfo('Cannot assign a coach to a coach.');
    return;
  } else if (rightRole === ROLE_COACH) {
    Utils.showInfo('The student is already assigned to a coach.');
    return;
  } else if (leftRole !== ROLE_STUDENT) {
    Utils.showInfo('Current row does not contain a student.');
    return;
  }

  const sourceRange = sheet.getRange(coachRowIdx, coachColIdx, 1, NUM_COLS);
  const targetRange = sheet.getRange(studentRowIdx, COL_REGISTERED_2, 1, NUM_COLS);
  sourceRange.moveTo(targetRange);

  sortByCurrentCriteria();

  PropertiesService.getScriptProperties().deleteProperty(PROP_COACH_ROW_INDEX);
  PropertiesService.getScriptProperties().deleteProperty(PROP_COACH_COL_INDEX);
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 4
function showPairings() {
  const { pairings, unpairedStudents, unpairedCoaches } = collectPairings();

  // Create HTML template from the pairings.html file
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile('templates/pairings');

  // Pass the data directly to the template
  template.pairings = pairings;
  template.unpairedStudents = unpairedStudents;
  template.unpairedCoaches = unpairedCoaches;

  const htmlOutput = template.evaluate()
    .setTitle('Workshop Pairings')
    .setWidth(700)
    .setHeight(600);

  ui.showModalDialog(htmlOutput, 'Pairings');
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 5
function showNumbers() {
  Utils.showInfo('Not yet implemented!');
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 3
function sortByCurrentCriteria() {
  const criteria = PropertiesService.getScriptProperties().getProperty(PROP_SORT_CRITERIA);
  switch (criteria) {
    case 'BY_NAME':
      sortByName();
      break;
    case 'BY_ROLE_NAME':
      sortByRoleName();
      break;
    case 'BY_GROUP_ROLE_NAME':
      sortByGroupRoleName();
      break;
    default:
      sortByName();
  }
}

function sortByName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const nameColIndex = data[0].indexOf('Name');

  if (nameColIndex !== -1) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort({ column: nameColIndex + 1, ascending: true });
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_NAME');
}

function sortByRoleName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const roleColIndex = data[0].indexOf('Role');
  const nameColIndex = data[0].indexOf('Name');

  if (roleColIndex !== -1 && nameColIndex !== -1) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      { column: roleColIndex + 1, ascending: true },
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_ROLE_NAME');
}

function sortByGroupRoleName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const groupColIndex = data[0].indexOf('Group');
  const roleColIndex = data[0].indexOf('Role');
  const nameColIndex = data[0].indexOf('Name');

  if (groupColIndex !== -1 && roleColIndex !== -1 && nameColIndex !== -1) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      { column: groupColIndex + 1, ascending: true },
      { column: roleColIndex + 1, ascending: true },
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_GROUP_ROLE_NAME');
}

function filterAll() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues([]).build();
  sheet.getFilter().setColumnFilterCriteria(1, criteria);
  sheet.getRange('A2').activate();
}

function filterPresentOnly() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['FALSE']).build();
  sheet.getFilter().setColumnFilterCriteria(1, criteria);
  sheet.getRange('A2').activate();
}

function filterAbsentOnly() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['TRUE']).build();
  sheet.getFilter().setColumnFilterCriteria(1, criteria);
  sheet.getRange('A2').activate();
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('templates/help')
    .setTitle('Help')
    .setWidth(600)
    .setHeight(500);
  ui.showSidebar(html);
}

function collectPairings() {
  const isStudent = (reg, name, role) => { return reg && role === ROLE_STUDENT && name !== ''; }
  const isCoach = (reg, name, role) => { return reg && role === ROLE_COACH && name !== ''; }

  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const pairings = {};
  let unpairedStudents = [];
  let unpairedCoaches = [];

  for (let i = 1; i < data.length; i++) {
    const group = data[i][COL_GROUP_1 - 1];
    const reg1 = data[i][COL_REGISTERED_1 - 1];
    const name1 = data[i][COL_NAME_1 - 1];
    const role1 = data[i][COL_ROLE_1 - 1];
    const reg2 = data[i][COL_REGISTERED_2 - 1];
    const name2 = data[i][COL_NAME_2 - 1];
    const role2 = data[i][COL_ROLE_2 - 1];

    if (isStudent(reg1, name1, role1) && isCoach(reg2, name2, role2)) {
      // Paired student and coach
      if (!pairings[group]) {
        pairings[group] = [];
      }
      // Find if the coach already exists in the group
      let coachEntry = pairings[group].find(entry => entry.coach === name2);
      if (!coachEntry) {
        coachEntry = { coach: name2, students: [] };
        pairings[group] = [...pairings[group], coachEntry];
      }
      coachEntry.students = [...coachEntry.students, name1];
    } else if (isStudent(reg1, name1, role1)) {
      unpairedStudents = [...unpairedStudents, name1];
    } else if (isCoach(reg1, name1, role1)) {
      unpairedCoaches = [...unpairedCoaches, name1];
    }
  }

  unpairedStudents.sort();
  unpairedCoaches.sort();

  return { pairings, unpairedStudents, unpairedCoaches };
}
