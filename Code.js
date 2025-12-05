function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Codebar 🏳️‍🌈')
    .addItem('Reset 🗑', 'reset')
    .addItem('Format 🖌️', 'format')
    // .addSeparator()
    // .addItem('Select coach ➖', 'selectCoach')
    // .addItem('Assign coach to student ➕', 'assignSelectedCoachToStudent')
    .addItem(`Show pairings ${ICONS.pair}`, 'showPairings')
    .addSeparator()
    .addItem('Sort by name', 'sortByName')
    .addItem('Sort by role/name', 'sortByRoleName')
    .addItem('Sort by group/role/name', 'sortByGroupRoleName')
    .addItem('Sort by role/group/name', 'sortByRoleGroupName')
    .addSeparator()
    .addItem('Filter all', 'filterAll')
    .addItem('Filter present only 👍', 'filterPresentOnly')
    .addItem('Filter absent only 👎', 'filterAbsentOnly')
    .addSeparator()
    .addSubMenu(
      ui.createMenu(`Tutorial ${ICONS.tutorial}`)
        .addItem('1 - Reset', 'Tutorial.step1ResetSheet')
        .addItem('2 - Paste CSV', 'Tutorial.step2PastePairingCsvData')
        .addItem('3 - Format', 'Tutorial.step3FormatCsv')
        .addItem('4 - Register participants', 'Tutorial.step4RegisterParticipants')
        .addItem('5 - Sort participants', 'Tutorial.step5SortParticipants')
        .addItem('6 - Assign coaches to groups', 'Tutorial.step6AssignCoachesToGroups')
        .addItem('7 - Sort participants', 'Tutorial.step7SortParticipants')
        .addItem('8 - Assign coaches to students', 'Tutorial.step8AssignCoachesToStudents')
        .addItem('9 - Show pairings', 'Tutorial.step9ShowPairings')
    )
    .addItem('Help 🛟', 'showHelp')
    .addToUi();
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
    .setText('Paste the output of Pairing CSV here then run the Format 🖌️ macro')
    //        0         1         2         3         4         5         6         7         8
    //        012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
    .setTextStyle(20, 31, bold)
    .setTextStyle(50, 56, bold)
    .build();
  sheet.getRange("A1").setRichTextValue(message);
  sheet.getRange("A1").setFontStyle('italic');
  sheet.getRange("A1").activate();
}

function format() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getRange(1, 1).getValue() !== HEADER_REGISTERED) {
      Format.splitCSVIntoColumns(); // structure
      Format.fillEmptyCells(); // content
      Format.flagNewcomersAndDeleteNewAttendeesColumn(); // content
      Format.flagCoachesAndStudents(); // content
      Format.mergeSkillsAndTutorialColumns(); // content
      Format.normalizeTechnologies('Skills/Tutorial', SKILLS_MAP); // content
      Format.normalizePronouns(); // content
      Format.insertRegisteredColumn(); // structure

      Format.insertGroupColumn(); // structure
      Format.setGroupForCoachesAndStudents(); // content
      Format.renameColumnHeaders(); // content

      Format.freezeTopRow(); // appearance
      Format.formatHeaderRow(); // appearance
      Format.duplicateHeaders(); // appearance
      Format.resizeColumnsToFit(); // appearance
      Format.clipColumns(); // appearance
      Format.addFilter();
      Format.addSummaryRow();
    }

    sortByCurrentCriteria(); // appearance

    // Utils.showInfo('Formatting completed successfully!');
  } catch (e) {
    Utils.showError(e);
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
  const { pairings, unpairedCoaches, unpairedStudents, unregisteredCoaches, unregisteredStudents } = collectPairings();

  // Create HTML template from the pairings.html file
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile('templates/pairings');

  // Pass the data directly to the template
  template.pairings = pairings;
  template.unpairedCoaches = unpairedCoaches;
  template.unpairedStudents = unpairedStudents;
  template.missingCoaches = unregisteredCoaches;
  template.missingStudents = unregisteredStudents;
  template.GROUPS = GROUPS;
  template.ICONS = ICONS;

  const htmlOutput = template.evaluate()
    .setTitle('Workshop Pairings')
    .setWidth(700)
    .setHeight(600);

  ui.showModalDialog(htmlOutput, 'Pairings');
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
    case 'BY_ROLE_GROUP_NAME':
      sortByRoleGroupName();
      break;
    default:
      sortByName();
  }
}

function sortByName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const nameColIndex = data[1].indexOf(HEADER_NAME);

  if (nameColIndex !== -1) {
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort({ column: nameColIndex + 1, ascending: true });
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_NAME');
  Format.applyConditionalFormatting();
}

function sortByRoleName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const roleColIndex = data[1].indexOf(HEADER_ROLE);
  const nameColIndex = data[1].indexOf(HEADER_NAME);

  if (roleColIndex !== -1 && nameColIndex !== -1) {
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      { column: roleColIndex + 1, ascending: true },
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_ROLE_NAME');
  Format.applyConditionalFormatting();
}

function sortByGroupRoleName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const groupColIndex = data[1].indexOf(HEADER_GROUP);
  const roleColIndex = data[1].indexOf(HEADER_ROLE);
  const nameColIndex = data[1].indexOf(HEADER_NAME);

  if (groupColIndex !== -1 && roleColIndex !== -1 && nameColIndex !== -1) {
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
    range.sort([
      { column: groupColIndex + 1, ascending: true },
      { column: roleColIndex + 1, ascending: true },
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_GROUP_ROLE_NAME');
  Format.applyConditionalFormatting();
}

function sortByRoleGroupName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const roleColIndex = data[1].indexOf(HEADER_ROLE);
  const groupColIndex = data[1].indexOf(HEADER_GROUP);
  const nameColIndex = data[1].indexOf(HEADER_NAME);
  
  if (roleColIndex !== -1 && groupColIndex !== -1 && nameColIndex !== -1) {
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
    range.sort([
      { column: roleColIndex + 1, ascending: true },
      { column: groupColIndex + 1, ascending: true },
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_ROLE_GROUP_NAME');
  Format.applyConditionalFormatting();
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
  const isCoach = (reg, name, role) => { return reg && role === ROLE_COACH && name !== ''; }
  const isStudent = (reg, name, role) => { return reg && role === ROLE_STUDENT && name !== ''; }

  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const pairings = {};
  let unpairedCoaches = [];
  let unpairedStudents = [];
  let unregisteredCoaches = [];
  let unregisteredStudents = [];

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
    } else if (!reg1 && role1 === ROLE_STUDENT && name1 !== '') {
      unregisteredStudents = [...unregisteredStudents, name1];
    } else if (!reg1 && role1 === ROLE_COACH && name1 !== '') {
      unregisteredCoaches = [...unregisteredCoaches, name1];
    }
  }

  unpairedCoaches.sort();
  unpairedStudents.sort();
  unregisteredCoaches.sort();
  unregisteredStudents.sort();

  return { pairings, unpairedCoaches, unpairedStudents, unregisteredCoaches, unregisteredStudents };
}
