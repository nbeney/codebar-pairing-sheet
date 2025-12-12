function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Codebar 🏳️‍🌈')
    .addItem('Reset 🗑', 'reset')
    .addItem('Format 🖌️', 'format')
    .addItem(`Show pairings ${ICONS.pair}`, 'showPairings')
    .addSeparator()
    .addSubMenu(
      ui.createMenu(`Sort list`)
        .addItem('By name', 'sortListByName')
        .addItem('By role/name', 'sortListByRoleName')
        .addItem('By group/role/name', 'sortListByGroupRoleName')
        .addItem('By role/group/name', 'sortListByRoleGroupName')
    )
    .addItem('Sort pairs', 'sortPairs')
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

  // Format.createPairsSheet();
  // sortPairs();
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
      Format.splitCSVIntoColumns();
      Format.fillEmptyCells();
      Format.flagNewcomersAndDeleteNewAttendeesColumn();
      Format.flagCoachesAndStudents();
      Format.mergeSkillsAndTutorialColumns();
      Format.normalizeTechnologies('Skills/Tutorial', SKILLS_MAP);
      Format.normalizePronouns();
      Format.insertRegisteredColumn();

      Format.insertGroupColumn();
      Format.setGroupForCoachesAndStudents();
      Format.renameColumnHeaders();

      Format.freezeTopRow();
      Format.formatHeaderRow();
      Format.resizeColumnsToFit();
      Format.clipColumns();
      Format.addFilter();
      Format.addSummaryRow();
      Format.renameSheetToList();
      Format.createPairsSheet();
    }

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIST).activate();
    sortListByName();
    sortPairs();
  } catch (e) {
    Utils.showError(e);
  }
}

// This macro should be imported and assigned to Ctrl-Alt-Shift 4
function showPairings() {
  // Activate Pairs sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAIRS);
  sheet.activate();
  
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

// This macro should be imported and assigned to Ctrl-Alt-Shift 1
function sortListByCurrentCriteria() {
  const sheet = SpreadsheetApp.getActiveSheet();
  switch (sheet.getName()) {
    case SHEET_LIST:
      const criteria = PropertiesService.getScriptProperties().getProperty(PROP_SORT_CRITERIA);
      switch (criteria) {
        case 'BY_NAME':
          sortListByName();
          break;
        case 'BY_ROLE_NAME':
          sortListByRoleName();
          break;
        case 'BY_GROUP_ROLE_NAME':
          sortListByGroupRoleName();
          break;
        case 'BY_ROLE_GROUP_NAME':
          sortListByRoleGroupName();
          break;
        default:
          sortListByName();
      }
      break;
    case SHEET_PAIRS:
      sortPairs();
      break;
    default:
      Utils.showError(`Sorting is only available on the "${SHEET_LIST}" or "${SHEET_PAIRS}" sheets.`);
  }
}

function sortListByName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIST);
  sheet.activate();
  const data = sheet.getDataRange().getValues();
  const nameColIndex = data[1].indexOf(HEADER_NAME);

  if (nameColIndex !== -1) {
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort({ column: nameColIndex + 1, ascending: true });
  }

  PropertiesService.getScriptProperties().setProperty(PROP_SORT_CRITERIA, 'BY_NAME');
  Format.applyConditionalFormatting();
}

function sortListByRoleName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIST);
  sheet.activate();
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

function sortListByGroupRoleName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIST);
  sheet.activate();
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

function sortListByRoleGroupName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIST);
  sheet.activate();
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

function sortPairs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAIRS);
  sheet.activate();

  // Add Key header in M2
  sheet.getRange(2, 13).setValue('Key').setFontWeight('bold');

  // Populate Key column with group (D & J) + role (C & I) + name (B & H)
  // Use getValues to read all data at once for performance
  // Use setValues to write all data at once for performance
  const data = sheet.getDataRange().getValues();
  const numRows = data.length;
  const keyValues = [];

  for (let i = 2; i < numRows; i++) {
    const leftGroup = data[i][COL_GROUP_1 - 1];
    const leftRole = data[i][COL_ROLE_1 - 1];
    const leftName = data[i][COL_NAME_1 - 1];
    const rightGroup = data[i][COL_GROUP_2 - 1];
    const rightRole = data[i][COL_ROLE_2 - 1];
    const rightName = data[i][COL_NAME_2 - 1];
    const key = `${leftGroup}${rightGroup}${leftRole}${rightRole}${leftName}${rightName}`;
    keyValues.push([key]);
  }

  sheet.getRange(3, 13, keyValues.length, 1).setValues(keyValues);

  // Sort by Key column (M)
  const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  range.sort({ column: 13, ascending: true });

  // Remove Key column
  sheet.deleteColumn(13);

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
  const EMPTY = 1;
  const ABSENT = 2;
  const PRESENT = 3;

  const status = (reg, name, role, targetRole) => {
    if (name === '') return EMPTY;
    if (!reg && role === targetRole) return ABSENT;
    if (reg && role === targetRole) return PRESENT;
    return EMPTY;
  };

  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const pairings = {};
  const unpairedCoaches = new Set();
  const unpairedStudents = new Set();
  const unregisteredCoaches = new Set();
  const unregisteredStudents = new Set();

  for (let i = 2; i < data.length; i++) {
    const group = data[i][COL_GROUP_1 - 1];
    const reg1 = data[i][COL_REGISTERED_1 - 1];
    const name1 = data[i][COL_NAME_1 - 1];
    const role1 = data[i][COL_ROLE_1 - 1];
    const studentStatus = status(reg1, name1, role1, ROLE_STUDENT);
    const reg2 = data[i][COL_REGISTERED_2 - 1];
    const name2 = data[i][COL_NAME_2 - 1];
    const role2 = data[i][COL_ROLE_2 - 1];
    const coachStatus = status(reg2, name2, role2, ROLE_COACH);

    // Fixed switch statement syntax - use string concatenation or if-else
    const combinedStatus = `${studentStatus}-${coachStatus}`;
    
    switch (combinedStatus) {
      case `${EMPTY}-${ABSENT}`:
        unregisteredCoaches.add(name2);
        break;
      case `${EMPTY}-${PRESENT}`:
        unpairedCoaches.add(name2);
        break;
      case `${ABSENT}-${EMPTY}`:
        unregisteredStudents.add(name1);
        break;
      case `${ABSENT}-${ABSENT}`:
        unregisteredStudents.add(name1);
        unregisteredCoaches.add(name2);
        break;
      case `${ABSENT}-${PRESENT}`:
        unregisteredStudents.add(name1);
        unpairedCoaches.add(name2);
        break;
      case `${PRESENT}-${EMPTY}`:
        unpairedStudents.add(name1);
        break;
      case `${PRESENT}-${ABSENT}`:
        unpairedStudents.add(name1);
        unregisteredCoaches.add(name2);
        break;
      case `${PRESENT}-${PRESENT}`:
        // Paired student and coach
        if (!pairings[group]) {
          pairings[group] = [];
        }
        let coachEntry = pairings[group].find(entry => entry.coach === name2);
        if (!coachEntry) {
          coachEntry = { coach: name2, students: [] };
          pairings[group].push(coachEntry);
        }
        coachEntry.students.push(name1);
        break;
    }
  }

  return { 
    pairings, 
    unpairedCoaches: [...unpairedCoaches].sort(), 
    unpairedStudents: [...unpairedStudents].sort(), 
    unregisteredCoaches: [...unregisteredCoaches].sort(), 
    unregisteredStudents: [...unregisteredStudents].sort() 
  };
}
