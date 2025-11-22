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
        .addItem('Paste sample Pairing CSV data 📚', 'pasteSamplePairingCsvData')
        .addItem('Register randomly 📚', 'registerAtRandom')
        .addItem('Pair randomly 📚', 'pairAtRandom')
    )
    .addSeparator()
    .addItem('Help 🛟', 'showHelp')
    .addToUi();

  // // Uncomment these 2 lines for development only
  // reset();
  // pasteSamplePairingCsvData();
  // formatCsv();
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
  const sheet = SpreadsheetApp.getActiveSheet();

  // Step 1: Split CSV into columns
  function splitCSVIntoColumns() {
    sheet.getActiveCell().getDataRegion().splitTextToColumns(",");
  }

  // Step 2: Fill the empty cells to avoid problems later
  function fillEmptyCells() {
    const data = sheet.getDataRange().getValues();
    const numRows = data.length;
    const numColumns = data[0].length;

    // Iterate through all cells in the data
    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < numColumns; j++) {
        // Check if the cell is empty (null, undefined, or empty string)
        if (data[i][j] === null || data[i][j] === undefined || data[i][j] === '') {
          data[i][j] = '-'; // Replace with empty marker string for consistency
        }
      }
    }

    // Write the modified data back to the sheet
    sheet.getRange(1, 1, numRows, numColumns).setValues(data);
  }
  // Step 3: Compact pronouns in Name column
  function compactPronouns() {
    const data = sheet.getDataRange().getValues();
    const nameColIndex = data[0].indexOf('Name');

    if (nameColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      let name = data[i][nameColIndex].toString();

      // Define pronoun sets
      const pronounSets = {
        H: ['he', 'him', 'his'],
        S: ['she', 'her', 'hers'],
        T: ['they', 'them', 'theirs']
      };

      // Extract and classify pronouns
      name = name.replace(/\([^)]+\)/gi, (match) => {
        // Extract words from parentheses, splitting on common separators
        const words = match
          .slice(1, -1) // Remove parentheses
          .toLowerCase()
          .split(/[/,\s]+/)
          .filter(w => w.length > 0);
        
        // Check which pronoun sets are present
        const hasH = words.some(w => pronounSets.H.includes(w));
        const hasS = words.some(w => pronounSets.S.includes(w));
        const hasT = words.some(w => pronounSets.T.includes(w));
        
        // Count how many sets are present
        const count = [hasH, hasS, hasT].filter(Boolean).length;
        
        // Return appropriate tag
        if (count === 0) return match; // Not pronouns, keep original
        if (count === 1) {
          if (hasH) return '[H]';
          if (hasS) return '[S]';
          if (hasT) return '[T]';
        }
        
        // Multiple sets - return combination
        if (hasH && hasS) return '[H/S]';
        if (hasH && hasT) return '[H/T]';
        if (hasS && hasT) return '[S/T]';
        
        return match; // Fallback
      });

      data[i][nameColIndex] = name;
    }

    sheet.getDataRange().setValues(data);
  }

  // Step 4: Flag newcomers with chick emoji
  function flagNewcomers() {
    const data = sheet.getDataRange().getValues();
    const newAttendeeColIndex = data[0].indexOf('New attendee');
    const nameColIndex = data[0].indexOf('Name');

    if (newAttendeeColIndex === -1 || nameColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][newAttendeeColIndex].toString().toLowerCase() === 'true') {
        data[i][nameColIndex] = data[i][nameColIndex] + ' 🐥';
      }
    }

    sheet.getDataRange().setValues(data);
  }

  // Step 5: Delete New attendees column
  function deleteNewAttendeesColumn() {
    const data = sheet.getDataRange().getValues();
    const newAttendeeColIndex = data[0].indexOf('New attendee');

    if (newAttendeeColIndex !== -1) {
      sheet.deleteColumn(newAttendeeColIndex + 1);
    }
  }

  // Step 6: Proper case for technologies
  function normalizeTechnologies(columnName, skillsMap) {
    const data = sheet.getDataRange().getValues();
    const colIndex = data[0].indexOf(columnName);

    if (colIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      let skills = data[i][colIndex].toString();

      if (skills && skills !== 'N/A') {
        Object.keys(skillsMap).forEach(key => {
          const regex = new RegExp('\\b' + key + '\\b', 'gi');
          skills = skills.replace(regex, skillsMap[key]);
        });

        data[i][colIndex] = skills;
      }
    }

    sheet.getDataRange().setValues(data);
  }

  // Step 7: Copy Skills to Tutorial for Coaches
  function copySkillsForCoaches() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');
    const skillsColIndex = data[0].indexOf('Skills');
    const tutorialColIndex = data[0].indexOf('Tutorial');

    if (roleColIndex === -1 || skillsColIndex === -1 || tutorialColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === ROLE_COACH) {
        data[i][tutorialColIndex] = data[i][skillsColIndex];
      }
    }

    sheet.getDataRange().setValues(data);
  }

  // Step 8: Delete Skills column and rename Tutorial -> Skills/Tutorial
  function deleteSkillsAndRenameTutorialColumn() {
    const data = sheet.getDataRange().getValues();
    const skillsColIndex = data[0].indexOf('Skills');

    if (skillsColIndex !== -1) {
      sheet.deleteColumn(skillsColIndex + 1);
    }

    // Refresh data after deletion
    const updatedData = sheet.getDataRange().getValues();
    const tutorialColIndex = updatedData[0].indexOf('Tutorial');

    if (tutorialColIndex !== -1) {
      sheet.getRange(1, tutorialColIndex + 1).setValue('Skills/Tutorial');
    }
  }

  // Step 9: Insert ?? column with checkboxes before Name
  function insertCheckboxColumn() {
    const data = sheet.getDataRange().getValues();
    const nameColIndex = data[0].indexOf('Name');

    if (nameColIndex !== -1) {
      sheet.insertColumnBefore(nameColIndex + 1);
      const newData = sheet.getDataRange().getValues();

      // Set header
      sheet.getRange(1, nameColIndex + 1).setValue('??');

      // Insert checkboxes for all data rows
      for (let i = 2; i <= newData.length; i++) {
        sheet.getRange(i, nameColIndex + 1).insertCheckboxes();
      }
    }
  }

  // Step 10: Insert Group column after Role
  function insertGroupColumn() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');

    if (roleColIndex !== -1) {
      sheet.insertColumnAfter(roleColIndex + 1);
      sheet.getRange(1, roleColIndex + 2).setValue('Group');
    }
  }

  // Step 11: Set group for students
  function setGroupForStudentsAndAddValidation() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');
    const groupColIndex = data[0].indexOf('Group');
    const skillsTutorialColIndex = data[0].indexOf('Skills/Tutorial');

    if (roleColIndex === -1 || groupColIndex === -1) return;

    // Data part

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === ROLE_STUDENT) {
        const skillsTutorial = data[i][skillsTutorialColIndex].toString();
        const group = TUTORIAL_GROUP_MAP[skillsTutorial];
        if (group !== undefined) {
          data[i][groupColIndex] = group;
        }
      }
    }
    sheet.getDataRange().setValues(data);

    // Validation part

    const groups = Array.from(new Set(Object.values(TUTORIAL_GROUP_MAP))).sort();
    const val = SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInList(groups, true)
      .build();
    sheet.getRange(2, groupColIndex + 1, sheet.getLastRow() - 1).setDataValidation(val);
  }

  // Step 12: Sort by Role and Name
  function sortAttendees() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');
    const nameColIndex = data[0].indexOf('Name');

    if (roleColIndex !== -1 && nameColIndex !== -1) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      range.sort([
        // { column: roleColIndex + 1, ascending: true },
        { column: nameColIndex + 1, ascending: true }
      ]);
    }
  }

  // Step 13: Freeze top row
  function freezeTopRow() {
    sheet.setFrozenRows(1);
  }

  // Step 14: Format header row
  function formatHeaderRow() {
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('bold');
    headerRange.setBackground(COLOR_HEADER);
  }

  // Step 15: Format Coach rows
  function formatCoachRows() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');

    if (roleColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === ROLE_COACH) {
        const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
        rowRange.setBackground(COLOR_COACH);
      }
    }
  }

  // Step 16: Format Student rows
  function formatStudentRows() {
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');

    if (roleColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === ROLE_STUDENT) {
        const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
        rowRange.setBackground(COLOR_STUDENT);
      }
    }
  }

  // Step 17: Duplicate headers
  function duplicateHeaders() {
    const source = sheet.getRange('A1').offset(0, 0, 1, NUM_COLS);
    const target = sheet.getRange('G1');
    source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // Step 18: Resize all columns to fit data
  function resizeColumnsToFit() {
    const numColumns = sheet.getLastColumn();
    for (let i = 1; i <= numColumns; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  // Step 19: Clip all the columns
  function clipColumns() {
    sheet.getActiveRange().getDataRegion().setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Step 20: Add filter
  function addFilter() {
    sheet.getDataRange().createFilter();
  }

  // Execute all steps in order
  try {
    splitCSVIntoColumns();
    fillEmptyCells();
    compactPronouns();
    flagNewcomers();
    deleteNewAttendeesColumn();
    normalizeTechnologies('Skills', SKILLS_MAP);
    normalizeTechnologies('Note', SKILLS_MAP);
    copySkillsForCoaches();
    deleteSkillsAndRenameTutorialColumn();
    insertCheckboxColumn();
    insertGroupColumn();
    setGroupForStudentsAndAddValidation();
    sortAttendees();
    freezeTopRow();
    formatHeaderRow();
    formatCoachRows();
    formatStudentRows();
    duplicateHeaders();
    resizeColumnsToFit();
    clipColumns();
    addFilter();

    Utils.Utils.showInfo('Formatting completed successfully!');
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

function pasteSamplePairingCsvData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(1, 1, SAMPLE_CSV_DATA.length, 1);
  range.setValues(SAMPLE_CSV_DATA);
}

function registerAtRandom() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const isRegistered = Math.random() < 0.7; // 70% chance of being registered
    data[i][0] = isRegistered ? 'TRUE' : 'FALSE';
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function pairAtRandom() {
  const sheet = SpreadsheetApp.getActiveSheet();
  let data = sheet.getDataRange().getValues();

  // Collect registered coaches and students that aren't already paired
  const availableCoaches = [];
  const availableStudents = [];
  
  for (let i = 1; i < data.length; i++) {
    const reg1 = data[i][COL_REGISTERED_1 - 1];
    const name1 = data[i][COL_NAME_1 - 1];
    const role1 = data[i][COL_ROLE_1 - 1];
    const reg2 = data[i][COL_REGISTERED_2 - 1];
    const name2 = data[i][COL_NAME_2 - 1];
    
    // Check if left side is registered and right side is empty
    const leftRegistered = reg1 === 'TRUE' || reg1 === true;
    const rightEmpty = !reg2 || reg2 === 'FALSE' || reg2 === false || name2 === '' || name2 === '-';
    
    if (leftRegistered && rightEmpty && name1 && name1 !== '-') {
      if (role1 === ROLE_COACH) {
        availableCoaches.push(i);
      } else if (role1 === ROLE_STUDENT) {
        availableStudents.push(i);
      }
    }
  }

  if (availableCoaches.length === 0) {
    Utils.showInfo('No available coaches found for pairing.');
    return;
  }
  
  if (availableStudents.length === 0) {
    Utils.showInfo('No available students found for pairing.');
    return;
  }

  const coachAssignments = {};
  let pairedCount = 0;
  
  // Shuffle students for random pairing
  const shuffledStudents = [...availableStudents].sort(() => Math.random() - 0.5);
  
  for (const studentRowIdx of shuffledStudents) {
    if (availableCoaches.length === 0) {
      break;
    }
    
    // Find coaches that haven't reached their limit (2 students max)
    const availableCoachesForPairing = availableCoaches.filter(coachRowIdx => {
      return (coachAssignments[coachRowIdx] || 0) < 2;
    });
    
    if (availableCoachesForPairing.length === 0) {
      break; // No more coaches available
    }
    
    // Pick a random coach from available ones
    const randomCoachIdx = Math.floor(Math.random() * availableCoachesForPairing.length);
    const coachRowIdx = availableCoachesForPairing[randomCoachIdx];
    
    // Move coach data to student's right side
    const sourceRange = sheet.getRange(coachRowIdx + 1, COL_REGISTERED_1, 1, NUM_COLS);
    const targetRange = sheet.getRange(studentRowIdx + 1, COL_REGISTERED_2, 1, NUM_COLS);
    
    // Copy data instead of moving to preserve source for multiple assignments
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    // Color the target range as coach color
    targetRange.setBackground(COLOR_COACH);
    
    // Track assignments
    coachAssignments[coachRowIdx] = (coachAssignments[coachRowIdx] || 0) + 1;
    pairedCount++;
    
    // If coach has reached limit, remove from available list
    if (coachAssignments[coachRowIdx] >= 2) {
      const indexToRemove = availableCoaches.indexOf(coachRowIdx);
      if (indexToRemove > -1) {
        availableCoaches.splice(indexToRemove, 1);
      }
    }
  }
  
  Utils.showInfo(`Successfully paired ${pairedCount} students with coaches!`);
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

function setColorfulBackgrounds() {
  // Define group constants as [name, color] tuples
  const GROUP_BEGINNER = ["Beginner", "#ffd6e8"];    // Soft pink
  const GROUP_HTML_CSS = ["HTML/CSS", "#c7e9ff"];    // Light blue
  const GROUP_JAVA = ["Java", "#ffe5cc"];            // Peach
  const GROUP_JS = ["JS", "#fff9c4"];                // Pale yellow
  const GROUP_OTHER = ["Other", "#e1d5f8"];          // Lavender
  const GROUP_PYTHON = ["Python", "#c8f0d4"];        // Mint green
  const GROUP_REACT = ["React", "#b2e8f0"];          // Aqua
  const GROUP_UNKNOWN = ["Unknown", "#e8e8e8"];      // Light grey
  
  // Create array with all groups
  const groups = [
    GROUP_BEGINNER,
    GROUP_HTML_CSS,
    GROUP_JAVA,
    GROUP_JS,
    GROUP_OTHER,
    GROUP_PYTHON,
    GROUP_REACT,
    GROUP_UNKNOWN
  ];
  
  // Get the active spreadsheet and sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Fill cells starting from A1
  groups.forEach((group, index) => {
  // for (const [group, index] of groups) {
    const [name, color] = group;
    const cell = sheet.getRange(`A${index + 1}`);
    cell.setValue(name).setBackground(color);
  });
  // }
}
