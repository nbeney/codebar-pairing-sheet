class Format {
  static splitCSVIntoColumns() {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getActiveCell().getDataRegion().splitTextToColumns(",");
  }

  static fillEmptyCells() {
    const sheet = SpreadsheetApp.getActiveSheet();
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

  static compactPronouns() {
    const sheet = SpreadsheetApp.getActiveSheet();
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

  static flagNewcomers() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const newAttendeeColIndex = data[0].indexOf('New attendee');
    const nameColIndex = data[0].indexOf('Name');

    if (newAttendeeColIndex === -1 || nameColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][newAttendeeColIndex].toString().toLowerCase() === 'true') {
        data[i][nameColIndex] = data[i][nameColIndex] + ' ' + ICONS.newcomer;
      }
    }

    sheet.getDataRange().setValues(data);
  }

  static flagCoachesAndStudents() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');

    if (roleColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === 'Coach') {
        data[i][roleColIndex] = ROLE_COACH;
      } else if (data[i][roleColIndex] === 'Student') {
        data[i][roleColIndex] = ROLE_STUDENT;
      }
    }

    sheet.getDataRange().setValues(data);
  }

  static deleteNewAttendeesColumn() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const newAttendeeColIndex = data[0].indexOf('New attendee');

    if (newAttendeeColIndex !== -1) {
      sheet.deleteColumn(newAttendeeColIndex + 1);
    }
  }

  static normalizeTechnologies(columnName, skillsMap) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const colIndex = data[0].indexOf(columnName);

    if (colIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      let skills = data[i][colIndex].toString();

      if (skills && skills !== 'N/A') {
        for (const key of Object.keys(skillsMap)) {
          const regex = new RegExp('\\b' + key + '\\b', 'gi');
          skills = skills.replace(regex, skillsMap[key]);
        }

        data[i][colIndex] = skills;
      }
    }

    sheet.getDataRange().setValues(data);
  }

  static copySkillsForCoaches() {
    const sheet = SpreadsheetApp.getActiveSheet();
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

  static deleteSkillsAndRenameTutorialColumn() {
    const sheet = SpreadsheetApp.getActiveSheet();
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

  static insertRegisteredColumn() {
    const sheet = SpreadsheetApp.getActiveSheet();
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
        sheet.getRange(i, nameColIndex + NUM_COLS).insertCheckboxes();
      }
    }
  }

  static insertGroupColumn() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');
    const groupColIndex = roleColIndex + 1;

    if (roleColIndex === -1) return;

    // Insert the column

    sheet.insertColumnAfter(roleColIndex + 1);
    sheet.getRange(1, roleColIndex + 2).setValue('Group');

    // Add the data validation

    const groups = GROUPS.map(g => g.name);
    const val = SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInList(groups, true)
      .build();
    sheet.getRange(2, groupColIndex + 1, sheet.getLastRow() - 1).setDataValidation(val);

    // Add the conditional formatting

    const rules = sheet.getConditionalFormatRules();

    for (const group of GROUPS) {
      // Create rules for both column ranges
      const columnRanges = [
        { range: sheet.getRange(2, 1, sheet.getLastRow() - 1, NUM_COLS), groupColIndex: groupColIndex },
        { range: sheet.getRange(2, NUM_COLS + 1, sheet.getLastRow() - 1, NUM_COLS), groupColIndex: groupColIndex + NUM_COLS }
      ];

      for (const { range, groupColIndex: colIndex } of columnRanges) {
        const cond = group.name === GROUP_TBD.name ? '<>""' : `="${group.name}"`;
        const formula = `=$${String.fromCharCode(65 + colIndex)}2${cond}`;
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setBackground(group.color)
          .setFontColor('#000000')
          .setRanges([range])
          .build();
        rules.push(rule);
      }
    }
    sheet.setConditionalFormatRules(rules);
  }

  static setGroupForCoachesAndStudents() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const roleColIndex = data[0].indexOf('Role');
    const groupColIndex = data[0].indexOf('Group');
    const skillsTutorialColIndex = data[0].indexOf('Skills/Tutorial');

    if (roleColIndex === -1 || groupColIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColIndex] === ROLE_COACH) {
          data[i][groupColIndex] = GROUP_TBD.name;
      } else if (data[i][roleColIndex] === ROLE_STUDENT) {
        const skillsTutorial = data[i][skillsTutorialColIndex].toString();
        const group = TUTORIAL_GROUP_MAP[skillsTutorial];
        if (group !== undefined) {
          data[i][groupColIndex] = group.name;
        }
      }
    }
    sheet.getDataRange().setValues(data);
  }

  static sortAttendees() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const nameColIndex = data[0].indexOf('Name');

    if (nameColIndex === -1) return;

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      { column: nameColIndex + 1, ascending: true }
    ]);
  }

  static freezeTopRow() {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.setFrozenRows(1);
  }

  static formatHeaderRow() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('bold');
    headerRange.setBackground(COLOR_HEADER);
  }

  static duplicateHeaders() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const source = sheet.getRange('A1').offset(0, 0, 1, NUM_COLS);
    const target = sheet.getRange('G1');
    source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  static resizeColumnsToFit() {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.autoResizeColumns(1, sheet.getLastColumn());

    // Adjust specific columns to be wider, by an extra 15 pixels
    const data = sheet.getDataRange().getValues();
    const columnsToAdjust = ['Name', 'Role'];

    for (const columnName of columnsToAdjust) {
      const colIndex = data[0].indexOf(columnName);
      if (colIndex !== -1) {
        const currentWidth = sheet.getColumnWidth(colIndex + 1);
        sheet.setColumnWidth(colIndex + 1, currentWidth + 15);
      }
    }

    // Make columns 5 and 6 narrow enough to fit on laptop screens
    sheet.setColumnWidth(5, 300);
    sheet.setColumnWidth(6, 300);

    // Make columns 7-12 the same width as columns 1-6
    for (let i = 1; i <= NUM_COLS; i++) {
      const sourceWidth = sheet.getColumnWidth(i);
      sheet.setColumnWidth(i + NUM_COLS, sourceWidth);
    }
  }

  static clipColumns() {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getActiveRange().getDataRegion().setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  static addFilter() {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getDataRange().createFilter();
  }
}