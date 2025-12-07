const DEMO_CSV_DATA = [
    ['New attendee,Name,Role,Tutorial,Note,Skills'],
    ['false,Adrian Awesome (he),Student,Other,"",N/A'],
    ['false,Bella Bright (She),Student,React Project,I need help understanding components and structure of the design,N/A'],
    ['false,Carl Cheerful (He),Student,Python,"",N/A'],
    ['false,Diana Delightful,Coach,N/A,"","Full-Stack JavaScript, functional programming, Game Development, java, python, distributed systems"'],
    ['false,Ethan Excellent (He),Coach,N/A,"","javascript, TypeScript, nodejs, ReactJS, Nextjs, python"'],
    ['true,Fiona Fabulous (she/her),Coach,N/A,"","html, css, javascript, node, React, Next.js, react-router, Redux, TypeScript, styled-components, GraphQL"'],
    ['false,Gabriel Giggly (he/him),Student,React Project,"",N/A'],
    ['false,"Hannah Happy (She, he)",Coach,N/A,"Java, python, data science, backend ","java, python, SQL, nosql, web, c, Linux, rust"'],
    ['false,Isaac Incredible (He/him),Student,Other,C#,N/A'],
    ['false,Jessica Jolly,Student,Java,"",N/A'],
    ['false,Kevin Kind (he/him/his),Coach,N/A,Junior software developer with experience with JavaScript/TypeScript,"javascript, React, Express, back-end, Front-end"'],
    ['false,Luna Lovely (she),Coach,N/A,"Python, Testing, Spark/Databricks/Data or can do most beginner stuff ",""'],
    ['false,Marcus Magnificent (He),Coach,N/A,"I am an android developer in Lloyds bank, knows Java, Kotlin and backend.","Android, java, Spring Boot, AWS, MySql, Angular, Ionic, kotlin"'],
    ['false,Nina Nimble (she/her),Student,Python,"",N/A'],
    ['false,"Oliver Optimistic (he, him)",Coach,N/A,"","Go, javascript, ruby, algorithms"'],
    ['false,Petra Playful (she),Coach,N/A,"",javascript'],
    ['false,Quinn Quirky (he),Student,React Project,changes not reflected on the deployed version,N/A'],
    ['false,Sophie Sparkling (She/her),Coach,N/A,"Happy to help any and all Java and JS projects, failing that just ask good questions.","java, javascript, ReactJS, SQL"'],
    ['false,Tyler Terrific (he/him),Student,Python,I\'d like help understanding classes and OOPs intuitively ,N/A'],
    ['false,Rosa Radiant,Coach,N/A,I specify in React on frontends and NodeJS on the backend,"html, css, javascript, React, node, REST, GraphQL, Git, Jest, Cypress, mocha, MonoDB, tailwind"'],
    ['true,Uma Unstoppable (she/her),Student,"I don\'t know, I\'m a complete beginner.","",N/A'],
    ['false,Victor Vibrant (He/Him),Student,"I don\'t know, I\'m a complete beginner.",First time coming to a codebar workshop so I\'m excited to join :),N/A'],
    ['false,Willow Wonderful,Student,Python,understanding API (for SQL server and for LLM),N/A'],
    ['false,Xander Xtraordinary (him),Coach,N/A,"","javascript/typescript, React, Remix.js/Next.js, HTML/CSS, version control"'],
    ['false,Yvonne Youthful,Student,Java,Java spring boot project ,N/A'],
    ['false,Zara Zippy (she/her),Student,React Project,"",N/A'],
];

class Tutorial {
    static doPasteCsvData() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const range = sheet.getRange(1, 1, DEMO_CSV_DATA.length, 1);
        range.setValues(DEMO_CSV_DATA);
    }

    static doRegisterParticipants() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const data = sheet.getDataRange().getValues();

        const namesToSkip = new Set([
            'Isaac Incredible [H]',
            'Sophie Sparkling [S]',
            'Willow Wonderful',
        ]);
        
        for (let i = 1; i < data.length; i++) {
            const name = data[i][COL_NAME_1 - 1];
            data[i][COL_REGISTERED_1 - 1] = !namesToSkip.has(name);
        }

        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

   static doAssignCoachesToGroups() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const data = sheet.getDataRange().getValues();

        const coachGroupMap = {
            'Diana Delightful': 'Python',
            'Ethan Excellent [H]': 'Python',
            'Fiona Fabulous': 'React',
            'Hannah Happy [H/S]': 'Python',
            'Kevin Kind [H]': 'Other',
            'Luna Lovely [S]': 'Python',
            'Marcus Magnificent [H]': 'Java',
            'Oliver Optimistic [H]': 'Beginner',
            'Petra Playful [S]': 'Other',
            'Rosa Radiant': 'React',
            'Xander Xtraordinary [H]': 'React',
        };

        for (let i = 1; i < data.length; i++) {
            const name = data[i][COL_NAME_1 - 1];
            if (coachGroupMap[name]) {
                data[i][COL_GROUP_1 - 1] = coachGroupMap[name];
            }
        }

        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    static doAssignCoachesToStudents() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const data = sheet.getDataRange().getValues();
        
        const studentToCoachMap = {
            // Student: Coach
            'Adrian Awesome [H]': 'Kevin Kind [H]',
            'Bella Bright [S]': 'Rosa Radiant',
            'Carl Cheerful [H]': 'Diana Delightful',
            'Gabriel Giggly [H]': 'Rosa Radiant',
            'Jessica Jolly': 'Marcus Magnificent [H]',
            'Nina Nimble [S]': 'Ethan Excellent [H]',
            'Quinn Quirky [H]': 'Xander Xtraordinary [H]',
            'Tyler Terrific [H]': 'Hannah Happy [H/S]',
            'Uma Unstoppable [S] 🐥': 'Oliver Optimistic [H]',
            'Victor Vibrant [H]': 'Oliver Optimistic [H]',
            'Yvonne Youthful': 'Marcus Magnificent [H]',
        };

        const assignedCoachRows = new Set();

        // Copy coach data to student rows
        for (let i = 1; i < data.length; i++) {
            const studentName = data[i][COL_NAME_1 - 1];
            const coachName = studentToCoachMap[studentName];
            if (coachName) {
                // Find the coach row
                for (let j = 1; j < data.length; j++) {
                    if (data[j][COL_NAME_1 - 1] === coachName) {
                        // Copy coach data to student columns
                        for (let col = 0; col < NUM_COLS; col++) {
                            data[i][COL_REGISTERED_2 - 1 + col] = data[j][COL_REGISTERED_1 - 1 + col];
                        }
                        
                        // Copy formatting from cols 1-6 to cols 7-12
                        const sourceRange = sheet.getRange(j + 1, 1, 1, NUM_COLS);
                        const targetRange = sheet.getRange(i + 1, COL_REGISTERED_2, 1, NUM_COLS);
                        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
                        
                        assignedCoachRows.add(j);
                        break;
                    }
                }
            }
        }

        // Clear unneeded coach rows, columns 1-6 if role is Coach and name is in the map
        for (let i = 1; i < data.length; i++) {
            const name = data[i][COL_NAME_1 - 1];
            const role = data[i][COL_ROLE_1 - 1];
            if (role === ROLE_COACH && Object.values(studentToCoachMap).includes(name)) {
                for (let col = 0; col < NUM_COLS; col++) {
                    data[i][col] = '';
                }
            }
        }

        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        
        // Clear formatting from assigned coach rows
        for (const rowIndex of assignedCoachRows) {
            const coachRange = sheet.getRange(rowIndex + 1, 1, 1, NUM_COLS);
            coachRange.clearFormat();
        }
    }

    static step1ResetSheet() {
        Tutorial.showStep(
            'Step 1: Reset',
            [
                'This will clear the entire sheet and prepare it for new data.',
                'It is unnecessary if you start with a blank sheet.',
                '',
                'Proceed with reset?'
            ],
            () => {
                reset();
                Tutorial.step2PastePairingCsvData();
            }
        );
    }

    static step2PastePairingCsvData() {
        Tutorial.showStep(
            'Step 2: Paste CSV',
            [
                'You will manually paste into A1 the participants list that you have exported from the workshop web page using the "Pairing CSV" function.',
                '',
                'In this tutorial we will use some sample data instead.',
                '',
                'Proceed with pasting data?'
            ],
            () => {
                Tutorial.doPasteCsvData();
                Tutorial.step3FormatCsv();
            }
        );
    }

    static step3FormatCsv() {
        Tutorial.showStep(
            'Step 3: Format',
            [
                'This will format the raw CSV data to make it easier to work with later',
                '',
                'Proceed with formatting?'
            ],
            () => {
                format();
                Tutorial.step4RegisterParticipants();
            }
        );
    }

    static step4RegisterParticipants() {
        Tutorial.showStep(
            'Step 4: Register participants',
            [
                'You will manually check the participants as they arrive to the workshop.',
                '',
                'In this tutorial we will use some sample data instead.',
                '',
                'Proceed with registration?'
            ],
            () => {
                Tutorial.doRegisterParticipants();
                Tutorial.step5SortParticipants();
            }
        );
    }

    static step5SortParticipants() {
        Tutorial.showStep(
            'Step 5: Sort participants',
            [
                'This will sort participants by Group > Role > Name to make pairing easier.',
                '',
                'Proceed with sorting?'
            ],
            () => {
                sortListByGroupRoleName();
                Tutorial.step6AssignCoachesToGroups();
            }
        );
    }

    static step6AssignCoachesToGroups() {
        Tutorial.showStep(
            'Step 6: Assign coaches to groups',
            [
                'You will manually assign the coaches to groups, based on their skills and what the students need.',
                '',
                'In this tutorial we will use some sample data instead.',
                '',
                'Proceed to next step?'
            ],
            () => {
                Tutorial.doAssignCoachesToGroups();
                Tutorial.step7SortParticipants();
            }
        );
    }

    static step7SortParticipants() {
        Tutorial.showStep(
            'Step 7: Sort participants',
            [
                'This will sort participants by Group > Role > Name to make pairing easier.',
                '',
                'Proceed with sorting?'
            ],
            () => {
                sortListByGroupRoleName();
                Tutorial.step8AssignCoachesToStudents();
            }
        );
    }

    static step8AssignCoachesToStudents() {
        Tutorial.showStep(
            'Step 8: Assign coaches to students',
            [
                'You will manually assign coaches to students based on their groups.', 
                '',
                'This usually involves selecting a coach with all its columns, cutting it with Ctrl-X and pasting it immediately to the right of a students with Ctrl-V.',
                '',
                'In this tutorial we will use some sample data instead.',
                '',
                'Proceed with pairing?'
            ],
            () => {
                Tutorial.doAssignCoachesToStudents();
                Tutorial.step9ShowPairings();
            }
        );
    }

    static step9ShowPairings() {
        Tutorial.showStep(
            'Step 9: Show pairings',
            [
                'This will display a summary of all pairings, unpaired participants, and missing participants.',
                '',
                'Proceed to show pairings?'
            ],
            () => {
                showPairings();
                const ui = SpreadsheetApp.getUi();
                ui.alert(
                    `${ICONS.tutorial} Tutorial complete!`,
                    'The tutorial is now complete. You can explore the pairings dialog and try other features.',
                    ui.ButtonSet.OK
                );
            }
        );
    }

    static showStep(title, lines, okCallback) {
        SpreadsheetApp.flush();
        const ui = SpreadsheetApp.getUi();
        const message = lines.join('\n');
        const result = ui.alert(
            `${ICONS.tutorial} ` + title,
            message,
            ui.ButtonSet.OK_CANCEL
        );

        if (result === ui.Button.OK) {
            okCallback();
        }
    }
}
