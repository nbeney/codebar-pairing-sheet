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
    static pasteSamplePairingCsvData() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const range = sheet.getRange(1, 1, DEMO_CSV_DATA.length, 1);
        range.setValues(DEMO_CSV_DATA);
    }

    static registerAtRandom() {
        const sheet = SpreadsheetApp.getActiveSheet();
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            const isRegistered = Math.random() < 0.7; // 70% chance of being registered
            data[i][0] = isRegistered ? 'TRUE' : 'FALSE';
        }

        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    static pairAtRandom() {
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

            // // Color the target range as coach color
            // targetRange.setBackground(COLOR_COACH);

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

    static step1ResetSheet() {
        Tutorial.showStep(
            'Step 1: Reset Sheet',
            [
                'This will clear the entire sheet and prepare it for new data.',
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
            'Step 2: Paste Pairing CSV Data',
            [
                'This will paste sample workshop registration data into the sheet.',
                '',
                'Proceed with pasting data?'
            ],
            () => {
                Tutorial.pasteSamplePairingCsvData();
                Tutorial.step3FormatCsv();
            }
        );
    }

    static step3FormatCsv() {
        Tutorial.showStep(
            'Step 3: Format CSV Data',
            [
                'This will format the raw CSV data by:',
                '• Splitting into columns',
                '• Adding icons for roles',
                '• Creating group assignments',
                '• Adding checkboxes for registration',
                '',
                'Proceed with formatting?'
            ],
            () => {
                formatCsv();
                Tutorial.step4RegisterParticipants();
            }
        );
    }

    static step4RegisterParticipants() {
        Tutorial.showStep(
            'Step 4: Register Participants',
            [
                'This will randomly check participants as "present" for the workshop (about 70% attendance rate).',
                '',
                'Proceed with registration?'
            ],
            () => {
                Tutorial.registerAtRandom();
                Tutorial.step5SortParticipants();
            }
        );
    }

    static step5SortParticipants() {
        Tutorial.showStep(
            'Step 5: Sort Participants',
            [
                'This will sort participants by Group > Role > Name to make pairing easier.',
                '',
                'Proceed with sorting?'
            ],
            () => {
                sortByGroupRoleName();
                Tutorial.step6AssignCoachesToGroups();
            }
        );
    }

    static step6AssignCoachesToGroups() {
        Tutorial.showStep(
            'Step 6: Assign Coaches to Groups',
            [
                'In a real workshop, you would manually assign coaches to appropriate tutorial groups.',
                '',
                'For this demo, coaches will remain in the "Unknown" group (highlighted in light red).',
                '',
                'Proceed to next step?'
            ],
            () => {
                Tutorial.step7SortParticipants();
            }
        );
    }

    static step7SortParticipants() {
        Tutorial.showStep(
            'Step 7: Sort Participants',
            [
                'This will sort participants by Role > Group > Name to optimize pairing workflow.',
                '',
                'Proceed with sorting?'
            ],
            () => {
                sortByRoleGroupName();
                Tutorial.step8AssignCoachesToStudents();
            }
        );
    }

    static step8AssignCoachesToStudents() {
        Tutorial.showStep(
            'Step 8: Assign Coaches to Students',
            [
                'This will randomly pair registered coaches with registered students.',
                '',
                'Proceed with pairing?'
            ],
            () => {
                Tutorial.pairAtRandom();
                Tutorial.step9ShowPairings();
            }
        );
    }

    static step9ShowPairings() {
        Tutorial.showStep(
            'Step 9: Show Pairings',
            [
                'This will display a summary of all pairings, unpaired participants, and missing participants.',
                '',
                'Proceed to show pairings?'
            ],
            () => {
                showPairings();
                const ui = SpreadsheetApp.getUi();
                ui.alert(
                    '👣 Demo Complete!',
                    'The step-by-step demo is now complete. You can explore the pairings dialog and try other features.',
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
            '👣 ' + title,
            message,
            ui.ButtonSet.OK_CANCEL
        );

        if (result === ui.Button.OK) {
            okCallback();
        }
    }
}
