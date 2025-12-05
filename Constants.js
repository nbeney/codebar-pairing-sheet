// Number of columns after the CSV data has been formatted
const NUM_COLS = 6;

const COL_REGISTERED_1 = 1;
const COL_NAME_1 = 2;
const COL_ROLE_1 = 3;
const COL_GROUP_1 = 4;
const COL_SKILLS_TUTORIAL_1 = 5;
const COL_NOTE_1 = 6;

const COL_REGISTERED_2 = COL_REGISTERED_1 + NUM_COLS;
const COL_NAME_2 = COL_NAME_1 + NUM_COLS;
const COL_ROLE_2 = COL_ROLE_1 + NUM_COLS;
const COL_GROUP_2 = COL_GROUP_1 + NUM_COLS;
const COL_SKILLS_TUTORIAL_2 = COL_SKILLS_TUTORIAL_1 + NUM_COLS;
const COL_NOTE_2 = COL_NOTE_1 + NUM_COLS;

const ICONS = {
  newcomer: '🐥',
  host: '🏠',
  coach: '😇',
  student: '🎓',
  pair: '😇🎓',
  tutorial: '👣',
};

const HEADER_REGISTERED = '??';
const HEADER_NAME = `Name (${ICONS.newcomer}New ${ICONS.host}Host)`;
const HEADER_ROLE = 'Role';
const HEADER_GROUP = 'Group';
const HEADER_SKILLS_TUTORIAL = 'Skills/Tutorial';
const HEADER_NOTE = 'Note';

const COLOR_HEADER = '#999999';

const ROLE_COACH = 'Coach ' + ICONS.coach;
const ROLE_STUDENT = 'Student ' + ICONS.student;

const PROP_SORT_CRITERIA = 'SORT_CRITERIA';
const PROP_COACH_ROW_INDEX = 'COACH_ROW_INDEX';
const PROP_COACH_COL_INDEX = 'COACH_COL_INDEX';

const SKILLS_MAP = {
  'android': 'Android',
  'angular': 'Angular',
  'aws': 'AWS',
  'c': 'C',
  'css': 'CSS',
  'cypress': 'Cypress',
  'express': 'Express',
  'git': 'Git',
  'go': 'Go',
  'golang': 'Go',
  'graphql': 'GraphQL',
  'html': 'HTML',
  'ionic': 'Ionic',
  'java': 'Java',
  'javascript': 'JS',
  'jest': 'Jest',
  'kotlin': 'Kotlin',
  'linux': 'Linux',
  'mocha': 'Mocha',
  'mongodb': 'MongoDB',
  'mysql': 'MySQL',
  'nextjs': 'NextJS',
  'next.js': 'NextJS',
  'node': 'Node',
  'nodejs': 'Node',
  'nosql': 'NoSQL',
  'python': 'Python',
  'react': 'React',
  'reactjs': 'React',
  'redux': 'Redux',
  'rest': 'REST',
  'ruby': 'Ruby',
  'rust': 'Rust',
  'sql': 'SQL',
  'tailwind': 'Tailwind',
  'typescript': 'TS',
};

class Group {
  constructor(name, color, tutorials) {
    this.name = name;
    this.color = color;
    this.tutorials = tutorials;
  }
}

const GROUP_TBD = new Group("TBD", "#e8e8e8", []);

const GROUPS = [
  new Group("Beginner", "#ffd6e8", [
    "I don't know, I'm a complete beginner.",
  ]),
  new Group("Git", "#d4f1f4", [
    'Introduction to version control',
  ]),
  new Group("HTML/CSS", "#c7e9ff", [
    'HTML 1: Introducing HTML',
    'HTML/CSS 2: Introducing CSS',
    'HTML/CSS 3: Beyond the basics',
    'HTML/CSS 5: Dive into HTML5 & CSS3',
    'HTML/CSS 6: Advanced HTML5',
    'HTML/CSS/JS Project',
  ]),
  new Group("Java", "#ffe5cc", [
    'Java',
  ]),
  new Group("JS", "#fff9c4", [
    'JS Project',
    'JS: Beginning JS',
    'JS: Building your own app',
    'JS: HTTP Requests, AJAX and APIs',
    'JS: Introduction to JS',
    'JS: Introduction to JQuery',
    'JS: Introduction to Testing',
  ]),
  new Group("Other", "#e1d5f8", [
    'Other',
    'Other programming language',
  ]),
  new Group("Python", "#c8f0d4", [
    'Python',
  ]),
  new Group("React", "#ffddc1", [
    'React Project',
  ]),
  GROUP_TBD,
];

const TUTORIAL_GROUP_MAP = {};
for (const group of GROUPS) {
  for (const tutorial of group.tutorials) {
    TUTORIAL_GROUP_MAP[tutorial] = group;
  }
}
