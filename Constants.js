// Number of columns after the CSV data has been formatted
const NUM_COLS = 6

const COL_REGISTERED_1 = 1
const COL_NAME_1 = 2
const COL_ROLE_1 = 3
const COL_GROUP_1 = 4
const COL_SKILLS_TUTORIAL_1 = 5
const COL_NOTE_1 = 6

const COL_REGISTERED_2 = COL_REGISTERED_1 + NUM_COLS
const COL_NAME_2 = COL_NAME_1 + NUM_COLS
const COL_ROLE_2 = COL_ROLE_1 + NUM_COLS
const COL_GROUP_2 = COL_GROUP_1 + NUM_COLS
const COL_SKILLS_TUTORIAL_2 = COL_SKILLS_TUTORIAL_1 + NUM_COLS
const COL_NOTE_2 = COL_NOTE_1 + NUM_COLS

const COLOR_HEADER = '#999999';
const COLOR_COACH = '#d9ead3';
const COLOR_STUDENT = '#cfe2f3';

const ROLE_COACH = 'Coach'
const ROLE_STUDENT = 'Student'

const PROP_SORT_CRITERIA = 'SORT_CRITERIA';
const PROP_COACH_ROW_INDEX = 'COACH_ROW_INDEX';
const PROP_COACH_COL_INDEX = 'COACH_COL_INDEX';

const SAMPLE_CSV_DATA = [
  ['New attendee,Name,Role,Tutorial,Note,Skills'],
  ['false,Adrian Awesome (he),Student,Other,"",N/A'],
  ['false,Bella Bright (She),Student,React Project,I need help understanding components and structure of the design,N/A'],
  ['false,Carl Cheerful (He),Student,Python,"",N/A'],
  ['false,Diana Delightful,Coach,N/A,"","Full-Stack JavaScript, functional programming, Game Development, java, python, distributed systems"'],
  ['false,Ethan Excellent (He),Coach,N/A,"","javascript, TypeScript, nodejs, ReactJS, Nextjs, python"'],
  ['true,Fiona Fabulous (she/her),Coach,N/A,"","html, css, javascript, node, React, Next.js, react-router, Redux, TypeScript, styled-components, GraphQL"'],
  ['false,Gabriel Giggly (he/him),Student,React Project,"",N/A'],
  ['false,Hannah Happy (She),Coach,N/A,"Java, python, data science, backend ","java, python, SQL, nosql, web, c, Linux, rust"'],
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

const TUTORIAL_GROUP_MAP = {
  'HTML 1: Introducing HTML': 'HTML/CSS',
  'HTML/CSS 2: Introducing CSS': 'HTML/CSS',
  'HTML/CSS 3: Beyond the basics': 'HTML/CSS',
  'HTML/CSS 5: Dive into HTML5 & CSS3': 'HTML/CSS',
  'HTML/CSS 6: Advanced HTML5': 'HTML/CSS',
  'HTML/CSS/JS Project': 'HTML/CSS',
  "I don't know, I'm a complete beginner.": 'Beginner',
  'Java': 'Java',
  'JavaScript Project': 'JS',
  'JS: Beginning JavaScript': 'JS',
  'JS: Building your own app': 'JS',
  'JS: HTTP Requests, AJAX and APIs': 'JS',
  'JS: Introduction to JavaScript': 'JS',
  'JS: Introduction to JQuery': 'JS',
  'JS: Introduction to Testing': 'JS',
  'Other': 'Other',
  'Other programming language': 'Other',
  'Python': 'Python',
  'React Project': 'React',
};

// Define group constants as [name, color] tuples
const GROUP_BEGINNER = ["Beginner", "#ffd6e8"];    // Soft pink
const GROUP_HTML_CSS = ["HTML/CSS", "#c7e9ff"];    // Light blue
const GROUP_JAVA = ["Java", "#ffe5cc"];            // Peach
const GROUP_JS = ["JS", "#fff9c4"];                // Pale yellow
const GROUP_OTHER = ["Other", "#e1d5f8"];          // Lavender
const GROUP_PYTHON = ["Python", "#c8f0d4"];        // Mint green
const GROUP_REACT = ["React", "#b2e8f0"];          // Aqua
const GROUP_UNKNOWN = ["Unknown", "#e8e8e8"];      // Light grey

const GROUPS = [
  GROUP_BEGINNER,
  GROUP_HTML_CSS,
  GROUP_JAVA,
  GROUP_JS,
  GROUP_OTHER,
  GROUP_PYTHON,
  GROUP_REACT,
  GROUP_UNKNOWN
];
