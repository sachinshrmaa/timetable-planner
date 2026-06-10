import { useState, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { db } from './firebase.js';
import { doc, setDoc, onSnapshot, collection, deleteDoc } from 'firebase/firestore';


// SVG Icons Components for Visual Excellence
const UploadIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l3 3m-3-3l-3 3M6.75 19.5a4.5 4.5 0 01-1.41-8.775 5.25 5.25 0 0110.233-2.33 3 3 0 013.758 3.848A3.752 3.752 0 0118 19.5H6.75z" />
  </svg>
);



const CheckCircleIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
  </svg>
);

const ExclamationTriangleIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
  </svg>
);

const InfoIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
  </svg>
);

const SearchIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
  </svg>
);

const DownloadIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
  </svg>
);

const RefreshIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 1121.21 7.89M9 11l3-3 3 3m-3-3v12" />
  </svg>
);

const CalendarIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 012.25-2.25h13.5A2.25 2.25 0 0121 7.5v11.25m-18 0A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75m-18 0v-7.5A2.25 2.25 0 015.25 9h13.5A2.25 2.25 0 0121 11.25v7.5" />
  </svg>
);

const SparklesIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M9.813 15.904L9 21l-1.813-5.096L2.096 15 7.2 13.187 9 8l1.813 5.187 5.096 1.813-5.096 1.904zM19.004 3.004L18.5 6l-.5-2.996L15 2.5l2.996-.5L18.5 0l.5 2.996L22 2.5l-2.996.5.004.004z" />
  </svg>
);

const ScissorsIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" />
  </svg>
);

const ArrowRightIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M14 5l7 7m0 0l-7 7m7-7H3" />
  </svg>
);

// Helper to generate initials-based course codes
const generateCourseCode = (name, id) => {
  if (!name) return `CRSE-${id}`;
  const words = name.split(/\s+/);
  const initials = words.map(w => w[0]).join('').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return `${initials.substring(0, 4)}-${100 + (id % 900)}`;
};

// Hashing helper for premium pastel color schemes based on subject name
const getSubjectColor = (courseName) => {
  if (!courseName) return { bg: 'bg-slate-50', border: 'border-slate-200', text: 'text-slate-800', badge: 'bg-slate-100 text-slate-700 border-slate-200' };
  
  let hash = 0;
  for (let i = 0; i < courseName.length; i++) {
    hash = courseName.charCodeAt(i) + ((hash << 5) - hash);
  }
  
  const palettes = [
    { bg: 'bg-indigo-50/70', border: 'border-indigo-200', text: 'text-indigo-900', badge: 'bg-indigo-100 text-indigo-700 border-indigo-200' },
    { bg: 'bg-emerald-50/70', border: 'border-emerald-200', text: 'text-emerald-900', badge: 'bg-emerald-100 text-emerald-700 border-emerald-200' },
    { bg: 'bg-amber-50/70', border: 'border-amber-200', text: 'text-amber-900', badge: 'bg-amber-100 text-amber-700 border-amber-200' },
    { bg: 'bg-rose-50/70', border: 'border-rose-200', text: 'text-rose-900', badge: 'bg-rose-100 text-rose-700 border-rose-200' },
    { bg: 'bg-sky-50/70', border: 'border-sky-200', text: 'text-sky-900', badge: 'bg-sky-100 text-sky-700 border-sky-200' },
    { bg: 'bg-fuchsia-50/70', border: 'border-fuchsia-200', text: 'text-fuchsia-900', badge: 'bg-fuchsia-100 text-fuchsia-700 border-fuchsia-200' },
    { bg: 'bg-teal-50/70', border: 'border-teal-200', text: 'text-teal-900', badge: 'bg-teal-100 text-teal-700 border-teal-200' },
    { bg: 'bg-orange-50/70', border: 'border-orange-200', text: 'text-orange-900', badge: 'bg-orange-100 text-orange-700 border-orange-200' }
  ];
  
  const index = Math.abs(hash) % palettes.length;
  return palettes[index];
};

const DEMO_DATA = [
  { id: 1, programme: "BCA 1st", semester: "Semester 1", courseName: "Programming in C", faculty: "Sachin Sharma", hours: 4, courseCode: "BCA-101", classroom: "Room 101", school: "School of Computing", mergeCode: "", sessionBlockSize: 0 },
  { id: 2, programme: "BCA 1st", semester: "Semester 1", courseName: "Computer Fundamentals", faculty: "Sachin Sharma", hours: 3, courseCode: "BCA-102", classroom: "Room 102", school: "School of Computing", mergeCode: "", sessionBlockSize: 0 },
  { id: 3, programme: "BCA 1st", semester: "Semester 1", courseName: "Programming Lab C", faculty: "Sachin Sharma", hours: 4, courseCode: "BCA-103L", classroom: "Lab 1", school: "School of Computing", mergeCode: "", sessionBlockSize: 0 },
  { id: 4, programme: "BCA 1st", semester: "Semester 1", courseName: "Technical Mathematics", faculty: "Prof. Carl Gauss", hours: 3, courseCode: "BCA-104", classroom: "Room 101", school: "School of Sciences", mergeCode: "", sessionBlockSize: 0 },
  
  { id: 5, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Data Structures", faculty: "Dr. Alan Turing", hours: 4, courseCode: "CSE-201", classroom: "Room 201", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 },
  { id: 6, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Database Systems", faculty: "Dr. Grace Hopper", hours: 3, courseCode: "CSE-202", classroom: "Room 202", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 },
  { id: 7, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Software Engineering Lab", faculty: "Dr. Alan Turing", hours: 4, courseCode: "CSE-203L", classroom: "Lab 2", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 },
  { id: 8, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Discrete Mathematics", faculty: "Prof. Carl Gauss", hours: 3, courseCode: "CSE-204", classroom: "Room 201", school: "School of Sciences", mergeCode: "", sessionBlockSize: 0 },
  
  { id: 9, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Machine Learning", faculty: "Dr. Geoffrey Hinton", hours: 4, courseCode: "AIM-501", classroom: "Room 301", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 },
  { id: 10, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Neural Networks Lab", faculty: "Dr. Geoffrey Hinton", hours: 4, courseCode: "AIM-502L", classroom: "Lab 3", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 },
  { id: 11, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Optimization Techniques", faculty: "Prof. Carl Gauss", hours: 3, courseCode: "AIM-503", classroom: "Room 302", school: "School of Sciences", mergeCode: "", sessionBlockSize: 0 },
  { id: 12, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Natural Language Processing", faculty: "Dr. Grace Hopper", hours: 3, courseCode: "AIM-504", classroom: "Room 301", school: "School of Engineering", mergeCode: "", sessionBlockSize: 0 }
];

// Pure wrapper helper for time-based merge code to satisfy linter purity checks
const getMergeCode = () => `MERGE-${Date.now().toString().slice(-4)}`;

const INITIAL_SLOTS = [
  { id: 0, label: "Slot 1", time: "09:30 AM - 10:20 AM" },
  { id: 1, label: "Slot 2", time: "10:20 AM - 11:10 AM" },
  { id: 2, label: "Slot 3", time: "11:20 AM - 12:10 PM" },
  { id: 3, label: "Slot 4", time: "12:10 PM - 01:00 PM" },
  { id: 4, label: "Slot 5", time: "01:50 PM - 02:40 PM" },
  { id: 5, label: "Slot 6", time: "02:40 PM - 03:30 PM" },
  { id: 6, label: "Slot 7", time: "03:30 PM - 04:20 PM" }
];

function App() {
  const [courses, setCourses] = useState([]);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedSchoolFilter, setSelectedSchoolFilter] = useState("All");

  // Timetable Scheduling State
  const [timetableData, setTimetableData] = useState(null);
  const [rawTimetableBackup, setRawTimetableBackup] = useState(null); // Keep backup for resets
  const [unscheduledLog, setUnscheduledLog] = useState([]);
  const [facultyWorkload, setFacultyWorkload] = useState({});
  const [selectedGroup, setSelectedGroup] = useState("");

  // Mode Selection State (Mode A: Student Programme, Mode B: Faculty)
  const [viewMode, setViewMode] = useState("programme"); // "programme" or "faculty"
  const [selectedFaculty, setSelectedFaculty] = useState("");

  // Interactive Re-arrangement State
  const [moveSource, setMoveSource] = useState(null); // { group, day, slot }
  const [moveMessage, setMoveMessage] = useState(null); // { type: 'success' | 'error', text }

  // New states for enhancements
  const [portalMode, setPortalMode] = useState("admin"); // "admin" or "student"
  const [slots, setSlots] = useState(INITIAL_SLOTS);
  const [fixedSlots, setFixedSlots] = useState([]);
  const [displayOptions, setDisplayOptions] = useState({
    showName: true,
    showCode: true,
    showFaculty: true,
    showClassroom: true,
    showSchool: false
  });
  const [editingCourse, setEditingCourse] = useState(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  const fileInputRef = useRef(null);
  const gridRef = useRef(null);
  const [pdfLoading, setPdfLoading] = useState(false);

  const TARGET_COLUMNS = {
    programme: 'Programme',
    semester: 'Semester',
    courseName: 'Course name',
    faculty: 'Teaching Faculty',
    hours: 'Contact teaching Hours/Week'
  };

  const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

  const findHeaderMapping = (headers) => {
    const mapping = {};
    const aliases = {
      programme: ['programme', 'program', 'course programme', 'branch', 'dept', 'department'],
      semester: ['semester', 'sem', 'term', 'year'],
      courseName: ['course name', 'coursename', 'course', 'subject name', 'subject', 'title'],
      faculty: ['teaching faculty', 'faculty', 'teacher', 'instructor', 'professor', 'lecturer'],
      hours: ['contact teaching hours/week', 'contact teaching hours', 'contact hours', 'hours/week', 'teaching hours', 'hours', 'credit hours', 'credits'],
      courseCode: ['course code', 'coursecode', 'code', 'subject code', 'subjectcode'],
      classroom: ['classroom', 'room', 'class room', 'room no', 'room number', 'venue'],
      school: ['school', 'school name', 'offering school', 'institution']
    };

    Object.keys(aliases).forEach(key => {
      const aliasList = aliases[key];
      const matchedHeader = headers.find(h => {
        const normalizedHeader = String(h).trim().toLowerCase();
        return aliasList.includes(normalizedHeader);
      });
      if (matchedHeader) {
        mapping[key] = matchedHeader;
      }
    });
    return mapping;
  };

  const parseExcelFile = (file) => {
    setLoading(true);
    setError(null);
    setFileName(file.name);
    setTimetableData(null);
    setRawTimetableBackup(null);
    setUnscheduledLog([]);
    setMoveSource(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        let targetSheetName = "Course_List";
        let sheetExists = workbook.SheetNames.includes(targetSheetName);

        if (!sheetExists) {
          if (workbook.SheetNames.length > 0) {
            targetSheetName = workbook.SheetNames[0];
          } else {
            throw new Error("Workbook contains no sheets.");
          }
        }

        const worksheet = workbook.Sheets[targetSheetName];
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (rawRows.length < 1) {
          throw new Error(`The sheet "${targetSheetName}" appears to be empty.`);
        }

        let headerRowIndex = 0;
        let mapping = null;
        let bestMatchCount = -1;

        // Scan the first 15 rows to find the row that matches the maximum number of headers
        const scanLimit = Math.min(15, rawRows.length);
        for (let r = 0; r < scanLimit; r++) {
          const row = rawRows[r];
          if (!row || !Array.isArray(row)) continue;
          
          const headersCandidate = row.map(h => String(h || "").trim());
          const mappingCandidate = findHeaderMapping(headersCandidate);
          const matchCount = Object.keys(mappingCandidate).length;
          
          if (matchCount > bestMatchCount) {
            bestMatchCount = matchCount;
            mapping = mappingCandidate;
            headerRowIndex = r;
          }
        }

        if (bestMatchCount <= 0 || !mapping) {
          throw new Error(`Could not detect the header row in sheet "${targetSheetName}". Please make sure the sheet contains the columns described in the template.`);
        }

        const headers = rawRows[headerRowIndex].map(h => String(h || "").trim());

        const missingColumns = [];
        Object.entries(TARGET_COLUMNS).forEach(([key, name]) => {
          if (!mapping[key]) {
            missingColumns.push(`"${name}"`);
          }
        });

        if (missingColumns.length > 0) {
          throw new Error(`Could not locate columns matching: ${missingColumns.join(', ')} in sheet "${targetSheetName}". Please align your headers with the template.`);
        }

        const extractedData = [];
        for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.every(val => val === "")) continue;

          const getValByHeader = (headerName) => {
            const headerIndex = headers.indexOf(headerName);
            return headerIndex !== -1 ? row[headerIndex] : "";
          };

          const rawName = String(getValByHeader(mapping.courseName)).trim();
          const courseItem = {
            id: i,
            programme: String(getValByHeader(mapping.programme)).trim(),
            semester: String(getValByHeader(mapping.semester)).trim(),
            courseName: rawName,
            faculty: String(getValByHeader(mapping.faculty)).trim(),
            hours: parseFloat(getValByHeader(mapping.hours)) || 0,
            courseCode: mapping.courseCode && getValByHeader(mapping.courseCode)
              ? String(getValByHeader(mapping.courseCode)).trim()
              : generateCourseCode(rawName, i),
            classroom: mapping.classroom && getValByHeader(mapping.classroom)
              ? String(getValByHeader(mapping.classroom)).trim()
              : "",
            school: mapping.school && getValByHeader(mapping.school)
              ? String(getValByHeader(mapping.school)).trim()
              : "",
            mergeCode: "",
            sessionBlockSize: 0 // 0 means auto-detect
          };

          if (courseItem.programme || courseItem.courseName || courseItem.faculty) {
            extractedData.push(courseItem);
          }
        }

        if (extractedData.length === 0) {
          throw new Error(`No course rows found in "${targetSheetName}" sheet.`);
        }

        setCourses(extractedData);
        updateFirestore({
          courses: extractedData,
          fileName: file.name,
          timetableData: null,
          rawTimetableBackup: null,
          unscheduledLog: [],
          facultyWorkload: {}
        });
      } catch (err) {
        console.error(err);
        setError(err.message || "An error occurred while parsing the file.");
        setCourses([]);
        updateFirestore({
          courses: [],
          timetableData: null,
          rawTimetableBackup: null,
          unscheduledLog: [],
          facultyWorkload: {}
        });
      } finally {
        setLoading(false);
      }
    };

    reader.onerror = () => {
      setError("Failed to read file.");
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      parseExcelFile(file);
    }
  };

  const handleFileSelect = (e) => {
    if (e.target.files && e.target.files[0]) {
      parseExcelFile(e.target.files[0]);
    }
  };

  const handleDownloadTemplate = () => {
    const headers = ['Programme', 'Semester', 'Course name', 'Teaching Faculty', 'Contact teaching Hours/Week'];
    const data = DEMO_DATA.map(item => [item.programme, item.semester, item.courseName, item.faculty, item.hours]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    ws['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 35 }, { wch: 25 }, { wch: 28 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Course_List");
    XLSX.writeFile(wb, "Timetable_Course_List_Template.xlsx");
  };

  const handleDownloadFixedSlotsTemplate = () => {
    const headers = [
      'Course name', 
      'Course code', 
      'Teaching Faculty', 
      'Classroom', 
      'School', 
      'Day', 
      'Slots', 
      'Target Type', 
      'Target Value'
    ];
    const demoFs = [
      ['University Clubs', 'CLUB-101', 'Clubs Coordinator', 'Auditorium', 'All Schools', 'Wednesday', 'Slot 5, Slot 6', 'all', ''],
      ['BCA Seminar', 'SEM-101', 'Dr. Alan Turing', 'Seminar Hall 1', 'School of Computing', 'Friday', 'Slot 3', 'group', 'BCA 1st - Semester 1'],
      ['B.Tech Engineering SkillDrill', 'SKILL-201', 'Skill Dept', 'Workshops', 'School of Engineering', 'Monday', 'Slot 6, Slot 7', 'semester', 'Semester 3']
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers, ...demoFs]);
    ws['!cols'] = [
      { wch: 25 }, { wch: 15 }, { wch: 22 }, { wch: 18 }, { wch: 22 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 25 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fixed_Schedules");
    XLSX.writeFile(wb, "Fixed_Schedules_Template.xlsx");
  };

  const parseFixedSlotsExcelFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        let targetSheetName = "Fixed_Schedules";
        let sheetExists = workbook.SheetNames.includes(targetSheetName);

        if (!sheetExists) {
          if (workbook.SheetNames.length > 0) {
            targetSheetName = workbook.SheetNames[0];
          } else {
            throw new Error("Workbook contains no sheets.");
          }
        }

        const worksheet = workbook.Sheets[targetSheetName];
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (rawRows.length < 1) {
          throw new Error(`The sheet "${targetSheetName}" appears to be empty.`);
        }

        let headerRowIndex = 0;
        let mapping = null;
        let bestMatchCount = -1;

        const aliases = {
          courseName: ['course name', 'name', 'activity name', 'activity', 'title'],
          courseCode: ['course code', 'code', 'activity code', 'id'],
          faculty: ['teaching faculty', 'faculty', 'coordinator', 'teacher', 'instructor'],
          classroom: ['classroom', 'room', 'venue', 'class room'],
          school: ['school', 'department', 'dept', 'offering school'],
          day: ['day', 'weekday', 'day of week'],
          slots: ['slots', 'time slots', 'slot numbers', 'slot index'],
          targetType: ['target type', 'targettype', 'type'],
          targetValue: ['target value', 'targetvalue', 'value']
        };

        const scanLimit = Math.min(15, rawRows.length);
        for (let r = 0; r < scanLimit; r++) {
          const row = rawRows[r];
          if (!row || !Array.isArray(row)) continue;
          const headersCandidate = row.map(h => String(h || "").trim().toLowerCase());
          
          const mappingCandidate = {};
          Object.keys(aliases).forEach(key => {
            const matchedHeader = headersCandidate.find(h => aliases[key].includes(h));
            if (matchedHeader) {
              mappingCandidate[key] = row[headersCandidate.indexOf(matchedHeader)];
            }
          });

          const matchCount = Object.keys(mappingCandidate).length;
          if (matchCount > bestMatchCount) {
            bestMatchCount = matchCount;
            mapping = mappingCandidate;
            headerRowIndex = r;
          }
        }

        if (bestMatchCount <= 2 || !mapping) {
          throw new Error("Could not detect the header row for fixed schedules. Please ensure columns match the template.");
        }

        const headers = rawRows[headerRowIndex].map(h => String(h || "").trim());
        const extractedLocks = [];

        for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.every(val => val === "")) continue;

          const getValByHeader = (headerName) => {
            const headerIndex = headers.indexOf(headerName);
            return headerIndex !== -1 ? row[headerIndex] : "";
          };

          const rawName = String(getValByHeader(mapping.courseName)).trim();
          const rawCode = String(getValByHeader(mapping.courseCode)).trim();
          const rawFaculty = String(getValByHeader(mapping.faculty)).trim();
          const rawDayStr = String(getValByHeader(mapping.day)).trim().toLowerCase();
          const rawSlotsStr = String(getValByHeader(mapping.slots)).trim();
          const rawTargetType = String(getValByHeader(mapping.targetType)).trim().toLowerCase() || "all";
          const rawTargetValue = String(getValByHeader(mapping.targetValue)).trim();

          if (!rawName || !rawDayStr || !rawSlotsStr) continue;

          const dayIndexMap = {
            monday: 0, mon: 0,
            tuesday: 1, tue: 1,
            wednesday: 2, wed: 2,
            thursday: 3, thu: 3,
            friday: 4, fri: 4
          };
          const dayIdx = dayIndexMap[rawDayStr];
          if (dayIdx === undefined) continue;

          const slotIndexes = [];
          const slotTokens = rawSlotsStr.split(/[,;+]/).map(s => s.trim().toLowerCase());
          
          slotTokens.forEach(tok => {
            const matchedSlot = slots.find(s => s.label.toLowerCase() === tok || s.label.toLowerCase().replace(/\s+/g, '') === tok);
            if (matchedSlot) {
              slotIndexes.push(matchedSlot.id);
            } else {
              const numVal = parseInt(tok.replace(/[^0-9]/g, ''));
              if (!isNaN(numVal)) {
                if (numVal >= 1 && numVal <= 7) {
                  slotIndexes.push(numVal - 1);
                } else if (numVal >= 0 && numVal <= 6) {
                  slotIndexes.push(numVal);
                }
              }
            }
          });

          if (slotIndexes.length === 0) continue;

          extractedLocks.push({
            id: `fs-bulk-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 4)}`,
            courseName: rawName,
            courseCode: rawCode || `FIXED-${i}`,
            faculty: rawFaculty || "University Coordinator",
            classroom: String(getValByHeader(mapping.classroom)).trim() || "TBD",
            school: String(getValByHeader(mapping.school)).trim() || "University Wide",
            day: dayIdx,
            slots: [...new Set(slotIndexes)].sort((a, b) => a - b),
            targetType: ['all', 'semester', 'group'].includes(rawTargetType) ? rawTargetType : "all",
            targetValue: rawTargetValue
          });
        }

        if (extractedLocks.length === 0) {
          throw new Error("No valid fixed schedule rows parsed. Make sure Day and Slots match template parameters.");
        }

        const nextFixedSlots = [...fixedSlots, ...extractedLocks];
        setFixedSlots(nextFixedSlots);
        updateFirestore({ fixedSlots: nextFixedSlots });
        setMoveMessage({ type: 'success', text: `Successfully imported ${extractedLocks.length} fixed schedule activities! Click "Generate Timetables" to apply changes.` });
        setTimeout(() => setMoveMessage(null), 5000);
      } catch (err) {
        console.error(err);
        alert(err.message || "An error occurred while parsing the fixed schedules file.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleLoadDemoData = () => {
    setCourses(DEMO_DATA);
    setFileName("Loaded Demo Dataset");
    setError(null);
    setTimetableData(null);
    setRawTimetableBackup(null);
    setUnscheduledLog([]);
    setMoveSource(null);
    updateFirestore({
      courses: DEMO_DATA,
      fileName: "Loaded Demo Dataset",
      timetableData: null,
      rawTimetableBackup: null,
      unscheduledLog: [],
      facultyWorkload: {}
    });
  };

  // Helper to check if a block crosses any of the breaks
  const crossesBreak = (startSlot, duration) => {
    // Break 1: after Slot 2 (index 1) -> between 1 and 2
    if (startSlot <= 1 && (startSlot + duration) > 2) return true;
    // Break 2: after Slot 4 (index 3) -> between 3 and 4
    if (startSlot <= 3 && (startSlot + duration) > 4) return true;
    return false;
  };

  // Helper to determine the session block sizes
  const getSessionBlocks = (c) => {
    if (c.sessionBlockSize && c.sessionBlockSize > 0) {
      return c.sessionBlockSize;
    }
    const nameLower = c.courseName.toLowerCase();
    if (nameLower.includes("graphics") || nameLower.includes("workshop")) {
      return 3;
    }
    if (nameLower.includes("lab") || nameLower.includes("applied")) {
      return 2;
    }
    return 1;
  };

  // --- SCHEDULING ENGINE ---
  const handleGenerateTimetable = () => {
    if (courses.length === 0) return;

    // Clear move variables
    setMoveSource(null);
    setMoveMessage(null);

    // 1. Find all unique groups (Programme + Semester)
    const groups = [...new Set(courses.map(c => `${c.programme} - ${c.semester}`))];

    // 2. Initialize Timetable grids
    const tempTimetables = {};
    groups.forEach(g => {
      tempTimetables[g] = Array(5).fill(null).map(() => Array(7).fill(null));
    });

    // 3. Initialize Faculty structures
    const tempFacultySchedules = {};
    const tempFacultyHours = {}; 
    const facultyTotalWeeklyHours = {};
    const uniqueFacultyList = [...new Set(courses.map(c => c.faculty))];

    uniqueFacultyList.forEach(fac => {
      tempFacultySchedules[fac] = Array(5).fill(null).map(() => Array(7).fill(null));
      tempFacultyHours[fac] = Array(5).fill(0);
      facultyTotalWeeklyHours[fac] = 0;
    });

    // 2b. Initialize Fixed slots for target groups/semesters
    fixedSlots.forEach(fs => {
      groups.forEach(g => {
        let applies = false;
        if (!fs.targetType || fs.targetType === 'all') {
          applies = true;
        } else if (fs.targetType === 'group' && g === fs.targetValue) {
          applies = true;
        } else if (fs.targetType === 'semester') {
          const sem = g.split(' - ')[1];
          if (sem && sem.trim().toLowerCase() === fs.targetValue.trim().toLowerCase()) {
            applies = true;
          }
        }

        if (applies) {
          fs.slots.forEach((sIdx) => {
            tempTimetables[g][fs.day][sIdx] = {
              courseName: fs.courseName,
              courseCode: fs.courseCode,
              faculty: fs.faculty,
              classroom: fs.classroom || 'TBD',
              school: fs.school || '',
              duration: 1,
              isPart2: false,
              isPart3: false,
              isFixed: true,
              targetType: fs.targetType,
              targetValue: fs.targetValue
            };
          });

          fs.slots.forEach(sIdx => {
            if (!tempFacultySchedules[fs.faculty]) {
              tempFacultySchedules[fs.faculty] = Array(5).fill(null).map(() => Array(7).fill(null));
            }
            const existing = tempFacultySchedules[fs.faculty][fs.day][sIdx] || [];
            if (Array.isArray(existing)) {
              tempFacultySchedules[fs.faculty][fs.day][sIdx] = [...new Set([...existing, g])];
            } else {
              tempFacultySchedules[fs.faculty][fs.day][sIdx] = [g];
            }
          });
        }
      });
    });

    // 4. Breakdown course hours into sessions, grouping merged courses together
    const sessionsToSchedule = [];
    const mergedGroups = {};
    const unmergedCourses = [];

    courses.forEach(c => {
      if (c.mergeCode && c.mergeCode.trim() !== "") {
        const code = c.mergeCode.trim();
        if (!mergedGroups[code]) {
          mergedGroups[code] = [];
        }
        mergedGroups[code].push(c);
      } else {
        unmergedCourses.push(c);
      }
    });

    // Process merged groups
    Object.keys(mergedGroups).forEach(code => {
      const groupCourses = mergedGroups[code];
      if (groupCourses.length === 0) return;
      const minHours = Math.min(...groupCourses.map(c => c.hours));
      let hoursRemaining = minHours;
      const blockSize = getSessionBlocks(groupCourses[0]);

      const createMergedSession = (dur) => {
        sessionsToSchedule.push({
          id: `merge-${code}-${hoursRemaining}-${dur}`,
          isMerged: true,
          mergeCode: code,
          courseName: groupCourses[0].courseName,
          courseCode: groupCourses[0].courseCode,
          faculty: groupCourses[0].faculty,
          classroom: groupCourses.map(c => c.classroom).filter(Boolean).join('/') || 'TBD',
          school: [...new Set(groupCourses.map(c => c.school).filter(Boolean))].join(', '),
          groups: groupCourses.map(c => `${c.programme} - ${c.semester}`),
          duration: dur,
          originalCourses: groupCourses
        });
      };

      if (blockSize === 3) {
        while (hoursRemaining >= 3) {
          createMergedSession(3);
          hoursRemaining -= 3;
        }
      }
      if (blockSize === 2 || (blockSize === 3 && hoursRemaining === 2)) {
        while (hoursRemaining >= 2) {
          createMergedSession(2);
          hoursRemaining -= 2;
        }
      }
      while (hoursRemaining >= 1) {
        createMergedSession(1);
        hoursRemaining -= 1;
      }

      // Excess hours scheduled unmerged
      groupCourses.forEach(c => {
        let rem = c.hours - minHours;
        if (rem > 0) {
          const groupName = `${c.programme} - ${c.semester}`;
          const cBlockSize = getSessionBlocks(c);
          if (cBlockSize === 3) {
            while (rem >= 3) {
              sessionsToSchedule.push({
                id: `${c.id}-rem3-${rem}`,
                courseName: c.courseName,
                courseCode: c.courseCode,
                faculty: c.faculty,
                classroom: c.classroom || 'TBD',
                school: c.school || '',
                group: groupName,
                duration: 3,
                originalCourse: c
              });
              rem -= 3;
            }
          }
          if (cBlockSize === 2 || (cBlockSize === 3 && rem === 2)) {
            while (rem >= 2) {
              sessionsToSchedule.push({
                id: `${c.id}-rem2-${rem}`,
                courseName: c.courseName,
                courseCode: c.courseCode,
                faculty: c.faculty,
                classroom: c.classroom || 'TBD',
                school: c.school || '',
                group: groupName,
                duration: 2,
                originalCourse: c
              });
              rem -= 2;
            }
          }
          while (rem >= 1) {
            sessionsToSchedule.push({
              id: `${c.id}-rem1-${rem}`,
              courseName: c.courseName,
              courseCode: c.courseCode,
              faculty: c.faculty,
              classroom: c.classroom || 'TBD',
              school: c.school || '',
              group: groupName,
              duration: 1,
              originalCourse: c
            });
            rem -= 1;
          }
        }
      });
    });

    // Process unmerged courses
    unmergedCourses.forEach(c => {
      const groupName = `${c.programme} - ${c.semester}`;
      let hoursRemaining = c.hours;
      const blockSize = getSessionBlocks(c);

      if (blockSize === 3) {
        while (hoursRemaining >= 3) {
          sessionsToSchedule.push({
            id: `${c.id}-block3-${hoursRemaining}`,
            courseName: c.courseName,
            courseCode: c.courseCode,
            faculty: c.faculty,
            classroom: c.classroom || 'TBD',
            school: c.school || '',
            group: groupName,
            duration: 3,
            originalCourse: c
          });
          hoursRemaining -= 3;
        }
      }
      if (blockSize === 2 || (blockSize === 3 && hoursRemaining === 2)) {
        while (hoursRemaining >= 2) {
          sessionsToSchedule.push({
            id: `${c.id}-block2-${hoursRemaining}`,
            courseName: c.courseName,
            courseCode: c.courseCode,
            faculty: c.faculty,
            classroom: c.classroom || 'TBD',
            school: c.school || '',
            group: groupName,
            duration: 2,
            originalCourse: c
          });
          hoursRemaining -= 2;
        }
      }
      while (hoursRemaining >= 1) {
        sessionsToSchedule.push({
          id: `${c.id}-block1-${hoursRemaining}`,
          courseName: c.courseName,
          courseCode: c.courseCode,
          faculty: c.faculty,
          classroom: c.classroom || 'TBD',
          school: c.school || '',
          group: groupName,
          duration: 1,
          originalCourse: c
        });
        hoursRemaining -= 1;
      }
    });

    // 5. Pre-calculate total faculty hours for sorting
    courses.forEach(c => {
      if (!facultyTotalWeeklyHours[c.faculty]) facultyTotalWeeklyHours[c.faculty] = 0;
      facultyTotalWeeklyHours[c.faculty] += c.hours;
    });

    // Heuristics: Labs first, then long courses, then busy faculty
    sessionsToSchedule.sort((a, b) => {
      if (b.duration !== a.duration) {
        return b.duration - a.duration;
      }
      const aHours = a.isMerged ? a.originalCourses[0].hours : a.originalCourse.hours;
      const bHours = b.isMerged ? b.originalCourses[0].hours : b.originalCourse.hours;
      if (bHours !== aHours) {
        return bHours - aHours;
      }
      const aFacHours = facultyTotalWeeklyHours[a.faculty] || 0;
      const bFacHours = facultyTotalWeeklyHours[b.faculty] || 0;
      return bFacHours - aFacHours;
    });

    const tempUnscheduledLog = [];

    const getGroupHoursOnDay = (grp, day) => {
      let count = 0;
      for (let s = 0; s < 7; s++) {
        if (tempTimetables[grp][day][s] !== null) count++;
      }
      return count;
    };

    // 6. Greedy Slot Allocation loop with Workload Balancing
    sessionsToSchedule.forEach(session => {
      const { isMerged, mergeCode, courseName, courseCode, faculty, classroom, school, duration } = session;
      const targetGroups = isMerged ? session.groups : [session.group];

      let bestDay = -1;
      let bestSlot = -1;
      let lowestScore = Infinity;
      const rejections = [];

      for (let day = 0; day < 5; day++) {
        for (let slot = 0; slot <= 7 - duration; slot++) {
          let isValid = true;
          const reasons = [];

          // Rule 1: Student Group Clash check
          for (const grp of targetGroups) {
            for (let offset = 0; offset < duration; offset++) {
              if (tempTimetables[grp][day][slot + offset] !== null) {
                isValid = false;
                reasons.push(`Group occupied: ${grp}`);
                break;
              }
            }
            if (!isValid) break;
          }

          // Rule 2: Faculty Clash check
          if (isValid) {
            for (let offset = 0; offset < duration; offset++) {
              const scheduledVal = tempFacultySchedules[faculty]?.[day]?.[slot + offset];
              if (scheduledVal) {
                if (scheduledVal === "University Fixed") {
                  isValid = false;
                  reasons.push("Faculty locked in Fixed slot");
                  break;
                }
                const hasExternalGroup = Array.isArray(scheduledVal) && scheduledVal.some(g => !targetGroups.includes(g));
                if (hasExternalGroup || !Array.isArray(scheduledVal)) {
                  isValid = false;
                  reasons.push("Faculty clash");
                  break;
                }
              }
            }
          }

          // Rule 3: Lab block crossover prevention
          if (isValid && crossesBreak(slot, duration)) {
            isValid = false;
            reasons.push("Block crosses break");
          }

          if (isValid) {
            let totalGroupDailyLoad = 0;
            targetGroups.forEach(grp => {
              totalGroupDailyLoad += getGroupHoursOnDay(grp, day);
            });
            const avgGroupLoad = totalGroupDailyLoad / targetGroups.length;
            const facultyDailyLoad = tempFacultyHours[faculty] ? tempFacultyHours[faculty][day] : 0;
            const score = avgGroupLoad + facultyDailyLoad;

            if (score < lowestScore) {
              lowestScore = score;
              bestDay = day;
              bestSlot = slot;
            }
          } else {
            rejections.push({ day, slot, reasons });
          }
        }
      }

      if (bestDay !== -1 && bestSlot !== -1) {
        for (let offset = 0; offset < duration; offset++) {
          targetGroups.forEach(grp => {
            const matchingCourse = isMerged 
              ? session.originalCourses.find(oc => `${oc.programme} - ${oc.semester}` === grp)
              : session.originalCourse;
            tempTimetables[grp][bestDay][bestSlot + offset] = {
              courseName: matchingCourse ? matchingCourse.courseName : courseName,
              courseCode: matchingCourse ? matchingCourse.courseCode : (courseCode || ''),
              faculty,
              classroom: matchingCourse ? (matchingCourse.classroom || 'TBD') : (classroom || 'TBD'),
              school: matchingCourse ? (matchingCourse.school || '') : (school || ''),
              duration,
              isPart2: offset === 1,
              isPart3: offset === 2,
              isMerged,
              mergeCode
            };
          });

          if (!tempFacultySchedules[faculty]) {
            tempFacultySchedules[faculty] = Array(5).fill(null).map(() => Array(7).fill(null));
          }
          const existingScheduled = tempFacultySchedules[faculty][bestDay][bestSlot + offset] || [];
          tempFacultySchedules[faculty][bestDay][bestSlot + offset] = [...new Set([...existingScheduled, ...targetGroups])];
        }
        if (!tempFacultyHours[faculty]) {
          tempFacultyHours[faculty] = Array(5).fill(0);
        }
        tempFacultyHours[faculty][bestDay] += duration;
      } else {
        const reasonCounts = {};
        rejections.forEach(r => {
          r.reasons.forEach(reason => {
            reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
          });
        });

        let primaryReason = "Over-constrained (no free slots)";
        if (reasonCounts["Faculty clash"] > 10) {
          primaryReason = `Faculty schedule clash`;
        } else if (reasonCounts["Faculty locked in Fixed slot"] > 0) {
          primaryReason = `Faculty busy in Fixed slot`;
        } else if (reasonCounts["Block crosses break"] > 10) {
          primaryReason = `Block crosses break line`;
        } else if (Object.keys(reasonCounts).some(r => r.startsWith("Group occupied"))) {
          primaryReason = `Student group timetable is full for this day`;
        }

        tempUnscheduledLog.push({
          courseName,
          faculty,
          group: targetGroups.join(', '),
          hours: duration,
          reason: primaryReason
        });
      }
    });

    // 8. Update UI State
    setTimetableData(tempTimetables);
    setRawTimetableBackup(JSON.parse(JSON.stringify(tempTimetables))); // Store deep copy backup
    setUnscheduledLog(tempUnscheduledLog);
    setFacultyWorkload(tempFacultyHours);
    updateFirestore({
      timetableData: JSON.stringify(tempTimetables),
      rawTimetableBackup: JSON.stringify(tempTimetables),
      unscheduledLog: tempUnscheduledLog,
      facultyWorkload: tempFacultyHours
    });
    
    if (groups.length > 0) {
      setSelectedGroup(groups[0]);
    }
    if (uniqueFacultyList.length > 0) {
      setSelectedFaculty(uniqueFacultyList[0]);
    }
  };

  // --- MANUAL SLOT REALLOCATION ---
  


  // Dry-run validator for moves
  const validateMove = (group, srcDay, srcSlot, destDay, destSlot) => {
    const cell = timetableData[group][srcDay][srcSlot];
    if (!cell) return "Source slot is empty.";
    if (cell.isFixed) return "University-wide fixed slots are locked and cannot be moved.";

    const startSlot = cell.isPart2 ? srcSlot - 1 : (cell.isPart3 ? srcSlot - 2 : srcSlot);
    const actualCell = timetableData[group][srcDay][startSlot];
    const dur = actualCell.duration;
    const faculty = actualCell.faculty;

    // 1. Boundary check
    if (destSlot + dur > 7) {
      return `A ${dur}-hour block goes beyond the 7 daily slots.`;
    }

    // 1b. Break crossover check
    if (crossesBreak(destSlot, dur)) {
      return `A ${dur}-hour block cannot cross the scheduled break lines.`;
    }

    // Identify all groups in this merged session
    const mergedGroupsList = [];
    if (actualCell.isMerged) {
      Object.keys(timetableData).forEach(grp => {
        const c = timetableData[grp][srcDay][startSlot];
        if (c && c.isMerged && c.mergeCode === actualCell.mergeCode) {
          mergedGroupsList.push(grp);
        }
      });
    } else {
      mergedGroupsList.push(group);
    }

    // 2. Student Group Clash check (for all groups involved)
    for (const grp of mergedGroupsList) {
      for (let o = 0; o < dur; o++) {
        const existing = timetableData[grp][destDay][destSlot + o];
        if (existing) {
          // Allow self-shifting on the same day/slots
          const isSelf = (destDay === srcDay) && (destSlot + o >= startSlot && destSlot + o < startSlot + dur);
          if (!isSelf) {
            return `Programme "${grp}" already has class "${existing.courseName}" scheduled at slot ${destSlot + o + 1} on ${DAYS[destDay]}.`;
          }
        }
      }
    }

    // 3. Faculty Clash check
    for (let o = 0; o < dur; o++) {
      const targetSlot = destSlot + o;
      for (const grp of Object.keys(timetableData)) {
        if (mergedGroupsList.includes(grp)) continue; // skip self groups
        
        const cellAtTarget = timetableData[grp][destDay][targetSlot];
        if (cellAtTarget && cellAtTarget.faculty === faculty) {
          const isSelf = (destDay === srcDay) && (targetSlot >= startSlot && targetSlot < startSlot + dur);
          if (!isSelf) {
            return `Lecturer ${faculty} is already teaching "${grp}" in that slot.`;
          }
        }
      }
    }

    return null; // Valid!
  };

  const handleCellClick = (dayIdx, slotIdx) => {
    if (portalMode === "student") return; // Students can't rearrange slots

    // If we haven't selected a move source yet, try to select one
    if (!moveSource) {
      if (viewMode !== "programme") return; // Only allow selecting source in Programme View
      const cell = timetableData[selectedGroup]?.[dayIdx][slotIdx];
      if (cell) {
        if (cell.isFixed) {
          setMoveMessage({ type: 'error', text: "Fixed slots are locked and cannot be moved." });
          setTimeout(() => setMoveMessage(null), 3000);
          return;
        }
        const startSlot = cell.isPart2 ? slotIdx - 1 : (cell.isPart3 ? slotIdx - 2 : slotIdx);
        setMoveSource({
          group: selectedGroup,
          day: dayIdx,
          slot: startSlot,
          cell: timetableData[selectedGroup][dayIdx][startSlot]
        });
        setMoveMessage(null);
      }
    } else {
      // We already have a moveSource. The click represents the destination slot
      const { group, day: srcDay, slot: srcSlot } = moveSource;
      
      // If they clicked the source coordinates, simply cancel selection
      if (group === selectedGroup && srcDay === dayIdx && srcSlot === slotIdx) {
        setMoveSource(null);
        return;
      }

      // Check move viability
      const errorMsg = validateMove(group, srcDay, srcSlot, dayIdx, slotIdx);
      if (errorMsg) {
        setMoveMessage({ type: 'error', text: errorMsg });
        return;
      }

      // Valid move -> Execute update
      const nextTimetable = JSON.parse(JSON.stringify(timetableData));
      const actualCell = nextTimetable[group][srcDay][srcSlot];
      const dur = actualCell.duration;

      // Identify all groups in this merged session
      const mergedGroupsList = [];
      if (actualCell.isMerged) {
        Object.keys(nextTimetable).forEach(grp => {
          const c = nextTimetable[grp][srcDay][srcSlot];
          if (c && c.isMerged && c.mergeCode === actualCell.mergeCode) {
            mergedGroupsList.push(grp);
          }
        });
      } else {
        mergedGroupsList.push(group);
      }

      // 1. Clear source slots for all groups
      mergedGroupsList.forEach(grp => {
        for (let o = 0; o < dur; o++) {
          nextTimetable[grp][srcDay][srcSlot + o] = null;
        }
      });

      // 2. Populate destination slots for all groups
      mergedGroupsList.forEach(grp => {
        for (let o = 0; o < dur; o++) {
          nextTimetable[grp][dayIdx][slotIdx + o] = {
            ...actualCell,
            isPart2: o === 1,
            isPart3: o === 2
          };
        }
      });

      // 3. Re-calculate workloads
      const nextWorkload = {};
      Object.keys(facultyWorkload).forEach(fac => {
        nextWorkload[fac] = Array(5).fill(0);
      });

      Object.keys(nextTimetable).forEach(g => {
        for (let d = 0; d < 5; d++) {
          for (let s = 0; s < 7; s++) {
            const c = nextTimetable[g][d][s];
            if (c && !c.isPart2 && !c.isPart3) {
              if (!nextWorkload[c.faculty]) {
                nextWorkload[c.faculty] = Array(5).fill(0);
              }
              nextWorkload[c.faculty][d] += c.duration;
            }
          }
        }
      });

      // Update states
      setTimetableData(nextTimetable);
      setFacultyWorkload(nextWorkload);
      setMoveSource(null);
      setMoveMessage({ type: 'success', text: `Successfully re-allocated "${actualCell.courseName}"!` });
      updateFirestore({
        timetableData: JSON.stringify(nextTimetable),
        facultyWorkload: nextWorkload
      });

      setTimeout(() => setMoveMessage(null), 3000);
    }
  };

  const handleResetTimetable = () => {
    if (!rawTimetableBackup) return;
    
    const restoredTimetable = JSON.parse(JSON.stringify(rawTimetableBackup));
    setTimetableData(restoredTimetable);
    
    // Re-calculate workloads
    const restoredWorkload = {};
    Object.keys(facultyWorkload).forEach(fac => {
      restoredWorkload[fac] = Array(5).fill(0);
    });

    Object.keys(restoredTimetable).forEach(g => {
      for (let d = 0; d < 5; d++) {
        for (let s = 0; s < 7; s++) {
          const c = restoredTimetable[g][d][s];
          if (c && !c.isPart2) {
            if (!restoredWorkload[c.faculty]) {
              restoredWorkload[c.faculty] = Array(5).fill(0);
            }
            restoredWorkload[c.faculty][d] += c.duration;
          }
        }
      }
    });

    setFacultyWorkload(restoredWorkload);
    setMoveSource(null);
    setMoveMessage({ type: 'success', text: "Reset timetable to generated state." });
    updateFirestore({
      timetableData: JSON.stringify(restoredTimetable),
      facultyWorkload: restoredWorkload
    });
    setTimeout(() => setMoveMessage(null), 3000);
  };

  const handleExportPDF = async () => {
    if (!gridRef.current) return;
    setPdfLoading(true);
    setMoveSource(null);
    setMoveMessage(null);
    
    try {
      await new Promise(resolve => setTimeout(resolve, 100));
      const element = gridRef.current;
      
      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });

      const imgData = canvas.toDataURL('image/jpeg', 1.0);
      
      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });

      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = pdf.internal.pageSize.getHeight();
      
      const margin = 10;
      const contentWidth = pdfWidth - (margin * 2);
      const contentHeight = (canvas.height * contentWidth) / canvas.width;
      
      let finalHeight = contentHeight;
      let finalWidth = contentWidth;
      let xOffset = margin;
      let yOffset = margin;
      
      if (contentHeight > pdfHeight - (margin * 2)) {
        finalHeight = pdfHeight - (margin * 2);
        finalWidth = (canvas.width * finalHeight) / canvas.height;
        xOffset = (pdfWidth - finalWidth) / 2;
      } else {
        yOffset = (pdfHeight - finalHeight) / 2;
      }

      pdf.addImage(imgData, 'JPEG', xOffset, yOffset, finalWidth, finalHeight);
      
      const docName = viewMode === "programme" 
        ? `${selectedGroup.replace(/\s+/g, '_')}_Timetable.pdf` 
        : `${selectedFaculty.replace(/\s+/g, '_')}_Schedule.pdf`;
         
      pdf.save(docName);
      setMoveMessage({ type: 'success', text: 'PDF Timetable downloaded successfully!' });
    } catch (err) {
      console.error("PDF Export error: ", err);
      setMoveMessage({ type: 'error', text: 'Failed to generate PDF. Check console logs.' });
    } finally {
      setPdfLoading(false);
      setTimeout(() => setMoveMessage(null), 3000);
    }
  };

  // --- MODE B: FACULTY GRID LOOKUP ---
  const facultyGrid = useMemo(() => {
    const grid = Array(5).fill(null).map(() => Array(7).fill(null));
    if (!timetableData || !selectedFaculty) return grid;

    Object.keys(timetableData).forEach(gName => {
      for (let d = 0; d < 5; d++) {
        for (let s = 0; s < 7; s++) {
          const cell = timetableData[gName][d][s];
          if (cell && cell.faculty === selectedFaculty) {
            grid[d][s] = {
              courseName: cell.courseName,
              group: gName,
              duration: cell.duration,
              isPart2: cell.isPart2
            };
          }
        }
      }
    });
    return grid;
  }, [timetableData, selectedFaculty]);

  const uniqueSchools = useMemo(() => {
    if (courses.length === 0) return [];
    return [...new Set(courses.map(c => c.school).filter(Boolean))];
  }, [courses]);

  const uniqueSemesters = useMemo(() => {
    if (courses.length === 0) return [];
    return [...new Set(courses.map(c => c.semester).filter(Boolean))];
  }, [courses]);

  const uniqueGroupsList = useMemo(() => {
    if (courses.length === 0) return [];
    return [...new Set(courses.map(c => `${c.programme} - ${c.semester}`))];
  }, [courses]);

  const filteredCourses = useMemo(() => {
    let result = courses;
    if (selectedSchoolFilter !== "All") {
      result = result.filter(c => c.school === selectedSchoolFilter);
    }
    if (!searchTerm.trim()) return result;
    const term = searchTerm.toLowerCase();
    return result.filter(item => 
      item.programme.toLowerCase().includes(term) ||
      item.semester.toLowerCase().includes(term) ||
      item.courseName.toLowerCase().includes(term) ||
      (item.courseCode && item.courseCode.toLowerCase().includes(term)) ||
      (item.school && item.school.toLowerCase().includes(term)) ||
      (item.classroom && item.classroom.toLowerCase().includes(term)) ||
      item.faculty.toLowerCase().includes(term)
    );
  }, [courses, searchTerm, selectedSchoolFilter]);

  // Unique lists for selectors
  const uniqueGroups = useMemo(() => {
    if (!timetableData) return [];
    return Object.keys(timetableData);
  }, [timetableData]);

  const uniqueFaculty = useMemo(() => {
    if (courses.length === 0) return [];
    return [...new Set(courses.map(c => c.faculty))];
  }, [courses]);

  const [settingsTab, setSettingsTab] = useState("slots");
  const [selectedCourseIds, setSelectedCourseIds] = useState([]);
  const [fsTargetType, setFsTargetType] = useState("all");
  const [fsTargetValue, setFsTargetValue] = useState("");
  const [fsSelectedSlots, setFsSelectedSlots] = useState([]);

  const [fsName, setFsName] = useState("");
  const [fsCode, setFsCode] = useState("");
  const [fsFaculty, setFsFaculty] = useState("");
  const [fsRoom, setFsRoom] = useState("");
  const [fsSchool, setFsSchool] = useState("");
  const [fsDay, setFsDay] = useState(0);
  const [editingFsId, setEditingFsId] = useState(null);

  const cancelFsEdit = () => {
    setEditingFsId(null);
    setFsName("");
    setFsCode("");
    setFsFaculty("");
    setFsRoom("");
    setFsSchool("");
    setFsDay(0);
    setFsTargetType("all");
    setFsTargetValue("");
    setFsSelectedSlots([]);
  };

  const [dbLoading, setDbLoading] = useState(true);
  const [plannersList, setPlannersList] = useState([]);
  const [activePlannerId, setActivePlannerId] = useState(() => {
    return localStorage.getItem("activePlannerId") || "default";
  });

  useEffect(() => {
    localStorage.setItem("activePlannerId", activePlannerId);
  }, [activePlannerId]);

  const updateFirestore = (updates) => {
    if (!activePlannerId) return;
    setDoc(doc(db, "planners", activePlannerId), updates, { merge: true })
      .catch(err => console.error("Firestore write error:", err));
  };

  // 1. Listen to all planners in real-time
  useEffect(() => {
    const unsub = onSnapshot(collection(db, "planners"), (snapshot) => {
      const list = [];
      snapshot.forEach(docSnap => {
        const data = docSnap.data();
        list.push({
          id: docSnap.id,
          name: data.name || docSnap.id,
          createdAt: data.createdAt || 0
        });
      });
      list.sort((a, b) => b.createdAt - a.createdAt);
      setPlannersList(list);
    }, (err) => {
      console.error("Failed to load planners list:", err);
    });
    return () => unsub();
  }, []);

  // 2. Listen to active planner's state in real-time
  useEffect(() => {
    if (!activePlannerId) return;

    Promise.resolve().then(() => setDbLoading(true));
    const unsub = onSnapshot(doc(db, "planners", activePlannerId), (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        setCourses(data.courses || []);
        
        let currentFixedSlots = data.fixedSlots || [];
        const hasDefaultFs = currentFixedSlots.some(fs => fs.id === 'fs-1');
        if (hasDefaultFs) {
          currentFixedSlots = currentFixedSlots.filter(fs => fs.id !== 'fs-1');
          setDoc(doc(db, "planners", activePlannerId), { fixedSlots: currentFixedSlots }, { merge: true })
            .catch(err => console.error("Error migrating default fixed slots:", err));
        }
        setFixedSlots(currentFixedSlots);
        setSlots(data.slots || INITIAL_SLOTS);
        
        let parsedTimetable = null;
        if (data.timetableData) {
          try {
            parsedTimetable = typeof data.timetableData === 'string'
              ? JSON.parse(data.timetableData)
              : data.timetableData;
          } catch (e) {
            console.error("Failed to parse timetableData JSON:", e);
          }
        }
        setTimetableData(parsedTimetable);

        let parsedBackup = null;
        if (data.rawTimetableBackup) {
          try {
            parsedBackup = typeof data.rawTimetableBackup === 'string'
              ? JSON.parse(data.rawTimetableBackup)
              : data.rawTimetableBackup;
          } catch (e) {
            console.error("Failed to parse rawTimetableBackup JSON:", e);
          }
        }
        setRawTimetableBackup(parsedBackup);

        setUnscheduledLog(data.unscheduledLog || []);
        setFacultyWorkload(data.facultyWorkload || {});
        setFileName(data.fileName || "");
      } else {
        // Initialize active planner document with default settings if empty
        const plannerName = activePlannerId === "default" ? "Default Planner" : activePlannerId;
        const stateRef = doc(db, "planners", activePlannerId);
        setDoc(stateRef, {
          name: plannerName,
          createdAt: Date.now(),
          courses: [],
          fixedSlots: [],
          slots: INITIAL_SLOTS,
          timetableData: null,
          rawTimetableBackup: null,
          unscheduledLog: [],
          facultyWorkload: {},
          fileName: ""
        });
      }
      setDbLoading(false);
    }, (err) => {
      console.error("Firestore active planner error:", err);
      setDbLoading(false);
    });

    return () => unsub();
  }, [activePlannerId]);

  const handleCreatePlanner = async () => {
    const name = prompt("Enter a name for the new planner (e.g. 'Summer Fall Planner', 'Winter Fall Planner'):");
    if (!name || !name.trim()) return;

    const newId = `planner-${Date.now()}`;
    try {
      await setDoc(doc(db, "planners", newId), {
        name: name.trim(),
        createdAt: Date.now(),
        courses: [],
        fixedSlots: [],
        slots: INITIAL_SLOTS,
        timetableData: null,
        rawTimetableBackup: null,
        unscheduledLog: [],
        facultyWorkload: {},
        fileName: ""
      });
      setActivePlannerId(newId);
    } catch (err) {
      console.error("Failed to create new planner:", err);
      alert("Error creating new planner.");
    }
  };

  const handleDeletePlanner = async (idToDelete) => {
    if (!window.confirm("Are you sure you want to delete this planner? This will delete all its courses, slots, and timetables permanently.")) return;
    try {
      await deleteDoc(doc(db, "planners", idToDelete));
      if (activePlannerId === idToDelete) {
        const remaining = plannersList.filter(p => p.id !== idToDelete);
        if (remaining.length > 0) {
          setActivePlannerId(remaining[0].id);
        } else {
          setActivePlannerId("default");
        }
      }
    } catch (err) {
      console.error("Failed to delete planner:", err);
      alert("Error deleting planner.");
    }
  };



  const handleSaveCourseEdit = (updated) => {
    const nextCourses = courses.map(c => c.id === updated.id ? updated : c);
    setCourses(nextCourses);
    
    // Dynamically update scheduled cells in the current timetableData
    if (timetableData) {
      const nextTimetable = JSON.parse(JSON.stringify(timetableData));
      const orig = courses.find(c => c.id === updated.id);
      
      Object.keys(nextTimetable).forEach(group => {
        for (let d = 0; d < 5; d++) {
          for (let s = 0; s < 7; s++) {
            const cell = nextTimetable[group][d][s];
            if (cell && (cell.courseCode === orig.courseCode || cell.courseName === orig.courseName)) {
              cell.courseName = updated.courseName;
              cell.courseCode = updated.courseCode;
              cell.faculty = updated.faculty;
              cell.classroom = updated.classroom || 'TBD';
              cell.school = updated.school || '';
              cell.mergeCode = updated.mergeCode;
            }
          }
        }
      });
      setTimetableData(nextTimetable);
      updateFirestore({
        courses: nextCourses,
        timetableData: JSON.stringify(nextTimetable)
      });
    } else {
      updateFirestore({
        courses: nextCourses
      });
    }
    setEditingCourse(null);
    setMoveMessage({ type: 'success', text: `Successfully updated course "${updated.courseName}"!` });
    setTimeout(() => setMoveMessage(null), 3000);
  };

  const handleMergeSelected = () => {
    const code = getMergeCode();
    const nextCourses = courses.map(c => {
      if (selectedCourseIds.includes(c.id)) {
        return { ...c, mergeCode: code };
      }
      return c;
    });
    setCourses(nextCourses);
    setSelectedCourseIds([]);
    updateFirestore({ courses: nextCourses });
    setMoveMessage({ type: 'success', text: `Merged selected courses with group code "${code}"! Click "Generate Timetables" to schedule.` });
    setTimeout(() => setMoveMessage(null), 5000);
  };

  return (
    <div className="max-w-7xl mx-auto py-4 px-2 sm:px-4 space-y-6">
      {/* Portal Mode Switching Header */}
      <div className="flex items-center justify-between border-b border-slate-200/80 pb-3.5 z-20 relative gap-3 flex-wrap">
        <div className="inline-flex rounded-lg bg-slate-100 p-0.5 border border-slate-200/80 shadow-sm">
          <button
            onClick={() => setPortalMode("admin")}
            className={`py-1 px-3.5 rounded text-[10px] font-bold transition-all flex items-center gap-1 cursor-pointer
              ${portalMode === "admin" 
                ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' 
                : 'text-slate-500 hover:text-slate-800'
              }`}
          >
            🛡️ Admin Dashboard
          </button>
          <button
            onClick={() => setPortalMode("student")}
            className={`py-1 px-3.5 rounded text-[10px] font-bold transition-all flex items-center gap-1 cursor-pointer
              ${portalMode === "student" 
                ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' 
                : 'text-slate-500 hover:text-slate-800'
              }`}
          >
            🎓 Student Portal
          </button>
        </div>

        {/* Planner Switcher Dropdown (Visible to both admin and student, but only admin can create/delete) */}
        <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg p-1 shadow-sm flex-wrap">
          <span className="text-[10px] font-bold text-slate-500 pl-1">📅 Active Planner:</span>
          <select 
            value={activePlannerId}
            onChange={(e) => setActivePlannerId(e.target.value)}
            className="px-2 py-0.5 border border-slate-250 bg-white hover:bg-slate-50 rounded-md text-[10px] font-bold text-slate-700 focus:outline-none cursor-pointer"
          >
            {plannersList.length === 0 ? (
              <option value={activePlannerId}>
                {activePlannerId === 'default' ? 'Default Planner' : 'Loading...'}
              </option>
            ) : (
              plannersList.map(p => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))
            )}
          </select>

          {portalMode === "admin" && (
            <div className="flex items-center gap-1">
              <button 
                onClick={handleCreatePlanner}
                className="px-2 py-1 bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 rounded-md text-[9px] font-bold transition-all cursor-pointer animate-pulse hover:animate-none"
                title="Create New Planner"
              >
                ➕ New
              </button>
              {activePlannerId !== "default" && (
                <button 
                  onClick={() => handleDeletePlanner(activePlannerId)}
                  className="px-2 py-1 bg-red-50 border border-red-200 text-red-700 hover:bg-red-100 rounded-md text-[9px] font-bold transition-all cursor-pointer"
                  title="Delete Current Planner"
                >
                  🗑️ Delete
                </button>
              )}
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          {dbLoading ? (
            <div className="flex items-center gap-1.5 px-2.5 py-1 bg-amber-50 border border-amber-200 text-amber-700 rounded-full text-[10px] font-bold shadow-sm select-none">
              <span className="w-1.5 h-1.5 bg-amber-500 rounded-full animate-pulse"></span>
              🔄 Connecting...
            </div>
          ) : (
            <div className="flex items-center gap-1.5 px-2.5 py-1 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-full text-[10px] font-bold shadow-sm select-none">
              <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-ping"></span>
              ☁️ Cloud Synced
            </div>
          )}

          {portalMode === "admin" && (
            <button
              onClick={() => setIsSettingsOpen(true)}
              className="py-1 px-3.5 rounded-lg border border-slate-250 bg-white hover:bg-slate-50 text-[10px] font-bold text-slate-700 shadow-sm flex items-center gap-1.5 transition-all cursor-pointer"
            >
              <span>⚙️</span> Scheduler Settings
            </button>
          )}
        </div>
      </div>

      {/* Header Area */}
      <header className="text-center relative z-10 py-3">
        <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-indigo-50 text-indigo-600 border border-indigo-100 mb-2 tracking-wider uppercase">
          {portalMode === "admin" ? "University Timetable Editor" : "Student Timetable Viewer"}
        </span>
        <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight font-display bg-clip-text text-transparent bg-gradient-to-r from-indigo-600 via-violet-600 to-indigo-700 pb-1">
          {portalMode === "admin" ? "Curriculum Scheduling Dashboard" : "Academic Timetable Portal"}
        </h1>
        <p className="mt-1 text-xs text-slate-500 max-w-xl mx-auto">
          {portalMode === "admin" 
            ? "Parse worksheets, lock fixed slots, merge classes, reassign faculty, and modify slots interactively."
            : "Search and view your programme timetable or faculty schedule. Save a digital copy via PDF download."}
        </p>
      </header>

      {/* Quick Actions Panel (Admin Only) */}
      {portalMode === "admin" && (
        <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
          {/* Template Card */}
          <div className="bg-white border border-slate-200/80 p-3.5 rounded-xl flex flex-col justify-between hover:border-slate-300 hover:shadow-sm transition-all duration-300">
            <div>
              <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5 font-display">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                Excel Schema Required
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Workbook sheet named <code className="text-indigo-600 px-0.5 bg-indigo-50/50 border border-indigo-100/30 rounded text-[10px] font-mono">Course_List</code>. Accepts optional Course Code, Room & School.
              </p>
            </div>
            <button 
              onClick={handleDownloadTemplate}
              className="mt-2.5 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-lg text-xs font-medium border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 transition-all cursor-pointer"
            >
              <DownloadIcon className="w-3.5 h-3.5" />
              Download Template
            </button>
          </div>

          {/* Quick Demo Card */}
          <div className="bg-white border border-slate-200/80 p-3.5 rounded-xl flex flex-col justify-between hover:border-slate-300 hover:shadow-sm transition-all duration-300">
            <div>
              <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5 font-display">
                <span className="w-1.5 h-1.5 rounded-full bg-indigo-500"></span>
                Evaluation Datasets
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Load mock curriculum (BCA & CSE courses, shared classrooms, department tags) to evaluate schedule creation.
              </p>
            </div>
            <button 
              onClick={handleLoadDemoData}
              className="mt-2.5 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-lg text-xs font-medium bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white shadow-sm shadow-indigo-600/10 hover:scale-[1.01] active:scale-[0.99] transition-all cursor-pointer"
            >
              <RefreshIcon className="w-3.5 h-3.5" />
              Load Demo Data
            </button>
          </div>

          {/* Scheduler Config Card */}
          <div className="bg-white border border-slate-200/80 p-3.5 rounded-xl flex flex-col justify-between hover:border-slate-300 hover:shadow-sm transition-all duration-300">
            <div>
              <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5 font-display">
                <span className="w-1.5 h-1.5 rounded-full bg-violet-500"></span>
                Scheduling Parameters
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Schedules 7 customizable daily slots. Automatic split for labs (2h blocks) and engineering workshops (3h blocks).
              </p>
            </div>
            <div className="mt-2.5 flex items-center gap-1.5 text-[10px] text-indigo-600 bg-indigo-50/50 border border-indigo-100/60 p-2 rounded-lg">
              <InfoIcon className="w-3.5 h-3.5 shrink-0" />
              <span>Modify time slots and university fixed locks in settings.</span>
            </div>
          </div>
        </section>
      )}

      {/* Upload Container Area (Admin Only) */}
      <main className="space-y-6">
        {portalMode === "admin" && (
          <div 
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current.click()}
            className={`relative group cursor-pointer overflow-hidden rounded-2xl border-2 border-dashed transition-all duration-300 p-6 text-center flex flex-col items-center justify-center min-h-[100px]
              ${dragActive 
                ? 'border-indigo-500 bg-indigo-50/30' 
                : 'border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50/60 shadow-sm'
              }
            `}
          >
            <input 
              ref={fileInputRef}
              type="file" 
              accept=".xlsx, .xls, .ods" 
              onChange={handleFileSelect}
              className="hidden" 
            />

            {loading ? (
              <div className="space-y-2">
                <div className="w-6 h-6 border-2 border-indigo-500/30 border-t-indigo-500 rounded-full animate-spin mx-auto" />
                <p className="text-indigo-650 font-medium text-xs font-display">Reading workbook locally...</p>
              </div>
            ) : (
              <div className="space-y-2">
                <div className="w-8 h-8 rounded-lg bg-slate-50 border border-slate-200 flex items-center justify-center mx-auto shadow-sm group-hover:scale-105 transition-all duration-300">
                  <UploadIcon className="w-4 h-4 text-indigo-500" />
                </div>
                <div>
                  <p className="text-sm font-semibold text-slate-700 font-display">Drag & drop your Excel file here</p>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    or <span className="text-indigo-650 font-semibold underline underline-offset-2">browse local files</span> (.xlsx, .xls)
                  </p>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Error Panel */}
        {error && (
          <div className="relative overflow-hidden bg-rose-50 border border-rose-100 text-rose-800 p-4 rounded-xl flex items-start gap-3">
            <div className="absolute top-0 bottom-0 left-0 w-1 bg-rose-500" />
            <ExclamationTriangleIcon className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <h4 className="font-semibold text-rose-900 font-display">Workbook Parsing Error</h4>
              <p className="text-sm text-rose-850">{error}</p>
            </div>
          </div>
        )}

        {/* Parsing success state banner + Generate (Admin Only) */}
        {courses.length > 0 && !error && !loading && portalMode === "admin" && (
          <div className="relative overflow-hidden bg-white border border-slate-200 p-3.5 rounded-xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-sm">
            <div className="flex items-center gap-2">
              <CheckCircleIcon className="w-5 h-5 text-emerald-600" />
              <div className="text-left">
                <span className="text-[9px] text-slate-400 uppercase font-semibold block leading-none">
                  {fileName ? "Uploaded Document" : "Active Planner Dataset"}
                </span>
                <span className="text-xs font-semibold font-mono text-slate-855 block mt-0.5">
                  {fileName || "Stored Course Catalog"}
                </span>
                <span className="text-[10px] text-slate-500 block mt-0.5">{courses.length} courses loaded</span>
              </div>
            </div>
            
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <button
                onClick={handleGenerateTimetable}
                className="w-full sm:w-auto flex items-center justify-center gap-1.5 py-1.5 px-4 rounded-lg text-xs font-bold bg-gradient-to-r from-indigo-500 to-violet-600 hover:from-indigo-400 hover:to-violet-500 text-white shadow-sm hover:scale-[1.01] active:scale-[0.99] transition-all cursor-pointer"
              >
                <SparklesIcon className="w-3.5 h-3.5" />
                Generate Timetables
              </button>
            </div>
          </div>
        )}

        {/* Unscheduled Warnings banner (Admin Only) */}
        {unscheduledLog.length > 0 && portalMode === "admin" && (
          <div className="relative overflow-hidden bg-amber-50 border border-amber-100 text-amber-800 p-3.5 rounded-xl space-y-2.5 shadow-sm">
            <div className="absolute top-0 bottom-0 left-0 w-1 bg-amber-500" />
            <div className="flex items-start gap-2.5">
              <ExclamationTriangleIcon className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-bold text-amber-900 font-display">Unscheduled Course Hours ({unscheduledLog.length})</h4>
                <p className="text-[10px] text-slate-550">
                  Some hours could not fit automatically. Review workloads, adjust fixed slots, or rearrange slots manually:
                </p>
              </div>
            </div>

            <div className="max-h-32 overflow-y-auto border border-slate-200 bg-white rounded-lg divide-y divide-slate-100 text-[11px]">
              {unscheduledLog.map((log, index) => (
                <div key={index} className="p-1.5 flex flex-col sm:flex-row sm:items-center justify-between gap-1 hover:bg-slate-50">
                  <span className="text-slate-800 font-medium font-display">
                    {log.courseName} <span className="text-slate-400 font-normal">({log.group})</span>
                  </span>
                  <div className="flex items-center gap-3">
                    <span className="text-slate-505">Faculty: {log.faculty}</span>
                    <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 border border-amber-200 text-[10px] font-semibold">
                      {log.reason}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Move operations toasts/alerts */}
        {moveMessage && (
          <div className={`relative overflow-hidden p-4 rounded-xl flex items-center gap-3 border shadow-sm animate-fadeIn
            ${moveMessage.type === 'error' 
              ? 'bg-rose-50 border-rose-100 text-rose-800' 
              : 'bg-emerald-50 border-emerald-100 text-emerald-800'
            }
          `}>
            {moveMessage.type === 'error' ? (
              <ExclamationTriangleIcon className="w-5 h-5 text-rose-600" />
            ) : (
              <CheckCircleIcon className="w-5 h-5 text-emerald-600" />
            )}
            <span className="text-sm font-medium">{moveMessage.text}</span>
          </div>
        )}

        {/* Slot Movement Status Banner */}
        {moveSource && (
          <div className="bg-indigo-50 border border-indigo-200 text-indigo-900 p-4 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-md animate-pulse">
            <div className="flex items-center gap-2 text-sm">
              <ScissorsIcon className="w-5 h-5 text-indigo-600 shrink-0" />
              <span>
                Moving <strong>{moveSource.cell.courseName}</strong> ({moveSource.cell.duration}h block) from <strong>{DAYS[moveSource.day]} {slots[moveSource.slot].label}</strong>.
              </span>
              <span className="hidden sm:inline text-indigo-400 font-bold">|</span>
              <span className="text-xs text-indigo-650">Click a green target cell to place it.</span>
            </div>
            <button
              onClick={() => {
                setMoveSource(null);
                setMoveMessage(null);
              }}
              className="py-1 px-3 text-xs bg-indigo-200 hover:bg-indigo-300 text-indigo-900 font-bold rounded-lg transition-colors cursor-pointer"
            >
              Cancel Move
            </button>
          </div>
        )}

        {/* Main Dashboard Display (Timetable Grid & Workloads) */}
        {timetableData && (
          <div className="space-y-6 animate-fadeIn">
            
            {/* View Filter Toggles & Dropdown Selection */}
            <div className="bg-white border border-slate-200 p-3 rounded-2xl flex flex-col md:flex-row items-center justify-between gap-3 shadow-sm">
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 w-full md:w-auto">
                <span className="text-[10px] font-bold text-slate-450 uppercase text-center sm:text-left shrink-0 self-center">View Calendar By:</span>
                
                {/* View Mode Toggle Buttons */}
                <div className="inline-flex rounded-lg bg-slate-100 p-0.5 border border-slate-200">
                  <button
                    onClick={() => {
                      setViewMode("programme");
                      setMoveSource(null);
                    }}
                    className={`py-1 px-3 rounded text-[10px] font-bold transition-all cursor-pointer
                      ${viewMode === "programme" 
                        ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' 
                        : 'text-slate-500 hover:text-slate-800'
                      }`}
                  >
                    Programme / Group
                  </button>
                  <button
                    onClick={() => {
                      setViewMode("faculty");
                      setMoveSource(null); // Cancel any active movement
                    }}
                    className={`py-1 px-3 rounded text-[10px] font-bold transition-all cursor-pointer
                      ${viewMode === "faculty" 
                        ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' 
                        : 'text-slate-500 hover:text-slate-800'
                      }`}
                  >
                    Faculty Lecturer
                  </button>
                </div>
              </div>

              {/* Selection Dropdown & Reset Actions */}
              <div className="flex flex-col sm:flex-row items-center gap-2.5 w-full md:w-auto">
                <div className="flex items-center gap-1.5 w-full sm:w-auto">
                  <span className="text-[10px] font-semibold text-slate-450 uppercase shrink-0">Selector:</span>
                  
                  {viewMode === "programme" ? (
                    <select
                      value={selectedGroup}
                      onChange={(e) => setSelectedGroup(e.target.value)}
                      className="w-full sm:w-48 px-2 py-1 bg-white text-slate-800 border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-xs transition-all shadow-sm"
                    >
                      {uniqueGroups.map(grp => (
                        <option key={grp} value={grp}>{grp}</option>
                      ))}
                    </select>
                  ) : (
                    <select
                      value={selectedFaculty}
                      onChange={(e) => setSelectedFaculty(e.target.value)}
                      className="w-full sm:w-48 px-2 py-1 bg-white text-slate-800 border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-xs transition-all shadow-sm"
                    >
                      {uniqueFaculty.map(fac => (
                        <option key={fac} value={fac}>{fac}</option>
                      ))}
                    </select>
                  )}
                </div>

                {/* Reset schedule modification button (Admin Only) */}
                {portalMode === "admin" && (
                  <button
                    onClick={handleResetTimetable}
                    className="w-full sm:w-auto px-3 py-1.5 text-[10px] font-bold text-slate-650 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg transition-all cursor-pointer"
                    title="Discard all slot modifications"
                  >
                    Reset Timetable
                  </button>
                )}

                {/* Download PDF button */}
                <button
                  onClick={handleExportPDF}
                  disabled={pdfLoading}
                  className={`w-full sm:w-auto px-3 py-1.5 text-[10px] font-bold rounded-lg border transition-all flex items-center justify-center gap-1 cursor-pointer
                    ${pdfLoading 
                      ? 'bg-slate-100 border-slate-200 text-slate-400 cursor-not-allowed' 
                      : 'bg-indigo-50 border-indigo-200 hover:bg-indigo-100 text-indigo-700 hover:border-indigo-300 shadow-sm'
                    }`}
                  title="Export current timetable as PDF"
                >
                  {pdfLoading ? (
                    <>
                      <div className="w-3.5 h-3.5 border-2 border-indigo-500/30 border-t-indigo-500 rounded-full animate-spin" />
                      <span>Generating PDF...</span>
                    </>
                  ) : (
                    <>
                      <CalendarIcon className="w-3.5 h-3.5" />
                      <span>Download PDF</span>
                    </>
                  )}
                </button>
              </div>
            </div>

            {/* Visual Options Inline Checkboxes */}
            <div className="bg-white border border-slate-200 p-3 rounded-2xl flex flex-wrap items-center gap-4 shadow-sm text-[11px] text-slate-600">
              <span className="font-bold text-slate-450 uppercase text-[9.5px]">Toggle Visibility:</span>
              <label className="flex items-center gap-1.5 cursor-pointer hover:text-indigo-650 select-none">
                <input 
                  type="checkbox" 
                  checked={displayOptions.showCode} 
                  onChange={(e) => setDisplayOptions({...displayOptions, showCode: e.target.checked})}
                  className="rounded text-indigo-600 focus:ring-indigo-500 border-slate-350 w-3.5 h-3.5"
                />
                Course Code
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer hover:text-indigo-650 select-none">
                <input 
                  type="checkbox" 
                  checked={displayOptions.showName} 
                  onChange={(e) => setDisplayOptions({...displayOptions, showName: e.target.checked})}
                  className="rounded text-indigo-600 focus:ring-indigo-500 border-slate-350 w-3.5 h-3.5"
                />
                Course Name
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer hover:text-indigo-650 select-none">
                <input 
                  type="checkbox" 
                  checked={displayOptions.showFaculty} 
                  onChange={(e) => setDisplayOptions({...displayOptions, showFaculty: e.target.checked})}
                  className="rounded text-indigo-600 focus:ring-indigo-500 border-slate-350 w-3.5 h-3.5"
                />
                Faculty Name
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer hover:text-indigo-650 select-none">
                <input 
                  type="checkbox" 
                  checked={displayOptions.showClassroom} 
                  onChange={(e) => setDisplayOptions({...displayOptions, showClassroom: e.target.checked})}
                  className="rounded text-indigo-600 focus:ring-indigo-500 border-slate-350 w-3.5 h-3.5"
                />
                Classroom
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer hover:text-indigo-650 select-none">
                <input 
                  type="checkbox" 
                  checked={displayOptions.showSchool} 
                  onChange={(e) => setDisplayOptions({...displayOptions, showSchool: e.target.checked})}
                  className="rounded text-indigo-600 focus:ring-indigo-500 border-slate-350 w-3.5 h-3.5"
                />
                Offering School
              </label>
            </div>

            {/* Timetable Grid Matrix */}
            <div ref={gridRef} className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
              
              {/* Context header indicator */}
              <div className="p-2.5 bg-slate-50/50 border-b border-slate-100 flex items-center justify-between text-[11px] text-slate-500 font-display">
                <span>
                  {viewMode === "programme" ? (
                    <>Weekly slots for: <strong>{selectedGroup}</strong></>
                  ) : (
                    <>Weekly teaching calendar for: <strong>{selectedFaculty}</strong></>
                  )}
                </span>
                {viewMode === "programme" && portalMode === "admin" && (
                  <span className="text-[9.5px] text-indigo-650 bg-indigo-50 px-2 py-0.5 rounded-full font-bold">
                    💡 Click on card to move | Click ✏️ to edit course details
                  </span>
                )}
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse table-fixed min-w-[760px]">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50/80 text-[9px] font-bold text-slate-500 uppercase tracking-wider">
                      <th className="py-1.5 px-1.5 w-[55px]">Day</th>
                      
                      {/* Slots 1 and 2 */}
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[0].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[0].time}</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[1].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[1].time}</span>
                      </th>
                      
                      <th className="py-1 px-0.5 border-l border-slate-200 text-center w-[22px] bg-slate-100/40">Break</th>
                      
                      {/* Slots 3 and 4 */}
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[2].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[2].time}</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[3].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[3].time}</span>
                      </th>
                      
                      <th className="py-1 px-0.5 border-l border-slate-200 text-center w-[22px] bg-slate-100/40">Lunch</th>
                      
                      {/* Slots 5, 6, 7 */}
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[4].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[4].time}</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[5].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[5].time}</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        {slots[6].label}
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">{slots[6].time}</span>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-sm">
                    {DAYS.map((day, dIdx) => {
                      return (
                        <tr key={day} className="hover:bg-slate-50/30">
                          {/* Row day header */}
                          <td className="py-1 px-1.5 font-semibold bg-slate-50/60 text-slate-700 border-r border-slate-200 font-display text-[10px]">
                            {day}
                          </td>

                          {/* Render columns: Slots 1-7 and breaks */}
                          {(() => {
                            const cells = [];

                            // Render cell helper
                            const renderCell = (sIdx) => {
                              // MODE A: Student Programme View
                              if (viewMode === "programme") {
                                const cell = timetableData[selectedGroup]?.[dIdx][sIdx];

                                // Skip cell render if it is a multi-hour block part
                                if (cell && (cell.isPart2 || cell.isPart3)) {
                                  return null;
                                }

                                const colSpan = cell ? cell.duration : 1;
                                const palette = cell ? (cell.isFixed ? { bg: 'bg-slate-100', border: 'border-slate-300', text: 'text-slate-700', badge: 'bg-slate-200 text-slate-650' } : getSubjectColor(cell.courseName)) : null;
                                const isSelectedForMove = moveSource && moveSource.group === selectedGroup && moveSource.day === dIdx && moveSource.slot === sIdx;

                                // Check if this slot is a valid target destination for an active move operation
                                let isValidMoveTarget = false;
                                if (moveSource) {
                                  const check = validateMove(moveSource.group, moveSource.day, moveSource.slot, dIdx, sIdx);
                                  isValidMoveTarget = (check === null);
                                }

                                if (!cell) {
                                  return (
                                    <td 
                                      key={`slot-${sIdx}`} 
                                      colSpan={colSpan}
                                      onClick={() => handleCellClick(dIdx, sIdx)}
                                      className={`py-0.5 px-0.5 border-l border-slate-200/60 align-top text-[9px] transition-all
                                        ${isValidMoveTarget 
                                          ? 'bg-emerald-50/65 border-2 border-dashed border-emerald-400 hover:bg-emerald-100/60 scale-[0.99] shadow-inner cursor-pointer' 
                                          : moveSource 
                                            ? 'bg-slate-50/20 opacity-40 cursor-not-allowed' 
                                            : 'bg-slate-50/20 text-slate-400 hover:bg-slate-50/85 cursor-pointer'
                                        }`}
                                    >
                                      {isValidMoveTarget ? (
                                        <div className="h-full flex flex-col justify-center items-center py-0.5 text-emerald-600 font-bold">
                                          <ArrowRightIcon className="w-3 h-3 mb-px animate-pulse" />
                                          <span className="text-[7.5px]">Place</span>
                                        </div>
                                      ) : (
                                        <span className="italic opacity-40 text-[8px] block py-0.5 text-center select-none">Empty</span>
                                      )}
                                    </td>
                                  );
                                }

                                return (
                                  <td
                                    key={`slot-${sIdx}`}
                                    colSpan={colSpan}
                                    onClick={() => handleCellClick(dIdx, sIdx)}
                                    className={`py-0.5 px-0.5 border-l border-slate-200/60 align-top transition-all
                                      ${portalMode === "student" || cell.isFixed
                                        ? '' 
                                        : isSelectedForMove 
                                          ? 'ring-2 ring-indigo-500 scale-[0.98] cursor-pointer' 
                                          : moveSource 
                                            ? 'opacity-40 hover:opacity-50 cursor-pointer' 
                                            : 'hover:scale-[1.01] cursor-pointer'
                                      }`}
                                  >
                                    <div className={`h-full flex flex-col justify-between rounded-md p-1 border shadow-sm transition-all
                                      ${palette.bg} ${palette.border} ${palette.text} 
                                      ${isSelectedForMove ? 'shadow-md ring-1 ring-indigo-300' : ''}
                                    `}>
                                      <div className="space-y-px">
                                        <div className="flex items-start justify-between gap-1">
                                          <div className="min-w-0 flex-1">
                                            {displayOptions.showCode && cell.courseCode && (
                                              <span className="text-[7.5px] font-mono font-bold tracking-tight block truncate opacity-75">
                                                {cell.courseCode}
                                              </span>
                                            )}
                                            {displayOptions.showName && (
                                              <span className="text-[9.5px] font-bold font-display leading-tight block truncate" title={cell.courseName}>
                                                {cell.courseName}
                                              </span>
                                            )}
                                          </div>

                                          {/* Lock indicator for Fixed, Edit indicator for regular courses */}
                                          {cell.isFixed ? (
                                            <span className="text-[8px] text-slate-400 select-none" title="University Locked">🔒</span>
                                          ) : (
                                            portalMode === "admin" && !moveSource && (
                                              <button
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  const targetCourse = courses.find(c => c.courseName === cell.courseName || c.courseCode === cell.courseCode);
                                                  if (targetCourse) {
                                                    setEditingCourse(targetCourse);
                                                  } else {
                                                    alert("Course reference not found in active catalog.");
                                                  }
                                                }}
                                                className="text-[8px] opacity-40 hover:opacity-100 hover:scale-110 shrink-0 select-none cursor-pointer"
                                                title="Edit course details"
                                              >
                                                ✏️
                                              </button>
                                            )
                                          )}
                                        </div>

                                        {displayOptions.showFaculty && cell.faculty && (
                                          <span className="text-[8px] opacity-85 block truncate" title={cell.faculty}>
                                            👨‍🏫 {cell.faculty}
                                          </span>
                                        )}

                                        {displayOptions.showClassroom && cell.classroom && (
                                          <span className="text-[8px] opacity-75 block truncate font-medium text-slate-550" title={cell.classroom}>
                                            🏫 {cell.classroom}
                                          </span>
                                        )}

                                        {displayOptions.showSchool && cell.school && (
                                          <span className="text-[7px] opacity-70 block truncate italic" title={cell.school}>
                                            🏛️ {cell.school}
                                          </span>
                                        )}
                                      </div>

                                      <div className="mt-1 flex items-center justify-between gap-1 flex-wrap">
                                        <span className={`px-0.5 py-px rounded-[3px] text-[6.5px] font-bold uppercase border ${palette.badge}`}>
                                          {cell.isFixed ? 'Fixed' : (cell.courseName.toLowerCase().includes("lab") || cell.courseName.toLowerCase().includes("applied") ? 'Applied' : 'Concept')}
                                        </span>
                                        {colSpan > 1 && (
                                          <span className="text-[6.5px] opacity-65 font-semibold font-mono">{colSpan}h Block</span>
                                        )}
                                        {cell.isMerged && (
                                          <span className="text-[6.5px] bg-indigo-100 text-indigo-700 px-0.5 rounded font-bold" title={`Merged: ${cell.mergeCode}`}>🔗 Merged</span>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                );
                              } 
                              
                              // MODE B: Faculty Calendar View
                              else {
                                const cell = facultyGrid[dIdx][sIdx];

                                if (cell && (cell.isPart2 || cell.isPart3)) {
                                  return null;
                                }

                                const colSpan = cell ? cell.duration : 1;
                                const palette = cell ? getSubjectColor(cell.courseName) : null;

                                if (!cell) {
                                  return (
                                    <td 
                                      key={`slot-${sIdx}`}
                                      colSpan={colSpan}
                                      className="py-0.5 px-0.5 border-l border-slate-200/60 align-top text-[9px] text-slate-450 bg-slate-50/10"
                                    >
                                      <span className="italic opacity-40 text-[8px] block py-0.5 text-center select-none">Free</span>
                                    </td>
                                  );
                                }

                                return (
                                  <td
                                    key={`slot-${sIdx}`}
                                    colSpan={colSpan}
                                    className="py-0.5 px-0.5 border-l border-slate-200/60 align-top"
                                  >
                                    <div className={`h-full flex flex-col justify-between rounded-md p-1 border bg-white shadow-sm
                                      ${palette.bg} ${palette.border} ${palette.text}`}
                                    >
                                      <div className="space-y-px">
                                        {displayOptions.showCode && cell.courseCode && (
                                          <span className="text-[7.5px] font-mono block opacity-75">{cell.courseCode}</span>
                                        )}
                                        <span className="text-[9.5px] font-bold font-display leading-tight block truncate" title={cell.courseName}>
                                          {cell.courseName}
                                        </span>
                                        <span className="text-[8px] opacity-85 block truncate" title={cell.group}>
                                          👥 {cell.group}
                                        </span>
                                        {displayOptions.showClassroom && cell.classroom && (
                                          <span className="text-[8.5px] opacity-75 block truncate font-mono">🏫 {cell.classroom}</span>
                                        )}
                                      </div>
                                      <div className="mt-1 flex items-center justify-between gap-1 flex-wrap">
                                        <span className={`px-0.5 py-px rounded-[3px] text-[6.5px] font-bold uppercase border ${palette.badge}`}>
                                          {cell.courseName.toLowerCase().includes("lab") || cell.courseName.toLowerCase().includes("applied") ? 'Applied' : 'Concept'}
                                        </span>
                                        {colSpan > 1 && (
                                          <span className="text-[6.5px] opacity-65 font-semibold font-mono">{colSpan}h Block</span>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                );
                              }
                            };

                            // Add slots 1 and 2
                            cells.push(renderCell(0));
                            cells.push(renderCell(1));

                            // Add Short Break
                            if (dIdx === 0) {
                              cells.push(
                                <td 
                                  key="short-break-col"
                                  rowSpan={5} 
                                  className="w-[22px] bg-slate-100/35 border-l border-r border-slate-200/60 text-center align-middle py-1 shrink-0 select-none"
                                >
                                  <div className="text-slate-400 font-bold uppercase tracking-widest text-[7px] py-1 mx-auto" style={{ writingMode: 'vertical-rl' }}>
                                    ☕ Break (11:10-11:20)
                                  </div>
                                </td>
                              );
                            }

                            // Add slots 3 and 4
                            cells.push(renderCell(2));
                            cells.push(renderCell(3));

                            // Add Lunch Break
                            if (dIdx === 0) {
                              cells.push(
                                <td 
                                  key="lunch-break-col"
                                  rowSpan={5} 
                                  className="w-[22px] bg-slate-100/35 border-l border-r border-slate-200/60 text-center align-middle py-1 shrink-0 select-none"
                                >
                                  <div className="text-slate-400 font-bold uppercase tracking-widest text-[7px] py-1 mx-auto" style={{ writingMode: 'vertical-rl' }}>
                                    🍽️ Lunch (01:00-01:50)
                                  </div>
                                </td>
                              );
                            }

                            // Add slots 5, 6, 7
                            cells.push(renderCell(4));
                            cells.push(renderCell(5));
                            cells.push(renderCell(6));

                            return cells;
                          })()}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div className="p-3 border-t border-slate-200 bg-slate-50/60 text-xs text-slate-400 text-center font-display">
                Time Matrix View represents standard weekly slots parameters.
              </div>
            </div>

            {/* Faculty Workload Grid */}
            <section className="bg-white border border-slate-200 p-4 rounded-2xl space-y-3 shadow-sm">
              <div className="text-center sm:text-left">
                <h3 className="text-base font-bold text-slate-800 font-display">Faculty Daily Workload Auditing</h3>
                <p className="text-xs text-slate-500">
                  Monitors the workload distribution across slots/day per faculty member.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                {Object.keys(facultyWorkload).map(faculty => {
                  const dailyLoads = facultyWorkload[faculty];
                  const weeklyTotal = dailyLoads.reduce((sum, val) => sum + val, 0);

                  return (
                    <div key={faculty} className="bg-slate-50/50 border border-slate-200/80 rounded-xl p-3 flex flex-col justify-between space-y-2 shadow-sm">
                      <div>
                        <span className="text-[11px] font-semibold text-slate-655 block truncate" title={faculty}>{faculty}</span>
                        <span className="text-sm font-bold text-slate-800 font-display mt-0.5 block">{weeklyTotal} hrs <span className="text-[10px] text-slate-400 font-medium">/ week</span></span>
                      </div>

                      {/* Daily bar indicators */}
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-550 uppercase tracking-wider font-semibold block">Daily Hours:</span>
                        <div className="grid grid-cols-5 gap-1">
                          {dailyLoads.map((hrs, idx) => {
                            const isAtLimit = hrs >= 6;
                            return (
                              <div key={idx} className="flex flex-col items-center">
                                <div className="w-full h-6 bg-slate-100 rounded overflow-hidden relative border border-slate-200">
                                  <div 
                                    className={`absolute bottom-0 w-full transition-all duration-300 
                                      ${isAtLimit 
                                        ? 'bg-rose-500' 
                                        : hrs >= 4 
                                          ? 'bg-amber-500' 
                                          : 'bg-indigo-500'
                                      }`}
                                    style={{ height: `${(hrs / 7) * 100}%` }}
                                  />
                                </div>
                                <span className="text-[7.5px] text-slate-400 font-mono mt-0.5 uppercase">{DAYS[idx].substring(0, 1)}</span>
                                <span className="text-[8px] font-bold font-mono text-slate-500">{hrs}</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>
          </div>
        )}

        {/* Catalog Table Preview */}
        {courses.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-3.5 border-b border-slate-200 flex flex-col md:flex-row items-center justify-between gap-3">
              <div className="space-y-1 text-center md:text-left">
                <h2 className="text-base font-bold text-slate-800 font-display">Source Course Catalog</h2>
                <p className="text-[11px] text-slate-500">
                  Imported database items. Select courses to merge w.r.t faculty.
                </p>
              </div>
              
              <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-end">
                {/* School Filter Dropdown */}
                <div className="flex items-center gap-1">
                  <span className="text-[9.5px] font-semibold text-slate-450 uppercase">School:</span>
                  <select
                    value={selectedSchoolFilter}
                    onChange={(e) => setSelectedSchoolFilter(e.target.value)}
                    className="px-2 py-1 bg-white border border-slate-200 rounded-lg text-xs"
                  >
                    <option value="All">All Schools</option>
                    {uniqueSchools.map(school => (
                      <option key={school} value={school}>{school}</option>
                    ))}
                  </select>
                </div>

                {/* Merge selected action */}
                {selectedCourseIds.length > 1 && portalMode === "admin" && (
                  <button
                    onClick={handleMergeSelected}
                    className="px-2.5 py-1 text-[10px] font-bold bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 shadow-sm flex items-center gap-1 transition-all cursor-pointer"
                  >
                    🔗 Merge Selected ({selectedCourseIds.length})
                  </button>
                )}

                <div className="relative w-full sm:w-48">
                  <SearchIcon className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                  <input 
                    type="text"
                    placeholder="Search courses..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full pl-8 pr-3 py-1.5 bg-white text-xs border border-slate-200 rounded-lg focus:border-indigo-500 focus:outline-none text-slate-855 placeholder-slate-450 transition-all shadow-inner"
                  />
                </div>
              </div>
            </div>

            <div className="overflow-x-auto max-h-64 overflow-y-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-[10px] font-semibold text-slate-500 uppercase tracking-wider sticky top-0 z-10 select-none">
                    {portalMode === "admin" && <th className="py-2 px-2 text-center w-[30px]">Merge</th>}
                    <th className="py-2 px-3 w-[45px]">#</th>
                    <th className="py-2 px-3 w-[75px]">Code</th>
                    <th className="py-2 px-4">Programme</th>
                    <th className="py-2 px-4">Semester</th>
                    <th className="py-2 px-4">Course Name</th>
                    <th className="py-2 px-4">Teaching Faculty</th>
                    <th className="py-2 px-4">Classroom</th>
                    <th className="py-2 px-4">School</th>
                    <th className="py-2 px-3 text-center">Hours</th>
                    {portalMode === "admin" && <th className="py-2 px-3 text-center">Actions</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-[10.5px]">
                  {filteredCourses.length > 0 ? (
                    filteredCourses.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-slate-50/50">
                        {portalMode === "admin" && (
                          <td className="py-1.5 px-2 text-center">
                            <input 
                              type="checkbox"
                              checked={selectedCourseIds.includes(item.id)}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedCourseIds([...selectedCourseIds, item.id]);
                                } else {
                                  setSelectedCourseIds(selectedCourseIds.filter(id => id !== item.id));
                                }
                              }}
                              className="rounded text-indigo-650 w-3 h-3 cursor-pointer"
                            />
                          </td>
                        )}
                        <td className="py-1.5 px-3 font-mono text-slate-400">{idx + 1}</td>
                        <td className="py-1.5 px-3 font-semibold text-indigo-650 font-mono">{item.courseCode}</td>
                        <td className="py-1.5 px-4 text-slate-600">{item.programme}</td>
                        <td className="py-1.5 px-4">
                          <span className="inline-flex items-center px-1.5 py-0.2 rounded text-[9px] font-medium bg-indigo-50 text-indigo-700 border border-indigo-100">
                            {item.semester}
                          </span>
                        </td>
                        <td className="py-1.5 px-4 font-medium text-slate-800">
                          {item.courseName}
                          {item.mergeCode && (
                            <span className="ml-1.5 px-1 py-[1px] bg-indigo-50 text-indigo-600 rounded font-bold font-mono text-[8.5px] border border-indigo-200/20" title={`Merge group key: ${item.mergeCode}`}>
                              🔗 {item.mergeCode}
                            </span>
                          )}
                        </td>
                        <td className="py-1.5 px-4 text-slate-600">{item.faculty}</td>
                        <td className="py-1.5 px-4 font-mono text-slate-500">{item.classroom || "TBD"}</td>
                        <td className="py-1.5 px-4 text-slate-500">{item.school || "Computing"}</td>
                        <td className="py-1.5 px-4 text-center font-mono font-semibold text-indigo-600">{item.hours} hrs</td>
                        
                        {portalMode === "admin" && (
                          <td className="py-1.5 px-3 text-center">
                            <button
                              onClick={() => setEditingCourse(item)}
                              className="px-2 py-0.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-[9px] font-medium text-slate-600 cursor-pointer"
                            >
                              ✏️ Edit
                            </button>
                          </td>
                        )}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={portalMode === "admin" ? "11" : "9"} className="py-8 text-center text-slate-400">No records match search.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {/* Course Editor Modal Dialog */}
      {editingCourse && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 max-w-md w-full shadow-2xl p-5 space-y-4 animate-scaleUp">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="text-sm font-bold text-slate-800 font-display">Edit Course Details</h3>
              <button 
                onClick={() => setEditingCourse(null)}
                className="text-slate-400 hover:text-slate-650 text-sm font-bold p-1 cursor-pointer"
              >
                ✕
              </button>
            </div>
            
            <div className="grid grid-cols-2 gap-3 text-xs">
              <div className="col-span-2">
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Course Name</label>
                <input 
                  type="text" 
                  value={editingCourse.courseName}
                  onChange={(e) => setEditingCourse({...editingCourse, courseName: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Course Code</label>
                <input 
                  type="text" 
                  value={editingCourse.courseCode}
                  onChange={(e) => setEditingCourse({...editingCourse, courseCode: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Teaching Faculty (Reassignment)</label>
                <input 
                  type="text" 
                  value={editingCourse.faculty}
                  onChange={(e) => setEditingCourse({...editingCourse, faculty: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Classroom / Venue</label>
                <input 
                  type="text" 
                  value={editingCourse.classroom}
                  onChange={(e) => setEditingCourse({...editingCourse, classroom: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">School / Department</label>
                <input 
                  type="text" 
                  value={editingCourse.school}
                  onChange={(e) => setEditingCourse({...editingCourse, school: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Hours / Week</label>
                <input 
                  type="number" 
                  value={editingCourse.hours}
                  onChange={(e) => setEditingCourse({...editingCourse, hours: parseFloat(e.target.value) || 0})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Merge Code</label>
                <input 
                  type="text" 
                  placeholder="e.g. MERGE-1"
                  value={editingCourse.mergeCode || ''}
                  onChange={(e) => setEditingCourse({...editingCourse, mergeCode: e.target.value})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                />
              </div>

              <div>
                <label className="block text-[10px] text-slate-400 font-semibold mb-1">Session Block Duration</label>
                <select 
                  value={editingCourse.sessionBlockSize || 0}
                  onChange={(e) => setEditingCourse({...editingCourse, sessionBlockSize: parseInt(e.target.value)})}
                  className="w-full px-2.5 py-1.5 border border-slate-250 rounded-lg text-xs"
                >
                  <option value={0}>Auto Detect (Applied: 2h, graphics: 3h, concept: 1h)</option>
                  <option value={1}>1 Hour Single Sessions</option>
                  <option value={2}>2 Hours Continuous Blocks</option>
                  <option value={3}>3 Hours Continuous Blocks</option>
                </select>
              </div>
            </div>

            <div className="bg-amber-50 text-amber-800 text-[10px] p-2 rounded-lg border border-amber-100 leading-normal">
              💡 <strong>Note:</strong> Modifying hours or session block size requires clicking <strong>Generate Timetables</strong> again. Faculty, room, code, and name edits apply immediately to scheduled slots!
            </div>

            <div className="flex justify-end gap-2 border-t border-slate-100 pt-3">
              <button 
                onClick={() => setEditingCourse(null)}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-650 rounded-lg text-xs cursor-pointer"
              >
                Cancel
              </button>
              <button 
                onClick={() => handleSaveCourseEdit(editingCourse)}
                className="px-3.5 py-1.5 bg-indigo-650 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold shadow-sm cursor-pointer"
              >
                Save Changes
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Settings Modal Dialog */}
      {isSettingsOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 max-w-xl w-full shadow-2xl p-5 space-y-4 animate-scaleUp max-h-[85vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="text-sm font-bold text-slate-800 font-display">University-wide Scheduler Settings</h3>
              <button 
                onClick={() => setIsSettingsOpen(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold p-1 cursor-pointer"
              >
                ✕
              </button>
            </div>

            {/* Inner Settings Tab switches */}
            <div className="flex rounded-lg bg-slate-100 p-0.5 border border-slate-200">
              <button
                onClick={() => setSettingsTab("slots")}
                className={`flex-1 py-1 text-[10px] font-bold rounded cursor-pointer
                  ${settingsTab === "slots" ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' : 'text-slate-500 hover:text-slate-800'}`}
              >
                📅 Customize Time Slots
              </button>
              <button
                onClick={() => setSettingsTab("fixed")}
                className={`flex-1 py-1 text-[10px] font-bold rounded cursor-pointer
                  ${settingsTab === "fixed" ? 'bg-white text-indigo-700 shadow-sm border border-slate-200/50' : 'text-slate-500 hover:text-slate-800'}`}
              >
                🔒 University Fixed Slots
              </button>
            </div>

            {/* Tab Contents */}
            {settingsTab === "slots" ? (
              <div className="space-y-3">
                <p className="text-[11px] text-slate-550">
                  Configure labels and time ranges for each of the 7 daily slots. Changes update timetable grid headers and PDF export layout instantly.
                </p>
                <div className="space-y-2 max-h-64 overflow-y-auto border border-slate-100 rounded-lg p-2 bg-slate-50/50">
                  {slots.map(slot => (
                    <div key={slot.id} className="grid grid-cols-3 gap-2 items-center text-xs bg-white p-2 border border-slate-200 rounded-lg">
                      <span className="font-bold text-slate-700">Slot Index {slot.id + 1}</span>
                      <input 
                        type="text" 
                        value={slot.label}
                        onChange={(e) => {
                          const next = slots.map(s => s.id === slot.id ? { ...s, label: e.target.value } : s);
                          setSlots(next);
                        }}
                        className="px-2 py-1 border border-slate-200 rounded text-xs focus:border-indigo-500 focus:outline-none"
                        placeholder="Label"
                      />
                      <input 
                        type="text" 
                        value={slot.time}
                        onChange={(e) => {
                          const next = slots.map(s => s.id === slot.id ? { ...s, time: e.target.value } : s);
                          setSlots(next);
                        }}
                        className="px-2 py-1 border border-slate-200 rounded text-xs focus:border-indigo-500 focus:outline-none"
                        placeholder="Time range"
                      />
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="space-y-4">
                <p className="text-[11px] text-slate-550">
                  Specify fixed slot activities university-wide (e.g. clubs, assembly, skilldrill). These populate all group timetables and block coordinator schedules automatically during timetable generation.
                </p>
                
                {/* List of current Fixed Slots */}
                <div className="space-y-2 max-h-36 overflow-y-auto border border-slate-100 rounded-lg p-2 bg-slate-50/50">
                  {fixedSlots.length === 0 ? (
                    <span className="text-xs text-slate-400 italic block text-center py-2 select-none">No university fixed slots defined yet.</span>
                  ) : (
                    fixedSlots.map(fs => (
                      <div key={fs.id} className="flex justify-between items-center text-xs bg-white p-2 border border-slate-200 rounded-lg">
                        <div>
                          <strong className="text-indigo-650">{fs.courseName}</strong> ({fs.courseCode})
                          <div className="text-[10px] text-slate-500 mt-0.5">
                            Day: {DAYS[fs.day]} | Slots: {fs.slots.map(s => s + 1).join(', ')} | Coordinator: {fs.faculty} | Venue: {fs.classroom || 'N/A'} | Target: {!fs.targetType || fs.targetType === 'all' ? 'All' : fs.targetType === 'semester' ? `Sem: ${fs.targetValue}` : `Group: ${fs.targetValue}`}
                          </div>
                        </div>
                        <div className="flex gap-2">
                          <button 
                            onClick={() => {
                              setEditingFsId(fs.id);
                              setFsName(fs.courseName || "");
                              setFsCode(fs.courseCode || "");
                              setFsFaculty(fs.faculty || "");
                              setFsRoom(fs.classroom || "");
                              setFsSchool(fs.school || "");
                              setFsDay(fs.day);
                              setFsTargetType(fs.targetType || "all");
                              setFsTargetValue(fs.targetValue || "");
                              setFsSelectedSlots(fs.slots || []);
                            }}
                            className="p-1 text-indigo-600 hover:bg-indigo-50 rounded text-xs cursor-pointer font-semibold"
                            title="Edit lock details"
                          >
                            ✏️ Edit
                          </button>
                          <button 
                            onClick={() => {
                              if (editingFsId === fs.id) {
                                cancelFsEdit();
                              }
                              const nextFixedSlots = fixedSlots.filter(x => x.id !== fs.id);
                              setFixedSlots(nextFixedSlots);
                              updateFirestore({ fixedSlots: nextFixedSlots });
                            }}
                            className="p-1 text-red-500 hover:bg-red-50 rounded text-xs cursor-pointer"
                            title="Remove lock"
                          >
                            🗑️ Delete
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>

                {/* Bulk Import / Template Section */}
                <div className="bg-indigo-50/45 border border-indigo-100 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                  <div className="text-left space-y-0.5">
                    <span className="font-bold text-slate-700 block">Bulk Import Fixed Schedules</span>
                    <span className="text-[10px] text-slate-500 block">Upload spreadsheet to lock activities in bulk.</span>
                  </div>
                  <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                    <button
                      onClick={handleDownloadFixedSlotsTemplate}
                      className="px-2.5 py-1 text-[10px] border border-indigo-200 bg-white text-indigo-755 rounded-lg hover:bg-indigo-50 font-bold transition-all cursor-pointer"
                    >
                      📥 Template
                    </button>
                    <label className="px-2.5 py-1 text-[10px] bg-indigo-650 hover:bg-indigo-700 text-white rounded-lg font-bold shadow-sm transition-all cursor-pointer text-center">
                      📁 Upload
                      <input 
                        type="file" 
                        accept=".xlsx, .xls"
                        onChange={(e) => {
                          if (e.target.files && e.target.files[0]) {
                            parseFixedSlotsExcelFile(e.target.files[0]);
                          }
                        }}
                        className="hidden" 
                      />
                    </label>
                  </div>
                </div>

                {/* Form to add or edit Fixed Slot */}
                 <div className="bg-slate-50 border border-slate-200 p-3 rounded-xl space-y-2 text-xs">
                   <div className="flex justify-between items-center mb-1">
                     <h4 className="font-bold text-slate-700">
                       {editingFsId ? "✏️ Edit Locked Fixed Slot" : "Add New Locked Fixed Slot"}
                     </h4>
                     {editingFsId && (
                       <button 
                         onClick={cancelFsEdit}
                         className="text-[10px] text-indigo-650 hover:text-indigo-800 font-bold cursor-pointer"
                       >
                         Cancel Edit
                       </button>
                     )}
                   </div>
                   <div className="grid grid-cols-2 gap-2">
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Course Name</label>
                       <input 
                         id="new-fs-name" 
                         type="text" 
                         value={fsName}
                         onChange={(e) => setFsName(e.target.value)}
                         placeholder="e.g. Clubs / Assembly" 
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       />
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Course Code</label>
                       <input 
                         id="new-fs-code" 
                         type="text" 
                         value={fsCode}
                         onChange={(e) => setFsCode(e.target.value)}
                         placeholder="e.g. CLUB-101" 
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       />
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Teaching Faculty / Coord.</label>
                       <input 
                         id="new-fs-faculty" 
                         type="text" 
                         value={fsFaculty}
                         onChange={(e) => setFsFaculty(e.target.value)}
                         placeholder="e.g. Coordinator Office" 
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       />
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Classroom / Venue</label>
                       <input 
                         id="new-fs-room" 
                         type="text" 
                         value={fsRoom}
                         onChange={(e) => setFsRoom(e.target.value)}
                         placeholder="e.g. Seminar Hall" 
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       />
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">School / Dept</label>
                       <input 
                         id="new-fs-school" 
                         type="text" 
                         value={fsSchool}
                         onChange={(e) => setFsSchool(e.target.value)}
                         placeholder="e.g. Inter-school" 
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       />
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Select Day</label>
                       <select 
                         id="new-fs-day" 
                         value={fsDay}
                         onChange={(e) => setFsDay(parseInt(e.target.value))}
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       >
                         {DAYS.map((d, i) => <option key={i} value={i}>{d}</option>)}
                       </select>
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Target Type</label>
                       <select 
                         id="new-fs-target-type" 
                         value={fsTargetType}
                         onChange={(e) => {
                           const type = e.target.value;
                           setFsTargetType(type);
                           if (type === 'semester') {
                             setFsTargetValue(uniqueSemesters[0] || "");
                           } else if (type === 'group') {
                             setFsTargetValue(uniqueGroupsList[0] || "");
                           } else {
                             setFsTargetValue("");
                           }
                         }}
                         className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                       >
                         <option value="all">All Semesters & Groups</option>
                         <option value="semester">Specific Semester</option>
                         <option value="group">Specific Programme / Group</option>
                       </select>
                     </div>
                     <div>
                       <label className="block text-[9px] text-slate-400 font-semibold mb-0.5">Target Value</label>
                       {fsTargetType === 'all' ? (
                         <select 
                           disabled
                           className="w-full px-2 py-1 bg-slate-100 border border-slate-200 rounded text-xs text-slate-400 focus:outline-none cursor-not-allowed"
                         >
                           <option value="">N/A (Applies to all)</option>
                         </select>
                       ) : fsTargetType === 'semester' ? (
                         <select 
                           id="new-fs-target-value"
                           value={fsTargetValue}
                           onChange={(e) => setFsTargetValue(e.target.value)}
                           className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                         >
                           {uniqueSemesters.length === 0 ? (
                             <option value="">No semesters (load data first)</option>
                           ) : (
                             uniqueSemesters.map(sem => (
                               <option key={sem} value={sem}>{sem}</option>
                             ))
                           )}
                         </select>
                       ) : (
                         <select 
                           id="new-fs-target-value"
                           value={fsTargetValue}
                           onChange={(e) => setFsTargetValue(e.target.value)}
                           className="w-full px-2 py-1 bg-white border border-slate-200 rounded text-xs focus:outline-none"
                         >
                           {uniqueGroupsList.length === 0 ? (
                             <option value="">No groups (load data first)</option>
                           ) : (
                             uniqueGroupsList.map(grp => (
                               <option key={grp} value={grp}>{grp}</option>
                             ))
                           )}
                         </select>
                       )}
                     </div>
                     <div className="col-span-2">
                       <label className="block text-[9px] text-slate-400 font-semibold mb-1">Select Slots (Check multiple for blocks)</label>
                       <div className="flex flex-wrap gap-3 py-1">
                         {slots.map(s => (
                           <label key={s.id} className="flex items-center gap-1 cursor-pointer">
                             <input 
                               type="checkbox" 
                               checked={fsSelectedSlots.includes(s.id)}
                               onChange={(e) => {
                                 if (e.target.checked) {
                                   setFsSelectedSlots([...fsSelectedSlots, s.id]);
                                 } else {
                                   setFsSelectedSlots(fsSelectedSlots.filter(id => id !== s.id));
                                 }
                               }}
                               className="rounded text-indigo-600 w-3.5 h-3.5"
                             />
                             {s.label}
                           </label>
                         ))}
                       </div>
                     </div>
                   </div>
                   <button 
                     onClick={() => {
                       const name = fsName.trim();
                       const code = fsCode.trim();
                       const faculty = fsFaculty.trim();
                       const room = fsRoom.trim();
                       const school = fsSchool.trim();
                       const day = fsDay;
 
                       if (!name || !code || !faculty || fsSelectedSlots.length === 0) {
                         alert("Please fill name, code, faculty coordinator, and select at least one slot.");
                         return;
                       }
 
                       if (fsTargetType !== 'all' && !fsTargetValue) {
                         alert("Please select/enter a valid target value (or load course data first).");
                         return;
                       }
 
                       if (editingFsId) {
                         // We are editing an existing slot
                         const updatedSlots = fixedSlots.map(x => x.id === editingFsId ? {
                           ...x,
                           courseName: name,
                           courseCode: code,
                           faculty,
                           classroom: room || 'TBD',
                           school: school || 'University',
                           day,
                           slots: [...fsSelectedSlots].sort((a, b) => a - b),
                           targetType: fsTargetType,
                           targetValue: fsTargetType === 'all' ? '' : fsTargetValue
                         } : x);
                         setFixedSlots(updatedSlots);
                         updateFirestore({ fixedSlots: updatedSlots });
                         cancelFsEdit();
                       } else {
                         // Adding a new one
                         const newLock = {
                           id: `fs-${Date.now()}`,
                           courseName: name,
                           courseCode: code,
                           faculty,
                           classroom: room || 'TBD',
                           school: school || 'University',
                           day,
                           slots: [...fsSelectedSlots].sort((a, b) => a - b),
                           targetType: fsTargetType,
                           targetValue: fsTargetType === 'all' ? '' : fsTargetValue
                         };
                         const nextFixedSlots = [...fixedSlots, newLock];
                         setFixedSlots(nextFixedSlots);
                         updateFirestore({ fixedSlots: nextFixedSlots });
                         cancelFsEdit();
                       }
                     }}
                     className="w-full py-1.5 bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 rounded-lg text-[10px] font-bold transition-all cursor-pointer mt-1"
                   >
                     {editingFsId ? "💾 Update Locked Fixed Slot (Requires Regeneration)" : "🔒 Add & Lock Fixed Slot (Requires Regeneration)"}
                   </button>
                 </div>
              </div>
            )}
            <div className="flex justify-end gap-2 border-t border-slate-100 pt-3">
              <button 
                onClick={() => {
                  setIsSettingsOpen(false);
                  updateFirestore({ slots, fixedSlots });
                }}
                className="px-4 py-1.5 bg-indigo-650 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold cursor-pointer shadow-sm"
              >
                Save Settings
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Footer Watermark */}
      <footer className="mt-8 border-t border-slate-205 pt-4 pb-2 text-center text-[10px] text-slate-400 font-medium tracking-wide font-display">
        Designed & Developed by <span className="font-bold bg-clip-text text-transparent bg-gradient-to-r from-indigo-650 to-violet-650">Sachin Sharma</span>
      </footer>
    </div>
  );
}

export default App;
