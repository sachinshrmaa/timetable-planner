import { useState, useRef, useMemo, Fragment } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';


// SVG Icons Components for Visual Excellence
const UploadIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l3 3m-3-3l-3 3M6.75 19.5a4.5 4.5 0 01-1.41-8.775 5.25 5.25 0 0110.233-2.33 3 3 0 013.758 3.848A3.752 3.752 0 0118 19.5H6.75z" />
  </svg>
);

const FileIcon = ({ className }) => (
  <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
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
  { id: 1, programme: "BCA 1st", semester: "Semester 1", courseName: "Programming in C", faculty: "Sachin Sharma", hours: 4 },
  { id: 2, programme: "BCA 1st", semester: "Semester 1", courseName: "Computer Fundamentals", faculty: "Sachin Sharma", hours: 3 },
  { id: 3, programme: "BCA 1st", semester: "Semester 1", courseName: "Programming Lab C", faculty: "Sachin Sharma", hours: 4 }, // Lab course
  { id: 4, programme: "BCA 1st", semester: "Semester 1", courseName: "Technical Mathematics", faculty: "Prof. Carl Gauss", hours: 3 },
  
  { id: 5, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Data Structures", faculty: "Dr. Alan Turing", hours: 4 },
  { id: 6, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Database Systems", faculty: "Dr. Grace Hopper", hours: 3 },
  { id: 7, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Software Engineering Lab", faculty: "Dr. Alan Turing", hours: 4 }, // Lab course
  { id: 8, programme: "B.Tech Computer Science", semester: "Semester 3", courseName: "Discrete Mathematics", faculty: "Prof. Carl Gauss", hours: 3 },
  
  { id: 9, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Machine Learning", faculty: "Dr. Geoffrey Hinton", hours: 4 },
  { id: 10, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Neural Networks Lab", faculty: "Dr. Geoffrey Hinton", hours: 4 }, // Lab course
  { id: 11, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Optimization Techniques", faculty: "Prof. Carl Gauss", hours: 3 },
  { id: 12, programme: "M.Tech Artificial Intelligence", semester: "Semester 1", courseName: "Natural Language Processing", faculty: "Dr. Grace Hopper", hours: 3 }
];

function App() {
  const [courses, setCourses] = useState([]);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");

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
  const SLOTS = [
    { id: 0, label: "Slot 1", time: "09:30 AM - 10:20 AM" },
    { id: 1, label: "Slot 2", time: "10:20 AM - 11:10 AM" },
    { id: 2, label: "Slot 3", time: "11:20 AM - 12:10 PM" },
    { id: 3, label: "Slot 4", time: "12:10 PM - 01:00 PM" },
    { id: 4, label: "Slot 5", time: "01:50 PM - 02:40 PM" },
    { id: 5, label: "Slot 6", time: "02:40 PM - 03:30 PM" },
    { id: 6, label: "Slot 7", time: "03:30 PM - 04:20 PM" }
  ];

  const findHeaderMapping = (headers) => {
    const mapping = {};
    const aliases = {
      programme: ['programme', 'program', 'course programme', 'branch', 'dept', 'department'],
      semester: ['semester', 'sem', 'term', 'year'],
      courseName: ['course name', 'coursename', 'course', 'subject name', 'subject', 'title'],
      faculty: ['teaching faculty', 'faculty', 'teacher', 'instructor', 'professor', 'lecturer'],
      hours: ['contact teaching hours/week', 'contact teaching hours', 'contact hours', 'hours/week', 'teaching hours', 'hours', 'credit hours', 'credits']
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

          const courseItem = {
            id: i,
            programme: String(getValByHeader(mapping.programme)).trim(),
            semester: String(getValByHeader(mapping.semester)).trim(),
            courseName: String(getValByHeader(mapping.courseName)).trim(),
            faculty: String(getValByHeader(mapping.faculty)).trim(),
            hours: parseFloat(getValByHeader(mapping.hours)) || 0
          };

          if (courseItem.programme || courseItem.courseName || courseItem.faculty) {
            extractedData.push(courseItem);
          }
        }

        if (extractedData.length === 0) {
          throw new Error(`No course rows found in "${targetSheetName}" sheet.`);
        }

        setCourses(extractedData);
      } catch (err) {
        console.error(err);
        setError(err.message || "An error occurred while parsing the file.");
        setCourses([]);
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

  const handleLoadDemoData = () => {
    setCourses(DEMO_DATA);
    setFileName("Loaded Demo Dataset");
    setError(null);
    setTimetableData(null);
    setRawTimetableBackup(null);
    setUnscheduledLog([]);
    setMoveSource(null);
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

    // 4. Breakdown course hours into sessions
    const sessionsToSchedule = [];
    courses.forEach(c => {
      const groupName = `${c.programme} - ${c.semester}`;
      const nameLower = c.courseName.toLowerCase();
      const isLabOrApplied = nameLower.includes("lab") || nameLower.includes("applied");
      let hoursRemaining = c.hours;

      if (isLabOrApplied) {
        while (hoursRemaining >= 2) {
          sessionsToSchedule.push({
            id: `${c.id}-lab-${hoursRemaining}`,
            courseName: c.courseName,
            faculty: c.faculty,
            group: groupName,
            duration: 2,
            originalCourse: c
          });
          hoursRemaining -= 2;
        }
      }

      while (hoursRemaining >= 1) {
        sessionsToSchedule.push({
          id: `${c.id}-lec-${hoursRemaining}`,
          courseName: c.courseName,
          faculty: c.faculty,
          group: groupName,
          duration: 1,
          originalCourse: c
        });
        hoursRemaining -= 1;
      }
    });

    // 5. Pre-calculate total faculty hours for sorting
    courses.forEach(c => {
      facultyTotalWeeklyHours[c.faculty] += c.hours;
    });

    // Heuristics: Labs first, then long courses, then busy faculty
    sessionsToSchedule.sort((a, b) => {
      if (b.duration !== a.duration) {
        return b.duration - a.duration;
      }
      if (b.originalCourse.hours !== a.originalCourse.hours) {
        return b.originalCourse.hours - a.originalCourse.hours;
      }
      return facultyTotalWeeklyHours[b.faculty] - facultyTotalWeeklyHours[a.faculty];
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
      const { group, faculty, duration, courseName } = session;
      let bestDay = -1;
      let bestSlot = -1;
      let lowestScore = Infinity;
      const rejections = [];

      for (let day = 0; day < 5; day++) {
        for (let slot = 0; slot <= 7 - duration; slot++) {
          let isValid = true;
          const reasons = [];

          // Rule 1: Student Group Clash check
          for (let offset = 0; offset < duration; offset++) {
            if (tempTimetables[group][day][slot + offset] !== null) {
              isValid = false;
              reasons.push("Group occupied");
              break;
            }
          }

          // Rule 2: Faculty Clash check
          if (isValid) {
            for (let offset = 0; offset < duration; offset++) {
              if (tempFacultySchedules[faculty][day][slot + offset] !== null) {
                isValid = false;
                reasons.push("Faculty clash");
                break;
              }
            }
          }

          // Rule 3: Lab block crossover prevention
          if (isValid && duration === 2 && (slot === 1 || slot === 3)) {
            isValid = false;
            reasons.push("Lab would cross break line");
          }

          if (isValid) {
            const groupDailyLoad = getGroupHoursOnDay(group, day);
            const facultyDailyLoad = tempFacultyHours[faculty][day];
            const score = groupDailyLoad + facultyDailyLoad;

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
          tempTimetables[group][bestDay][bestSlot + offset] = {
            courseName,
            faculty,
            duration,
            isPart2: offset === 1
          };
          tempFacultySchedules[faculty][bestDay][bestSlot + offset] = group;
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
        } else if (reasonCounts["Group occupied"] > 10) {
          primaryReason = `Student group timetable is full for this day`;
        } else if (reasonCounts["Lab would cross break line"] > 10) {
          primaryReason = `Lab block cannot cross break line`;
        } else if (reasonCounts["Group occupied"] > 10) {
          primaryReason = `Student group timetable is full for this day`;
        }

        tempUnscheduledLog.push({
          courseName,
          faculty,
          group,
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
    
    if (groups.length > 0) {
      setSelectedGroup(groups[0]);
    }
    if (uniqueFacultyList.length > 0) {
      setSelectedFaculty(uniqueFacultyList[0]);
    }
  };

  // --- MANUAL SLOT REALLOCATION ---
  
  // Calculate total hours a faculty member teaches on a given day
  const getFacultyHoursOnDay = (fac, day, currentTimetable) => {
    let count = 0;
    Object.keys(currentTimetable).forEach(g => {
      for (let s = 0; s < 7; s++) {
        const c = currentTimetable[g][day][s];
        if (c && c.faculty === fac && !c.isPart2) {
          count += c.duration;
        }
      }
    });
    return count;
  };

  // Dry-run validator for moves
  const validateMove = (group, srcDay, srcSlot, destDay, destSlot) => {
    const cell = timetableData[group][srcDay][srcSlot];
    if (!cell) return "Source slot is empty.";

    const startSlot = cell.isPart2 ? srcSlot - 1 : srcSlot;
    const actualCell = timetableData[group][srcDay][startSlot];
    const dur = actualCell.duration;
    const faculty = actualCell.faculty;
    const courseName = actualCell.courseName;

    // 1. Boundary check
    if (destSlot + dur > 7) {
      return `A ${dur}-hour block goes beyond the 7 daily slots.`;
    }

    // 1b. Break crossover check
    if (dur === 2 && (destSlot === 1 || destSlot === 3)) {
      return `A 2-hour block cannot start at Slot 2 or Slot 4 to avoid crossing break lines.`;
    }

    // 2. Student Group Clash check
    for (let o = 0; o < dur; o++) {
      const existing = timetableData[group][destDay][destSlot + o];
      if (existing) {
        // Allow it if it's the class we're moving (shifting on the same day/slots)
        const isSelf = (destDay === srcDay) && (destSlot + o >= startSlot && destSlot + o < startSlot + dur);
        if (!isSelf) {
          return `Student group already has class "${existing.courseName}" scheduled at slot ${destSlot + o + 1} on ${DAYS[destDay]}.`;
        }
      }
    }

    // 3. Faculty Clash check
    for (let o = 0; o < dur; o++) {
      const targetSlot = destSlot + o;
      // Scan all other groups
      for (const grp of Object.keys(timetableData)) {
        const cellAtTarget = timetableData[grp][destDay][targetSlot];
        if (cellAtTarget && cellAtTarget.faculty === faculty) {
          const isSelf = (grp === group) && (destDay === srcDay) && (targetSlot >= startSlot && targetSlot < startSlot + dur);
          if (!isSelf) {
            return `Lecturer ${faculty} is already teaching "${grp}" in that slot.`;
          }
        }
      }
    }

    return null; // Valid!
  };

  const handleCellClick = (dayIdx, slotIdx) => {
    // If we haven't selected a move source yet, try to select one
    if (!moveSource) {
      if (viewMode !== "programme") return; // Only allow selecting source in Programme View
      const cell = timetableData[selectedGroup]?.[dayIdx][slotIdx];
      if (cell) {
        const startSlot = cell.isPart2 ? slotIdx - 1 : slotIdx;
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

      // 1. Clear source slots
      for (let o = 0; o < dur; o++) {
        nextTimetable[group][srcDay][srcSlot + o] = null;
      }

      // 2. Populate destination slots
      for (let o = 0; o < dur; o++) {
        nextTimetable[group][dayIdx][slotIdx + o] = {
          courseName: actualCell.courseName,
          faculty: actualCell.faculty,
          duration: dur,
          isPart2: o === 1
        };
      }

      // 3. Re-calculate workloads
      const nextWorkload = {};
      Object.keys(facultyWorkload).forEach(fac => {
        nextWorkload[fac] = Array(5).fill(0);
      });

      Object.keys(nextTimetable).forEach(g => {
        for (let d = 0; d < 5; d++) {
          for (let s = 0; s < 7; s++) {
            const c = nextTimetable[g][d][s];
            if (c && !c.isPart2) {
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

  const filteredCourses = useMemo(() => {
    if (!searchTerm.trim()) return courses;
    const term = searchTerm.toLowerCase();
    return courses.filter(item => 
      item.programme.toLowerCase().includes(term) ||
      item.semester.toLowerCase().includes(term) ||
      item.courseName.toLowerCase().includes(term) ||
      item.faculty.toLowerCase().includes(term)
    );
  }, [courses, searchTerm]);

  // Unique lists for selectors
  const uniqueGroups = useMemo(() => {
    if (!timetableData) return [];
    return Object.keys(timetableData);
  }, [timetableData]);

  const uniqueFaculty = useMemo(() => {
    if (courses.length === 0) return [];
    return [...new Set(courses.map(c => c.faculty))];
  }, [courses]);

  return (
    <div className="max-w-7xl mx-auto py-4 px-2 sm:px-4 space-y-6">
      {/* Header Area */}
      <header className="text-center relative z-10 py-3">
        <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-indigo-50 text-indigo-600 border border-indigo-100 mb-2 tracking-wider uppercase">
          University Timetable Generator
        </span>
        <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight font-display bg-clip-text text-transparent bg-gradient-to-r from-indigo-600 via-violet-600 to-indigo-700 pb-1">
          Curriculum Scheduling Dashboard
        </h1>
        <p className="mt-1 text-xs text-slate-500 max-w-xl mx-auto">
          Parse worksheets, generate clash-free schedules, and interactively adjust time slots for groups or faculty lecturers.
        </p>
      </header>

      {/* Quick Actions Panel */}
      <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
        {/* Template Card */}
        <div className="bg-white border border-slate-200/80 p-3.5 rounded-xl flex flex-col justify-between hover:border-slate-300 hover:shadow-sm transition-all duration-300">
          <div>
            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5 font-display">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
              Excel Schema Required
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Your workbook sheet must be named <code className="text-indigo-600 px-0.5 bg-indigo-50/50 border border-indigo-100/30 rounded text-[10px] font-mono">Course_List</code>.
            </p>
          </div>
          <button 
            onClick={handleDownloadTemplate}
            className="mt-2.5 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-lg text-xs font-medium border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 transition-all"
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
              Load our mock curriculum (BCA & B.Tech modules, overlapping faculty) to test schedule generation.
            </p>
          </div>
          <button 
            onClick={handleLoadDemoData}
            className="mt-2.5 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-lg text-xs font-medium bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white shadow-sm shadow-indigo-600/10 hover:scale-[1.01] active:scale-[0.99] transition-all"
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
              5 days × 7 slots with short & lunch breaks. Bundles "Lab"/"Applied" into 2 consecutive hours.
            </p>
          </div>
          <div className="mt-2.5 flex items-center gap-1.5 text-[10px] text-indigo-600 bg-indigo-50/50 border border-indigo-100/60 p-2 rounded-lg">
            <InfoIcon className="w-3.5 h-3.5 shrink-0" />
            <span>Rearranging dynamically audits constraints.</span>
          </div>
        </div>
      </section>

      {/* Upload Container Area */}
      <main className="space-y-6">
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

        {/* Parsing success state banner + Generate */}
        {fileName && !error && !loading && (
          <div className="relative overflow-hidden bg-white border border-slate-200 p-3.5 rounded-xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-sm">
            <div className="flex items-center gap-2">
              <CheckCircleIcon className="w-5 h-5 text-emerald-600" />
              <div className="text-left">
                <span className="text-[9px] text-slate-400 uppercase font-semibold block leading-none">Uploaded Document</span>
                <span className="text-xs font-semibold font-mono text-slate-855 block mt-0.5">{fileName}</span>
                <span className="text-[10px] text-slate-500 block mt-0.5">{courses.length} courses loaded from Course_List</span>
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

        {/* Unscheduled Warnings banner */}
        {unscheduledLog.length > 0 && (
          <div className="relative overflow-hidden bg-amber-50 border border-amber-100 text-amber-800 p-3.5 rounded-xl space-y-2.5 shadow-sm">
            <div className="absolute top-0 bottom-0 left-0 w-1 bg-amber-500" />
            <div className="flex items-start gap-2.5">
              <ExclamationTriangleIcon className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-bold text-amber-900 font-display">Unscheduled Course Hours ({unscheduledLog.length})</h4>
                <p className="text-[10px] text-slate-550">
                  Some hours could not fit automatically. Review workloads or rearrange slots manually:
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
                Moving <strong>{moveSource.cell.courseName}</strong> ({moveSource.cell.duration}h block) from <strong>{DAYS[moveSource.day]} {SLOTS[moveSource.slot].label}</strong>.
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
          <div className="space-y-8 animate-fadeIn">
            
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
                        ? 'bg-white text-indigo-750 shadow-sm border border-slate-200/50' 
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
                        ? 'bg-white text-indigo-750 shadow-sm border border-slate-200/50' 
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
                      className="w-full sm:w-48 px-2 py-1 bg-white text-slate-800 border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-xs transition-all"
                    >
                      {uniqueGroups.map(grp => (
                        <option key={grp} value={grp}>{grp}</option>
                      ))}
                    </select>
                  ) : (
                    <select
                      value={selectedFaculty}
                      onChange={(e) => setSelectedFaculty(e.target.value)}
                      className="w-full sm:w-48 px-2 py-1 bg-white text-slate-800 border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-xs transition-all"
                    >
                      {uniqueFaculty.map(fac => (
                        <option key={fac} value={fac}>{fac}</option>
                      ))}
                    </select>
                  )}
                </div>

                {/* Reset schedule modification button */}
                <button
                  onClick={handleResetTimetable}
                  className="w-full sm:w-auto px-3 py-1.5 text-[10px] font-bold text-slate-650 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg transition-all cursor-pointer"
                  title="Discard all slot modifications"
                >
                  Reset Timetable
                </button>

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
                {viewMode === "programme" && (
                  <span className="text-[9.5px] text-indigo-650 bg-indigo-50 px-2 py-0.5 rounded-full font-bold">
                    💡 Click on card to move/rearrange
                  </span>
                )}
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse table-fixed min-w-[760px]">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50/80 text-[9px] font-bold text-slate-500 uppercase tracking-wider">
                      <th className="py-1 px-1.5 w-[55px]">Day</th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 1
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">09:30 AM-10:20 AM</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 2
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">10:20 AM-11:10 AM</span>
                      </th>
                      <th className="py-1 px-0.5 border-l border-slate-200 text-center w-[22px] bg-slate-100/40">Break</th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 3
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">11:20 AM-12:10 PM</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 4
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">12:10 PM-01:00 PM</span>
                      </th>
                      <th className="py-1 px-0.5 border-l border-slate-200 text-center w-[22px] bg-slate-100/40">Lunch</th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 5
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">01:50 PM-02:40 PM</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 6
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">02:40 PM-03:30 PM</span>
                      </th>
                      <th className="py-1 px-1 border-l border-slate-200 text-center w-[95px]">
                        Slot 7
                        <span className="block text-[7px] text-slate-400 font-mono font-normal mt-px">03:30 PM-04:20 PM</span>
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

                                // If second hour of a 2-hour lab, skip cell render
                                if (cell && cell.duration === 2 && cell.isPart2) {
                                  return null;
                                }

                                const colSpan = (cell && cell.duration === 2 && !cell.isPart2) ? 2 : 1;
                                const palette = cell ? getSubjectColor(cell.courseName) : null;
                                const isSelectedForMove = moveSource && moveSource.group === selectedGroup && moveSource.day === dIdx && moveSource.slot === (cell?.isPart2 ? sIdx - 1 : sIdx);

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
                                      className={`py-0.5 px-0.5 border-l border-slate-200/60 align-top text-[9px] transition-all cursor-pointer
                                        ${isValidMoveTarget 
                                          ? 'bg-emerald-50/65 border-2 border-dashed border-emerald-400 hover:bg-emerald-100/60 scale-[0.99] shadow-inner' 
                                          : moveSource 
                                            ? 'bg-slate-50/20 opacity-40 cursor-not-allowed' 
                                            : 'bg-slate-50/20 text-slate-400 hover:bg-slate-50/85'
                                        }`}
                                    >
                                      {isValidMoveTarget ? (
                                        <div className="h-full flex flex-col justify-center items-center py-0.5 text-emerald-600 font-bold">
                                          <ArrowRightIcon className="w-3 h-3 mb-px animate-pulse" />
                                          <span className="text-[7.5px]">Place</span>
                                        </div>
                                      ) : (
                                        <span className="italic opacity-40 text-[8px] block py-0.5 text-center">Empty</span>
                                      )}
                                    </td>
                                  );
                                }

                                const isLab = cell.courseName.toLowerCase().includes("lab") || cell.courseName.toLowerCase().includes("applied");

                                return (
                                  <td
                                    key={`slot-${sIdx}`}
                                    colSpan={colSpan}
                                    onClick={() => handleCellClick(dIdx, sIdx)}
                                    className={`py-0.5 px-0.5 border-l border-slate-200/60 align-top transition-all cursor-pointer
                                      ${isSelectedForMove 
                                        ? 'ring-2 ring-indigo-500 scale-[0.98]' 
                                        : moveSource 
                                          ? 'opacity-40 hover:opacity-50' 
                                          : 'hover:scale-[1.01]'
                                      }`}
                                  >
                                    <div className={`h-full flex flex-col justify-between rounded-md p-1 border shadow-sm transition-all
                                      ${palette.bg} ${palette.border} ${palette.text} 
                                      ${isSelectedForMove ? 'shadow-md ring-1 ring-indigo-300' : ''}
                                    `}>
                                      <div className="space-y-px">
                                        <div className="flex items-start justify-between gap-0.5">
                                          <span className="text-[9.5px] font-bold font-display leading-tight block truncate" title={cell.courseName}>
                                            {cell.courseName}
                                          </span>
                                          {viewMode === "programme" && !moveSource && (
                                            <span className="text-[8px] opacity-40 hover:opacity-100 shrink-0 select-none">
                                              ✏️
                                            </span>
                                          )}
                                        </div>
                                        <span className="text-[8px] opacity-85 block truncate mt-0.5" title={cell.faculty}>
                                          👨‍🏫 {cell.faculty}
                                        </span>
                                      </div>
                                      <div className="mt-0.5 flex items-center justify-between">
                                        <span className={`px-0.5 py-px rounded-[3px] text-[7px] font-bold uppercase border ${palette.badge}`}>
                                          {isLab ? 'Lab' : 'Lecture'}
                                        </span>
                                        {colSpan > 1 && (
                                          <span className="text-[7px] opacity-65 font-semibold font-mono">2h Block</span>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                );
                              } 
                              
                              // MODE B: Faculty Calendar View
                              else {
                                const cell = facultyGrid[dIdx][sIdx];

                                // If second hour of a 2-hour lab block, skip cell render
                                if (cell && cell.duration === 2 && cell.isPart2) {
                                  return null;
                                }

                                const colSpan = (cell && cell.duration === 2 && !cell.isPart2) ? 2 : 1;
                                const palette = cell ? getSubjectColor(cell.courseName) : null;

                                if (!cell) {
                                  return (
                                    <td 
                                      key={`slot-${sIdx}`}
                                      colSpan={colSpan}
                                      className="py-0.5 px-0.5 border-l border-slate-200/60 align-top text-[9px] text-slate-450 bg-slate-50/10"
                                    >
                                      <span className="italic opacity-40 text-[8px] block py-0.5 text-center">Free</span>
                                    </td>
                                  );
                                }

                                const isLab = cell.courseName.toLowerCase().includes("lab") || cell.courseName.toLowerCase().includes("applied");

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
                                        <span className="text-[9.5px] font-bold font-display leading-tight block truncate" title={cell.courseName}>
                                          {cell.courseName}
                                        </span>
                                        <span className="text-[8px] opacity-85 block truncate mt-0.5" title={cell.group}>
                                          👥 {cell.group}
                                        </span>
                                      </div>
                                      <div className="mt-0.5 flex items-center justify-between">
                                        <span className={`px-0.5 py-px rounded-[3px] text-[7px] font-bold uppercase border ${palette.badge}`}>
                                          {isLab ? 'Lab' : 'Lecture'}
                                        </span>
                                        {colSpan > 1 && (
                                          <span className="text-[7px] opacity-65 font-semibold font-mono">2h Block</span>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                );
                              }
                            };

                            // Add slots 1 and 2 (indices 0 and 1)
                            cells.push(renderCell(0));
                            cells.push(renderCell(1));

                            // Add Short Break (10m) column
                            if (dIdx === 0) {
                              cells.push(
                                <td 
                                  key="short-break-col"
                                  rowSpan={5} 
                                  className="w-[22px] bg-slate-100/30 border-l border-r border-slate-200/60 text-center align-middle py-1 shrink-0 select-none"
                                >
                                  <div className="flex flex-col items-center justify-center h-full text-slate-400 font-bold uppercase tracking-widest text-[7px] py-1" style={{ writingMode: 'vertical-rl' }}>
                                    ☕ Break (11:10-11:20)
                                  </div>
                                </td>
                              );
                            }

                            // Add slots 3 and 4 (indices 2 and 3)
                            cells.push(renderCell(2));
                            cells.push(renderCell(3));

                            // Add Lunch Break (50m) column
                            if (dIdx === 0) {
                              cells.push(
                                <td 
                                  key="lunch-break-col"
                                  rowSpan={5} 
                                  className="w-[22px] bg-slate-100/30 border-l border-r border-slate-200/60 text-center align-middle py-1 shrink-0 select-none"
                                >
                                  <div className="flex flex-col items-center justify-center h-full text-slate-400 font-bold uppercase tracking-widest text-[7px] py-1" style={{ writingMode: 'vertical-rl' }}>
                                    🍽️ Lunch (01:00-01:50)
                                  </div>
                                </td>
                              );
                            }

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
                  Monitors the workload distribution across 7 slots/day per faculty member.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                {Object.keys(facultyWorkload).map(faculty => {
                  const dailyLoads = facultyWorkload[faculty];
                  const weeklyTotal = dailyLoads.reduce((sum, val) => sum + val, 0);

                  return (
                    <div key={faculty} className="bg-slate-50/50 border border-slate-200/80 rounded-xl p-3 flex flex-col justify-between space-y-2 shadow-sm">
                      <div>
                        <span className="text-[11px] font-semibold text-slate-650 block truncate" title={faculty}>{faculty}</span>
                        <span className="text-sm font-bold text-slate-800 font-display mt-0.5 block">{weeklyTotal} hrs <span className="text-[10px] text-slate-400 font-medium">/ week</span></span>
                      </div>

                      {/* Daily bar indicators */}
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-500 uppercase tracking-wider font-semibold block">Daily Hours:</span>
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
            <div className="p-3.5 border-b border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
              <div className="space-y-0.5 text-center sm:text-left">
                <h2 className="text-base font-bold text-slate-800 font-display">Source Course Catalog</h2>
                <p className="text-[11px] text-slate-500">
                  Imported database items parsed from worksheet.
                </p>
              </div>
              <div className="relative w-full sm:w-64">
                <SearchIcon className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Search catalog courses..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-8 pr-3 py-1.5 bg-white text-xs border border-slate-200 rounded-lg focus:border-indigo-500 focus:outline-none text-slate-855 placeholder-slate-450 transition-all"
                />
              </div>
            </div>

            <div className="overflow-x-auto max-h-48 overflow-y-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-[10px] font-semibold text-slate-500 uppercase tracking-wider sticky top-0 z-10">
                    <th className="py-2 px-4">#</th>
                    <th className="py-2 px-4">Programme</th>
                    <th className="py-2 px-4">Semester</th>
                    <th className="py-2 px-4">Course Name</th>
                    <th className="py-2 px-4">Teaching Faculty</th>
                    <th className="py-2 px-4 text-center">Hours</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-[10.5px]">
                  {filteredCourses.length > 0 ? (
                    filteredCourses.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-slate-50/50">
                        <td className="py-1.5 px-4 font-mono text-slate-400">{idx + 1}</td>
                        <td className="py-1.5 px-4 text-slate-600">{item.programme}</td>
                        <td className="py-1.5 px-4">
                          <span className="inline-flex items-center px-1.5 py-0.2 rounded text-[9px] font-medium bg-indigo-50 text-indigo-700 border border-indigo-100">
                            {item.semester}
                          </span>
                        </td>
                        <td className="py-1.5 px-4 font-medium text-slate-800">{item.courseName}</td>
                        <td className="py-1.5 px-4 text-slate-600">{item.faculty}</td>
                        <td className="py-1.5 px-4 text-center font-mono font-semibold text-indigo-600">{item.hours} hrs</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="6" className="py-8 text-center text-slate-400">No records match search.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
