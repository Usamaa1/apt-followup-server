import ExcelJS from 'exceljs';

const cleanText = (val) => {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.text) return String(val.text).trim();
    if (val.result !== undefined) return String(val.result).trim();
    if (Array.isArray(val.richText)) {
      return val.richText.map((r) => r.text).join('').trim();
    }
  }
  return String(val).trim();
};

// Normalize header text: strip newlines, apostrophes, periods, and parenthetical notes
const normHeader = (val) =>
  cleanText(val)
    .replace(/\r?\n/g, ' ')
    .toUpperCase()
    .replace(/['’.]/g, '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\s+/g, ' ')
    .trim();

const normalizeEnrollment = (raw) => {
  const s = cleanText(raw);
  if (!s) return null;
  const upper = s.toUpperCase();
  if (
    upper === 'NAN' ||
    upper === 'N/A' ||
    upper === '-' ||
    upper.includes('NOT ENROLLED') ||
    upper.includes('NO ID')
  ) {
    return null;
  }
  return s.toUpperCase();
};

const cleanStudentName = (raw) => cleanText(raw).replace(/\s+\.$/, '').trim();

const extractScheduleType = (text) => {
  const upper = cleanText(text).toUpperCase();
  if (upper.includes('TTS')) return 'TTS';
  return 'MWF';
};

const formatExcelDate = (val) => {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return val.toISOString().split('T')[0];
  }
  const parsed = new Date(val);
  if (!isNaN(parsed)) return parsed.toISOString().split('T')[0];
  return cleanText(val);
};

// Calculate which week of the month (1-4) a YYYY-MM-DD date falls into
const getWeekOfMonth = (dateStr) => {
  const parts = dateStr.split('-');
  if (parts.length !== 3) return 1;
  const day = parseInt(parts[2], 10);
  if (day <= 7) return 1;
  if (day <= 14) return 2;
  if (day <= 21) return 3;
  return 4;
};

// Extract batchCode (e.g. 2512G1, 2406C2) & facultyName from filename as fallback
const extractMetadataFromFilename = (filename = '') => {
  const base = String(filename || '').replace(/\.xlsx?$/i, '');
  const batchMatch = base.match(/\b(\d{4}[A-Z]\d+)\b/i);
  const batchCode = batchMatch ? batchMatch[1].toUpperCase() : '';

  let facultyName = '';
  const sirMatch = base.match(/^(?:sir|miss|ms|mr)\.?\s+([^_|-]+)/i);
  if (sirMatch) {
    facultyName = sirMatch[1].trim();
  }

  return { batchCode, facultyName };
};

export const parseUploadedWorkbook = async (buffer, originalFilename = '') => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const fileMeta = extractMetadataFromFilename(originalFilename);
  const firstSheet = workbook.worksheets[0];
  const firstRowValues = [];
  firstSheet.getRow(1).eachCell((cell) => firstRowValues.push(cleanText(cell.value)));
  const firstRowJoined = firstRowValues.join(' | ').toUpperCase();

  // 1. Detect Aptrack Batches Export
  if (firstRowJoined.includes('BATCH ID') && firstRowJoined.includes('TOTAL STUDENT REGISTERED')) {
    return parseAptrackBatches(firstSheet);
  }

  // 2. Detect Manual Attendance Sheet
  let isAttendanceSheet = false;
  for (let r = 1; r <= 8; r++) {
    const rowText = [];
    firstSheet.getRow(r).eachCell((cell) => rowText.push(cleanText(cell.value).toUpperCase()));
    if (rowText.join(' ').includes('ATTENDANCE SHEET')) {
      isAttendanceSheet = true;
      break;
    }
  }
  if (isAttendanceSheet) {
    return parseManualAttendanceSheet(firstSheet);
  }

  // 3. Detect Aptech eProject Submission Workbook ('Submission Format' tab or 'TEAM/GROUP' + 'PROJECT NAME')
  const eProjectSheet =
    workbook.worksheets.find((ws) => ws.name.trim().toUpperCase() === 'SUBMISSION FORMAT') ||
    workbook.worksheets.find((ws) => {
      for (let r = 1; r <= 12; r++) {
        const vals = [];
        ws.getRow(r).eachCell((cell) => vals.push(normHeader(cell.value)));
        const joined = vals.join(' | ');
        if (joined.includes('PROJECT NAME') && (joined.includes('TEAM/GROUP') || joined.includes('VIVA MARKS'))) {
          return true;
        }
      }
      return false;
    });

  if (eProjectSheet) {
    return parseEProjectSubmissionSheet(eProjectSheet, fileMeta);
  }

  // 4. Default: Multi-tab or Single-tab Batch Roster / Student Contact Details (both banner & column formats)
  return parseRosterOrContactWorkbook(workbook, fileMeta);
};

function parseAptrackBatches(sheet) {
  const batches = [];
  const headerMap = {};

  sheet.getRow(1).eachCell((cell, colNumber) => {
    headerMap[cleanText(cell.value).toUpperCase()] = colNumber;
  });

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const aptrackBatchId = cleanText(row.getCell(headerMap['BATCH ID'] || 1).value);
    const rawName = cleanText(row.getCell(headerMap['NAME'] || 2).value);
    const facultyName = cleanText(row.getCell(headerMap['FACULTY'] || 5).value);
    const startDateRaw = row.getCell(headerMap['START DATE'] || 8).value;
    const statusRaw = cleanText(row.getCell(headerMap['STATUS'] || 10).value);

    if (!rawName) return;

    const codeMatch = rawName.match(/_(\d{4}[A-Z]\d+)_/i);
    const batchCode = codeMatch ? codeMatch[1].toUpperCase() : aptrackBatchId.toUpperCase();

    batches.push({
      batchCode,
      aptrackBatchId,
      scheduleType: 'MWF',
      timings: 'MWF (TBD)',
      facultyName: facultyName || 'Unassigned',
      coordinatorName: '',
      courseTitle: rawName,
      startDate: formatExcelDate(startDateRaw) || null,
      status: statusRaw === 'Ongoing' ? 'Ongoing' : 'Ongoing',
    });
  });

  return {
    detectedType: 'aptrack_batches',
    summary: `Detected ${batches.length} batches from Aptrack export.`,
    batches,
    students: [],
    absenteeAlerts: [],
    eProjects: [],
  };
}

function parseManualAttendanceSheet(sheet) {
  let batchCode = '';
  let timings = '';
  let facultyName = '';
  let coordinatorName = '';
  let headerRowIdx = -1;

  for (let r = 1; r <= 10; r++) {
    const cells = [];
    sheet.getRow(r).eachCell((cell, col) => {
      cells.push({ col, text: cleanText(cell.value) });
    });

    for (let i = 0; i < cells.length; i++) {
      const upper = cells[i].text.toUpperCase();
      if (upper.includes('BATCH CODE')) {
        const inline = cells[i].text.split(':')[1]?.trim();
        batchCode = (inline || cells[i + 1]?.text || '').toUpperCase();
      }
      if (upper.includes('TIMINGS')) {
        const inline = cells[i].text.split(':')[1]?.trim();
        timings = inline || cells[i + 1]?.text || '';
      }
      if (upper.startsWith('FACULTY')) {
        const inline = cells[i].text.split(':')[1]?.trim();
        facultyName = inline || cells[i + 1]?.text || '';
      }
      if (upper.includes('COORDINATOR')) {
        const inline = cells[i].text.split(':')[1]?.trim();
        coordinatorName = inline || cells[i + 1]?.text || '';
      }
      if (upper.includes('STD. ID') || upper.includes('STUDENT NAME')) {
        headerRowIdx = r;
      }
    }
  }

  if (headerRowIdx === -1) headerRowIdx = 8;

  const headerRow = sheet.getRow(headerRowIdx);
  let idCol = 2;
  let nameCol = 3;
  let courseCol = 4;
  const dateColumns = [];

  headerRow.eachCell((cell, col) => {
    const val = cell.value;
    const upper = cleanText(val).toUpperCase();
    if (upper.includes('ID') || upper.includes('ENROLLMENT')) idCol = col;
    else if (upper.includes('NAME')) nameCol = col;
    else if (upper.includes('COURSE')) courseCol = col;
    else if (val instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(cleanText(val))) {
      const formatted = formatExcelDate(val);
      if (formatted) dateColumns.push({ col, date: formatted });
    }
  });

  const students = [];
  const absenteeAlerts = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIdx) return;
    const studentName = cleanStudentName(row.getCell(nameCol).value);
    if (!studentName || studentName.toUpperCase() === 'STUDENT NAME') return;

    const enrollmentNo = normalizeEnrollment(row.getCell(idCol).value);
    const courseCode = cleanText(row.getCell(courseCol).value);

    const weeklyMap = {};
    let isFreeze = false;

    dateColumns.forEach(({ col, date }) => {
      const mark = cleanText(row.getCell(col).value).toUpperCase();
      if (mark === 'FRZ') isFreeze = true;

      const month = date.slice(0, 7);
      const weekNumber = getWeekOfMonth(date);
      const key = `${month}_W${weekNumber}`;

      if (!weeklyMap[key]) {
        weeklyMap[key] = { month, weekNumber, days: [], abCount: 0 };
      }
      const normalizedStatus = mark === 'AB' ? 'AB' : mark === 'P' ? 'P' : '';
      weeklyMap[key].days.push({ date, status: normalizedStatus });
      if (mark === 'AB') weeklyMap[key].abCount += 1;
    });

    students.push({
      enrollmentNo,
      studentName,
      courseCode,
      batchCode,
      status: isFreeze ? 'Freeze' : 'Active',
    });

    if (!isFreeze) {
      Object.values(weeklyMap).forEach((wk) => {
        if (wk.abCount >= 2) {
          absenteeAlerts.push({
            batchCode,
            enrollmentNo,
            studentName,
            month: wk.month,
            weekNumber: wk.weekNumber,
            abCount: wk.abCount,
            attendanceDates: wk.days.slice(0, 3),
          });
        }
      });
    }
  });

  const batches = batchCode
    ? [
        {
          batchCode,
          scheduleType: extractScheduleType(timings),
          timings: timings || `${extractScheduleType(timings)} (TBD)`,
          facultyName: facultyName || 'Unassigned',
          coordinatorName,
          status: 'Ongoing',
        },
      ]
    : [];

  return {
    detectedType: 'manual_attendance',
    summary: `Parsed ${students.length} students in ${batchCode || 'sheet'}; flagged ${absenteeAlerts.length} weekly follow-up cases (2+ AB).`,
    batches,
    students,
    absenteeAlerts,
    eProjects: [],
  };
}

function parseEProjectSubmissionSheet(sheet, fileMeta) {
  let headerRowIdx = -1;
  let colMap = {};
  let projectStartDate = '';
  let projectEndDate = '';

  // Scan top 15 rows for metadata and Row 7 header
  for (let r = 1; r <= Math.min(15, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    const label = normHeader(row.getCell(1).value);
    const metaVal = formatExcelDate(row.getCell(3).value);
    if (label.includes('PROJECT START DATE')) projectStartDate = metaVal;
    if (label.includes('PROJECT END DATE')) projectEndDate = metaVal;

    const candidate = {};
    row.eachCell((cell, col) => {
      const h = normHeader(cell.value);
      if (!h) return;
      if (h.includes('TEAM') || h.includes('GROUP')) candidate.teamGroup = col;
      else if (h.includes('PROJECT NAME') || h.includes('PROJECT TITLE')) candidate.projectName = col;
      else if (h.includes('ENROLLMENT')) candidate.enrollment = col;
      else if (h === 'STUDENT NAME' || h === 'NAME') candidate.name = col;
      else if (h.includes('COURSE FAMILY')) candidate.courseFamily = col;
      else if (h.includes('COURSE NAME')) candidate.courseName = col;
      else if (h.includes('SEMESTER')) candidate.semesterNumber = col;
      else if (h.includes('VIVA MARKS') || h === 'MARKS') candidate.vivaMarks = col;
      else if (h.includes('NAME OF STUDENT') || h.includes('SUBMITTED BY')) candidate.submittedBy = col;
      else if (h.includes('SUBMISSION LINK')) candidate.submissionLink = col;
      else if (h.includes('BATCH')) candidate.batchCode = col;
    });

    if (candidate.name && candidate.enrollment) {
      headerRowIdx = r;
      colMap = candidate;
      break;
    }
  }

  if (headerRowIdx === -1) headerRowIdx = 7;

  const batchMap = new Map();
  const students = [];
  const eProjects = [];

  // Forward-fill state for vertically merged group cells (Team/Group, Project Name, Submitted By, Submission Link)
  let lastGroup = '';
  let lastProjectName = '';
  let lastSubmittedBy = '';
  let lastSubmissionLink = '';

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIdx) return;

    const studentName = cleanStudentName(row.getCell(colMap.name || 5).value);
    if (!studentName || normHeader(studentName) === 'STUDENT NAME') return;

    const rowBatch = colMap.batchCode ? cleanText(row.getCell(colMap.batchCode).value).toUpperCase() : '';
    const batchCode = rowBatch || fileMeta.batchCode;
    if (!batchCode) return;

    if (!batchMap.has(batchCode)) {
      batchMap.set(batchCode, {
        batchCode,
        scheduleType: 'MWF',
        timings: 'MWF (TBD)',
        facultyName: fileMeta.facultyName || 'Unassigned',
        coordinatorName: '',
        status: 'Ongoing',
      });
    }

    const enrollmentNo = normalizeEnrollment(row.getCell(colMap.enrollment || 4).value);
    const courseFamily = colMap.courseFamily ? cleanText(row.getCell(colMap.courseFamily).value) : '';
    const courseName = colMap.courseName ? cleanText(row.getCell(colMap.courseName).value) : '';
    const semesterNumber = colMap.semesterNumber ? cleanText(row.getCell(colMap.semesterNumber).value) : '';
    const vivaMarks = colMap.vivaMarks ? cleanText(row.getCell(colMap.vivaMarks).value) : '';

    const grp = colMap.teamGroup ? cleanText(row.getCell(colMap.teamGroup).value) : '';
    const proj = colMap.projectName ? cleanText(row.getCell(colMap.projectName).value).replace(/\r?\n/g, ' ').trim() : '';
    const subBy = colMap.submittedBy ? cleanStudentName(row.getCell(colMap.submittedBy).value) : '';
    const link = colMap.submissionLink ? cleanText(row.getCell(colMap.submissionLink).value) : '';

    if (grp) lastGroup = grp;
    if (proj) lastProjectName = proj;
    if (subBy) lastSubmittedBy = subBy;
    if (link) lastSubmissionLink = link;

    students.push({
      enrollmentNo,
      studentName,
      courseCode: courseName,
      ovStatus: courseFamily,
      batchCode,
      status: 'Active',
    });

    eProjects.push({
      batchCode,
      teamGroup: lastGroup,
      projectName: lastProjectName,
      enrollmentNo,
      studentName,
      courseFamily,
      courseName,
      semesterNumber,
      vivaMarks,
      submittedBy: lastSubmittedBy,
      submissionLink: lastSubmissionLink,
      projectStartDate,
      projectEndDate,
    });
  });

  const batches = Array.from(batchMap.values());
  return {
    detectedType: 'eproject_submission',
    summary: `Parsed ${eProjects.length} eProject student entries across ${batches.length} batch (${batches[0]?.batchCode || 'N/A'}).`,
    batches,
    students,
    absenteeAlerts: [],
    eProjects,
  };
}

function parseRosterOrContactWorkbook(workbook, fileMeta = {}) {
  const batchMap = new Map();
  const students = [];
  const skipTabs = [
    'BATCH END',
    'DROP OUT',
    'FREEZE',
    'BATCH MERGED',
    'BATCH DISSOLVE',
    'DISSOLVE BATCHES',
    'EXTRA STUDENT DATA',
    'DATA',
  ];

  workbook.worksheets.forEach((sheet) => {
    const tabUpper = sheet.name.trim().toUpperCase();
    if (skipTabs.some((s) => tabUpper.includes(s))) return;

    let currentBatchCode = fileMeta.batchCode || '';
    let currentTimings = '';
    let currentFaculty = fileMeta.facultyName || '';
    let currentCoordinator = '';
    let currentCourseTitle = '';
    let colMap = {};

    sheet.eachRow((row) => {
      const rowCells = [];
      row.eachCell((cell, col) => {
        rowCells.push({ col, text: cleanText(cell.value) });
      });
      if (rowCells.length === 0) return;

      const joined = rowCells.map((c) => c.text).join(' | ');
      const joinedUpper = joined.toUpperCase();

      // 1. Detect Table Header Row FIRST (handles flat column-based rosters like 2406C2.xlsx & 2410B2.xlsx)
      const isHeaderRow =
        (joinedUpper.includes('ENROLLMENT') || joinedUpper.includes('STUDENT NAME') || joinedUpper.includes('NAME')) &&
        (joinedUpper.includes('CONTACT') ||
          joinedUpper.includes('FATHER') ||
          joinedUpper.includes('COURSE') ||
          joinedUpper.includes('SEMESTER') ||
          joinedUpper.includes('EMAIL') ||
          joinedUpper.includes('TIME SLOT'));

      if (isHeaderRow) {
        colMap = {};
        rowCells.forEach(({ col, text }) => {
          const u = normHeader(text);
          if (u === 'BATCH CODE' || u === 'BATCH' || u === 'BATCH ID') colMap.batchCode = col;
          else if (u === 'TIME SLOT' || u === 'TIMINGS' || u === 'TIMING' || u === 'SLOT') colMap.timings = col;
          else if (u.includes('COORDINATOR')) colMap.coordinator = col;
          else if (u === 'FACULTY' || u === 'FACULTY NAME') colMap.faculty = col;
          else if (u.includes('ENROLLMENT') || u.includes('STD ID')) colMap.enrollment = col;
          else if (u === 'NAME' || u.includes('STUDENT NAME')) colMap.name = col;
          else if (u.includes('FATHER') || u.includes('GUARDIAN')) colMap.father = col;
          else if (u.includes('STUDENT CONTACT') || u.includes('STUDENTS CONTACT') || u === 'CONTACT 1') colMap.studentPhone = col;
          else if (u.includes('PARENT CONTACT') || u.includes('PARENTS CONTACT') || u === 'CONTACT 2') colMap.parentsPhone = col;
          else if (u.includes('COURSE') || u === 'SEMESTER') colMap.course = col;
          else if (u.includes('ADDRESS')) colMap.address = col;
          else if (u.includes('EMAIL')) colMap.email = col;
          else if (u.includes('OV STATUS')) colMap.ovStatus = col;
        });
        return;
      }

      // 2. Check if row is a Batch Metadata Banner (e.g., "Batch Code: 2406C2" in multi-batch sheets)
      if (joinedUpper.includes('BATCH CODE') && !colMap.name) {
        const codeMatch = joined.match(/BATCH CODE\s*:?\s*([A-Z0-9']+)/i);
        if (codeMatch) {
          currentBatchCode = codeMatch[1].replace(/'/g, '').toUpperCase();
        }

        const timingMatch = joined.match(/((?:MWF|TTS)\s*\(?[0-9\s-]+\)?)/i);
        if (timingMatch) {
          currentTimings = timingMatch[1].trim();
        } else {
          currentTimings = `${extractScheduleType(joined)} (TBD)`;
        }

        const coordMatch = joined.match(/BATCH COORDINATOR\s*:?\s*([^|]+)/i);
        if (coordMatch) currentCoordinator = coordMatch[1].trim();

        const facMatch = joined.match(/FACULTY NAME\s*:?\s*([^|]+)/i);
        if (facMatch) {
          currentFaculty = facMatch[1].trim();
        } else {
          const trailingPerson = joined.match(/\)\s*-?\s*((?:MISS|SIR|MS\.?|MR\.?)\s+[A-Z\s.]+)/i);
          if (trailingPerson) currentFaculty = trailingPerson[1].trim();
        }

        if (currentBatchCode && !batchMap.has(currentBatchCode)) {
          batchMap.set(currentBatchCode, {
            batchCode: currentBatchCode,
            scheduleType: extractScheduleType(currentTimings || joined),
            timings: currentTimings,
            facultyName: currentFaculty || 'Unassigned',
            coordinatorName: currentCoordinator,
            courseTitle: currentCourseTitle,
            status: 'Ongoing',
          });
        }
        return;
      }

      // Capture top title line for course title (only if row 1 is not already a header)
      if (row.number === 1 && !colMap.name) {
        currentCourseTitle = rowCells[0]?.text || '';
        return;
      }

      // 3. Parse Student Data Row (supports batch metadata either from row columns OR banner above)
      if (colMap.name) {
        const studentName = cleanStudentName(row.getCell(colMap.name).value);
        if (!studentName || studentName.toUpperCase().includes('STUDENT NAME')) return;

        const rowBatchCode = colMap.batchCode
          ? cleanText(row.getCell(colMap.batchCode).value).replace(/'/g, '').toUpperCase()
          : currentBatchCode;

        if (!rowBatchCode) return;

        const rowTimings = colMap.timings
          ? cleanText(row.getCell(colMap.timings).value)
          : currentTimings;
        const rowFaculty = colMap.faculty
          ? cleanText(row.getCell(colMap.faculty).value)
          : currentFaculty;
        const rowCoordinator = colMap.coordinator
          ? cleanText(row.getCell(colMap.coordinator).value)
          : currentCoordinator;

        if (!batchMap.has(rowBatchCode)) {
          batchMap.set(rowBatchCode, {
            batchCode: rowBatchCode,
            scheduleType: extractScheduleType(rowTimings),
            timings: rowTimings || `${extractScheduleType(rowTimings)} (TBD)`,
            facultyName: rowFaculty || 'Unassigned',
            coordinatorName: rowCoordinator,
            courseTitle: currentCourseTitle,
            status: 'Ongoing',
          });
        }

        students.push({
          enrollmentNo: colMap.enrollment ? normalizeEnrollment(row.getCell(colMap.enrollment).value) : null,
          studentName,
          fatherName: colMap.father ? cleanText(row.getCell(colMap.father).value) : '',
          studentPhone: colMap.studentPhone ? cleanText(row.getCell(colMap.studentPhone).value) : '',
          parentsPhone: colMap.parentsPhone ? cleanText(row.getCell(colMap.parentsPhone).value) : '',
          courseCode: colMap.course ? cleanText(row.getCell(colMap.course).value) : '',
          address: colMap.address ? cleanText(row.getCell(colMap.address).value) : '',
          email: colMap.email ? cleanText(row.getCell(colMap.email).value).toLowerCase() : '',
          ovStatus: colMap.ovStatus ? cleanText(row.getCell(colMap.ovStatus).value) : '',
          batchCode: rowBatchCode,
          status: 'Active',
        });
      }
    });
  });

  const batches = Array.from(batchMap.values());
  return {
    detectedType: 'roster_or_contacts',
    summary: `Parsed ${batches.length} batches and ${students.length} students across ${workbook.worksheets.length} sheet(s).`,
    batches,
    students,
    absenteeAlerts: [],
    eProjects: [],
  };
}