import ExcelJS from 'exceljs';

const cleanText = (val) => {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && val.text) return String(val.text).trim();
  if (typeof val === 'object' && val.result !== undefined) return String(val.result).trim();
  return String(val).trim();
};

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
  return s;
};

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

export const parseUploadedWorkbook = async (buffer) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

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

  // 3. Default: Multi-tab or Single-tab Batch Roster / Student Contact Details
  return parseRosterOrContactWorkbook(workbook);
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

    // Extract 6-char batch code like 2410E2 or 2609C3 from T_0726_2410E2_HDSE_3D...
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
  };
}

function parseManualAttendanceSheet(sheet) {
  let batchCode = '';
  let timings = '';
  let facultyName = '';
  let coordinatorName = '';
  let headerRowIdx = -1;

  // Scan top 10 rows for metadata and table header
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
    const studentName = cleanText(row.getCell(nameCol).value);
    if (!studentName || studentName.toUpperCase() === 'STUDENT NAME') return;

    const enrollmentNo = normalizeEnrollment(row.getCell(idCol).value);
    const courseCode = cleanText(row.getCell(courseCol).value);

    // Group attendance by YYYY-MM + Week #
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
  };
}

function parseRosterOrContactWorkbook(workbook) {
  const batchMap = new Map();
  const students = [];
  const skipTabs = ['BATCH END', 'DROP OUT', 'FREEZE', 'BATCH MERGED', 'BATCH DISSOLVE', 'DISSOLVE BATCHES', 'EXTRA STUDENT DATA'];

  workbook.worksheets.forEach((sheet) => {
    const tabUpper = sheet.name.trim().toUpperCase();
    if (skipTabs.some((s) => tabUpper.includes(s))) return;

    let currentBatchCode = '';
    let currentTimings = '';
    let currentFaculty = '';
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

      // Check if row defines Batch Metadata (e.g., "Batch Code: 2406C2")
      if (joinedUpper.includes('BATCH CODE')) {
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
          // Format like "Batch Code: 2308B2 (MWF 9-11)"
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

      // Capture top title line for course title
      if (row.number === 1 && !joinedUpper.includes('ENROLLMENT')) {
        currentCourseTitle = rowCells[0]?.text || '';
        return;
      }

      // Detect table header row
      if (
        (joinedUpper.includes('ENROLLMENT') || joinedUpper.includes('STUDENT NAME') || joinedUpper.includes('NAME')) &&
        (joinedUpper.includes('CONTACT') || joinedUpper.includes('FATHER') || joinedUpper.includes('COURSES'))
      ) {
        colMap = {};
        rowCells.forEach(({ col, text }) => {
          const u = text.toUpperCase();
          if (u.includes('ENROLLMENT') || u.includes('STD. ID')) colMap.enrollment = col;
          else if (u === 'NAME' || u.includes('STUDENT NAME')) colMap.name = col;
          else if (u.includes('FATHER') || u.includes('GUARDIAN')) colMap.father = col;
          else if (u.includes('STUDENT CONTACT') || u === 'CONTACT 1') colMap.studentPhone = col;
          else if (u.includes('PARENTS CONTACT') || u === 'CONTACT 2') colMap.parentsPhone = col;
          else if (u.includes('COURSE')) colMap.course = col;
          else if (u.includes('ADDRESS')) colMap.address = col;
          else if (u.includes('EMAIL')) colMap.email = col;
          else if (u.includes('OV STATUS')) colMap.ovStatus = col;
        });
        return;
      }

      // Parse Student Data Row
      if (colMap.name && currentBatchCode) {
        const studentName = cleanText(row.getCell(colMap.name).value);
        if (!studentName || studentName.toUpperCase().includes('STUDENT NAME')) return;

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
          batchCode: currentBatchCode,
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
  };
}