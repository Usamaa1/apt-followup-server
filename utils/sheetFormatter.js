import { getWeekClassDates } from './weekDates.js';

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const SHORT_MONTHS = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
];

export const getMonthOnlyName = (monthStr) => {
  const [, m] = monthStr.split('-');
  const idx = parseInt(m, 10) - 1;
  return MONTH_NAMES[idx] || monthStr;
};

// Formats "2026-11-02" -> "2-Nov-26"
export const formatShortExcelDate = (isoDate) => {
  if (!isoDate || !isoDate.includes('-')) return isoDate || '';
  const [y, m, d] = isoDate.split('-');
  const dayNum = parseInt(d, 10);
  const monShort = SHORT_MONTHS[parseInt(m, 10) - 1] || m;
  const yrShort = y.slice(-2);
  return `${dayNum}-${monShort}-${yrShort}`;
};

export const normalizeDayMark = (status) => {
  if (!status) return '';
  const u = status.toUpperCase();
  if (u === 'AB') return 'AB';
  if (u === 'P') return 'P';
  if (u === 'LV') return 'LV';
  return u;
};

export const TABLE_HEADERS = [
  'S.#',
  'Enrollment Number',
  'Student Name',
  'Batch Code',
  'Batch Timings',
  'Coordinator Name',
  'Absent Info',
  '',
  '',
  'Faculty Remarks/Reasons',
  'Follow up Date',
  'Medium',
  'Faculty Phone Number',
  'Dialed Contact',
  'Coordinator Remarks/Reasons',
  'ScreenShot',
  'Manager Verification',
];

/**
 * Generates exact rows + merge instructions + styling metadata matching the institute's template
 */
export const buildStyledFollowUpBlueprint = ({ month, batches, followUps }) => {
  const monthOnly = getMonthOnlyName(month);
  const rows = [];
  const rowStyles = []; // Metadata per 1-based row index
  const merges = []; // { sRow, sCol, eRow, eCol } (1-based)
  const screenshotMarkers = []; // { rowIndex, driveUrl, driveFileId }

  // Row 1: Black Title Banner (Merged A1:Q1)
  rows.push([`Students Follow Up Sheet For The Month Of ${monthOnly}`]);
  rowStyles.push({ type: 'titleBanner' });
  merges.push({ sRow: 1, sCol: 1, eRow: 1, eCol: 17 });

  // Row 2: Grey Column Headers (A2:Q2, with G2:I2 merged for "Absent Info")
  rows.push(TABLE_HEADERS);
  rowStyles.push({ type: 'columnHeader' });
  merges.push({ sRow: 2, sCol: 7, eRow: 2, eCol: 9 });

  const mwfBatches = batches.filter((b) => b.scheduleType === 'MWF');
  const ttsBatches = batches.filter((b) => b.scheduleType === 'TTS');

  const renderScheduleGroup = (groupBatches, week, scheduleType) => {
    const classDates = getWeekClassDates(month, week, scheduleType);
    const dateRowIdx = rows.length + 1;

    // Yellow Date Header Row in G:I (Cols 7, 8, 9)
    rows.push([
      '',
      '',
      '',
      '',
      '',
      '',
      formatShortExcelDate(classDates[0]),
      formatShortExcelDate(classDates[1]),
      formatShortExcelDate(classDates[2]),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
    rowStyles.push({ type: 'dateHeader', rowIndex: dateRowIdx });

    if (groupBatches.length === 0) {
      const emptyIdx = rows.length + 1;
      rows.push(['', '', '', `No ${scheduleType} Batches`, '', '', 'NO FOLLOW-UP', '', '', '', '', '', '', '', '', '', '']);
      rowStyles.push({ type: 'dataRow', rowIndex: emptyIdx });
      return;
    }

    for (const batch of groupBatches) {
      const entries = followUps.filter(
        (f) =>
          f.weekNumber === week &&
          String(f.batchId) === String(batch._id) &&
          f.studentId
      );

      // Keep at least 3 rows per batch block (matching your template's visual block height)
      const totalBlockRows = Math.max(entries.length, 3);
      const blockStartRow = rows.length + 1;
      const blockEndRow = blockStartRow + totalBlockRows - 1;

      for (let i = 0; i < totalBlockRows; i++) {
        const currentRowIdx = rows.length + 1;
        const entry = entries[i] || null;

        if (entry) {
          const stu = entry.studentId;
          const enrollment =
            stu?.enrollmentNo && stu.enrollmentNo.trim() !== ''
              ? stu.enrollmentNo
              : 'No ID provided';

          const d0 = normalizeDayMark(entry.attendanceDates?.[0]?.status);
          const d1 = normalizeDayMark(entry.attendanceDates?.[1]?.status);
          const d2 = normalizeDayMark(entry.attendanceDates?.[2]?.status);

          if (entry.screenshotDriveUrl) {
            screenshotMarkers.push({
              rowIndex: currentRowIdx,
              driveUrl: entry.screenshotDriveUrl,
              driveFileId: entry.screenshotDriveFileId,
            });
          }

          rows.push([
            '', // S.# blank like your template (or i + 1 if preferred)
            enrollment,
            stu?.studentName || '',
            i === 0 ? batch.batchCode : '',
            i === 0 ? batch.timings : '',
            i === 0 ? batch.coordinatorName || '' : '',
            d0,
            d1,
            d2,
            entry.facultyRemarks || '',
            entry.followUpDate ? formatShortExcelDate(entry.followUpDate) : '',
            entry.medium || '',
            entry.facultyPhone || '',
            entry.dialedContact || '',
            entry.coordinatorRemarks || '',
            entry.screenshotDriveUrl ? `=IMAGE("${entry.screenshotDriveUrl}")` : '',
            entry.managerVerified ? 'Verified' : '',
          ]);
        } else {
          // Blank peach template row inside the batch block
          rows.push([
            '',
            '',
            '',
            i === 0 ? batch.batchCode : '',
            i === 0 ? batch.timings : '',
            i === 0 ? batch.coordinatorName || '' : '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
          ]);
        }

        rowStyles.push({ type: 'dataRow', rowIndex: currentRowIdx });
      }

      // Vertically merge Batch Code (Col 4), Batch Timings (Col 5), Coordinator Name (Col 6)
      if (blockEndRow > blockStartRow) {
        merges.push({ sRow: blockStartRow, sCol: 4, eRow: blockEndRow, eCol: 4 });
        merges.push({ sRow: blockStartRow, sCol: 5, eRow: blockEndRow, eCol: 5 });
        merges.push({ sRow: blockStartRow, sCol: 6, eRow: blockEndRow, eCol: 6 });
      }
    }
  };

  for (let week = 1; week <= 4; week++) {
    // Green Week Banner (Merged A:Q)
    const weekBannerIdx = rows.length + 1;
    rows.push([`Week # ${week}`]);
    rowStyles.push({ type: 'weekBanner', rowIndex: weekBannerIdx });
    merges.push({ sRow: weekBannerIdx, sCol: 1, eRow: weekBannerIdx, eCol: 17 });

    // MWF Block
    renderScheduleGroup(mwfBatches, week, 'MWF');

    // TTS Block
    renderScheduleGroup(ttsBatches, week, 'TTS');
  }

  return { rows, rowStyles, merges, screenshotMarkers };
};