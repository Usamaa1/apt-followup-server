// Generates the 3 class dates (YYYY-MM-DD) for Week # 1..4 of a given YYYY-MM
export const getWeekClassDates = (monthStr, weekNumber, scheduleType = 'MWF') => {
  const [yearStr, monStr] = monthStr.split('-');
  const year = parseInt(yearStr, 10);
  const monthIdx = parseInt(monStr, 10) - 1; // 0-indexed

  // Target JS getDay() values: Mon=1, Wed=3, Fri=5 | Tue=2, Thu=4, Sat=6
  const targetDays = scheduleType === 'TTS' ? [2, 4, 6] : [1, 3, 5];

  const firstDayOfMonth = new Date(Date.UTC(year, monthIdx, 1));
  const daysInMonth = new Date(Date.UTC(year, monthIdx + 1, 0)).getUTCDate();

  // Find the first Monday (for MWF) or first Tuesday (for TTS) of the month
  const anchorDay = targetDays[0];
  let firstAnchorDate = 1;
  while (
    firstAnchorDate <= 7 &&
    new Date(Date.UTC(year, monthIdx, firstAnchorDate)).getUTCDay() !== anchorDay
  ) {
    firstAnchorDate++;
  }

  const weekStartDate = firstAnchorDate + (weekNumber - 1) * 7;

  return targetDays.map((targetDow, idx) => {
    const offset = idx * 2; // Mon->Wed (+2)->Fri (+4), Tue->Thu (+2)->Sat (+4)
    const dayNum = Math.min(weekStartDate + offset, daysInMonth);
    const d = new Date(Date.UTC(year, monthIdx, dayNum));
    return d.toISOString().split('T')[0];
  });
};