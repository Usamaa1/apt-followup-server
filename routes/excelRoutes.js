import express from 'express';
import multer from 'multer';
import Batch from '../models/Batch.js';
import Student from '../models/Student.js';
import User from '../models/User.js';
import FollowUp from '../models/FollowUp.js';
import { requireAuth } from '../middleware/auth.js';
import { parseUploadedWorkbook } from '../utils/excelParser.js';

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });

// 1. Preview Parsed Excel File
router.post('/preview', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please select an .xlsx file to upload.' });
    }
    const parsed = await parseUploadedWorkbook(req.file.buffer);
    res.json(parsed);
  } catch (err) {
    console.error('Excel Preview Error:', err);
    res.status(400).json({ message: `Could not parse Excel file: ${err.message}` });
  }
});

// 2. Commit Parsed Data into MongoDB (Smart Merge)
router.post('/commit', requireAuth, async (req, res) => {
  try {
    const { batches = [], students = [], absenteeAlerts = [] } = req.body;

    const allUsers = await User.find();
    const batchIdByCode = new Map();
    let batchesUpserted = 0;
    let studentsMerged = 0;
    let followUpsCreated = 0;

   // Step A: Upsert Batches & link Faculty User ID if displayName matches
    for (const b of batches) {
      if (!b.batchCode) continue;
      const code = b.batchCode.toUpperCase().trim();

      const matchedUser = allUsers.find(
        (u) =>
          u.displayName &&
          b.facultyName &&
          u.displayName.toLowerCase().trim() === b.facultyName.toLowerCase().trim()
      );

      const hasValidTimings = b.timings && b.timings !== 'MWF (TBD)';
      const hasValidFaculty = b.facultyName && b.facultyName !== 'Unassigned';

      const updateFields = {
        batchCode: code,
        scheduleType: b.scheduleType || 'MWF',
        ...(hasValidTimings && { timings: b.timings }),
        ...(hasValidFaculty && { facultyName: b.facultyName }),
        ...(matchedUser && { facultyId: matchedUser._id }),
        ...(b.coordinatorName && { coordinatorName: b.coordinatorName }),
        ...(b.courseTitle && { courseTitle: b.courseTitle }),
        ...(b.aptrackBatchId && { aptrackBatchId: b.aptrackBatchId }),
        status: b.status || 'Ongoing',
      };

      const insertDefaults = {
        ...(!hasValidTimings && { timings: b.timings || 'MWF (TBD)' }),
        ...(!hasValidFaculty && { facultyName: b.facultyName || 'Unassigned' }),
      };

      const savedBatch = await Batch.findOneAndUpdate(
        { batchCode: code },
        {
          $set: updateFields,
          ...(Object.keys(insertDefaults).length > 0 && { $setOnInsert: insertDefaults }),
        },
        { upsert: true, new: true }
      );

      batchIdByCode.set(code, savedBatch);
      batchesUpserted += 1;
    }
    // Step B: Merge Students (Match by Enrollment No first, fallback to Student Name + Phone)
    for (const s of students) {
      if (!s.studentName) continue;
      const code = (s.batchCode || '').toUpperCase().trim();
      let batchDoc = batchIdByCode.get(code);
      if (!batchDoc && code) {
        batchDoc = await Batch.findOne({ batchCode: code });
        if (batchDoc) batchIdByCode.set(code, batchDoc);
      }
      if (!batchDoc) continue;

      let existing = null;
      if (s.enrollmentNo) {
        existing = await Student.findOne({ enrollmentNo: s.enrollmentNo });
      }
      if (!existing) {
        const nameRegex = new RegExp(`^${s.studentName.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i');
        const query = { studentName: nameRegex };
        if (s.studentPhone) query.studentPhone = s.studentPhone.trim();
        existing = await Student.findOne(query);
      }

      if (existing) {
        if (s.enrollmentNo && !existing.enrollmentNo) existing.enrollmentNo = s.enrollmentNo;
        if (s.studentPhone) existing.studentPhone = s.studentPhone.trim();
        if (s.parentsPhone) existing.parentsPhone = s.parentsPhone.trim();
        if (s.email) existing.email = s.email.trim().toLowerCase();
        if (s.fatherName) existing.fatherName = s.fatherName.trim();
        if (s.address) existing.address = s.address.trim();
        if (s.courseCode) existing.courseCode = s.courseCode.trim();
        if (s.ovStatus) existing.ovStatus = s.ovStatus.trim();
        if (s.status) existing.status = s.status;
        existing.currentBatchId = batchDoc._id;
        await existing.save();
      } else {
        await Student.create({
          enrollmentNo: s.enrollmentNo || null,
          studentName: s.studentName.trim(),
          studentPhone: s.studentPhone?.trim() || '',
          parentsPhone: s.parentsPhone?.trim() || '',
          email: s.email?.trim().toLowerCase() || '',
          fatherName: s.fatherName?.trim() || '',
          address: s.address?.trim() || '',
          courseCode: s.courseCode?.trim() || '',
          ovStatus: s.ovStatus?.trim() || '',
          currentBatchId: batchDoc._id,
          status: s.status || 'Active',
        });
      }
      studentsMerged += 1;
    }

    // Step C: Auto-create Weekly FollowUp entries for 2+ AB alerts from Manual Attendance Sheet
    for (const alert of absenteeAlerts) {
      const code = (alert.batchCode || '').toUpperCase().trim();
      const batchDoc = batchIdByCode.get(code) || (await Batch.findOne({ batchCode: code }));
      if (!batchDoc) continue;

      let studentDoc = null;
      if (alert.enrollmentNo) {
        studentDoc = await Student.findOne({ enrollmentNo: alert.enrollmentNo });
      }
      if (!studentDoc) {
        const nameRegex = new RegExp(`^${alert.studentName.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i');
        studentDoc = await Student.findOne({ studentName: nameRegex, currentBatchId: batchDoc._id });
      }
      if (!studentDoc) continue;

      const facultyOwnerId = batchDoc.facultyId || req.user._id;

      await FollowUp.findOneAndUpdate(
        {
          month: alert.month,
          weekNumber: alert.weekNumber,
          batchId: batchDoc._id,
          studentId: studentDoc._id,
        },
        {
          $setOnInsert: {
            month: alert.month,
            weekNumber: alert.weekNumber,
            batchId: batchDoc._id,
            facultyId: facultyOwnerId,
            studentId: studentDoc._id,
            attendanceDates: alert.attendanceDates,
            contactType: 'None',
            dialedContact: studentDoc.studentPhone || '',
          },
        },
        { upsert: true, new: true }
      );
      followUpsCreated += 1;
    }

    res.json({
      message: 'Excel data imported and merged successfully.',
      stats: { batchesUpserted, studentsMerged, followUpsCreated },
    });
  } catch (err) {
    console.error('Excel Commit Error:', err);
    res.status(500).json({ message: 'Failed to save imported Excel data.' });
  }
});

export default router;