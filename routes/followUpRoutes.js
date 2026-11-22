import express from 'express';
import multer from 'multer';
import Batch from '../models/Batch.js';
import Student from '../models/Student.js';
import FollowUp from '../models/FollowUp.js';
import User from '../models/User.js';
import { requireAuth, requireAdmin, sanitizeStudentForRole } from '../middleware/auth.js';
import { getWeekClassDates } from '../utils/weekDates.js';
import { uploadScreenshotToFacultyDrive } from '../utils/googleDriveUpload.js';

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// 1. Get Full Workspace Data for a Month, Week (1-4), and Faculty Sheet
router.get('/workspace', requireAuth, async (req, res) => {
  try {
    const month = req.query.month || new Date().toISOString().slice(0, 7);
    const weekNumber = Math.min(4, Math.max(1, parseInt(req.query.week || '1', 10)));

    // Admin can switch bottom faculty tabs via ?facultyId=...
    let targetUser = req.user;
    if (req.user.role === 'admin' && req.query.facultyId) {
      const found = await User.findById(req.query.facultyId);
      if (found) targetUser = found;
    }

    // Load all faculty users for Admin bottom sheet tabs
    const facultyTabs =
      req.user.role === 'admin'
        ? await User.find({ isActive: true }).select('_id displayName name email role googleSheetUrl lastSyncedAt').sort({ displayName: 1 })
        : [
            {
              _id: req.user._id,
              displayName: req.user.displayName || req.user.name,
              email: req.user.email,
              role: req.user.role,
              googleSheetUrl: req.user.googleSheetUrl,
              lastSyncedAt: req.user.lastSyncedAt,
            },
          ];

    // Fetch batches belonging to targetUser
    const batches = await Batch.find({
      status: 'Ongoing',
      $or: [
        { facultyId: targetUser._id },
        { facultyName: new RegExp(`^${(targetUser.displayName || targetUser.name || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i') },
      ],
    }).sort({ scheduleType: 1, batchCode: 1 });

    const batchSections = [];

    for (const batch of batches) {
      const weekDates = getWeekClassDates(month, weekNumber, batch.scheduleType);

      // Load active students in this batch (sanitized for role)
      const rawStudents = await Student.find({
        currentBatchId: batch._id,
        status: { $ne: 'Freeze' },
      }).sort({ studentName: 1 });
      const students = rawStudents.map((s) => sanitizeStudentForRole(s, req.user.role));

      // Load existing follow-up rows for this batch/month/week
      let rows = await FollowUp.find({
        month,
        weekNumber,
        batchId: batch._id,
      })
        .populate('studentId')
        .sort({ createdAt: 1 });

      // Ensure at least ONE empty starter row exists per batch (Excel-style ready grid)
      if (rows.length === 0) {
        const starter = await FollowUp.create({
          month,
          weekNumber,
          batchId: batch._id,
          facultyId: batch.facultyId || targetUser._id,
          studentId: null,
          attendanceDates: weekDates.map((d) => ({ date: d, status: '' })),
          contactType: 'None', // "No WhatsApp Message" selected by default
          dialedContact: '',
        });
        rows = [starter];
      }

      const formattedRows = rows.map((r) => {
        const obj = r.toObject();
        // Sync dates if missing
        if (!obj.attendanceDates || obj.attendanceDates.length !== 3) {
          obj.attendanceDates = weekDates.map((d) => ({ date: d, status: '' }));
        }
        if (obj.studentId) {
          obj.student = sanitizeStudentForRole(obj.studentId, req.user.role);
          obj.studentId = obj.student._id;
        } else {
          obj.student = null;
        }
        return obj;
      });

      batchSections.push({
        batch,
        weekDates,
        students,
        rows: formattedRows,
      });
    }

    res.json({
      month,
      weekNumber,
      activeFaculty: {
        _id: targetUser._id,
        displayName: targetUser.displayName || targetUser.name,
        googleSheetUrl: targetUser.googleSheetUrl,
        lastSyncedAt: targetUser.lastSyncedAt,
        whatsappTemplate: targetUser.whatsappTemplate,
      },
      facultyTabs,
      batchSections,
    });
  } catch (err) {
    console.error('Workspace Load Error:', err);
    res.status(500).json({ message: 'Failed to load follow-up workspace.' });
  }
});

// 2. Add an Extra Blank Row to a Batch Table
router.post('/rows', requireAuth, async (req, res) => {
  try {
    const { month, weekNumber, batchId } = req.body;
    const batch = await Batch.findById(batchId);
    if (!batch) return res.status(404).json({ message: 'Batch not found.' });

    const weekDates = getWeekClassDates(month, weekNumber, batch.scheduleType);

    const row = await FollowUp.create({
      month,
      weekNumber,
      batchId: batch._id,
      facultyId: batch.facultyId || req.user._id,
      studentId: null,
      attendanceDates: weekDates.map((d) => ({ date: d, status: '' })),
      contactType: 'None',
      dialedContact: '',
    });

    res.json({ row: { ...row.toObject(), student: null } });
  } catch (err) {
    res.status(500).json({ message: 'Could not add new row.' });
  }
});

// 3. Update a Follow-Up Row Inline (Auto-fills student info on student selection)
router.patch('/rows/:id', requireAuth, async (req, res) => {
  try {
    const row = await FollowUp.findById(req.params.id);
    if (!row) return res.status(404).json({ message: 'Row not found.' });

    const {
      studentId,
      attendanceDates,
      facultyRemarks,
      followUpDate,
      medium,
      facultyPhone,
      contactType,
      dialedContact,
      coordinatorRemarks,
      managerVerified,
    } = req.body;

    // Handle Student Selection -> Auto-populate dialedContact & keep contactType default 'None'
    if (studentId !== undefined) {
      if (!studentId) {
        row.studentId = null;
        row.dialedContact = '';
        row.contactType = 'None';
      } else {
        const student = await Student.findById(studentId);
        if (!student) return res.status(404).json({ message: 'Student not found.' });
        row.studentId = student._id;
        // Prioritize student phone by default while keeping WhatsApp selector at 'None'
        row.dialedContact = student.studentPhone || student.parentsPhone || '';
        if (!row.followUpDate) {
          row.followUpDate = new Date().toISOString().split('T')[0];
        }
      }
    }

    // Handle Contact Type Switch ('None' | 'Student' | 'Parents')
    if (contactType !== undefined) {
      row.contactType = contactType;
      if (row.studentId) {
        const stu = await Student.findById(row.studentId);
        if (stu) {
          if (contactType === 'Parents') {
            row.dialedContact = stu.parentsPhone || stu.studentPhone || '';
          } else if (contactType === 'Student' || contactType === 'None') {
            row.dialedContact = stu.studentPhone || stu.parentsPhone || '';
          }
        }
      }
    }

    if (attendanceDates !== undefined) row.attendanceDates = attendanceDates;
    if (facultyRemarks !== undefined) row.facultyRemarks = facultyRemarks;
    if (followUpDate !== undefined) row.followUpDate = followUpDate;
    if (medium !== undefined) row.medium = medium;
    if (facultyPhone !== undefined) row.facultyPhone = facultyPhone;
    if (dialedContact !== undefined) row.dialedContact = dialedContact;
    if (coordinatorRemarks !== undefined) row.coordinatorRemarks = coordinatorRemarks;

    if (typeof managerVerified === 'boolean' && req.user.role === 'admin') {
      row.managerVerified = managerVerified;
    }

    await row.save();

    const populated = await FollowUp.findById(row._id).populate('studentId');
    const obj = populated.toObject();
    if (obj.studentId) {
      obj.student = sanitizeStudentForRole(obj.studentId, req.user.role);
      obj.studentId = obj.student._id;
    } else {
      obj.student = null;
    }

    res.json({ row: obj });
  } catch (err) {
    if (err.code === 11000) {
      return res.status(400).json({ message: 'This student is already added in this week.' });
    }
    res.status(500).json({ message: 'Failed to update follow-up row.' });
  }
});

// 4. Delete a Row
router.delete('/rows/:id', requireAuth, async (req, res) => {
  try {
    await FollowUp.findByIdAndDelete(req.params.id);
    res.json({ message: 'Row removed.' });
  } catch (err) {
    res.status(500).json({ message: 'Failed to delete row.' });
  }
});

// 5. Upload WhatsApp Screenshot to Faculty's Google Drive
router.post('/rows/:id/screenshot', requireAuth, upload.single('screenshot'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose an image file.' });
    }

    const row = await FollowUp.findById(req.params.id).populate('studentId');
    if (!row) return res.status(404).json({ message: 'Follow-up row not found.' });

    const studentLabel = row.studentId?.studentName?.replace(/\s+/g, '_') || 'Student';
    const fileName = `FollowUp_${studentLabel}_${row.month}_W${row.weekNumber}_${Date.now()}.png`;

    const driveResult = await uploadScreenshotToFacultyDrive(
      req.user._id,
      req.file.buffer,
      req.file.mimetype,
      fileName
    );

    row.screenshotDriveFileId = driveResult.fileId;
    row.screenshotDriveUrl = driveResult.directImageUrl;
    await row.save();

    res.json({
      screenshotDriveFileId: driveResult.fileId,
      screenshotDriveUrl: driveResult.directImageUrl,
    });
  } catch (err) {
    console.error('Drive Screenshot Upload Error:', err);
    res.status(500).json({ message: err.message || 'Failed to upload screenshot to Google Drive.' });
  }
});

export default router;