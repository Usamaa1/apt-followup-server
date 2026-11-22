import express from 'express';
import Batch from '../models/Batch.js';
import Student from '../models/Student.js';
import User from '../models/User.js';
import { requireAuth, requireAdmin, sanitizeStudentForRole } from '../middleware/auth.js';
import FollowUp from '../models/FollowUp.js';

const router = express.Router();

// 1. Get Batches (Staff sees their own batches; Admin sees all or filtered by facultyId)
router.get('/batches', requireAuth, async (req, res) => {
  try {
    const query = {};
    if (req.user.role !== 'admin') {
      query.$or = [
        { facultyId: req.user._id },
        { facultyName: new RegExp(`^${req.user.displayName}$`, 'i') },
      ];
    } else if (req.query.facultyId) {
      query.facultyId = req.query.facultyId;
    }

    const batches = await Batch.find(query).sort({ batchCode: 1 });
    res.json({ batches });
  } catch (err) {
    res.status(500).json({ message: 'Failed to load batches.' });
  }
});

// 2. Create or Update a Batch (Admin only)
router.post('/batches', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { batchCode, scheduleType, timings, facultyId, facultyName, coordinatorName, courseTitle, status } = req.body;
    if (!batchCode || !scheduleType || !timings) {
      return res.status(400).json({ message: 'Batch Code, Schedule Type (MWF/TTS), and Timings are required.' });
    }

    let resolvedFacultyName = facultyName || 'Unassigned';
    if (facultyId) {
      const facUser = await User.findById(facultyId);
      if (facUser) resolvedFacultyName = facUser.displayName || facUser.name;
    }

    const batch = await Batch.findOneAndUpdate(
      { batchCode: batchCode.toUpperCase().trim() },
      {
        batchCode: batchCode.toUpperCase().trim(),
        scheduleType,
        timings: timings.trim(),
        facultyId: facultyId || null,
        facultyName: resolvedFacultyName,
        coordinatorName: coordinatorName?.trim() || '',
        courseTitle: courseTitle?.trim() || '',
        status: status || 'Ongoing',
      },
      { upsert: true, new: true }
    );

    res.json({ batch });
  } catch (err) {
    res.status(500).json({ message: 'Error saving batch.' });
  }
});
// 2B. Update Existing Batch Information (Timings, Schedule, Coordinator, Faculty, Status, Batch Code)
router.patch('/batches/:id', requireAuth, async (req, res) => {
  try {
    const batch = await Batch.findById(req.params.id);
    if (!batch) return res.status(404).json({ message: 'Batch not found.' });

    const {
      batchCode,
      scheduleType,
      timings,
      facultyId,
      facultyName,
      coordinatorName,
      courseTitle,
      status,
    } = req.body;

    if (batchCode !== undefined && batchCode.trim() !== '') {
      batch.batchCode = batchCode.toUpperCase().trim();
    }
    if (scheduleType !== undefined) {
      batch.scheduleType = scheduleType;
    }
    if (timings !== undefined && timings.trim() !== '') {
      batch.timings = timings.trim();
    }
    if (coordinatorName !== undefined) {
      batch.coordinatorName = coordinatorName.trim();
    }
    if (courseTitle !== undefined) {
      batch.courseTitle = courseTitle.trim();
    }
    if (status !== undefined) {
      batch.status = status;
    }

    // Only Admin can reassign batch ownership to a different faculty member
    if (req.user.role === 'admin') {
      if (facultyId !== undefined) {
        batch.facultyId = facultyId || null;
        if (facultyId) {
          const facUser = await User.findById(facultyId);
          if (facUser) batch.facultyName = facUser.displayName || facUser.name;
        }
      }
      if (facultyName !== undefined && facultyName.trim() !== '') {
        batch.facultyName = facultyName.trim();
      }
    }

    await batch.save();
    res.json({ message: `Batch ${batch.batchCode} updated.`, batch });
  } catch (err) {
    if (err.code === 11000) {
      return res.status(400).json({ message: 'Another batch already uses that Batch Code.' });
    }
    res.status(500).json({ message: 'Failed to update batch.' });
  }
});

// 2C. Delete Batch (Admin only; cleans up linked follow-up rows)
router.delete('/batches/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const batch = await Batch.findById(req.params.id);
    if (!batch) return res.status(404).json({ message: 'Batch not found.' });

    await FollowUp.deleteMany({ batchId: batch._id });
    await Student.updateMany({ currentBatchId: batch._id }, { $set: { currentBatchId: null } });
    await Batch.findByIdAndDelete(batch._id);

    res.json({ message: `Deleted batch ${batch.batchCode}.` });
  } catch (err) {
    res.status(500).json({ message: 'Failed to delete batch.' });
  }
});

// 3. Global Student Search by Name or Enrollment ID (Auto-fill helper for Add / Transfer)
router.get('/students/search', requireAuth, async (req, res) => {
  try {
    const q = (req.query.q || '').trim();
    if (!q) return res.json({ students: [] });

    const regex = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
    const students = await Student.find({
      $or: [{ studentName: regex }, { enrollmentNo: regex }, { studentPhone: regex }],
    })
      .populate('currentBatchId', 'batchCode timings facultyName')
      .limit(15);

    const sanitized = students.map((s) => sanitizeStudentForRole(s, req.user.role));
    res.json({ students: sanitized });
  } catch (err) {
    res.status(500).json({ message: 'Student search failed.' });
  }
});

// 4. Get Students in a Specific Batch
router.get('/batches/:batchId/students', requireAuth, async (req, res) => {
  try {
    const students = await Student.find({ currentBatchId: req.params.batchId }).sort({
      studentName: 1,
    });
    const sanitized = students.map((s) => sanitizeStudentForRole(s, req.user.role));
    res.json({ students: sanitized });
  } catch (err) {
    res.status(500).json({ message: 'Failed to load batch students.' });
  }
});

// 5. Add New Student OR Transfer Existing Student to Target Batch
router.post('/students/assign-or-transfer', requireAuth, async (req, res) => {
  try {
    const {
      existingStudentId,
      targetBatchId,
      enrollmentNo,
      studentName,
      studentPhone,
      parentsPhone,
      email,
      fatherName,
      address,
      courseCode,
      ovStatus,
    } = req.body;

    if (!targetBatchId) {
      return res.status(400).json({ message: 'Target batch is required.' });
    }

    const targetBatch = await Batch.findById(targetBatchId);
    if (!targetBatch) return res.status(404).json({ message: 'Target batch not found.' });

    const cleanEnrollment =
      enrollmentNo &&
      enrollmentNo.trim() !== '' &&
      enrollmentNo.trim().toLowerCase() !== 'no id provided'
        ? enrollmentNo.trim()
        : null;

    // Case A: Transferring an existing student to a new batch
    if (existingStudentId) {
      const student = await Student.findById(existingStudentId).populate('currentBatchId', 'batchCode');
      if (!student) return res.status(404).json({ message: 'Student not found.' });

      const oldBatchCode = student.currentBatchId?.batchCode || 'Unknown';
      if (String(student.currentBatchId?._id) !== String(targetBatch._id)) {
        student.batchHistory.push({
          fromBatchCode: oldBatchCode,
          toBatchCode: targetBatch.batchCode,
          transferredAt: new Date(),
          transferredBy: req.user._id,
        });
        student.currentBatchId = targetBatch._id;
      }

      if (cleanEnrollment !== null) student.enrollmentNo = cleanEnrollment;
      if (studentName) student.studentName = studentName.trim();
      if (studentPhone !== undefined) student.studentPhone = studentPhone.trim();
      if (parentsPhone !== undefined) student.parentsPhone = parentsPhone.trim();
      if (email !== undefined) student.email = email.trim().toLowerCase();

      if (req.user.role === 'admin') {
        if (fatherName !== undefined) student.fatherName = fatherName.trim();
        if (address !== undefined) student.address = address.trim();
        if (courseCode !== undefined) student.courseCode = courseCode.trim();
        if (ovStatus !== undefined) student.ovStatus = ovStatus.trim();
      }

      await student.save();
      return res.json({
        message: `Student transferred to ${targetBatch.batchCode}.`,
        student: sanitizeStudentForRole(student, req.user.role),
      });
    }

    // Case B: Creating a brand-new student in targetBatch
    if (!studentName || !studentName.trim()) {
      return res.status(400).json({ message: 'Student Name is required.' });
    }

    const newStudent = await Student.create({
      enrollmentNo: cleanEnrollment,
      studentName: studentName.trim(),
      studentPhone: studentPhone?.trim() || '',
      parentsPhone: parentsPhone?.trim() || '',
      email: email?.trim().toLowerCase() || '',
      currentBatchId: targetBatch._id,
      fatherName: req.user.role === 'admin' ? fatherName?.trim() || '' : '',
      address: req.user.role === 'admin' ? address?.trim() || '' : '',
      courseCode: req.user.role === 'admin' ? courseCode?.trim() || '' : '',
      ovStatus: req.user.role === 'admin' ? ovStatus?.trim() || '' : '',
    });

    res.json({
      message: `Student added to ${targetBatch.batchCode}.`,
      student: sanitizeStudentForRole(newStudent, req.user.role),
    });
  } catch (err) {
    res.status(500).json({ message: 'Failed to save or transfer student.' });
  }
});

// 6. Update Student Record
router.patch('/students/:id', requireAuth, async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ message: 'Student not found.' });

    const {
      enrollmentNo,
      studentName,
      studentPhone,
      parentsPhone,
      email,
      status,
      fatherName,
      address,
      courseCode,
      ovStatus,
    } = req.body;

    if (enrollmentNo !== undefined) {
      student.enrollmentNo =
        enrollmentNo.trim() === '' || enrollmentNo.trim().toLowerCase() === 'no id provided'
          ? null
          : enrollmentNo.trim();
    }
    if (studentName !== undefined) student.studentName = studentName.trim();
    if (studentPhone !== undefined) student.studentPhone = studentPhone.trim();
    if (parentsPhone !== undefined) student.parentsPhone = parentsPhone.trim();
    if (email !== undefined) student.email = email.trim().toLowerCase();
    if (status !== undefined) student.status = status;

    if (req.user.role === 'admin') {
      if (fatherName !== undefined) student.fatherName = fatherName.trim();
      if (address !== undefined) student.address = address.trim();
      if (courseCode !== undefined) student.courseCode = courseCode.trim();
      if (ovStatus !== undefined) student.ovStatus = ovStatus.trim();
    }

    await student.save();
    res.json({ student: sanitizeStudentForRole(student, req.user.role) });
  } catch (err) {
    res.status(500).json({ message: 'Failed to update student.' });
  }
});
// 6B. Delete Student Record (and clean up any linked follow-up rows)
router.delete('/students/:id', requireAuth, async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ message: 'Student not found.' });

    await FollowUp.deleteMany({ studentId: student._id });
    await Student.findByIdAndDelete(student._id);

    res.json({ message: `Deleted ${student.studentName} successfully.` });
  } catch (err) {
    res.status(500).json({ message: 'Failed to delete student.' });
  }
});

// 7. Admin-Only Full Directory (includes Father Name, Address, OV Status, Batch History)
router.get('/admin/students', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { batchId, q } = req.query;
    const filter = {};
    if (batchId) filter.currentBatchId = batchId;
    if (q) {
      const regex = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
      filter.$or = [{ studentName: regex }, { enrollmentNo: regex }, { studentPhone: regex }];
    }

    const students = await Student.find(filter)
      .populate('currentBatchId', 'batchCode timings facultyName coordinatorName')
      .sort({ studentName: 1 })
      .limit(300);

    res.json({
      students: students.map((s) => sanitizeStudentForRole(s, 'admin')),
    });
  } catch (err) {
    res.status(500).json({ message: 'Failed to load admin student directory.' });
  }
});

export default router;