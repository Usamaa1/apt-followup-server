import mongoose from 'mongoose';

const attendanceDaySchema = new mongoose.Schema(
  {
    date: { type: String, required: true }, // YYYY-MM-DD
    status: { type: String, enum: ['A', 'P','LV', ''], default: '' },
  },
  { _id: false }
);

const followUpSchema = new mongoose.Schema(
  {
    month: { type: String, required: true, index: true }, // e.g., "2026-11"
    weekNumber: { type: Number, required: true, min: 1, max: 4, index: true },
    batchId: { type: mongoose.Schema.Types.ObjectId, ref: 'Batch', required: true, index: true },
    facultyId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },

    // Null for the default starter empty row under each batch until a student is chosen
    studentId: { type: mongoose.Schema.Types.ObjectId, ref: 'Student', default: null },

    // 3 class dates for the week (Mon/Wed/Fri or Tue/Thu/Sat)
    attendanceDates: {
      type: [attendanceDaySchema],
      default: [],
    },

    facultyRemarks: { type: String, trim: true, default: '' },
    followUpDate: { type: String, default: '' }, // YYYY-MM-DD
    medium: { type: String, trim: true, default: '' },
    facultyPhone: { type: String, trim: true, default: '' },

    // Contact & WhatsApp selection ("None" selected by default)
    contactType: {
      type: String,
      enum: ['None', 'Student', 'Parents'],
      default: 'None',
    },
    dialedContact: { type: String, trim: true, default: '' },

    coordinatorRemarks: { type: String, trim: true, default: '' },

    // Google Drive screenshot reference (stored in faculty's Drive, NOT MongoDB)
    screenshotDriveFileId: { type: String, default: '' },
    screenshotDriveUrl: { type: String, default: '' },

    // Admin/Manager verification status
    managerVerified: { type: Boolean, default: false },
  },
  { timestamps: true }
);

// Prevent duplicate follow-up entries for the same student in the same batch/month/week
followUpSchema.index(
  { month: 1, weekNumber: 1, batchId: 1, studentId: 1 },
  { unique: true, partialFilterExpression: { studentId: { $type: 'objectId' } } }
);

export default mongoose.model('FollowUp', followUpSchema);