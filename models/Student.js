import mongoose from 'mongoose';

const batchHistorySchema = new mongoose.Schema(
  {
    fromBatchCode: String,
    toBatchCode: String,
    transferredAt: { type: Date, default: Date.now },
    transferredBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

const studentSchema = new mongoose.Schema(
  {
    // Stored as null when missing/not enrolled; formatted as "No ID provided" in UI & exports
    enrollmentNo: { type: String, trim: true, default: null, index: true },
    studentName: { type: String, required: true, trim: true, index: true },
    studentPhone: { type: String, trim: true, default: '' },
    parentsPhone: { type: String, trim: true, default: '' },
    email: { type: String, trim: true, lowercase: true, default: '' },
    currentBatchId: { type: mongoose.Schema.Types.ObjectId, ref: 'Batch', required: true, index: true },
    status: {
      type: String,
      enum: ['Active', 'Freeze', 'Dropout', 'Completed'],
      default: 'Active',
    },

    // Admin-only restricted fields (stripped from faculty API responses)
    fatherName: { type: String, trim: true, default: '' },
    address: { type: String, trim: true, default: '' },
    courseCode: { type: String, trim: true, default: '' },
    ovStatus: { type: String, trim: true, default: '' },
    batchHistory: [batchHistorySchema],
  },
  {
    timestamps: true,
    toJSON: { virtuals: true },
    toObject: { virtuals: true },
  }
);

// Virtual helper so frontend/exports always receive "No ID provided" when enrollmentNo is empty
studentSchema.virtual('displayEnrollmentNo').get(function () {
  return this.enrollmentNo && this.enrollmentNo.trim() !== ''
    ? this.enrollmentNo
    : 'No ID provided';
});

export default mongoose.model('Student', studentSchema);