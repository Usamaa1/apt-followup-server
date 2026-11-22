import mongoose from 'mongoose';

const batchSchema = new mongoose.Schema(
  {
    batchCode: { type: String, required: true, unique: true, uppercase: true, trim: true },
    aptrackBatchId: { type: String, trim: true, default: '' },
    scheduleType: {
      type: String,
      enum: ['MWF', 'TTS'],
      required: true,
    },
    timings: { type: String, required: true, trim: true }, // e.g., "MWF (9 - 11)"
    facultyId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
    facultyName: { type: String, required: true, trim: true }, // e.g., "Miss Urooj Rehman"
    coordinatorName: { type: String, trim: true, default: '' },
    courseTitle: { type: String, trim: true, default: '' },
    startDate: { type: Date, default: null },
    status: {
      type: String,
      enum: ['Ongoing', 'Ended', 'Dissolved'],
      default: 'Ongoing',
    },
  },
  { timestamps: true }
);

export default mongoose.model('Batch', batchSchema);