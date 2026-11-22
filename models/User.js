import mongoose from 'mongoose';

const userSchema = new mongoose.Schema(
  {
    googleId: { type: String, unique: true, sparse: true },
    email: { type: String, required: true, unique: true, lowercase: true, trim: true },
    name: { type: String, trim: true },
    // Exact name matching the Google Sheet / Excel tab 
    displayName: { type: String, trim: true },
    avatar: { type: String },
    role: {
      type: String,
      enum: ['admin', 'staff'],
      default: 'staff',
    },
    // Hybrid mode: true if pre-added by Admin OR auto-provisioned on valid Google login
    isWhitelisted: { type: Boolean, default: true },
    isActive: { type: Boolean, default: true },

    // Per-faculty Google Sheet link for manual "Sync to Google Sheet" action
    googleSheetUrl: { type: String, trim: true, default: '' },
    googleSheetId: { type: String, trim: true, default: '' },
    lastSyncedAt: { type: Date, default: null },

    // Customizable WhatsApp template with auto-absent-dates default
 whatsappTemplate: {
      type: String,
      default: 'Why are you absent on {{absentDates}}? Please tell the reason',
    },
    emailSubjectTemplate: {
      type: String,
      default: 'Attendance Follow-Up ({{batchCode}}) - {{studentName}}',
    },
    emailBodyTemplate: {
      type: String,
      default:
        'Dear {{studentName}},\n\nYou were marked absent in batch {{batchCode}} on {{absentDates}}. Please share the reason for your absence at your earliest convenience.\n\nRegards,\n{{facultyName}}',
    },

    // OAuth tokens for uploading screenshots to faculty's own Google Drive & syncing Sheets
    googleAccessToken: { type: String, select: false },
    googleRefreshToken: { type: String, select: false },
    tokenExpiryDate: { type: Number, select: false },
  },
  { timestamps: true }
);

export default mongoose.model('User', userSchema);