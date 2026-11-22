import mongoose from 'mongoose';

const lookupDataSchema = new mongoose.Schema(
  {
    key: { type: String, default: 'GLOBAL_LOOKUPS', unique: true },
    mediums: {
      type: [String],
      default: [
        'WhatsApp',
        'In Class',
        'Already Informed',
        'Faculty Phone',
        'Faculty Phone WhatsApp',
        'Email',
        'Through Coordinator',
      ],
    },
    facultyPhones: {
      type: [String],
      default: [
        'Faculty Phone 01 - 0321-9212669',
        'Faculty Phone 02 - 0315-3000668',
        'Faculty Phone 03 - 0315-3000679',
        'Faculty Phone 04 - 0321-9212681',
        'Faculty Phone 05 - 0320-3000616',
        'Faculty Phone 06 - 0320-3000611',
        'Faculty Smart Phone - 0321-9212678',
      ],
    },
    coordinators: {
      type: [String],
      default: [
        'Miss Warda Siddiqui',
        'Miss Maha Haider',
        'Miss Sana Yousuf',
        'Miss Munazzah Khan',
        'Miss Maham Haider',
        'Miss Tehreem Fatima',
      ],
    },
  },
  { timestamps: true }
);

export default mongoose.model('LookupData', lookupDataSchema);