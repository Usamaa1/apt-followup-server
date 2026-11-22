import dotenv from 'dotenv';
import { connectDB } from '../config/db.js';
import LookupData from '../models/LookupData.js';

dotenv.config();

const seed = async () => {
  await connectDB();
  await LookupData.findOneAndUpdate(
    { key: 'GLOBAL_LOOKUPS' },
    { key: 'GLOBAL_LOOKUPS' },
    { upsert: true, new: true }
  );
  console.log('Default LookupData (Mediums, Faculty Phones 01-06, Coordinators) seeded.');
  process.exit(0);
};

seed();