import express from 'express';
import dotenv from 'dotenv';
import cors from 'cors';
import cookieParser from 'cookie-parser';
import { connectDB } from './config/db.js';
import authRoutes from './routes/authRoutes.js';
import adminRoutes from './routes/adminRoutes.js';
import batchStudentRoutes from './routes/batchStudentRoutes.js';
import excelRoutes from './routes/excelRoutes.js';
import followUpRoutes from './routes/followUpRoutes.js';
import syncRoutes from './routes/syncRoutes.js';

dotenv.config();

const app = express();

app.use(
  cors({
    origin: ['http://localhost:5173', process.env.CLIENT_URL].filter(Boolean),
    credentials: true,
  })
);
app.use(express.json({ limit: '4.5mb' }));
app.use(cookieParser());

// Ensure MongoDB connection is ready on every serverless invocation
app.use(async (req, res, next) => {
  try {
    await connectDB();
    next();
  } catch (err) {
    console.error('MongoDB connection error:', err);
    res.status(500).json({ message: 'Database connection failed.' });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Follow-Up Portal API running' });
});

app.use('/api/auth', authRoutes);
app.use('/api/admin', adminRoutes);
app.use('/api/portal', batchStudentRoutes);
app.use('/api/excel', excelRoutes);
app.use('/api/followup', followUpRoutes);
app.use('/api/sync', syncRoutes);

if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
}

export default app;