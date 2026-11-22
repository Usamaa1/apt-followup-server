import jwt from 'jsonwebtoken';
import User from '../models/User.js';

export const requireAuth = async (req, res, next) => {
  try {
    const token =
      req.cookies?.token ||
      (req.headers.authorization?.startsWith('Bearer ')
        ? req.headers.authorization.split(' ')[1]
        : null);

    if (!token) {
      return res.status(401).json({ message: 'Not authenticated. Please sign in with Google.' });
    }

    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    const user = await User.findById(decoded.userId);

    if (!user || !user.isActive) {
      return res.status(403).json({ message: 'Account disabled or unauthorized.' });
    }

    req.user = user;
    next();
  } catch (err) {
    return res.status(401).json({ message: 'Session expired. Please log in again.' });
  }
};

export const requireAdmin = (req, res, next) => {
  if (req.user?.role !== 'admin') {
    return res.status(403).json({ message: 'Admin/Manager privileges required.' });
  }
  next();
};

// Strips admin-only student fields when requested by non-admin staff
export const sanitizeStudentForRole = (studentDoc, role) => {
  const obj = studentDoc.toObject ? studentDoc.toObject({ virtuals: true }) : { ...studentDoc };
  obj.displayEnrollmentNo =
    obj.enrollmentNo && obj.enrollmentNo.trim() !== '' ? obj.enrollmentNo : 'No ID provided';

  if (role !== 'admin') {
    delete obj.fatherName;
    delete obj.address;
    delete obj.courseCode;
    delete obj.ovStatus;
    delete obj.batchHistory;
  }
  return obj;
};