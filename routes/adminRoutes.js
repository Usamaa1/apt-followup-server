import express from 'express';
import User from '../models/User.js';
import LookupData from '../models/LookupData.js';
import { requireAuth, requireAdmin } from '../middleware/auth.js';

const router = express.Router();

// Get all users (Admin only)
router.get('/users', requireAuth, requireAdmin, async (req, res) => {
  const users = await User.find().sort({ role: 1, displayName: 1 });
  res.json({ users });
});

// Pre-register / whitelist a faculty or coordinator Gmail with exact sheet tab name
router.post('/users/whitelist', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { email, displayName, role = 'staff', googleSheetUrl = '' } = req.body;
    if (!email || !displayName) {
      return res.status(400).json({ message: 'Email and Sheet Tab Display Name are required.' });
    }

    const match = googleSheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const googleSheetId = match ? match[1] : '';

    const user = await User.findOneAndUpdate(
      { email: email.toLowerCase().trim() },
      {
        email: email.toLowerCase().trim(),
        displayName: displayName.trim(),
        role,
        isWhitelisted: true,
        isActive: true,
        ...(googleSheetUrl && { googleSheetUrl: googleSheetUrl.trim(), googleSheetId }),
      },
      { upsert: true, new: true }
    );

    res.json({ user });
  } catch (err) {
    res.status(500).json({ message: 'Error saving whitelisted user.' });
  }
});

// Update user role / active status / display name
router.patch('/users/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { role, displayName, isActive } = req.body;
    const user = await User.findByIdAndUpdate(
      req.params.id,
      { ...(role && { role }), ...(displayName && { displayName }), ...(typeof isActive === 'boolean' && { isActive }) },
      { new: true }
    );
    res.json({ user });
  } catch (err) {
    res.status(500).json({ message: 'Error updating user.' });
  }
});

// Get Lookup Data (Mediums, Faculty Phones, Coordinators)
router.get('/lookups', requireAuth, async (req, res) => {
  let lookups = await LookupData.findOne({ key: 'GLOBAL_LOOKUPS' });
  if (!lookups) {
    lookups = await LookupData.create({ key: 'GLOBAL_LOOKUPS' });
  }
  res.json({ lookups });
});


// Update Lookup Data (Mediums, Faculty Phones 01-06, Coordinators)
router.put('/lookups', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { mediums, facultyPhones, coordinators } = req.body;
    const lookups = await LookupData.findOneAndUpdate(
      { key: 'GLOBAL_LOOKUPS' },
      {
        ...(Array.isArray(mediums) && { mediums }),
        ...(Array.isArray(facultyPhones) && { facultyPhones }),
        ...(Array.isArray(coordinators) && { coordinators }),
      },
      { upsert: true, new: true }
    );
    res.json({ lookups });
  } catch (err) {
    res.status(500).json({ message: 'Failed to update lookup lists.' });
  }
});

// GET /api/admin/settings
router.get('/settings', requireAuth, async (req, res) => {
  try {
    const user = await User.findById(req.user._id).select('-password');
    res.json({ user });
  } catch (err) {
    res.status(500).json({ message: 'Failed to load settings.' });
  }
});

// PATCH /api/admin/settings
router.patch('/settings', requireAuth, async (req, res) => {
  try {
    const targetId =
      req.user.role === 'admin' && req.body.facultyId ? req.body.facultyId : req.user._id;

    const user = await User.findById(targetId);
    if (!user) {
      return res.status(404).json({ message: 'User not found.' });
    }

    if (req.body.googleSheetUrl !== undefined) {
      user.googleSheetUrl = String(req.body.googleSheetUrl).trim();
    }
    if (req.body.googleSheetId !== undefined) {
      user.googleSheetId = String(req.body.googleSheetId).trim();
    }
    if (req.body.displayName !== undefined) {
      user.displayName = String(req.body.displayName).trim();
    }
    if (req.body.whatsappTemplate !== undefined) {
      user.whatsappTemplate = req.body.whatsappTemplate;
    }

    await user.save();

    res.json({
      message: 'Settings saved successfully.',
      user: {
        _id: user._id,
        name: user.name,
        displayName: user.displayName,
        email: user.email,
        role: user.role,
        googleSheetUrl: user.googleSheetUrl,
        googleSheetId: user.googleSheetId,
        whatsappTemplate: user.whatsappTemplate,
      },
    });
  } catch (err) {
    console.error('PATCH /api/admin/settings Error:', err);
    res.status(500).json({ message: err.message || 'Failed to save settings.' });
  }
});

export default router;