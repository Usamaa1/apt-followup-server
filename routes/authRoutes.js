import express from 'express';
import { google } from 'googleapis';
import jwt from 'jsonwebtoken';
import User from '../models/User.js';
import { requireAuth } from '../middleware/auth.js';

const router = express.Router();

const createOAuthClient = () =>
  new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );

const SCOPES = [
  'https://www.googleapis.com/auth/userinfo.profile',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/spreadsheets',
];

// 1. Initiate Google OAuth (supports silent prompt='none' for auto-login if already signed into Gmail)
router.get('/google', (req, res) => {
  const oauth2Client = createOAuthClient();
  const silent = req.query.silent === 'true';

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: silent ? 'none' : 'select_account consent',
  });

  res.redirect(authUrl);
});

// 2. Google OAuth Callback
router.get('/google/callback', async (req, res) => {
  const { code, error } = req.query;

  // If silent login fails because user isn't signed into Google or hasn't consented yet, redirect cleanly to /login
  if (error || !code) {
    return res.redirect(`${process.env.CLIENT_URL}/login?authError=${error || 'access_denied'}`);
  }

  try {
    const oauth2Client = createOAuthClient();
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
    const { data: profile } = await oauth2.userinfo.get();

    const email = profile.email.toLowerCase();
    const isInitialAdmin =
      process.env.INITIAL_ADMIN_EMAIL &&
      email === process.env.INITIAL_ADMIN_EMAIL.toLowerCase();

    // Hybrid Login: check if Admin pre-whitelisted this email, otherwise auto-provision as 'staff'
    let user = await User.findOne({ email }).select('+googleRefreshToken');

    if (!user) {
      user = await User.create({
        googleId: profile.id,
        email,
        name: profile.name,
        displayName: isInitialAdmin
          ? process.env.INITIAL_ADMIN_DISPLAY_NAME || profile.name
          : profile.name,
        avatar: profile.picture,
        role: isInitialAdmin ? 'admin' : 'staff',
        isWhitelisted: true,
        googleAccessToken: tokens.access_token,
        googleRefreshToken: tokens.refresh_token || '',
        tokenExpiryDate: tokens.expiry_date || null,
      });
    } else {
      // Link Google profile to pre-whitelisted record & refresh tokens
      user.googleId = profile.id;
      user.name = profile.name;
      if (!user.displayName) user.displayName = profile.name;
      user.avatar = profile.picture;
      if (isInitialAdmin) user.role = 'admin';
      if (tokens.access_token) user.googleAccessToken = tokens.access_token;
      if (tokens.refresh_token) user.googleRefreshToken = tokens.refresh_token;
      if (tokens.expiry_date) user.tokenExpiryDate = tokens.expiry_date;
      await user.save();
    }

    if (!user.isActive) {
      return res.redirect(`${process.env.CLIENT_URL}/login?authError=account_disabled`);
    }

    const sessionToken = jwt.sign({ userId: user._id }, process.env.JWT_SECRET, {
      expiresIn: '7d',
    });

    res.cookie('token', sessionToken, {
      httpOnly: true,
      sameSite: 'lax',
      secure: process.env.NODE_ENV === 'production',
      maxAge: 7 * 24 * 60 * 60 * 1000,
    });

    res.redirect(`${process.env.CLIENT_URL}/followup`);
  } catch (err) {
    console.error('Google OAuth Callback Error:', err);
    res.redirect(`${process.env.CLIENT_URL}/login?authError=oauth_failed`);
  }
});

// 3. Get Current Logged-in User
router.get('/me', requireAuth, async (req, res) => {
  res.json({ user: req.user });
});

// 4. Update Personal Settings (Manager's Google Sheet URL & WhatsApp Message Template)
router.patch('/settings', requireAuth, async (req, res) => {
  try {
    const {
      googleSheetUrl,
      whatsappTemplate,
      emailSubjectTemplate,
      emailBodyTemplate,
      displayName,
    } = req.body;

    if (displayName !== undefined && displayName.trim() !== '') {
      req.user.displayName = displayName.trim();
    }

    if (googleSheetUrl !== undefined) {
      const trimmed = googleSheetUrl.trim();
      req.user.googleSheetUrl = trimmed;
      const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      req.user.googleSheetId = match ? match[1] : '';
    }

    if (whatsappTemplate !== undefined) {
      req.user.whatsappTemplate = whatsappTemplate;
    }

    if (emailSubjectTemplate !== undefined) {
      req.user.emailSubjectTemplate = emailSubjectTemplate;
    }

    if (emailBodyTemplate !== undefined) {
      req.user.emailBodyTemplate = emailBodyTemplate;
    }

    await req.user.save();
    res.json({
      message: 'Settings saved successfully.',
      user: req.user,
    });
  } catch (err) {
    res.status(500).json({ message: 'Failed to save settings.' });
  }
});

// 5. Logout
router.post('/logout', (req, res) => {
  res.clearCookie('token');
  res.json({ message: 'Logged out successfully.' });
});

export default router;