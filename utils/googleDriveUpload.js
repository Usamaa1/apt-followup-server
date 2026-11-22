import { Readable } from 'stream';
import { google } from 'googleapis';
import User from '../models/User.js';

export const getUserOAuthClient = async (userId) => {
  const user = await User.findById(userId).select(
    '+googleAccessToken +googleRefreshToken +tokenExpiryDate'
  );
  if (!user || (!user.googleAccessToken && !user.googleRefreshToken)) {
    throw new Error('Google Drive authorization missing. Please sign out and log in with Google again.');
  }

  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );

  oauth2Client.setCredentials({
    access_token: user.googleAccessToken,
    refresh_token: user.googleRefreshToken,
    expiry_date: user.tokenExpiryDate,
  });

  // Persist refreshed tokens automatically
  oauth2Client.on('tokens', async (tokens) => {
    const updates = {};
    if (tokens.access_token) updates.googleAccessToken = tokens.access_token;
    if (tokens.refresh_token) updates.googleRefreshToken = tokens.refresh_token;
    if (tokens.expiry_date) updates.tokenExpiryDate = tokens.expiry_date;
    if (Object.keys(updates).length > 0) {
      await User.findByIdAndUpdate(userId, updates);
    }
  });

  return { oauth2Client, user };
};

export const uploadScreenshotToFacultyDrive = async (userId, fileBuffer, mimeType, fileName) => {
  const { oauth2Client } = await getUserOAuthClient(userId);
  const drive = google.drive({ version: 'v3', auth: oauth2Client });

  // Find or create "FollowUp_Screenshots" folder in faculty's Google Drive
  const folderName = 'FollowUp_Screenshots';
  const folderSearch = await drive.files.list({
    q: `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
  });

  let folderId = folderSearch.data.files?.[0]?.id;
  if (!folderId) {
    const createdFolder = await drive.files.create({
      requestBody: {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
      },
      fields: 'id',
    });
    folderId = createdFolder.data.id;
  }

  const stream = Readable.from(fileBuffer);
  const uploaded = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [folderId],
    },
    media: {
      mimeType,
      body: stream,
    },
    fields: 'id, webViewLink, webContentLink',
  });

  const fileId = uploaded.data.id;

  // Make file viewable via link so Google Sheets =IMAGE() and ExcelJS export can render it
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: 'reader',
      type: 'anyone',
    },
  });

  const directImageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

  return {
    fileId,
    directImageUrl,
    webViewLink: uploaded.data.webViewLink,
  };
};