import { initializeApp, getApp, getApps } from 'firebase/app';
import { getAuth, signInWithPopup, signInWithRedirect, getRedirectResult, GoogleAuthProvider, onAuthStateChanged, User } from 'firebase/auth';
import defaultFirebaseConfig from '../../firebase-applet-config.json';

// Bulletproof Storage Utility to bypass SecurityError issues in sandboxed frames / incognito
const safeLocalStorage = {
  getItem: (key: string): string | null => {
    try {
      return typeof window !== 'undefined' && window.localStorage ? window.localStorage.getItem(key) : null;
    } catch (e) {
      return null;
    }
  },
  setItem: (key: string, value: string): void => {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(key, value);
      }
    } catch (e) {}
  },
  removeItem: (key: string): void => {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.removeItem(key);
      }
    } catch (e) {}
  }
};

// Get active Firebase configuration dynamically (custom from localStorage, or default)
const getActiveConfig = () => {
  const customStr = safeLocalStorage.getItem('kanooz_custom_firebase_config');
  if (customStr) {
    try {
      const parsed = JSON.parse(customStr);
      if (parsed.apiKey && parsed.authDomain && parsed.projectId) {
        return parsed;
      }
    } catch (e) {
      console.error('Failed to parse custom Firebase config:', e);
    }
  }
  return defaultFirebaseConfig;
};

const activeConfig = getActiveConfig();

let app: any = null;
let resolvedAuth: any = null;

try {
  // Initialize the app cleanly
  app = getApps().length === 0 ? initializeApp(activeConfig) : getApp();
  resolvedAuth = getAuth(app);
} catch (e) {
  console.error('Firebase initialization error gracefully handled to avoid page crash:', e);
  resolvedAuth = {
    app: {
      options: activeConfig
    },
    signOut: async () => {},
    onAuthStateChanged: (cb: any) => {
      return () => {};
    }
  };
}

export const auth = resolvedAuth;

const provider = new GoogleAuthProvider();
provider.addScope('https://www.googleapis.com/auth/drive');

let isSigningIn = false;
let cachedAccessToken: string | null = safeLocalStorage.getItem('kanooz_google_drive_token');

// Initialize auth state listener
export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  try {
    // If we have a mock auth object, use its custom listener
    if (auth && typeof auth.onAuthStateChanged === 'function' && !('currentUser' in auth)) {
      if (onAuthFailure) onAuthFailure();
      return auth.onAuthStateChanged(() => {});
    }
    return onAuthStateChanged(auth, async (user: User | null) => {
      if (user) {
        const storedToken = safeLocalStorage.getItem('kanooz_google_drive_token');
        if (storedToken) {
          cachedAccessToken = storedToken;
          if (onAuthSuccess) onAuthSuccess(user, storedToken);
        } else if (cachedAccessToken) {
          if (onAuthSuccess) onAuthSuccess(user, cachedAccessToken);
        } else if (!isSigningIn) {
          cachedAccessToken = null;
          if (onAuthFailure) onAuthFailure();
        }
      } else {
        cachedAccessToken = null;
        safeLocalStorage.removeItem('kanooz_google_drive_token');
        if (onAuthFailure) onAuthFailure();
      }
    });
  } catch (e) {
    console.error('onAuthStateChanged error handled safely:', e);
    if (onAuthFailure) onAuthFailure();
    return () => {};
  }
};

// Initiate Google authentication via popup web flow
export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Failed to obtain Google Drive permissions from Firebase Auth.');
    }

    cachedAccessToken = credential.accessToken;
    safeLocalStorage.setItem('kanooz_google_drive_token', cachedAccessToken);
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Sign in error:', error);
    throw error;
  } finally {
    isSigningIn = false;
  }
};

// Initiate Google authentication via page redirect
export const googleSignInRedirect = async (): Promise<void> => {
  isSigningIn = true;
  await signInWithRedirect(auth, provider);
};

// Retrieve results/credential from previous redirect login
export const checkRedirectResult = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    const result = await getRedirectResult(auth);
    if (result) {
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (credential?.accessToken) {
        cachedAccessToken = credential.accessToken;
        safeLocalStorage.setItem('kanooz_google_drive_token', cachedAccessToken);
        return { user: result.user, accessToken: cachedAccessToken };
      }
    }
    return null;
  } catch (error: any) {
    console.error('Redirect result retrieve error:', error);
    throw error;
  }
};

export const getAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken;
};

export const logout = async () => {
  await auth.signOut();
  cachedAccessToken = null;
  safeLocalStorage.removeItem('kanooz_google_drive_token');
};

// File constant name
const BACKUP_FILENAME = 'Kanooz_Master_Planning_Backup.json';

/**
 * Enhanced Google Drive response error parser and handler.
 * Provides a user-friendly error description if the Google Drive API is disabled on a custom project.
 */
async function handleResponseError(res: Response, fallbackPrefix: string): Promise<never> {
  const errText = await res.text();
  let parsedMsg = '';
  try {
    const json = JSON.parse(errText);
    if (json?.error?.message) {
      parsedMsg = json.error.message;
    }
  } catch (e) {}

  const finalMsg = parsedMsg || errText;

  if (
    finalMsg.includes('Google Drive API has not been used') || 
    finalMsg.includes('drive.googleapis.com/overview') ||
    errText.includes('Google Drive API has not been used') ||
    errText.includes('drive.googleapis.com/overview')
  ) {
    // Attempt to extract the project number
    const projMatch = finalMsg.match(/project=(\d+)/) || errText.match(/project=(\d+)/) || finalMsg.match(/project_number:(\d+)/);
    const projectIdOrNum = projMatch ? projMatch[1] : '';
    const enableUrl = projectIdOrNum 
      ? `https://console.cloud.google.com/apis/library/drive.googleapis.com?project=${projectIdOrNum}`
      : 'https://console.cloud.google.com/apis/library/drive.googleapis.com';
    
    throw new Error(
      `🔒 Google Drive API is disabled or not activated in your private Firebase/Google Cloud project!\n\n` +
      `To resolve this, please open the following official link in a new tab to enable it (ensure you are logged in using the owner account):\n\n` +
      `👉 ${enableUrl}\n\n` +
      `Simply click the "Enable" button, wait 1-2 minutes for Google's systems to sync, then try uploading or restoring files again!`
    );
  }

  throw new Error(`${fallbackPrefix}: ${finalMsg}`);
}

/**
 * Searches for the specific backup file in the user's Google Drive.
 */
async function findBackupFile(token: string): Promise<string | null> {
  const query = encodeURIComponent(`name = '${BACKUP_FILENAME}' and trashed = false`);
  const response = await fetch(`https://www.googleapis.com/drive/v3/files?q=${query}&spaces=drive`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });
  
  if (!response.ok) {
    await handleResponseError(response, 'Google Drive Search failed');
  }
  
  const data = await response.json();
  if (data.files && data.files.length > 0) {
    return data.files[0].id;
  }
  return null;
}

/**
 * Saves database state directly into Google Drive (overwriting if exists, creating if new).
 */
export const saveToDrive = async (payload: any): Promise<{ id: string; name: string; updated: boolean }> => {
  const token = cachedAccessToken;
  if (!token) {
    throw new Error('User is not signed in with Google Drive permissions.');
  }

  // Find if backup file already exists
  const fileId = await findBackupFile(token);
  const jsonString = JSON.stringify(payload, null, 2);

  if (fileId) {
    // Overwrite the file content
    const res = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: jsonString
    });

    if (!res.ok) {
      await handleResponseError(res, 'Google Drive write/overwrite failed');
    }

    return { id: fileId, name: BACKUP_FILENAME, updated: true };
  } else {
    // Create new metadata file in Drive
    const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        name: BACKUP_FILENAME,
        mimeType: 'application/json'
      })
    });

    if (!createRes.ok) {
      await handleResponseError(createRes, 'Google Drive metadata initialization failed');
    }

    const fileMeta = await createRes.json();
    const newFileId = fileMeta.id;

    // Upload the file content to newly created ID
    const uploadRes = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${newFileId}?uploadType=media`, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: jsonString
    });

    if (!uploadRes.ok) {
      await handleResponseError(uploadRes, 'Google Drive content populate failed');
    }

    return { id: newFileId, name: BACKUP_FILENAME, updated: false };
  }
};

/**
 * Loads database state from Google Drive if it exists.
 */
export const loadFromDrive = async (): Promise<any | null> => {
  const token = cachedAccessToken;
  if (!token) {
    throw new Error('User is not signed in with Google Drive permissions.');
  }

  const fileId = await findBackupFile(token);
  if (!fileId) {
    return null; // File doesn't exist yet
  }

  const res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!res.ok) {
    await handleResponseError(res, 'Google Drive file retrieve failed');
  }

  return await res.json();
};

/**
 * Lists the user's files from Google Drive that are spreadsheets (.xlsx/.xls) or files
 */
export const listDriveFiles = async (): Promise<Array<{ id: string; name: string; mimeType: string; modifiedTime?: string }>> => {
  const token = cachedAccessToken;
  if (!token) {
    throw new Error('User is not signed in with Google Drive permissions.');
  }

  // We query files that are spreadsheets or JSON files and not trashed
  const query = encodeURIComponent("trashed = false and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel' or mimeType = 'application/json' or name contains '.xlsx' or name contains '.xls' or name contains '.json')");
  const response = await fetch(`https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name,mimeType,modifiedTime)&spaces=drive&pageSize=50`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!response.ok) {
    await handleResponseError(response, 'Google Drive list failed');
  }

  const data = await response.json();
  return data.files || [];
};

/**
 * Downloads a binary file from Google Drive as ArrayBuffer
 */
export const downloadFileFromDrive = async (fileId: string): Promise<ArrayBuffer> => {
  const token = cachedAccessToken;
  if (!token) {
    throw new Error('User is not signed in with Google Drive permissions.');
  }

  const res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!res.ok) {
    await handleResponseError(res, 'Google Drive file retrieve failed');
  }

  return await res.arrayBuffer();
};

/**
 * Uploads any custom file (like a generated Excel file) to the user's Google Drive.
 */
export const uploadCustomFileToDrive = async (
  fileName: string,
  content: ArrayBuffer | Blob | string,
  mimeType: string
): Promise<{ id: string; name: string }> => {
  const token = cachedAccessToken;
  if (!token) {
    throw new Error('User is not signed in with Google Drive permissions.');
  }

  // Create the metadata file record
  const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name: fileName,
      mimeType: mimeType
    })
  });

  if (!createRes.ok) {
    await handleResponseError(createRes, 'Google Drive file metadata creation failed');
  }

  const fileMeta = await createRes.json();
  const fileId = fileMeta.id;

  // Upload the media content
  const uploadRes = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': mimeType
    },
    body: content
  });

  if (!uploadRes.ok) {
    await handleResponseError(uploadRes, 'Google Drive content upload failed');
  }

  return { id: fileId, name: fileName };
};

