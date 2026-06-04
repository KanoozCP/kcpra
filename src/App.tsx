import React, { useState, useEffect, ChangeEvent } from 'react';
import { 
  BarChart3, 
  Users, 
  Briefcase, 
  Calendar, 
  GanttChart as GanttIcon, 
  AlertTriangle, 
  UserMinus,
  FileText,
  Plus,
  RefreshCw,
  LogOut,
  ChevronRight,
  ShieldCheck,
  Search,
  Filter,
  Download,
  Trash2,
  CheckCircle2,
  Menu,
  RotateCcw,
  Upload,
  Printer,
  Moon,
  Sun,
  Sliders,
  Key
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';

import { Manpower, Project, Assignment, Tab } from './types';
import { runAutoAssignment, calculateShortages } from './lib/assignment-engine';
import { cn } from './lib/utils';
import { formatToExcelDate, parseExcelDate } from './lib/dateUtils';

// Google Drive Sync Service
import { 
  initAuth as initGoogleAuth, 
  googleSignIn, 
  googleSignInRedirect,
  checkRedirectResult,
  logout as googleLogout, 
  saveToDrive, 
  loadFromDrive,
  auth
} from './lib/googleDriveService';
import { User as FirebaseUser } from 'firebase/auth';

// Sub-components (to be extracted or kept inline if small)
import Dashboard from './components/Dashboard';
import ManpowerPool from './components/ManpowerPool';
import ProjectManagement from './components/Projects';
import AssignmentList from './components/Assignments';
import GanttView from './components/GanttChart';
import ShortageScreen from './components/ShortageAnalysis';
import UnassignedManpower from './components/UnassignedManpower';
import Login from './components/Login';

// Bulletproof Storage Utilities to bypass SecurityErrors in sandboxed previews
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

const safeSessionStorage = {
  getItem: (key: string): string | null => {
    try {
      return typeof window !== 'undefined' && window.sessionStorage ? window.sessionStorage.getItem(key) : null;
    } catch (e) {
      return null;
    }
  },
  setItem: (key: string, value: string): void => {
    try {
      if (typeof window !== 'undefined' && window.sessionStorage) {
        window.sessionStorage.setItem(key, value);
      }
    } catch (e) {}
  },
  removeItem: (key: string): void => {
    try {
      if (typeof window !== 'undefined' && window.sessionStorage) {
        window.sessionStorage.removeItem(key);
      }
    } catch (e) {}
  }
};

const STORAGE_KEY = 'kanooz_system_data';

export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [activeTab, setActiveTab] = useState<Tab>(Tab.DASHBOARD);
  const [manpower, setManpower] = useState<Manpower[]>([]);
  const [projects, setProjects] = useState<Project[]>([]);
  const [assignments, setAssignments] = useState<Assignment[]>([]);

  const [isAddingManpower, setIsAddingManpower] = useState(false);
  const [isAddingProject, setIsAddingProject] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  // Custom Reset & Backup States
  const [showResetModal, setShowResetModal] = useState(false);
  const [showBackupModal, setShowBackupModal] = useState(false);
  const [resetManpowerOpt, setResetManpowerOpt] = useState(true);
  const [resetProjectsOpt, setResetProjectsOpt] = useState(true);
  const [resetAssignmentsOpt, setResetAssignmentsOpt] = useState(true);

  // Credentials Modifying States
  const [showCredentialsModal, setShowCredentialsModal] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [currentPasswordConfirm, setCurrentPasswordConfirm] = useState('');

  // Google Drive Cloud Sync States
  const [googleUser, setGoogleUser] = useState<FirebaseUser | null>(null);
  const [googleToken, setGoogleToken] = useState<string | null>(null);
  const [isDriveSyncing, setIsDriveSyncing] = useState(false);
  const [driveSyncMessage, setDriveSyncMessage] = useState<string | null>(null);
  const [showCustomConfig, setShowCustomConfig] = useState(false);
  const [customConfigStr, setCustomConfigStr] = useState(() => {
    return safeLocalStorage.getItem('kanooz_custom_firebase_config') || '';
  });

  // Check Redirect Result on mount to handle redirect auth flow
  useEffect(() => {
    const handleRedirect = async () => {
      try {
        const result = await checkRedirectResult();
        if (result) {
          setGoogleUser(result.user);
          setGoogleToken(result.accessToken);
          setDriveSyncMessage('Successfully connected to Google Drive!');
        }
      } catch (err: any) {
        console.error('Redirect result exception:', err);
        const code = err?.code || '';
        const message = err?.message || '';
        const isDomainError = code === 'auth/unauthorized-domain' || message.includes('unauthorized-domain');
        const isOpenError = code === 'auth/operation-not-allowed' || message.includes('operation-not-allowed') ||
                            code === 'auth/configuration-not-found' || message.includes('configuration-not-found');
        const currentDomain = window.location.hostname;
        
        let activeProjectId = 'spiritual-amplifier-307pf';
        try {
          if (auth && auth.app && auth.app.options) {
            activeProjectId = auth.app.options.projectId || activeProjectId;
          }
        } catch (e) {}

        const hasCustomConfig = !!safeLocalStorage.getItem('kanooz_custom_firebase_config');

        if (isDomainError || isOpenError) {
          if (!hasCustomConfig) {
            alert(
              "🔒 Google Drive Backup Setup Required for External Deployments!\n\n" +
              `You are hosting this app on your own domain "${currentDomain}" but are currently using the default system database.\n\n` +
              `Because the default Firebase project (${activeProjectId}) is owned by the development sandbox, you do NOT have permission to manage its settings or authorize "${currentDomain}". This is why you see permission/configuration errors.\n\n` +
              "💡 EASY SOLUTION (Takes 2 Minutes):\n" +
              "To connect Google Drive sync, you just need to link this interface to your own free Firebase project:\n\n" +
              "1. Scroll to the bottom of the 'Backup & Sync' tab.\n" +
              "2. Open 'Custom Firebase Project (Advanced)' and expand 'SHOW ▼'.\n" +
              "3. Follow the simple steps to create your free project, enable Google auth, and paste your config block here.\n" +
              "4. Click 'Save & Apply' and you are ready to sync securely!"
            );
          } else {
            if (isDomainError) {
              alert(
                "🔒 Google Auth (Redirect): Unauthorized Client Domain!\n\n" +
                `Your custom Firebase project (${activeProjectId}) requires the domain "${currentDomain}" to be authorized.\n\n` +
                "To fix this, please follow these simple steps:\n" +
                `1. Open Firebase Authentication Settings:\n` +
                `   https://console.firebase.google.com/project/${activeProjectId}/authentication/settings\n\n` +
                "2. Find 'Authorized domains' and click 'Add domain'.\n" +
                `3. Add your current domain:\n` +
                `   👉 ${currentDomain}\n\n` +
                "Then return here and connect again!"
              );
            } else if (isOpenError) {
              alert(
                "🔒 Google Sign-In Method Disabled or Not Configured!\n\n" +
                `Your custom Firebase project (${activeProjectId}) has not enabled Google as a safe sign-in provider, or Google's configuration is missing.\n\n` +
                "How to enable Google Provider (takes 1 minute):\n" +
                `1. Open project settings provider tab:\n` +
                `   https://console.firebase.google.com/project/${activeProjectId}/authentication/providers\n\n` +
                "2. Click the 'Add new provider' button and choose 'Google'.\n" +
                "3. Toggle 'Enable', specify a support email, and click 'Save'.\n\n" +
                "Once saved, return here and retry!"
              );
            }
          }
        } else {
          // If the error indicates a storage or security-sandbox block (common in iframes, incognito, etc),
          // don't show a loud blocking alert on mount to avoid worsening user experience. Just log it.
          const isStorageIssue = message.includes('storage') || message.includes('SecurityError') || code.includes('storage');
          if (!isStorageIssue) {
            alert(`Google Redirect Login Failed: ${message || 'Please check your configurations.'}`);
          }
        }
      }
    };
    handleRedirect();
  }, []);

  // Initialize Google Auth state listener
  useEffect(() => {
    const unsubscribe = initGoogleAuth(
      (user, token) => {
        setGoogleUser(user);
        setGoogleToken(token);
      },
      () => {
        setGoogleUser(null);
        setGoogleToken(null);
      }
    );
    return () => {
      if (unsubscribe) unsubscribe();
    };
  }, []);

  // Force Light Theme (Dark Mode Disabled Throughout)
  useEffect(() => {
    document.documentElement.classList.remove('dark');
    safeLocalStorage.removeItem('kanooz_theme');
  }, []);

  // Load data
  useEffect(() => {
    const loggedInStatus = safeSessionStorage.getItem('kanooz_logged_in');
    if (loggedInStatus === 'true') {
      setIsLoggedIn(true);
    }

    const saved = safeLocalStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setManpower(parsed.manpower || []);
        setProjects(parsed.projects || []);
        setAssignments(parsed.assignments || []);
      } catch (err) {
        console.error('Failed to parse saved localStorage planning data:', err);
        setManpower([]);
        setProjects([]);
        setAssignments([]);
      }
    }
  }, []);

  // Save data
  useEffect(() => {
    safeLocalStorage.setItem(STORAGE_KEY, JSON.stringify({ manpower, projects, assignments }));
  }, [manpower, projects, assignments]);

  const handleUpdateCredentials = (e: React.FormEvent) => {
    e.preventDefault();
    
    let currentCreds = { username: 'Admin', password: 'Admin' };
    try {
      const stored = safeLocalStorage.getItem('kanooz_admin_credentials');
      if (stored) {
        currentCreds = JSON.parse(stored);
      }
    } catch (err) {
      console.error(err);
    }

    if (currentPasswordConfirm !== currentCreds.password) {
      alert('Authentication failure: The typed current password is incorrect.');
      return;
    }

    if (!newUsername.trim()) {
      alert('Error: Username cannot be blank.');
      return;
    }

    if (!newPassword.trim()) {
      alert('Error: Password cannot be blank.');
      return;
    }

    if (newPassword !== confirmPassword) {
      alert('Error: The new password and confirmed password do not match.');
      return;
    }

    const nextCreds = {
      username: newUsername.trim(),
      password: newPassword.trim()
    };

    safeLocalStorage.setItem('kanooz_admin_credentials', JSON.stringify(nextCreds));
    alert('Security Success: Administrator credentials modified successfully. Please keep them safe.');
    setShowCredentialsModal(false);
    
    // Reset inputs
    setNewUsername('');
    setNewPassword('');
    setConfirmPassword('');
    setCurrentPasswordConfirm('');
  };

  const handleLogout = () => {
    if (window.confirm('Are you sure you would like to sign out of the planning portal?')) {
      safeSessionStorage.removeItem('kanooz_logged_in');
      setIsLoggedIn(false);
    }
  };

  const handleReset = () => {
    // Open the custom reset configuration dialog instead
    setShowResetModal(true);
  };

  const executeConfiguredReset = () => {
    if (resetManpowerOpt) setManpower([]);
    if (resetProjectsOpt) setProjects([]);
    if (resetAssignmentsOpt) setAssignments([]);
    
    // Clear localStorage entirely or save partially
    const nextManpower = resetManpowerOpt ? [] : manpower;
    const nextProjects = resetProjectsOpt ? [] : projects;
    const nextAssignments = resetAssignmentsOpt ? [] : assignments;

    if (resetManpowerOpt && resetProjectsOpt && resetAssignmentsOpt) {
      safeLocalStorage.removeItem(STORAGE_KEY);
    } else {
      safeLocalStorage.setItem(STORAGE_KEY, JSON.stringify({ 
        manpower: nextManpower, 
        projects: nextProjects, 
        assignments: nextAssignments 
      }));
    }

    setShowResetModal(false);
    alert('System configuration reset executed successfully according to selected scope.');
  };

  const downloadFullBackup = () => {
    const dataObj = { manpower, projects, assignments, backupDate: dayjs().format('YYYY-MM-DD'), creator: 'Kanooz Central Planning' };
    const str = JSON.stringify(dataObj, null, 2);
    const blob = new Blob([str], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Kanooz_Master_Planning_Backup_${dayjs().format('DD-MMM-YY')}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleBackupUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const contents = event.target?.result as string;
        const parsed = JSON.parse(contents);
        
        if (!parsed.manpower || !parsed.projects || !parsed.assignments) {
          alert('Invalid backup file structure. File must contain manpower, projects, and assignments lists.');
          return;
        }

        setManpower(parsed.manpower);
        setProjects(parsed.projects);
        setAssignments(parsed.assignments);
        alert(`Success! Loaded ${parsed.manpower.length} workers, ${parsed.projects.length} projects, and preserved ${parsed.assignments.length} assignments.`);
        setShowBackupModal(false);
      } catch (err) {
        console.error(err);
        alert('Failed to parse the backup file. Please ensure it is a valid Kanooz Central Planning .json backup package.');
      }
    };
    reader.readAsText(file);
    e.target.value = ''; // Reset the input file path
  };

  const handleGoogleConnect = async () => {
    setIsDriveSyncing(true);
    setDriveSyncMessage('Connecting to Google...');
    try {
      const result = await googleSignIn();
      if (result) {
        setGoogleUser(result.user);
        setGoogleToken(result.accessToken);
        setDriveSyncMessage('Successfully connected to Google Drive!');
      }
    } catch (err: any) {
      console.error(err);
      const code = err?.code || '';
      const message = err?.message || '';
      const isPopupError = code === 'auth/popup-closed-by-user' || 
                           message.includes('popup-closed-by-user') ||
                           code === 'auth/popup-blocked' ||
                           message.includes('popup-blocked');
      const isDomainError = code === 'auth/unauthorized-domain' || message.includes('unauthorized-domain');
      const isOpenError = code === 'auth/operation-not-allowed' || message.includes('operation-not-allowed') ||
                          code === 'auth/configuration-not-found' || message.includes('configuration-not-found');
      const currentDomain = window.location.hostname;
      
      let activeProjectId = 'spiritual-amplifier-307pf';
      try {
        if (auth && auth.app && auth.app.options) {
          activeProjectId = auth.app.options.projectId || activeProjectId;
        }
      } catch (e) {}
      
      const hasCustomConfig = !!safeLocalStorage.getItem('kanooz_custom_firebase_config');

      if (isDomainError || isOpenError) {
        if (!hasCustomConfig) {
          alert(
            "🔒 Custom Firebase Setup Required for Google Drive on GitHub Pages/External Domains!\n\n" +
            `You are viewing this app from "${currentDomain}" but are currently using the default Google-owned sandbox project (${activeProjectId}).\n\n` +
            `Because this default project is owned by the development sandbox, you do NOT have permission to add custom domains to its authorized list. This is why you see permission/configuration errors on GitHub Pages.\n\n` +
            "💡 EASY SOLUTION (Takes 2 Minutes):\n" +
            "To connect your Google Drive sync on your own domain, you must link this UI to your own free Firebase project:\n\n" +
            "1. Scroll to the bottom of the 'Backup & Sync' tab.\n" +
            "2. Open 'Custom Firebase Project (Advanced)' and expand 'SHOW ▼'.\n" +
            "3. Follow the simple steps to create your free project, enable Google auth, and paste your config block here.\n" +
            "4. Click 'Save & Apply' and you are ready to sync your projects securely!"
          );
          setDriveSyncMessage(`Authentication setup required for: ${currentDomain}`);
        } else {
          if (isDomainError) {
            alert(
              "🔒 Google Auth: Unauthorized Client Domain!\n\n" +
              `Your private Firebase project (${activeProjectId}) requires the domain "${currentDomain}" to be authorized before signing in.\n\n` +
              "To resolve this, please follow these simple steps:\n" +
              `1. Open your own Firebase Console settings page:\n` +
              `   https://console.firebase.google.com/project/${activeProjectId}/authentication/settings\n\n` +
              "2. Find the 'Authorized domains' section and click 'Add domain'.\n" +
              "3. Copy and paste your current environment domain:\n" +
              `   👉 ${currentDomain}\n\n` +
              "Once registered, refresh the page and try connecting again!"
            );
          } else if (isOpenError) {
            alert(
              "🔒 Google Sign-In Method Disabled or Not Configured!\n\n" +
              `Your custom Firebase project (${activeProjectId}) has not enabled Google as a safe sign-in provider yet, or its Google configuration is incomplete.\n\n` +
              "HOW TO ENABLE GOOGLE PROVIDER (Takes 1 Minute):\n" +
              `1. Open your Firebase project console directly:\n` +
              `   https://console.firebase.google.com/project/${activeProjectId}/authentication/providers\n\n` +
              "2. Click the 'Add new provider' button and select 'Google' from the list.\n" +
              "3. Toggle the switch to 'Enable', choose a project support email, and click 'Save'.\n\n" +
              "After saving the Firebase console changes, return here and try connecting again!"
            );
          }
          setDriveSyncMessage('Custom Firebase configuration error.');
        }
      } else if (isPopupError) {
        alert(
          "⚠️ Google Auth Popup Blocked or Closed\n\n" +
          "This error occurs when standard login popups are restricted or blocked by your browser settings or inside sandboxed previews.\n\n" +
          "💡 EASY WORKAROUND:\n" +
          "Please click the 'Use Redirect Flow' button below instead! It bypasses popup blocking and works perfectly on all devices and browsers."
        );
        setDriveSyncMessage('Popup blocked. Please use Redirect Flow.');
      } else {
        alert(`Google Connection Failed: ${message || 'Check your internet connection or browser settings.'}`);
        setDriveSyncMessage(null);
      }
    } finally {
      setIsDriveSyncing(false);
    }
  };

  const handleGoogleConnectRedirect = async () => {
    setIsDriveSyncing(true);
    setDriveSyncMessage('Redirecting to Google Sign-In...');
    try {
      await googleSignInRedirect();
    } catch (err: any) {
      console.error(err);
      const code = err?.code || '';
      const message = err?.message || '';
      const isDomainError = code === 'auth/unauthorized-domain' || message.includes('unauthorized-domain');
      const isOpenError = code === 'auth/operation-not-allowed' || message.includes('operation-not-allowed') ||
                          code === 'auth/configuration-not-found' || message.includes('configuration-not-found');
      
      let activeProjectId = 'spiritual-amplifier-307pf';
      try {
        if (auth && auth.app && auth.app.options) {
          activeProjectId = auth.app.options.projectId || activeProjectId;
        }
      } catch (e) {}

      const hasCustomConfig = !!safeLocalStorage.getItem('kanooz_custom_firebase_config');

      if (isDomainError || isOpenError) {
        const currentDomain = window.location.hostname;
        if (!hasCustomConfig) {
          alert(
            "🔒 Custom Firebase Setup Required for Google Drive on GitHub Pages/External Domains!\n\n" +
            `You are viewing this app from "${currentDomain}" but are currently using the default Google-owned sandbox project (${activeProjectId}).\n\n` +
            `Because this default project is owned by the development sandbox, you do NOT have permission to add custom domains to its authorized list. This is why you see permission/configuration errors on GitHub Pages.\n\n` +
            "💡 EASY SOLUTION (Takes 2 Minutes):\n" +
            "To connect your Google Drive sync on your own domain, you must link this UI to your own free Firebase project:\n\n" +
            "1. Scroll to the bottom of the 'Backup & Sync' tab.\n" +
            "2. Open 'Custom Firebase Project (Advanced)' and expand 'SHOW ▼'.\n" +
            "3. Follow the simple steps to create your free project, enable Google auth, and paste your config block here.\n" +
            "4. Click 'Save & Apply' and you are ready to sync your projects securely!"
          );
        } else {
          if (isDomainError) {
            alert(
              "🔒 Google Auth (Redirect): Unauthorized Client Domain!\n\n" +
              `Your custom Firebase project (${activeProjectId}) requires the domain "${currentDomain}" to be authorized before signing in.\n\n` +
              "To resolve this, please follow these simple steps:\n" +
              `1. Open your own Firebase Console settings page:\n` +
              `   https://console.firebase.google.com/project/${activeProjectId}/authentication/settings\n\n` +
              "2. Find the 'Authorized domains' section and click 'Add domain'.\n\n" +
              "3. Copy and paste your current environment domain:\n" +
              `   👉 ${currentDomain}\n\n` +
              "Once registered, refresh the page and try connecting again!"
            );
          } else if (isOpenError) {
            alert(
              "🔒 Google Sign-In Method Disabled or Not Configured!\n\n" +
              `Your custom Firebase project (${activeProjectId}) has not enabled Google as a safe sign-in provider yet, or its Google configuration is incomplete.\n\n` +
              "HOW TO ENABLE GOOGLE PROVIDER (Takes 1 Minute):\n" +
              `1. Open your Firebase project console directly:\n` +
              `   https://console.firebase.google.com/project/${activeProjectId}/authentication/providers\n\n` +
              "2. Click the 'Add new provider' button and select 'Google' from the list.\n" +
              "3. Toggle the switch to 'Enable', choose a project support email, and click 'Save'.\n\n" +
              "After saving the Firebase console changes, return here and try connecting again!"
            );
          }
        }
      } else {
        alert(`Could not start redirect flow: ${message || 'Error occurred.'}`);
      }
      setIsDriveSyncing(false);
      setDriveSyncMessage(null);
    }
  };

  const handleGoogleDisconnect = async () => {
    if (window.confirm('Are you sure you want to disconnect your Google Drive? Your local data will remain safe.')) {
      setIsDriveSyncing(true);
      try {
        await googleLogout();
        setGoogleUser(null);
        setGoogleToken(null);
        setDriveSyncMessage('Signed out from Google.');
        setTimeout(() => setDriveSyncMessage(null), 3000);
      } catch (err: any) {
        console.error(err);
        alert(`Disconnection Error: ${err.message}`);
      } finally {
        setIsDriveSyncing(false);
      }
    }
  };

  const handleSaveCustomConfig = () => {
    if (!customConfigStr.trim()) {
      if (safeLocalStorage.getItem('kanooz_custom_firebase_config')) {
        safeLocalStorage.removeItem('kanooz_custom_firebase_config');
        alert('Custom Firebase configuration cleared! Resetting back to default system credentials.');
        window.location.reload();
      }
      return;
    }

    let parsed: any = null;

    // 1. Try standard JSON parsing
    try {
      parsed = JSON.parse(customConfigStr);
    } catch (e) {
      // 2. If JSON fails, extract using regex to support JS object copy-pasting
      const input = customConfigStr;
      const extractField = (field: string, text: string): string | null => {
        const regex = new RegExp(`['"]?${field}['"]?\\s*:\\s*['"]([^'"]+)['"]`);
        const match = text.match(regex);
        return match ? match[1] : null;
      };

      const apiKey = extractField('apiKey', input);
      const authDomain = extractField('authDomain', input);
      const projectId = extractField('projectId', input);
      const storageBucket = extractField('storageBucket', input);
      const messagingSenderId = extractField('messagingSenderId', input);
      const appId = extractField('appId', input);
      const measurementId = extractField('measurementId', input);

      if (apiKey && authDomain && projectId) {
        parsed = {
          apiKey,
          authDomain,
          projectId,
          ...(storageBucket ? { storageBucket } : {}),
          ...(messagingSenderId ? { messagingSenderId } : {}),
          ...(appId ? { appId } : {}),
          ...(measurementId ? { measurementId } : {})
        };
      }
    }

    if (parsed && parsed.apiKey && parsed.authDomain && parsed.projectId) {
      safeLocalStorage.setItem('kanooz_custom_firebase_config', JSON.stringify(parsed, null, 2));
      alert('🎉 Success! Custom Firebase credentials applied.\n\nThe page will now reload to initialize your Firebase project.');
      window.location.reload();
    } else {
      alert(
        'Could not parse the configurations. Please copy and paste the entire config block from the Firebase console!\n\n' +
        'It should look like this (you can paste the whole block directly):\n' +
        'const firebaseConfig = {\n' +
        '  apiKey: "...",\n' +
        '  authDomain: "...",\n' +
        '  projectId: "..."\n' +
        '};'
      );
    }
  };

  const handleBackupToDrive = async () => {
    if (!googleUser) {
      alert('Connection error: Google Drive is not paired. Please connect your Drive first.');
      return;
    }

    const confirmed = window.confirm(
      'Are you sure you want to backup your active database to Google Drive?\n\nThis will write "Kanooz_Master_Planning_Backup.json" directly onto your connected Google Drive storage. Overwriting any previous backend backups.'
    );
    if (!confirmed) return;

    setIsDriveSyncing(true);
    setDriveSyncMessage('Uploading backup pack to Google Drive...');
    try {
      const payload = {
        manpower,
        projects,
        assignments,
        backupDate: dayjs().format('YYYY-MM-DD'),
        creator: 'Kanooz Central Planning'
      };
      
      const result = await saveToDrive(payload);
      
      setDriveSyncMessage(
        result.updated 
          ? 'Backup successfully updated on Google Drive!' 
          : 'New backup package registered on Google Drive!'
      );
      
      alert(
        `Cloud Synchronize Successful!\n\n${
          result.updated 
            ? 'Your online Google Drive file has been updated successfully.' 
            : 'A new file "Kanooz_Master_Planning_Backup.json" was successfully created on your Google Drive.'
        }`
      );
    } catch (err: any) {
      console.error(err);
      alert(`Google Drive Upload failed: ${err.message || err}`);
      setDriveSyncMessage('Upload failed. Try again.');
    } finally {
      setIsDriveSyncing(false);
      setTimeout(() => setDriveSyncMessage(null), 5000);
    }
  };

  const handleRestoreFromDrive = async () => {
    if (!googleUser) {
      alert('Connection error: Google Drive is not paired. Please connect your Drive first.');
      return;
    }

    const confirmed = window.confirm(
      'Are you sure you want to restore data from Google Drive?\n\nThis will download your cloud master backup file "Kanooz_Master_Planning_Backup.json" and completely replace your current local manpower pool, projects list, and confirmed assignments. Local data will be overwritten!'
    );
    if (!confirmed) return;

    setIsDriveSyncing(true);
    setDriveSyncMessage('Retrieving database from Google Drive...');
    try {
      const parsed = await loadFromDrive();
      if (!parsed) {
        alert('Could not find any file named "Kanooz_Master_Planning_Backup.json" on your Google Drive. Make a backup first.');
        setDriveSyncMessage('No backup file found on Google Drive.');
        return;
      }

      if (!parsed.manpower || !parsed.projects || !parsed.assignments) {
        alert('Validation error: The backup file on your Google Drive doesn\'t match the required database schema standards.');
        setDriveSyncMessage('Invalid file schema.');
        return;
      }

      setManpower(parsed.manpower);
      setProjects(parsed.projects);
      setAssignments(parsed.assignments);
      
      setDriveSyncMessage('Database restored successfully!');
      alert(`Success!\n\nDatabase fully synchronized. Loaded ${parsed.manpower.length} workers, ${parsed.projects.length} projects, and ${parsed.assignments.length} assignments directly from your cloud storage.`);
    } catch (err: any) {
      console.error(err);
      alert(`Google Drive Retrieve failed: ${err.message || err}`);
      setDriveSyncMessage('Retrieve failed.');
    } finally {
      setIsDriveSyncing(false);
      setTimeout(() => setDriveSyncMessage(null), 5000);
    }
  };

  const handleAutoAssign = () => {
    try {
      if (manpower.length === 0) {
        alert('Manpower pool is empty. Please add workers first.');
        return;
      }
      if (projects.length === 0) {
        alert('No projects found. Please create a project first.');
        return;
      }

      const projectsWithReqs = projects.filter(p => p.requirements.some(r => r.qty > 0));
      if (projectsWithReqs.length === 0) {
        alert('No requirements found. Please use "Define Requirements" in the Projects tab to set quantities.');
        setActiveTab(Tab.PROJECTS);
        return;
      }

      const newArr = runAutoAssignment(manpower, projects);
      
      if (newArr.length > 0) {
        setAssignments(newArr);
        const shortages = calculateShortages(manpower, projects, newArr);
        
        if (shortages.length > 0) {
          alert(`Successfully assigned ${newArr.length} workers.\n\nNote: ${shortages.length} requirement(s) could not be fully met. Check the "Shortages" tab for details.`);
        } else {
          alert(`Optimization complete! ${newArr.length} workers assigned. All project requirements fully satisfied.`);
        }
        setActiveTab(Tab.ASSIGNMENTS);
      } else {
        alert(`The auto-assign engine found 0 matches.\n\nPlease check:\n1. Worker crafts must match requirement crafts exactly.\n2. Workers must be available for the ENTIRE duration of the requirement.\n3. Verify worker Join Date and Contract End (Release) dates.`);
      }
    } catch (error) {
      console.error('Auto-Assign Error:', error);
      alert('The assignment engine encountered an unexpected error. Please check your data formatting and try again.');
    }
  };

  const tabs = [
    { id: Tab.DASHBOARD, label: 'Dashboard', icon: BarChart3 },
    { id: Tab.MANPOWER, label: 'Manpower Pool', icon: Users },
    { id: Tab.PROJECTS, label: 'Projects', icon: Briefcase },
    { id: Tab.ASSIGNMENTS, label: 'Assignments', icon: Calendar },
    { id: Tab.GANTT, label: 'Gantt View', icon: GanttIcon },
    { id: Tab.SHORTAGE, label: 'Shortages', icon: AlertTriangle },
    { id: Tab.UNASSIGNED, label: 'Not Assigned', icon: UserMinus },
    { id: Tab.REPORTS, label: 'Reports', icon: FileText },
  ];

  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    
    // 1. Manpower Sheet
    const wsManpower = XLSX.utils.json_to_sheet(manpower.map(m => ({
      'Name': m.name,
      'Badge Number': m.badgeNo,
      'Passport / Iqama': m.passportIqama || '',
      'Craft': m.craft,
      'Employment Type': m.employmentType,
      'Joining Date': formatToExcelDate(m.joinDate),
      'Contract End Date': formatToExcelDate(m.releaseDate),
      'Vacation Start': formatToExcelDate(m.vacationStart),
      'Vacation Finish': formatToExcelDate(m.vacationEnd)
    })));
    XLSX.utils.book_append_sheet(wb, wsManpower, "Manpower Pool");
    
    // 2. Projects Master Sheet
    const wsProjects = XLSX.utils.json_to_sheet(projects.map(p => ({
      Code: p.code,
      Name: p.name,
      Location: p.location,
      'Start Date': formatToExcelDate(p.startDate),
      'End Date': formatToExcelDate(p.endDate),
      'Total Required': p.requirements.reduce((sum, r) => sum + r.qty, 0)
    })));
    XLSX.utils.book_append_sheet(wb, wsProjects, "Projects Master");

    // 3. Requirements Detail Sheet
    const flatRequirements = projects.flatMap(p => 
      p.requirements.map(r => ({
        'Project Code': p.code,
        'Project Name': p.name,
        'Craft': r.craft,
        'Phase': r.phase,
        'Qty Required': r.qty,
        'Requirement Start': formatToExcelDate(r.startDate),
        'Requirement End': formatToExcelDate(r.endDate)
      }))
    );
    const wsReqs = XLSX.utils.json_to_sheet(flatRequirements);
    XLSX.utils.book_append_sheet(wb, wsReqs, "Demand Details");

    // 4. Assignments Sheet
    const flatAssignments = assignments.map(a => {
      const p = projects.find(proj => proj.id === a.projectId);
      const m = manpower.find(man => man.id === a.workerId);
      return {
        'Project Code': p?.code || 'N/A',
        'Project Name': p?.name || 'Unknown',
        'Worker Name': m?.name || 'Unknown',
        'Badge No': m?.badgeNo || 'Unknown',
        'Craft': a.craft,
        'Start Date': formatToExcelDate(a.startDate),
        'End Date': formatToExcelDate(a.endDate),
        'Date Assigned': formatToExcelDate(dayjs().format('YYYY-MM-DD'))
      };
    });
    const wsAss = XLSX.utils.json_to_sheet(flatAssignments);
    XLSX.utils.book_append_sheet(wb, wsAss, "Confirmed Assignments");

    // 5. Gap Analysis (Shortages)
    const gapAnalysis = projects.flatMap(p => 
      p.requirements.map(req => {
        const assignedCount = assignments.filter(a => a.projectId === p.id && a.craft === req.craft).length;
        const gap = req.qty - assignedCount;
        return {
          'Project': p.name,
          'Craft': req.craft,
          'Required': req.qty,
          'Assigned': assignedCount,
          'Shortage': gap > 0 ? gap : 0,
          'Start Date': formatToExcelDate(req.startDate),
          'End Date': formatToExcelDate(req.endDate),
          'Status': gap <= 0 ? 'FULFILLED' : 'UNDERSTAFFED'
        };
      })
    );
    const wsGaps = XLSX.utils.json_to_sheet(gapAnalysis);
    XLSX.utils.book_append_sheet(wb, wsGaps, "Gap Analysis");

    XLSX.writeFile(wb, `Kanooz_Master_Planning_${dayjs().format('DD-MMM-YY')}.xlsx`);
  };

  if (!isLoggedIn) {
    return <Login onLoginSuccess={() => setIsLoggedIn(true)} />;
  }

  return (
    <div className="min-h-screen bg-[#FDFDFD] text-[#1A1A1A] font-sans">
      {/* Sidebar / Nav */}
      <div className="flex h-screen overflow-hidden">
        <aside className={cn(
          "bg-white border-r border-[#E5E5E5] flex flex-col shrink-0 transition-all duration-300 ease-in-out z-20",
          isSidebarOpen ? "w-64" : "w-16"
        )}>
          <div className="h-16 bg-white border-b border-[#E5E5E5] flex items-center px-4 justify-end shrink-0">
            <button 
              onClick={() => setIsSidebarOpen(!isSidebarOpen)}
              className="p-1.5 hover:bg-gray-50 rounded-lg transition-colors text-gray-400 hover:text-indigo-600 focus:outline-none"
              title={isSidebarOpen ? "Collapse Sidebar" : "Expand Sidebar"}
            >
              <Menu className="w-4 h-4" />
            </button>
          </div>

          <nav className="flex-1 p-2 space-y-1 overflow-y-auto">
            {tabs.map((tab) => (
              <button
                key={tab.id}
                id={`tab-${tab.id}`}
                onClick={() => setActiveTab(tab.id as Tab)}
                className={cn(
                  "w-full flex items-center gap-3 px-3 py-2 rounded-lg text-sm font-medium transition-all group relative",
                  activeTab === tab.id 
                    ? "bg-indigo-50 text-indigo-700 shadow-sm" 
                    : "text-[#666] hover:bg-gray-50 hover:text-[#1A1A1A]"
                )}
                title={!isSidebarOpen ? tab.label : ""}
              >
                <tab.icon className={cn(
                  "w-4 h-4 shrink-0",
                  activeTab === tab.id ? "text-indigo-600" : "text-[#A0A0A0] group-hover:text-[#666]"
                )} />
                {isSidebarOpen && <span className="truncate">{tab.label}</span>}
              </button>
            ))}
          </nav>

          <div className="flex-none p-2 space-y-1 border-t border-slate-100 bg-slate-50/50">
            {isSidebarOpen ? (
              <button
                onClick={() => {
                  let currentCreds = { username: 'Admin', password: 'Admin' };
                  try {
                    const stored = safeLocalStorage.getItem('kanooz_admin_credentials');
                    if (stored) currentCreds = JSON.parse(stored);
                  } catch (e) {}
                  setNewUsername(currentCreds.username);
                  setShowCredentialsModal(true);
                }}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-lg text-xs font-semibold text-slate-650 hover:bg-white hover:text-indigo-600 border border-transparent hover:border-slate-200 transition-all text-left"
              >
                <ShieldCheck className="w-4 h-4 text-slate-400 shrink-0" />
                <span className="truncate">Credentials & Access</span>
              </button>
            ) : (
              <button
                onClick={() => {
                  let currentCreds = { username: 'Admin', password: 'Admin' };
                  try {
                    const stored = safeLocalStorage.getItem('kanooz_admin_credentials');
                    if (stored) currentCreds = JSON.parse(stored);
                  } catch (e) {}
                  setNewUsername(currentCreds.username);
                  setShowCredentialsModal(true);
                }}
                className="w-full flex items-center justify-center p-2 rounded-lg text-slate-500 hover:bg-white hover:text-indigo-600 border border-transparent hover:border-slate-200 transition-all"
                title="Credentials & Access"
              >
                <ShieldCheck className="w-4 h-4 text-slate-400 shrink-0" />
              </button>
            )}

            {isSidebarOpen ? (
              <button
                onClick={handleLogout}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-lg text-xs font-semibold text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-100 transition-all text-left"
              >
                <LogOut className="w-4 h-4 text-rose-500 shrink-0" />
                <span className="truncate">Sign Out</span>
              </button>
            ) : (
              <button
                onClick={handleLogout}
                className="w-full flex items-center justify-center p-2 rounded-lg text-rose-500 hover:bg-rose-50 border border-transparent hover:border-rose-100 transition-all"
                title="Sign Out"
              >
                <LogOut className="w-4 h-4 text-rose-500 shrink-0" />
              </button>
            )}
          </div>
        </aside>

        {/* Main Content */}
        <main className="flex-1 flex flex-col overflow-hidden bg-[#FAFAFA] relative">
          {/* Watermark Background */}
          <div className="absolute inset-0 pointer-events-none flex items-center justify-center opacity-[0.03] z-0 overflow-hidden">
            <img 
              src="https://lh3.googleusercontent.com/d/1oQr4RoR9ON5vK6U_oIN2hW0HdnpEloO8" 
              alt="Background Logo" 
              className="w-1/2 max-w-2xl grayscale"
              referrerPolicy="no-referrer"
            />
          </div>

          <header className="h-16 bg-white border-b border-[#E5E5E5] flex items-center justify-between px-6 shrink-0 z-20 sticky top-0">
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2.5">
                <img src="https://kanooz.com/assets/images/logo.png" alt="Logo" className="h-6 w-auto" referrerPolicy="no-referrer" />
                <div className="h-6 w-[1px] bg-[#E5E5E5]" />
                <div>
                  <h2 className="text-xs font-bold text-indigo-900 leading-tight">Central Planning Portal</h2>
                  <p className="text-[9px] text-[#888] font-medium uppercase tracking-widest leading-none">{activeTab} View</p>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-4">
              <div className="hidden lg:flex items-center gap-4">
                <span className="text-[10px] text-[#888] font-bold uppercase tracking-wider">{dayjs().format('dddd, D MMMM YYYY')}</span>
                <div className="h-4 w-[1px] bg-[#E5E5E5]" />
              </div>
              
              <div className="flex items-center gap-2">
                <button 
                  onClick={handleReset}
                  className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all no-print"
                  title="System Reset"
                >
                  <RotateCcw className="w-4 h-4" />
                </button>

                {activeTab === Tab.MANPOWER && (
                  <button 
                    onClick={() => setIsAddingManpower(!isAddingManpower)}
                    className="bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-2 hover:bg-indigo-700 shadow-sm transition-all"
                  >
                    <Plus className="w-4 h-4" />
                    Add Worker
                  </button>
                )}
                {activeTab === Tab.PROJECTS && (
                  <button 
                    onClick={() => setIsAddingProject(!isAddingProject)}
                    className="bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-2 hover:bg-indigo-700 shadow-sm transition-all"
                  >
                    <Plus className="w-4 h-4" />
                    Add Project
                  </button>
                )}
              </div>
            </div>
          </header>

          <div className="flex-1 overflow-y-auto p-4 md:p-6 z-10">
            <AnimatePresence mode="wait">
              <motion.div
                key={activeTab}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                transition={{ duration: 0.2 }}
                className="max-w-7xl mx-auto h-full"
              >
                {activeTab === Tab.DASHBOARD && (
                  <Dashboard manpower={manpower} projects={projects} assignments={assignments} />
                )}
                {activeTab === Tab.MANPOWER && (
                  <ManpowerPool 
                    manpower={manpower} 
                    setManpower={setManpower} 
                    isAdding={isAddingManpower} 
                    onCloseAdd={() => setIsAddingManpower(false)} 
                    googleUser={googleUser}
                    onGoogleConnect={handleGoogleConnect}
                  />
                )}
                {activeTab === Tab.PROJECTS && (
                  <ProjectManagement 
                    projects={projects} 
                    setProjects={setProjects} 
                    isAdding={isAddingProject}
                    onCloseAdd={() => setIsAddingProject(false)}
                    googleUser={googleUser}
                    onGoogleConnect={handleGoogleConnect}
                  />
                )}
                {activeTab === Tab.ASSIGNMENTS && (
                  <AssignmentList 
                    assignments={assignments} 
                    manpower={manpower} 
                    projects={projects}
                    onAutoAssign={handleAutoAssign}
                    setAssignments={setAssignments}
                  />
                )}
                {activeTab === Tab.GANTT && (
                  <GanttView manpower={manpower} projects={projects} assignments={assignments} />
                )}
                {activeTab === Tab.SHORTAGE && (
                  <ShortageScreen manpower={manpower} projects={projects} assignments={assignments} />
                )}
                {activeTab === Tab.UNASSIGNED && (
                  <UnassignedManpower manpower={manpower} assignments={assignments} />
                )}
                 {activeTab === Tab.REPORTS && (
                  <div className="space-y-6">
                    <div className="bg-white rounded-2xl border border-[#E5E5E5] p-5 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div>
                        <h3 className="text-base font-bold text-[#1A1A1A]">Reports & Backups</h3>
                        <p className="text-xs text-gray-500 mt-0.5">Consolidated spreadsheet generation and system database backups</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-4xl mx-auto py-4">
                      {/* Excel Reports Card */}
                      <div className="bg-white p-8 rounded-2xl border border-[#E5E5E5] text-center shadow-sm flex flex-col justify-between">
                        <div className="mb-6">
                          <FileText className="w-12 h-12 text-indigo-100 mx-auto mb-4" />
                          <h3 className="text-lg font-bold mb-2">Detailed Reports Export</h3>
                          <p className="text-xs text-[#666] leading-relaxed">
                            Generate multi-tab Excel reports containing consolidated planning records, manpower pools, demand lists, shortages, and assignments in <b>dd-mmm-yy</b> format.
                          </p>
                        </div>
                        <button 
                          onClick={handleExport}
                          className="bg-indigo-600 text-white w-full py-3 rounded-xl font-bold hover:bg-indigo-700 shadow-lg shadow-indigo-100 transition-all inline-flex items-center justify-center gap-2 text-sm"
                        >
                          <Download className="w-4 h-4" />
                          Download Excel Report
                        </button>
                      </div>

                      {/* Backup Master Packs Card */}
                      <div className="bg-white p-8 rounded-2xl border border-indigo-100 text-center shadow-sm flex flex-col justify-between bg-gradient-to-br from-white to-indigo-50/10">
                        <div className="mb-6">
                          <RefreshCw className="w-12 h-12 text-indigo-500 mx-auto mb-4" />
                          <h3 className="text-lg font-bold mb-2 text-indigo-950">Backup & Sync Master Pack</h3>
                          <p className="text-xs text-[#666] leading-relaxed">
                            Export or reload the <b>Entire Database (including all earlier assignments)</b> as a single JSON pack. Perfect to save locally or within your shared directory.
                          </p>
                        </div>
                        <div className="space-y-3">
                          <button 
                            onClick={downloadFullBackup}
                            className="bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 w-full py-2.5 rounded-xl font-bold transition-all text-xs inline-flex items-center justify-center gap-2"
                          >
                            <Download className="w-3.5 h-3.5" />
                            Download Master Pack (.json)
                          </button>
                          
                          <label className="border border-dashed border-indigo-200 bg-indigo-50/30 hover:bg-indigo-50/60 hover:border-indigo-300 w-full py-2.5 rounded-xl font-bold cursor-pointer transition-all text-xs inline-flex items-center justify-center gap-2 text-indigo-600">
                            <Upload className="w-3.5 h-3.5" />
                            Restore Backup Pack (.json)
                            <input 
                              type="file" 
                              accept=".json" 
                              onChange={handleBackupUpload} 
                              className="hidden" 
                            />
                          </label>
                        </div>
                      </div>
                    </div>
                  </div>
                 )}
              </motion.div>
            </AnimatePresence>
          </div>

          {/* Custom Reset Configuration Dialog */}
          {showResetModal && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-[2px] animate-in fade-in duration-200">
              <div className="bg-white rounded-2xl border border-gray-100 shadow-2xl p-6 w-full max-w-md mx-4 animate-in zoom-in-95 duration-200">
                <div className="flex items-center gap-2 mb-2 text-red-600">
                  <RotateCcw className="w-5 h-5 shrink-0" />
                  <h3 className="text-base font-bold text-gray-900">Configure System Reset</h3>
                </div>
                <p className="text-xs text-gray-500 mb-4 leading-relaxed">
                  Select database elements to clear. Keep <b>Confirmed Assignments</b> unchecked if you want your matches preserved for reloaded data!
                </p>
                
                <div className="space-y-3 mb-6 bg-gray-50 p-4 rounded-xl border border-gray-100">
                  <label className="flex items-start gap-3 cursor-pointer select-none">
                    <input 
                      type="checkbox" 
                      className="mt-0.5 w-4 h-4 text-indigo-650 border-gray-300 rounded focus:ring-indigo-500 accent-indigo-600"
                      checked={resetManpowerOpt}
                      onChange={e => setResetManpowerOpt(e.target.checked)}
                    />
                    <div>
                      <span className="text-xs font-bold text-gray-800">Clear Manpower Pool</span>
                      <p className="text-[10px] text-gray-400">Deletes the active workers database</p>
                    </div>
                  </label>

                  <label className="flex items-start gap-3 cursor-pointer select-none">
                    <input 
                      type="checkbox" 
                      className="mt-0.5 w-4 h-4 text-indigo-650 border-gray-300 rounded focus:ring-indigo-500 accent-indigo-600"
                      checked={resetProjectsOpt}
                      onChange={e => setResetProjectsOpt(e.target.checked)}
                    />
                    <div>
                      <span className="text-xs font-bold text-gray-800">Clear Projects & Site Master</span>
                      <p className="text-[10px] text-gray-400">Deletes site requisitions, requirements, and targets</p>
                    </div>
                  </label>

                  <label className="flex items-start gap-3 cursor-pointer select-none">
                    <input 
                      type="checkbox" 
                      className="mt-0.5 w-4 h-4 text-indigo-650 border-gray-300 rounded focus:ring-indigo-500 accent-indigo-600"
                      checked={resetAssignmentsOpt}
                      onChange={e => setResetAssignmentsOpt(e.target.checked)}
                    />
                    <div>
                      <span className="text-xs font-bold text-gray-800">Clear Confirmed Assignments</span>
                      <p className="text-[10px] text-gray-400">Removes the optimized worker assignments</p>
                    </div>
                  </label>
                </div>

                <div className="flex items-center justify-end gap-3 text-xs font-semibold">
                  <button 
                    onClick={() => setShowResetModal(false)}
                    className="px-4 py-2 border border-gray-200 hover:bg-gray-50 text-gray-600 rounded-xl transition-colors"
                  >
                    Cancel
                  </button>
                  <button 
                    onClick={executeConfiguredReset}
                    className="px-5 py-2 bg-red-650 hover:bg-red-700 bg-red-600 text-white rounded-xl transition-colors shadow-md shadow-red-50"
                  >
                    Execute Reset
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Change Security Credentials Dialog */}
          {showCredentialsModal && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-[2px] animate-in fade-in duration-200">
              <div className="bg-white rounded-2xl border border-gray-100 shadow-2xl p-6 w-full max-w-md mx-4 animate-in zoom-in-95 duration-200">
                <div className="flex items-center gap-2 mb-2 text-indigo-750 text-indigo-700">
                  <ShieldCheck className="w-5 h-5 shrink-0" />
                  <h3 className="text-base font-bold text-gray-900">Modify Security Credentials</h3>
                </div>
                <p className="text-xs text-gray-500 mb-4 leading-relaxed">
                  Change portal credentials. Ensure you write them down safely so you do not lock yourself out!
                </p>

                <form onSubmit={handleUpdateCredentials} className="space-y-4">
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wide block">
                      New Username
                    </label>
                    <input 
                      type="text" 
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-xl text-xs focus:ring-2 focus:ring-indigo-100 outline-none text-gray-800 font-medium"
                      placeholder="e.g. Admin"
                      value={newUsername}
                      onChange={e => setNewUsername(e.target.value)}
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wide block">
                      New Password
                    </label>
                    <input 
                      type="password" 
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-xl text-xs focus:ring-2 focus:ring-indigo-100 outline-none text-gray-800 font-mono"
                      placeholder="••••••••"
                      value={newPassword}
                      onChange={e => setNewPassword(e.target.value)}
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wide block">
                      Confirm New Password
                    </label>
                    <input 
                      type="password" 
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-xl text-xs focus:ring-2 focus:ring-indigo-100 outline-none text-gray-800 font-mono"
                      placeholder="••••••••"
                      value={confirmPassword}
                      onChange={e => setConfirmPassword(e.target.value)}
                    />
                  </div>

                  <div className="border-t border-gray-100 pt-3 mt-2 space-y-1">
                    <label className="text-[10px] font-bold text-rose-600 uppercase tracking-wide block">
                      Current Password (To Authorize)
                    </label>
                    <input 
                      type="password" 
                      required
                      className="w-full px-4 py-2 border border-rose-200 rounded-xl text-xs focus:ring-2 focus:ring-rose-100 outline-none text-gray-800 font-mono bg-rose-50/20"
                      placeholder="Enter active password"
                      value={currentPasswordConfirm}
                      onChange={e => setCurrentPasswordConfirm(e.target.value)}
                    />
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-3 text-xs font-semibold border-t border-gray-100">
                    <button 
                      type="button"
                      onClick={() => {
                        setShowCredentialsModal(false);
                        setNewUsername('');
                        setNewPassword('');
                        setConfirmPassword('');
                        setCurrentPasswordConfirm('');
                      }}
                      className="px-4 py-2 border border-gray-200 hover:bg-gray-50 text-gray-600 rounded-xl transition-colors"
                    >
                      Cancel
                    </button>
                    <button 
                      type="submit"
                      className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl transition-colors shadow-md shadow-indigo-50"
                    >
                      Apply Modify
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          <footer className="h-12 bg-white border-t border-[#E5E5E5] flex items-center justify-between px-8 text-[11px] text-[#888] font-medium shrink-0 z-10">
            <div className="flex items-center gap-2">
              <span>© 2026 Kanooz Central Planning. All Rights Reserved.</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="bg-green-50 text-green-600 px-2 py-0.5 rounded-full text-[10px] uppercase font-bold tracking-tight">System Online</span>
            </div>
          </footer>
        </main>
      </div>
    </div>
  );
}
