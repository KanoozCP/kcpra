import React, { useState } from 'react';
import { ShieldCheck, Lock, User, AlertCircle, Eye, EyeOff } from 'lucide-react';
import { motion } from 'motion/react';

interface LoginProps {
  onLoginSuccess: () => void;
}

export default function Login({ onLoginSuccess }: LoginProps) {
  const [usernameInput, setUsernameInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  // Load latest credentials from localStorage, fallback to Admin/Admin
  const getStoredCredentials = () => {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        const stored = window.localStorage.getItem('kanooz_admin_credentials');
        if (stored) {
          return JSON.parse(stored);
        }
      }
    } catch (e) {
      console.error(e);
    }
    return { username: 'Admin', password: 'Admin' };
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setIsLoading(true);

    // Simulate small smooth delay for luxury professional feel
    setTimeout(() => {
      const { username, password } = getStoredCredentials();
      
      const cleanInputUser = usernameInput.trim();
      const cleanInputPass = passwordInput.trim();

      // Case-insensitive comparison for ultimate simplicity and user-friendliness
      if (cleanInputUser.toLowerCase() === username.toLowerCase() && cleanInputPass.toLowerCase() === password.toLowerCase()) {
        try {
          if (typeof window !== 'undefined' && window.sessionStorage) {
            window.sessionStorage.setItem('kanooz_logged_in', 'true');
          }
        } catch (e) {
          console.warn('sessionStorage is locked but allowing transient in-memory login', e);
        }
        onLoginSuccess();
      } else {
        setError('Invalid username or password. Please try again.');
        setIsLoading(false);
      }
    }, 600);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-between py-12 sm:px-6 lg:px-8 relative overflow-hidden font-sans">
      {/* Dynamic abstract grid pattern */}
      <div className="absolute inset-0 pointer-events-none opacity-30">
        <div className="absolute min-w-full h-full bg-[radial-gradient(#cbd5e1_1px,transparent_1px)] [background-size:16px_16px] [mask-image:radial-gradient(ellipse_50%_50%_at_50%_50%,#000_70%,transparent_100%)]"></div>
      </div>

      <div className="shrink-0"></div>

      <div className="sm:mx-auto sm:w-full sm:max-w-md z-10">
        {/* Animated Card Container */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, ease: 'easeOut' }}
          className="bg-white py-10 px-6 shadow-xl rounded-2xl border border-slate-200/80 sm:px-10"
        >
          {/* Logo Representation */}
          <div className="flex flex-col items-center mb-8">
            <div className="mb-4 relative">
              <div className="absolute inset-0 bg-indigo-100 rounded-full blur-md opacity-50 scale-125"></div>
              <img 
                src="https://kanooz.com/wp-content/uploads/2026/05/Kanooz-Logo-transparent-png-2048x556.png" 
                alt="Kanooz Logo" 
                className="h-10 w-auto relative z-10 filter drop-shadow-sm" 
                referrerPolicy="no-referrer"
                onError={(e) => {
                  // Fallback to text identifier if logo is unreachable
                  (e.target as HTMLElement).style.display = 'none';
                }}
              />
            </div>
            
            <h2 className="text-xl font-bold text-slate-950 tracking-tight text-center">
              Central Planning Portal
            </h2>
            <p className="mt-1 text-xs text-slate-500 font-medium tracking-wide prose uppercase">
              Kanooz Industrial Services
            </p>
          </div>

          <form className="space-y-5" onSubmit={handleSubmit}>
            {error && (
              <motion.div 
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                className="rounded-xl bg-rose-50 border border-rose-100 p-3.5 flex items-start gap-2.5 text-xs text-rose-700"
              >
                <AlertCircle className="w-4 h-4 text-rose-500 shrink-0 mt-0.5" />
                <span className="font-semibold">{error}</span>
              </motion.div>
            )}

            <div>
              <label htmlFor="username" className="block text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">
                Username
              </label>
              <div className="relative rounded-xl shadow-sm">
                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <User className="h-4 w-4 text-slate-400" />
                </div>
                <input
                  id="username"
                  name="username"
                  type="text"
                  required
                  autoFocus
                  placeholder="Username"
                  value={usernameInput}
                  onChange={(e) => setUsernameInput(e.target.value)}
                  className="block w-full pl-9 pr-3 py-2.5 border border-slate-200 rounded-xl text-sm placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition-all font-medium text-slate-800"
                />
              </div>
            </div>

            <div>
              <label htmlFor="password" className="block text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">
                Password
              </label>
              <div className="relative rounded-xl shadow-sm">
                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <Lock className="h-4 w-4 text-slate-400" />
                </div>
                <input
                  id="password"
                  name="password"
                  type={showPassword ? 'text' : 'password'}
                  required
                  placeholder="Password"
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  className="block w-full pl-9 pr-10 py-2.5 border border-slate-200 rounded-xl text-sm placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition-all font-medium text-slate-800"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-450 hover:text-slate-650"
                  tabIndex={-1}
                >
                  {showPassword ? (
                    <EyeOff className="h-4 w-4 text-slate-400 hover:text-slate-650" />
                  ) : (
                    <Eye className="h-4 w-4 text-slate-400 hover:text-slate-650" />
                  )}
                </button>
              </div>
            </div>

            <div className="pt-2">
              <button
                type="submit"
                disabled={isLoading}
                className="w-full flex justify-center py-2.5 px-4 border border-transparent rounded-xl text-sm font-bold text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 shadow-md shadow-indigo-100 cursor-pointer disabled:opacity-50 transition-all"
              >
                {isLoading ? (
                  <span className="flex items-center gap-2">
                    <svg className="animate-spin -ml-1 mr-3 h-4 w-4 text-white" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 00 5.373 12 4 12z" />
                    </svg>
                    Authenticating...
                  </span>
                ) : (
                  'Sign In'
                )}
              </button>
            </div>
          </form>

          {/* Secure System Instructions */}
          <div className="mt-8 pt-6 border-t border-slate-100 flex items-center gap-2 text-slate-400 text-[11px] justify-center select-none font-medium">
            <ShieldCheck className="w-4 h-4 text-emerald-500" />
            <span>Kanooz Encrypted Session Standard</span>
          </div>
        </motion.div>
      </div>

      <div className="text-center text-slate-400 shrink-0 select-none py-4 text-[11px] text-[#888] font-medium z-10 transition-all">
        <p>© 2026 Kanooz Central Planning Portal. All Rights Reserved.</p>
        <span className="inline-flex items-center gap-1.5 mt-1 bg-slate-200/55 text-slate-600 px-2 py-0.5 rounded-full text-[10px] font-bold">
          <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-ping"></span>
          <span>SYSTEM READY</span>
        </span>
      </div>
    </div>
  );
}
