import { useState, useRef, ChangeEvent } from 'react';
import { Search, Filter, Trash2, Edit2, UserPlus, X, Save, Download, Upload, AlertCircle, Cloud, RefreshCw } from 'lucide-react';
import { Manpower, Craft, EmploymentType } from '../types';
import { cn } from '../lib/utils';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import { formatToExcelDate, parseExcelDate } from '../lib/dateUtils';
import { DEFAULT_CRAFTS } from '../lib/constants';
import { listDriveFiles, downloadFileFromDrive, uploadCustomFileToDrive } from '../lib/googleDriveService';
import { User as FirebaseUser } from 'firebase/auth';

interface Props {
  manpower: Manpower[];
  setManpower: (m: Manpower[]) => void;
  isAdding?: boolean;
  onCloseAdd?: () => void;
  googleUser?: FirebaseUser | null;
  onGoogleConnect?: () => Promise<void>;
}

const EMPLOYMENT_TYPES: EmploymentType[] = ['Direct', 'Qiwa', 'Local Hire'];

export default function ManpowerPool({ manpower, setManpower, isAdding, onCloseAdd, googleUser, onGoogleConnect }: Props) {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState<'name' | 'badgeNo' | 'craft' | 'joinDate'>('name');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [originalWorkerState, setOriginalWorkerState] = useState<Manpower | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleStartEdit = (worker: Manpower) => {
    setEditingId(worker.id);
    setOriginalWorkerState({ ...worker });
  };

  const handleCancelEdit = () => {
    if (originalWorkerState) {
      setManpower(manpower.map(m => m.id === originalWorkerState.id ? originalWorkerState : m));
    }
    setEditingId(null);
    setOriginalWorkerState(null);
  };

  const handleSaveEdit = () => {
    setEditingId(null);
    setOriginalWorkerState(null);
  };

  // Google Drive state hooks
  const [isDriveModalOpen, setIsDriveModalOpen] = useState(false);
  const [driveFiles, setDriveFiles] = useState<any[]>([]);
  const [isFetchLoading, setIsFetchLoading] = useState(false);
  const [driveSearchQuery, setDriveSearchQuery] = useState('');
  const [selectedDriveFileId, setSelectedDriveFileId] = useState<string | null>(null);
  const [isImportLoading, setIsImportLoading] = useState(false);

  const handleOpenDriveImport = async () => {
    setIsDriveModalOpen(true);
    if (googleUser) {
      await fetchDriveFiles();
    }
  };

  const fetchDriveFiles = async () => {
    setIsFetchLoading(true);
    try {
      const files = await listDriveFiles();
      setDriveFiles(files);
    } catch (err: any) {
      console.error(err);
      alert(`Google Drive error: ${err.message || err}`);
    } finally {
      setIsFetchLoading(false);
    }
  };

  const handleImportFromSelectedDriveFile = async (fileId: string, fileName: string) => {
    setIsImportLoading(true);
    try {
      const buffer = await downloadFileFromDrive(fileId);
      
      if (fileName.toLowerCase().endsWith('.json')) {
        // Handle parsing of JSON file backup
        const textDecoder = new TextDecoder('utf-8');
        const contents = textDecoder.decode(buffer);
        const parsed = JSON.parse(contents);
        
        let personnel: Manpower[] = [];
        if (parsed.manpower && Array.isArray(parsed.manpower)) {
          personnel = parsed.manpower;
        } else if (Array.isArray(parsed)) {
          personnel = parsed;
        } else {
          throw new Error("No manpower database array found. Verify the backup scheme.");
        }
        
        if (personnel.length === 0) {
          alert('Import warning: Selected JSON file has no manpower records inside.');
          return;
        }

        // Merge: overwrite existing badges
        const map = new Map(manpower.map(p => [p.badgeNo.toUpperCase().trim(), p]));
        personnel.forEach(item => {
          if (item.name && item.badgeNo) {
            map.set(item.badgeNo.toUpperCase().trim(), {
              ...item,
              id: item.id || `W-${item.badgeNo.toUpperCase().trim()}`
            });
          }
        });
        setManpower(Array.from(map.values()));
        alert(`Successfully imported/restored ${personnel.length} personnel records from Google Drive backup.`);
        setIsDriveModalOpen(false);

      } else {
        // Handle sheet array buffer
        const dataArr = new Uint8Array(buffer);
        const wb = XLSX.read(dataArr, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        const imported: Manpower[] = data.map((row: any) => {
          const rawBadge = String(row['Badge Number'] || row['badgeNo'] || '').trim();
          const cleanBadgeToken = rawBadge.toUpperCase();
          const rawStrength = String(row['Strength/Skill'] || row['Strength'] || row['strength'] || 'Good').trim();
          let parsedStrength: 'Good' | 'Average' | 'Excellent' = 'Good';
          if (/exc/i.test(rawStrength)) parsedStrength = 'Excellent';
          else if (/avg|aver/i.test(rawStrength)) parsedStrength = 'Average';
          
          return {
            id: `W-${cleanBadgeToken}`,
            name: row['Name'] || row['name'] || '',
            badgeNo: rawBadge,
            craft: row['Craft'] || row['craft'] || '',
            strength: parsedStrength,
            passportIqama: String(row['Passport / Iqama'] || row['passportIqama'] || ''),
            employmentType: (row['Employment Type'] || row['employmentType'] || 'Direct') as EmploymentType,
            joinDate: parseExcelDate(row['Joining Date'] || row['joinDate']) || dayjs().format('YYYY-MM-DD'),
            releaseDate: parseExcelDate(row['Contract End Date'] || row['releaseDate']) || dayjs().add(1, 'year').format('YYYY-MM-DD'),
            vacationStart: parseExcelDate(row['Vacation Start'] || row['vacationStart']) || '',
            vacationEnd: parseExcelDate(row['Vacation Finish'] || row['vacationEnd']) || ''
          };
        }).filter(w => w.name && w.badgeNo);

        if (imported.length === 0) {
          alert('No valid personnel records found in sheet. Please verify row columns (Name, Badge Number).');
          return;
        }

        // Merge: overwrite existing badges
        const map = new Map(manpower.map(p => [p.badgeNo.toUpperCase().trim(), p]));
        imported.forEach(item => {
          map.set(item.badgeNo.toUpperCase().trim(), item);
        });
        setManpower(Array.from(map.values()));
        alert(`Successfully imported/synchronized ${imported.length} personnel records from Google Drive file: ${fileName}`);
        setIsDriveModalOpen(false);
      }
    } catch (err: any) {
      console.error(err);
      alert(`Could not import selected Google Drive file: ${err.message || err}`);
    } finally {
      setIsImportLoading(false);
    }
  };

  const [newWorker, setNewWorker] = useState<Partial<Manpower>>({
    name: '',
    badgeNo: '',
    craft: '',
    passportIqama: '',
    joinDate: dayjs().format('YYYY-MM-DD'),
    releaseDate: dayjs().add(1, 'year').format('YYYY-MM-DD'),
    employmentType: 'Direct',
    vacationStart: '',
    vacationEnd: '',
    strength: 'Good'
  });

  const badgeCounts = manpower.reduce((acc, m) => {
    acc[m.badgeNo] = (acc[m.badgeNo] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const handleCreate = () => {
    if (!newWorker.name || !newWorker.badgeNo || !newWorker.craft) {
      alert('Please fill all required fields');
      return;
    }
    const cleanBadge = String(newWorker.badgeNo).trim();
    const tokenBadge = cleanBadge.toUpperCase();
    if (manpower.some(m => m.badgeNo.toUpperCase().trim() === tokenBadge)) {
      alert(`Error: A worker with Badge Number "${cleanBadge}" already exists in the manpower pool.`);
      return;
    }

    const worker: Manpower = {
      id: `W-${tokenBadge}`,
      name: newWorker.name!,
      badgeNo: cleanBadge,
      craft: newWorker.craft!,
      passportIqama: newWorker.passportIqama || '',
      joinDate: newWorker.joinDate!,
      releaseDate: newWorker.releaseDate!,
      employmentType: (newWorker.employmentType as EmploymentType) || 'Direct',
      vacationStart: newWorker.vacationStart,
      vacationEnd: newWorker.vacationEnd,
      strength: (newWorker.strength as any) || 'Good',
    };
    setManpower([...manpower, worker]);
    setNewWorker({
      name: '',
      badgeNo: '',
      craft: '',
      passportIqama: '',
      joinDate: dayjs().format('YYYY-MM-DD'),
      releaseDate: dayjs().add(1, 'year').format('YYYY-MM-DD'),
      employmentType: 'Direct',
      vacationStart: '',
      vacationEnd: '',
      strength: 'Good'
    });
    // Do not close automatically to allow multiple entries
  };

  const updateWorkerField = (worker: Manpower) => {
    setManpower(manpower.map(m => m.id === worker.id ? worker : m));
  };

  const deleteWorker = (id: string) => {
    setManpower(manpower.filter(m => m.id !== id));
    setConfirmDeleteId(null);
  };

  const handleExport = () => {
    const ws = XLSX.utils.json_to_sheet(manpower.map(m => ({
      'Name': m.name,
      'Badge Number': m.badgeNo,
      'Passport / Iqama': m.passportIqama || '',
      'Craft': m.craft,
      'Strength/Skill': m.strength || 'Good',
      'Employment Type': m.employmentType,
      'Joining Date': formatToExcelDate(m.joinDate),
      'Contract End Date': formatToExcelDate(m.releaseDate),
      'Vacation Start': formatToExcelDate(m.vacationStart || ''),
      'Vacation Finish': formatToExcelDate(m.vacationEnd || '')
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Manpower");
    XLSX.writeFile(wb, `Kanooz_Manpower_Pool_${formatToExcelDate(new Date())}.xlsx`);
  };

  const handleExportToDrive = async () => {
    if (!googleUser) {
      alert("Please connect Google Drive first (available in the import modal or Reports tab) to export.");
      setIsDriveModalOpen(true);
      return;
    }

    const ws = XLSX.utils.json_to_sheet(manpower.map(m => ({
      'Name': m.name,
      'Badge Number': m.badgeNo,
      'Passport / Iqama': m.passportIqama || '',
      'Craft': m.craft,
      'Strength/Skill': m.strength || 'Good',
      'Employment Type': m.employmentType,
      'Joining Date': formatToExcelDate(m.joinDate),
      'Contract End Date': formatToExcelDate(m.releaseDate),
      'Vacation Start': formatToExcelDate(m.vacationStart || ''),
      'Vacation Finish': formatToExcelDate(m.vacationEnd || '')
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Manpower");
    
    try {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const fileName = `Kanooz_Manpower_Pool_${formatToExcelDate(new Date())}.xlsx`;
      const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      
      const fileBuffer = new Uint8Array(wbout);
      const result = await uploadCustomFileToDrive(fileName, fileBuffer, mimeType);
      alert(`🎉 Exported successfully to Google Drive!\nFile saved as: ${result.name}`);
    } catch (err: any) {
      console.error(err);
      alert(`Could not export to Google Drive: ${err.message || err}`);
    }
  };

  const handleImport = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const bstr = event.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        const imported: Manpower[] = data.map((row: any) => {
          const rawBadge = String(row['Badge Number'] || row['badgeNo'] || '').trim();
          const cleanBadgeToken = rawBadge.toUpperCase();
          const rawStrength = String(row['Strength/Skill'] || row['Strength'] || row['strength'] || 'Good').trim();
          let parsedStrength: 'Good' | 'Average' | 'Excellent' = 'Good';
          if (/exc/i.test(rawStrength)) parsedStrength = 'Excellent';
          else if (/avg|aver/i.test(rawStrength)) parsedStrength = 'Average';
          
          return {
            id: `W-${cleanBadgeToken}`,
            name: row['Name'] || row['name'] || '',
            badgeNo: rawBadge,
            craft: row['Craft'] || row['craft'] || '',
            strength: parsedStrength,
            passportIqama: String(row['Passport / Iqama'] || row['passportIqama'] || ''),
            employmentType: (row['Employment Type'] || row['employmentType'] || 'Direct') as EmploymentType,
            joinDate: parseExcelDate(row['Joining Date'] || row['joinDate']) || dayjs().format('YYYY-MM-DD'),
            releaseDate: parseExcelDate(row['Contract End Date'] || row['releaseDate']) || dayjs().add(1, 'year').format('YYYY-MM-DD'),
            vacationStart: parseExcelDate(row['Vacation Start'] || row['vacationStart']) || '',
            vacationEnd: parseExcelDate(row['Vacation Finish'] || row['vacationEnd']) || ''
          };
        }).filter(w => w.name && w.badgeNo);

        if (imported.length === 0) {
          alert('No valid personnel records found in sheet. Please verify row columns (Name, Badge Number).');
          return;
        }

        // Merge: overwrite existing badges so earlier assignments remain beautifully intact
        const map = new Map(manpower.map(p => [p.badgeNo.toUpperCase().trim(), p]));
        imported.forEach(item => {
          map.set(item.badgeNo.toUpperCase().trim(), item);
        });
        setManpower(Array.from(map.values()));

        if (fileInputRef.current) fileInputRef.current.value = '';
        alert(`Successfully imported/synchronized ${imported.length} personnel records.`);
      } catch (err) {
        console.error(err);
        alert('Failed to parse Excel sheet. Ensure correct headers (Name, Badge Number, Craft).');
      }
    };
    reader.readAsBinaryString(file);
  };

  const filteredAndSorted = (() => {
    const list = manpower.filter(m => 
      m.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
      m.badgeNo.toLowerCase().includes(searchTerm.toLowerCase()) ||
      m.craft.toLowerCase().includes(searchTerm.toLowerCase())
    );

    return [...list].sort((a, b) => {
      if (sortBy === 'badgeNo') {
        return a.badgeNo.localeCompare(b.badgeNo, undefined, { numeric: true });
      } else if (sortBy === 'craft') {
        return (a.craft || '').localeCompare(b.craft || '');
      } else if (sortBy === 'joinDate') {
        return dayjs(a.joinDate).valueOf() - dayjs(b.joinDate).valueOf();
      } else {
        return a.name.localeCompare(b.name);
      }
    });
  })();

  return (
    <div className="space-y-6">
      {isAdding && (
        <div className="bg-white p-6 rounded-2xl border border-indigo-200 shadow-lg animate-in fade-in slide-in-from-top-4 duration-300">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-lg font-bold text-indigo-900 flex items-center gap-2">
              <UserPlus className="w-5 h-5" /> 
              Register New Personnel
            </h3>
            <button onClick={onCloseAdd} className="p-2 hover:bg-gray-100 rounded-full text-gray-400">
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Full Name</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. John Doe"
                value={newWorker.name}
                onChange={e => setNewWorker({...newWorker, name: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Badge Number</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. K-90210"
                value={newWorker.badgeNo}
                onChange={e => setNewWorker({...newWorker, badgeNo: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Passport / Iqama</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. 1029384756"
                value={newWorker.passportIqama}
                onChange={e => setNewWorker({...newWorker, passportIqama: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Primary Craft</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. Pipe Fitter"
                list="crafts-list"
                value={newWorker.craft}
                onChange={e => setNewWorker({...newWorker, craft: e.target.value})}
              />
              <datalist id="crafts-list">
                {DEFAULT_CRAFTS.map((craft) => (
                  <option key={craft} value={craft} />
                ))}
              </datalist>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Employment Type</label>
              <select 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none bg-white"
                value={newWorker.employmentType}
                onChange={e => setNewWorker({...newWorker, employmentType: e.target.value as EmploymentType})}
              >
                {EMPLOYMENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Skill/Strength Rating</label>
              <select 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none bg-white font-semibold text-[#1A1A1A]"
                value={newWorker.strength || 'Good'}
                onChange={e => setNewWorker({...newWorker, strength: e.target.value as any})}
              >
                <option value="Excellent">Excellent 🌟</option>
                <option value="Good">Good ✅</option>
                <option value="Average">Average ⚖️</option>
              </select>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Joining Date</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newWorker.joinDate}
                onChange={e => setNewWorker({...newWorker, joinDate: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Contract End Date</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newWorker.releaseDate}
                onChange={e => setNewWorker({...newWorker, releaseDate: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Vacation Start</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newWorker.vacationStart}
                onChange={e => setNewWorker({...newWorker, vacationStart: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Vacation Finish</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newWorker.vacationEnd}
                onChange={e => setNewWorker({...newWorker, vacationEnd: e.target.value})}
              />
            </div>
            <div className="flex items-end">
              <button 
                onClick={handleCreate}
                className="w-full bg-indigo-600 text-white py-2 rounded-xl font-bold flex items-center justify-center gap-2 hover:bg-indigo-700 transition-all shadow-md"
              >
                <Save className="w-4 h-4" />
                Save Personnel
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="bg-white rounded-2xl border border-[#E5E5E5] p-5 shadow-sm flex flex-col gap-4 mb-6">
        <div>
          <h3 className="text-base font-bold text-[#1A1A1A]">Manpower Pool</h3>
          <p className="text-xs text-gray-500 mt-0.5">Active workforce registry, qualifications, and mobilization status</p>
        </div>
        
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 border-t border-slate-100 pt-3">
          <div className="relative flex-1 min-w-[200px] max-w-xs">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Search by name, badge, or craft..." 
              className="w-full pl-10 pr-4 py-1.5 bg-[#F9FAFB] border border-[#E5E5E5] rounded-xl text-xs font-medium focus:ring-2 focus:ring-indigo-100 outline-none transition-all"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          
          <div className="flex flex-wrap items-center gap-2.5">
            <div className="flex items-center gap-1.5 shrink-0">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest block leading-none">Sort:</span>
              <select
                value={sortBy}
                onChange={(e) => setSortBy(e.target.value as any)}
                className="px-2.5 py-1.5 bg-white border border-[#E5E5E5] rounded-xl text-[11px] font-bold text-[#333] outline-none cursor-pointer hover:bg-slate-50"
              >
                <option value="name">Name (A-Z)</option>
                <option value="badgeNo">Badge Number</option>
                <option value="craft">Craft (A-Z)</option>
                <option value="joinDate">Join Date (Asc)</option>
              </select>
            </div>
            
            <div className="flex items-center gap-1.5 shrink-0">
              <input 
                type="file" 
                ref={fileInputRef} 
                onChange={handleImport} 
                className="hidden" 
                accept=".xlsx,.xls" 
              />
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-indigo-700 border border-indigo-100 bg-indigo-50 rounded-lg hover:bg-indigo-100 transition-colors"
              >
                <Upload className="w-3.5 h-3.5" />
                Import Excel
              </button>
              <button 
                onClick={handleExport}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-650 border border-slate-200 rounded-lg bg-white hover:bg-slate-50 transition-colors"
              >
                <Download className="w-3.5 h-3.5" />
                Export Excel
              </button>
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-[#E5E5E5] shadow-sm overflow-hidden min-h-[600px] flex flex-col">


      <div className="flex-1 overflow-x-auto custom-scrollbar">
        <table className="w-full text-left border-collapse min-w-[820px]">
          <thead>
            <tr className="bg-[#F9FAFB] border-b border-[#E5E5E5]">
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[140px] max-w-[155px]">Badge & Name</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[100px] max-w-[110px]">Passport / Iqama</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[130px] max-w-[145px]">Craft Trade</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[75px] max-w-[85px]">Type</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider text-center min-w-[75px] max-w-[85px]">Skill Rating</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[115px] max-w-[125px]">Contract period</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider min-w-[115px] max-w-[125px]">Vacation period</th>
              <th className="px-3.5 py-2.5 text-[10.5px] font-bold text-gray-400 uppercase tracking-wider text-right w-[95px] min-w-[95px]">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#F2F2F2]">
            {filteredAndSorted.map((m) => {
              const isDuplicate = badgeCounts[m.badgeNo] > 1;
              const isEditing = editingId === m.id;
              
              if (isEditing) {
                return (
                  <tr key={m.id} className="bg-indigo-50/40">
                    <td className="px-3.5 py-2">
                      <div className="space-y-1 max-w-[140px]">
                        <input 
                          className="w-full px-2 py-0.5 text-xs border border-slate-200 rounded-lg focus:ring-1 focus:ring-indigo-500 bg-white font-medium outline-hidden" 
                          value={m.name} 
                          onChange={e => updateWorkerField({...m, name: e.target.value})}
                          placeholder="Worker Name"
                        />
                        <input 
                          className="w-full px-2 py-0.5 text-[11px] border border-slate-200 rounded-lg focus:ring-1 focus:ring-indigo-500 bg-white font-semibold font-mono outline-hidden" 
                          value={m.badgeNo} 
                          onChange={e => updateWorkerField({...m, badgeNo: e.target.value})}
                          placeholder="Badge No"
                        />
                      </div>
                    </td>
                    <td className="px-3.5 py-2">
                      <input 
                        className="w-full px-2 py-0.5 text-xs border border-slate-200 rounded-lg focus:ring-1 focus:ring-indigo-500 bg-white font-medium outline-hidden" 
                        value={m.passportIqama || ''} 
                        onChange={e => updateWorkerField({...m, passportIqama: e.target.value})}
                        placeholder="Passport/Iqama"
                      />
                    </td>
                    <td className="px-3.5 py-2">
                      <input 
                        className="w-full px-2 py-0.5 text-xs border border-slate-200 rounded-lg focus:ring-1 focus:ring-indigo-500 bg-white font-bold outline-hidden" 
                        value={m.craft} 
                        onChange={e => updateWorkerField({...m, craft: e.target.value})}
                        placeholder="Craft Trade"
                      />
                    </td>
                    <td className="px-3.5 py-2">
                      <select 
                        className="w-full px-1.5 py-0.5 text-xs border border-slate-200 rounded-lg bg-white font-semibold text-slate-700 outline-hidden" 
                        value={m.employmentType} 
                        onChange={e => updateWorkerField({...m, employmentType: e.target.value as EmploymentType})}
                      >
                        {EMPLOYMENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </td>
                    <td className="px-3.5 py-2">
                      <select
                        className="w-full px-1.5 py-0.5 text-xs border border-slate-200 rounded-lg bg-white font-extrabold text-slate-700 outline-hidden"
                        value={m.strength || 'Good'}
                        onChange={e => updateWorkerField({...m, strength: e.target.value as any})}
                      >
                        <option value="Excellent">Excellent</option>
                        <option value="Good">Good</option>
                        <option value="Average">Average</option>
                      </select>
                    </td>
                    <td className="px-3.5 py-2">
                      <div className="space-y-1">
                        <div className="flex items-center gap-1">
                          <span className="text-[9px] font-bold text-gray-400 w-8 uppercase shrink-0">Start:</span>
                          <input 
                            type="date" 
                            className="px-1 py-0.5 text-[10px] border border-slate-200 rounded-md font-semibold font-mono outline-hidden w-full" 
                            value={m.joinDate} 
                            onChange={e => updateWorkerField({...m, joinDate: e.target.value})}
                          />
                        </div>
                        <div className="flex items-center gap-1">
                          <span className="text-[9px] font-bold text-gray-400 w-8 uppercase shrink-0">End:</span>
                          <input 
                            type="date" 
                            className="px-1 py-0.5 text-[10px] border border-slate-200 rounded-md font-semibold font-mono outline-hidden w-full" 
                            value={m.releaseDate} 
                            onChange={e => updateWorkerField({...m, releaseDate: e.target.value})}
                          />
                        </div>
                      </div>
                    </td>
                    <td className="px-3.5 py-2">
                      <div className="space-y-1">
                        <div className="flex items-center gap-1">
                          <span className="text-[9px] font-bold text-gray-400 w-8 uppercase shrink-0">Start:</span>
                          <input 
                            type="date" 
                            className="px-1 py-0.5 text-[10px] border border-slate-200 rounded-md font-semibold font-mono outline-hidden w-full" 
                            value={m.vacationStart || ''} 
                            onChange={e => updateWorkerField({...m, vacationStart: e.target.value})}
                          />
                        </div>
                        <div className="flex items-center gap-1">
                          <span className="text-[9px] font-bold text-gray-400 w-8 uppercase shrink-0">End:</span>
                          <input 
                            type="date" 
                            className="px-1 py-0.5 text-[10px] border border-slate-200 rounded-md font-semibold font-mono outline-hidden w-full" 
                            value={m.vacationEnd || ''} 
                            onChange={e => updateWorkerField({...m, vacationEnd: e.target.value})}
                          />
                        </div>
                      </div>
                    </td>
                    <td className="px-3.5 py-2 text-right w-[95px] min-w-[95px]">
                      <div className="flex flex-col gap-1 w-full">
                        <button 
                          onClick={handleSaveEdit} 
                          className="px-1 py-0.5 bg-indigo-600 text-white rounded text-[10px] font-bold hover:bg-indigo-700 transition-colors shadow-3xs flex items-center justify-center gap-1 cursor-pointer border border-indigo-200 w-full shrink-0"
                          title="Save changes"
                        >
                          <Save className="w-3 h-3" />
                          Save
                        </button>
                        <button 
                          onClick={handleCancelEdit} 
                          className="px-1 py-0.5 bg-white border border-slate-200 text-slate-600 rounded text-[10px] font-bold hover:bg-slate-50 transition-colors cursor-pointer shadow-3xs w-full shrink-0"
                          title="Discard changes"
                        >
                          Cancel
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              }

              return (
                <tr key={m.id} className={cn("hover:bg-slate-50/60 transition-colors group", isDuplicate && "bg-red-50/30")}>
                  <td className="px-3.5 py-2.5">
                    <div className="flex flex-col gap-0.5 max-w-[140px]">
                      <div className="flex items-center gap-1.5">
                        <span className={cn(
                          "px-1.5 py-0.5 rounded text-[9px] font-extrabold tracking-tight",
                          isDuplicate ? "bg-red-100 text-red-700 border border-red-200" : "bg-slate-100 text-slate-700"
                        )}>
                          #{m.badgeNo}
                        </span>
                        {isDuplicate && <AlertCircle className="w-3.5 h-3.5 text-red-500 animate-pulse shrink-0" title="Duplicate Badge Number" />}
                      </div>
                      <div className="text-xs font-bold text-[#1A1A1A] leading-tight break-words whitespace-normal">
                        {m.name}
                      </div>
                    </div>
                  </td>
                  <td className="px-3.5 py-2.5 text-xs text-slate-600 font-medium break-all whitespace-normal max-w-[105px]">
                    {m.passportIqama || <span className="text-slate-300 italic text-[11px]">—</span>}
                  </td>
                  <td className="px-3.5 py-2.5">
                    <span className="bg-indigo-50 text-indigo-700 border border-indigo-150 rounded px-2 py-0.5 text-[10px] font-bold tracking-tight block w-fit whitespace-normal break-words max-w-[130px] leading-tight">
                      {m.craft}
                    </span>
                  </td>
                  <td className="px-3.5 py-2.5 text-xs text-slate-600 font-semibold">
                    {m.employmentType}
                  </td>
                  <td className="px-3.5 py-2.5 text-center">
                    <span className={`px-2 py-0.5 rounded text-[10px] font-extrabold shadow-3xs border block w-fit mx-auto ${
                      m.strength === 'Excellent' ? 'bg-purple-50 text-purple-700 border-purple-150' :
                      m.strength === 'Average' ? 'bg-amber-50 text-amber-700 border-amber-150' :
                      'bg-emerald-50 text-emerald-700 border-emerald-150'
                    }`}>
                      {m.strength || 'Good'}
                    </span>
                  </td>
                  <td className="px-3.5 py-2.5">
                    <div className="flex flex-col gap-0.5 font-mono max-w-[115px]">
                      <div className="flex items-center gap-1.5 text-[10px] text-slate-500 font-medium">
                        <span className="text-[9px] text-slate-400 font-bold uppercase w-8 shrink-0">Start:</span>
                        <span>{dayjs(m.joinDate).format('DD-MMM-YY')}</span>
                      </div>
                      <div className="flex items-center gap-1.5 text-[10px] text-slate-500 font-medium">
                        <span className="text-[9px] text-slate-400 font-bold uppercase w-8 shrink-0">End:</span>
                        <span>{dayjs(m.releaseDate).format('DD-MMM-YY')}</span>
                      </div>
                    </div>
                  </td>
                  <td className="px-3.5 py-2.5">
                    {m.vacationStart ? (
                      <div className="flex flex-col gap-0.5 font-mono max-w-[115px]">
                        <div className="flex items-center gap-1.5 text-[10px] text-orange-600 font-extrabold">
                          <span className="text-[9px] text-orange-400 font-bold uppercase w-8 shrink-0">Start:</span>
                          <span>{dayjs(m.vacationStart).format('DD-MMM-YY')}</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-[10px] text-orange-600 font-extrabold">
                          <span className="text-[9px] text-orange-400 font-bold uppercase w-8 shrink-0">End:</span>
                          <span>{dayjs(m.vacationEnd).format('DD-MMM-YY')}</span>
                        </div>
                      </div>
                    ) : (
                      <span className="text-[10px] text-slate-400 font-medium italic">—</span>
                    )}
                  </td>
                  <td className="px-3.5 py-2.5 text-right w-[95px] min-w-[95px]">
                    <div className="flex items-center justify-end gap-1 text-xs w-full">
                      {confirmDeleteId === m.id ? (
                        <div className="flex flex-col gap-1 items-end w-full">
                          <button 
                            onClick={() => deleteWorker(m.id)}
                            className="px-1.5 py-0.5 bg-red-600 text-white rounded text-[9px] font-bold hover:bg-red-750 transition-colors shadow-3xs border border-red-200 block w-full text-center cursor-pointer"
                          >
                            Confirm
                          </button>
                          <button 
                            onClick={() => setConfirmDeleteId(null)}
                            className="px-1.5 py-0.5 bg-white border border-slate-200 text-slate-600 rounded text-[9px] font-bold hover:bg-slate-50 transition-colors shadow-3xs block w-full text-center cursor-pointer"
                          >
                            Cancel
                          </button>
                        </div>
                      ) : (
                        <div className="flex items-center gap-0.5 justify-end w-full">
                          <button onClick={() => handleStartEdit(m)} className="p-1 px-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all cursor-pointer" title="Edit">
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button 
                            onClick={() => setConfirmDeleteId(m.id)}
                            className="p-1 px-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all cursor-pointer"
                            title="Delete"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        
        {filteredAndSorted.length === 0 && (
          <div className="p-20 text-center">
            <UserPlus className="w-12 h-12 text-[#E5E5E5] mx-auto mb-4" />
            <h4 className="text-lg font-bold text-[#1A1A1A] mb-1">No manpower records found</h4>
            <p className="text-sm text-[#888]">Start by adding workers to the pool.</p>
          </div>
        )}
      </div>
    </div>
  </div>
  );
}
