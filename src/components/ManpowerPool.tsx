import { useState, useRef, ChangeEvent } from 'react';
import { Search, Filter, Trash2, Edit2, UserPlus, X, Save, Download, Upload, AlertCircle } from 'lucide-react';
import { Manpower, Craft, EmploymentType } from '../types';
import { cn } from '../lib/utils';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import { formatToExcelDate, parseExcelDate } from '../lib/dateUtils';
import { DEFAULT_CRAFTS } from '../lib/constants';

interface Props {
  manpower: Manpower[];
  setManpower: (m: Manpower[]) => void;
  isAdding?: boolean;
  onCloseAdd?: () => void;
}

const EMPLOYMENT_TYPES: EmploymentType[] = ['Direct', 'Qiwa', 'Local Hire'];

export default function ManpowerPool({ manpower, setManpower, isAdding, onCloseAdd }: Props) {
  const [searchTerm, setSearchTerm] = useState('');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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

  const filtered = manpower.filter(m => 
    m.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    m.badgeNo.toLowerCase().includes(searchTerm.toLowerCase()) ||
    m.craft.toLowerCase().includes(searchTerm.toLowerCase())
  );

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

      <div className="bg-white rounded-2xl border border-[#E5E5E5] shadow-sm overflow-hidden min-h-[600px] flex flex-col">
      <div className="p-6 border-b border-[#E5E5E5] flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#888]" />
          <input 
            type="text" 
            placeholder="Search by name, badge, or craft..." 
            className="w-full pl-10 pr-4 py-2 bg-[#F9FAFB] border border-[#E5E5E5] rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none transition-all"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <div className="flex items-center gap-2">
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleImport} 
            className="hidden" 
            accept=".xlsx,.xls" 
          />
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-indigo-700 border border-indigo-100 bg-indigo-50 rounded-lg hover:bg-indigo-100 transition-colors"
          >
            <Upload className="w-3.5 h-3.5" />
            Import Excel
          </button>
          <button 
            onClick={handleExport}
            className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-[#666] border border-[#E5E5E5] rounded-lg hover:bg-gray-50 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Export Excel
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-[#F9FAFB] border-b border-[#E5E5E5]">
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Worker Details</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Passport / Iqama</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Craft Info</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Type</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Contract Period</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider">Vacation</th>
              <th className="px-4 py-3 text-[10px] font-bold text-[#888] uppercase tracking-wider text-right">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#F0F0F0]">
            {filtered.map((m) => {
              const isDuplicate = badgeCounts[m.badgeNo] > 1;
              const isEditing = editingId === m.id;
              
              if (isEditing) {
                return (
                  <tr key={m.id} className="bg-indigo-50/50">
                    <td className="px-4 py-2.5">
                      <div className="space-y-1">
                        <input 
                          className="w-full px-2 py-1 text-xs border rounded" 
                          value={m.name} 
                          onChange={e => updateWorkerField({...m, name: e.target.value})}
                        />
                        <input 
                          className="w-full px-2 py-0.5 text-[10px] border rounded" 
                          value={m.badgeNo} 
                          onChange={e => updateWorkerField({...m, badgeNo: e.target.value})}
                        />
                      </div>
                    </td>
                    <td className="px-4 py-2.5">
                      <input 
                        className="w-full px-2 py-1 text-xs border rounded" 
                        value={m.passportIqama || ''} 
                        onChange={e => updateWorkerField({...m, passportIqama: e.target.value})}
                        placeholder="Passport / Iqama"
                      />
                    </td>
                    <td className="px-4 py-2.5">
                      <div className="space-y-1">
                        <input 
                          className="w-full px-2 py-1 text-xs border rounded" 
                          value={m.craft} 
                          onChange={e => updateWorkerField({...m, craft: e.target.value})}
                        />
                        <select
                          className="w-full px-2 py-1 text-[10px] border rounded bg-white font-semibold text-slate-800"
                          value={m.strength || 'Good'}
                          onChange={e => updateWorkerField({...m, strength: e.target.value as any})}
                        >
                          <option value="Excellent">Excellent</option>
                          <option value="Good">Good</option>
                          <option value="Average">Average</option>
                        </select>
                      </div>
                    </td>
                    <td className="px-4 py-2.5">
                      <select 
                        className="w-full px-2 py-1 text-xs border rounded bg-white" 
                        value={m.employmentType} 
                        onChange={e => updateWorkerField({...m, employmentType: e.target.value as EmploymentType})}
                      >
                        {EMPLOYMENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </td>
                    <td className="px-4 py-2.5">
                      <div className="space-y-1">
                        <input 
                          type="date" 
                          className="w-full px-2 py-0.5 text-[10px] border rounded" 
                          value={m.joinDate} 
                          onChange={e => updateWorkerField({...m, joinDate: e.target.value})}
                        />
                        <input 
                          type="date" 
                          className="w-full px-2 py-0.5 text-[10px] border rounded" 
                          value={m.releaseDate} 
                          onChange={e => updateWorkerField({...m, releaseDate: e.target.value})}
                        />
                      </div>
                    </td>
                    <td className="px-4 py-2.5">
                      <div className="space-y-1">
                        <input 
                          type="date" 
                          className="w-full px-2 py-0.5 text-[10px] border rounded" 
                          value={m.vacationStart || ''} 
                          onChange={e => updateWorkerField({...m, vacationStart: e.target.value})}
                        />
                        <input 
                          type="date" 
                          className="w-full px-2 py-0.5 text-[10px] border rounded" 
                          value={m.vacationEnd || ''} 
                          onChange={e => updateWorkerField({...m, vacationEnd: e.target.value})}
                        />
                      </div>
                    </td>
                    <td className="px-4 py-2.5 text-right">
                      <button onClick={() => setEditingId(null)} className="p-1.5 text-indigo-600 hover:bg-white rounded-lg">
                        <Save className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  </tr>
                );
              }

              return (
                <tr key={m.id} className={cn("hover:bg-indigo-50/30 transition-colors group", isDuplicate && "bg-red-50/50")}>
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-3">
                      <div className={cn(
                        "w-8 h-8 rounded-full flex items-center justify-center text-white font-bold text-[10px]",
                        isDuplicate ? "bg-red-500" : "bg-indigo-600"
                      )}>
                        {m.name.split(' ').map(n => n[0]).join('')}
                      </div>
                      <div>
                        <div className="flex items-center gap-1.5">
                          <p className="text-xs font-bold text-[#1A1A1A]">{m.name}</p>
                          {isDuplicate && <AlertCircle className="w-3 h-3 text-red-500" title="Duplicate Badge Number" />}
                        </div>
                        <p className={cn("text-[10px] font-bold", isDuplicate ? "text-red-600" : "text-[#888]")}>{m.badgeNo}</p>
                      </div>
                    </div>
                  </td>
                  <td className="px-4 py-3">
                    <span className="text-[10px] font-semibold text-[#444]">{m.passportIqama || '--'}</span>
                  </td>
                  <td className="px-4 py-3">
                    <div className="flex flex-col gap-1">
                      <span className="px-2 py-0.5 bg-[#EEF2FF] text-indigo-700 text-[9px] font-bold uppercase rounded border border-indigo-100 w-fit">
                        {m.craft}
                      </span>
                      <span className={cn(
                        "px-1.5 py-0.5 text-[9px] font-bold rounded border w-fit uppercase tracking-wider",
                        (m.strength || 'Good') === 'Excellent' ? 'bg-purple-50 text-purple-700 border-purple-200' :
                        (m.strength || 'Good') === 'Average' ? 'bg-amber-50 text-amber-700 border-amber-205' :
                        'bg-emerald-50 text-emerald-700 border-emerald-200'
                      )}>
                        {m.strength || 'Good'}
                      </span>
                    </div>
                  </td>
                  <td className="px-4 py-3">
                    <span className="text-[10px] font-semibold text-[#666]">{m.employmentType}</span>
                  </td>
                  <td className="px-4 py-3">
                    <div className="text-[10px] font-bold text-[#444]">
                      {dayjs(m.joinDate).format('DD MMM YY')} — {dayjs(m.releaseDate).format('DD MMM YY')}
                    </div>
                  </td>
                  <td className="px-4 py-3">
                    {m.vacationStart ? (
                      <div className="text-[9px] font-bold text-orange-600 bg-orange-50 px-1.5 py-0.5 rounded border border-orange-100">
                        {dayjs(m.vacationStart).format('DD MMM')} — {dayjs(m.vacationEnd).format('DD MMM YY')}
                      </div>
                    ) : (
                      <span className="text-[9px] text-[#AAA] font-medium italic">No vacation set</span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <div className="flex items-center justify-end gap-1 text-xs">
                      {confirmDeleteId === m.id ? (
                        <>
                          <button 
                            onClick={() => deleteWorker(m.id)}
                            className="px-2 py-0.5 bg-red-500 text-white rounded text-[9px] font-bold hover:bg-red-600 transition-colors"
                          >
                            Confirm
                          </button>
                          <button 
                            onClick={() => setConfirmDeleteId(null)}
                            className="px-2 py-0.5 bg-gray-200 text-gray-700 rounded text-[9px] font-bold hover:bg-gray-300 transition-colors"
                          >
                            Cancel
                          </button>
                        </>
                      ) : (
                        <div className="flex items-center gap-0.5">
                          <button onClick={() => setEditingId(m.id)} className="p-1.5 text-[#888] hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all" title="Edit">
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button 
                            onClick={() => setConfirmDeleteId(m.id)}
                            className="p-1.5 text-[#888] hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
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
        
        {filtered.length === 0 && (
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
