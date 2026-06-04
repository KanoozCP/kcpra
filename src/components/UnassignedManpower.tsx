import { useState, useMemo } from 'react';
import { 
  UserMinus, 
  Search, 
  Download, 
  ShieldCheck, 
  CheckCircle,
  HelpCircle,
  AlertOctagon,
  Briefcase
} from 'lucide-react';
import { Manpower, Assignment } from '../types';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';

interface Props {
  manpower: Manpower[];
  assignments: Assignment[];
}

export default function UnassignedManpower({ manpower, assignments }: Props) {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState<'name' | 'badgeNo' | 'craft' | 'joinDate'>('name');

  // Filter manpower to find workers who have NO assignments at all
  const unassignedWorkers = useMemo(() => {
    const assignedIds = new Set(assignments.map(a => a.workerId));
    return manpower.filter(m => !assignedIds.has(m.id));
  }, [manpower, assignments]);

  // Search filter
  const filteredWorkers = useMemo(() => {
    const query = searchTerm.toLowerCase().trim();
    let result = unassignedWorkers;

    if (query) {
      result = result.filter(m => 
        m.name.toLowerCase().includes(query) ||
        m.badgeNo.toLowerCase().includes(query) ||
        (m.passportIqama && m.passportIqama.toLowerCase().includes(query)) ||
        m.craft.toLowerCase().includes(query)
      );
    }

    // Sort
    return [...result].sort((a, b) => {
      if (sortBy === 'name') return a.name.localeCompare(b.name);
      if (sortBy === 'badgeNo') return a.badgeNo.localeCompare(b.badgeNo);
      if (sortBy === 'craft') return a.craft.localeCompare(b.craft);
      if (sortBy === 'joinDate') return dayjs(a.joinDate).valueOf() - dayjs(b.joinDate).valueOf();
      return 0;
    });
  }, [unassignedWorkers, searchTerm, sortBy]);

  const exportToExcel = () => {
    const data = filteredWorkers.map(m => ({
      'Badge Number': m.badgeNo,
      'Name': m.name,
      'Passport / Iqama': m.passportIqama || '',
      'Craft': m.craft,
      'Employment Type': m.employmentType,
      'Skill/Strength Rating': m.strength || 'Good',
      'Joining Date': m.joinDate ? dayjs(m.joinDate).format('DD-MMM-YY') : '',
      'Release Date': m.releaseDate ? dayjs(m.releaseDate).format('DD-MMM-YY') : ''
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Not Assigned Manpower");

    const wscols = [
      { wch: 15 }, // Badge
      { wch: 25 }, // Name
      { wch: 18 }, // Iqama
      { wch: 20 }, // Craft
      { wch: 15 }, // Type
      { wch: 15 }, // strength
      { wch: 15 }, // joinDate
      { wch: 15 }  // releaseDate
    ];
    ws['!cols'] = wscols;

    XLSX.writeFile(wb, `Kanooz_Unassigned_Manpower_${dayjs().format('DD-MMM-YY')}.xlsx`);
  };

  return (
    <div className="space-y-6 font-sans">
      <div className="bg-white rounded-2xl border border-[#E5E5E5] p-5 shadow-sm flex flex-col gap-4">
        {/* Row 1: Title & Subtitle */}
        <div>
          <h3 className="text-base font-bold text-[#1A1A1A]">Not Assigned Manpower</h3>
          <p className="text-xs text-gray-500 mt-0.5">Personnel from the manpower pool who currently have no active project assignments or role mobilizations</p>
        </div>
        
        {/* Row 2: Search, Sort and Buttons */}
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 border-t border-slate-100 pt-3">
          <div className="relative flex-1 min-w-[200px] max-w-xs">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Search by name, badge, craft or iqama..." 
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
                <option value="joinDate">Join Date</option>
              </select>
            </div>
            
            <button 
              onClick={exportToExcel}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-650 border border-slate-200 rounded-lg bg-white hover:bg-slate-50 transition-colors shrink-0"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              Export Excel
            </button>
          </div>
        </div>
      </div>

      {/* KPI Stats Block */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5">
          <div className="p-3 bg-indigo-55 bg-indigo-50 rounded-xl shrink-0">
            <UserMinus className="text-indigo-600 w-5 h-5" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Unassigned Crew Size</p>
            <h4 className="text-xl font-bold text-indigo-900 mt-0.5">{unassignedWorkers.length} Workers Available</h4>
          </div>
        </div>

        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5">
          <div className="p-3 bg-teal-50 rounded-xl shrink-0">
            <ShieldCheck className="text-teal-600 w-5 h-5" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Standby Craft Categories</p>
            <h4 className="text-xl font-bold text-teal-800 mt-0.5">
              {new Set(unassignedWorkers.map(w => w.craft)).size} Unique Trades
            </h4>
          </div>
        </div>

        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5 md:col-span-2 lg:col-span-1">
          <div className="p-3 bg-amber-50 rounded-xl shrink-0">
            <Briefcase className="text-amber-600 w-5 h-5" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Total Pool Size</p>
            <h4 className="text-xl font-bold text-slate-800 mt-0.5">
              {manpower.length} Registered Crew members
            </h4>
          </div>
        </div>
      </div>

      {/* Table Container */}
      <div className="bg-white rounded-2xl border border-[#E5E5E5] shadow-sm overflow-hidden min-h-[400px] flex flex-col">
        {filteredWorkers.length === 0 ? (
          <div className="flex-1 flex flex-col items-center justify-center p-8 text-center bg-slate-50/40">
            <div className="p-4 bg-slate-100 rounded-full mb-3">
              <HelpCircle className="w-8 h-8 text-slate-400" />
            </div>
            <h3 className="text-sm font-bold text-[#1A1A1A]">No unassigned workers found</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-sm">All personnel are currently assigned to active project phases, or match criteria filters.</p>
          </div>
        ) : (
          <div className="flex-1 overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-[#F9FAFB] border-b border-[#E5E5E5]">
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Badge & Name</th>
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Passport / Iqama</th>
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Craft Trade</th>
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Type</th>
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider text-center">Skill Rating</th>
                  <th className="px-5 py-3 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Join Date</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#F2F2F2]">
                {filteredWorkers.map(m => (
                  <tr key={m.id} className="hover:bg-slate-50/60 transition-colors">
                    <td className="px-5 py-3.5">
                      <div className="flex items-center gap-2.5">
                        <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[10px] font-bold tracking-tight">#{m.badgeNo}</span>
                        <div className="text-xs font-bold text-[#1A1A1A]">{m.name}</div>
                      </div>
                    </td>
                    <td className="px-5 py-3.5 text-xs text-slate-600 font-medium">
                      {m.passportIqama || <span className="text-slate-350 italic text-[11px]">—</span>}
                    </td>
                    <td className="px-5 py-3.5">
                      <span className="bg-indigo-50 text-indigo-700 border border-indigo-100 rounded px-2 py-0.5 text-[10px] font-bold tracking-tight">
                        {m.craft}
                      </span>
                    </td>
                    <td className="px-5 py-3.5 text-xs text-slate-600 font-medium">
                      {m.employmentType}
                    </td>
                    <td className="px-5 py-3.5 text-center">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-extrabold shadow-3xs border ${
                        m.strength === 'Excellent' ? 'bg-emerald-50 text-emerald-700 border-emerald-150' :
                        m.strength === 'Average' ? 'bg-amber-50 text-amber-700 border-amber-150' :
                        'bg-blue-50 text-blue-700 border-blue-150'
                      }`}>
                        {m.strength || 'Good'}
                      </span>
                    </td>
                    <td className="px-5 py-3.5 text-xs text-slate-500 font-semibold font-mono">
                      {m.joinDate ? dayjs(m.joinDate).format('DD-MMM-YY') : 'N/A'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
