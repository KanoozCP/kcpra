import { useState, useRef, ChangeEvent } from 'react';
import { Plus, Trash2, LayoutGrid, List, MapPin, Hash, Calculator, X, Save, Layers, Download, Upload, AlertCircle, Building2, Calendar, Search, ChevronDown, ChevronRight, FolderMinus, FolderPlus, Edit2, Cloud, RefreshCw } from 'lucide-react';
import { Project, ProjectRequirement, Craft, ProjectPhase } from '../types';
import { cn } from '../lib/utils';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import { formatToExcelDate, parseExcelDate, getProjectActualStatus } from '../lib/dateUtils';
import { DEFAULT_CRAFTS } from '../lib/constants';
import { listDriveFiles, downloadFileFromDrive, uploadCustomFileToDrive } from '../lib/googleDriveService';
import { User as FirebaseUser } from 'firebase/auth';

interface Props {
  projects: Project[];
  setProjects: (p: Project[]) => void;
  isAdding?: boolean;
  onCloseAdd?: () => void;
  googleUser?: FirebaseUser | null;
  onGoogleConnect?: () => Promise<void>;
}


export default function ProjectManagement({ projects, setProjects, isAdding, onCloseAdd, googleUser, onGoogleConnect }: Props) {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState<'startDate' | 'status'>('startDate');
  const [confirmDeleteProjectId, setConfirmDeleteProjectId] = useState<string | null>(null);
  const [confirmDeleteReqId, setConfirmDeleteReqId] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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
        
        let importedProjects: Project[] = [];
        if (parsed.projects && Array.isArray(parsed.projects)) {
          importedProjects = parsed.projects;
        } else if (Array.isArray(parsed)) {
          importedProjects = parsed;
        } else {
          throw new Error("No projects database array found. Verify the backup scheme.");
        }
        
        if (importedProjects.length === 0) {
          alert('Import warning: Selected JSON file has no project records inside.');
          return;
        }

        // Merge: overwrite existing projects by code
        const map = new Map(projects.map(p => [p.code.toUpperCase().trim(), p]));
        importedProjects.forEach(item => {
          if (item.name && item.code) {
            map.set(item.code.toUpperCase().trim(), {
              ...item,
              id: item.id || `P-${item.code.toUpperCase().trim()}`,
              requirements: item.requirements || []
            });
          }
        });
        setProjects(Array.from(map.values()));
        alert(`Successfully imported/restored ${importedProjects.length} projects from Google Drive backup.`);
        setIsDriveModalOpen(false);

      } else {
        // Handle sheet array buffer
        const dataArr = new Uint8Array(buffer);
        const wb = XLSX.read(dataArr, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        const projectMap: Record<string, Project> = {};

        data.forEach((row: any) => {
          const rawCode = String(row['Project Code'] || row['code'] || '').trim();
          const cleanCodeToken = rawCode.toUpperCase();
          const name = row['Project Title'] || row['name'] || '';
          if (!rawCode || !name) return;

          if (!projectMap[cleanCodeToken]) {
            projectMap[cleanCodeToken] = {
              id: `P-${cleanCodeToken}`,
              name: name,
              code: rawCode,
              location: row['Location'] || row['location'] || 'Main Site',
              department: row['Department'] || row['department'] || '',
              startDate: parseExcelDate(row['Start Date'] || row['startDate']) || dayjs().format('YYYY-MM-DD'),
              endDate: parseExcelDate(row['Finish Date'] || row['endDate']) || dayjs().add(6, 'month').format('YYYY-MM-DD'),
              requirements: []
            };
          }

          const rawCraft = row['Required Craft'] || row['craft'];
          const rawQty = row['Required Qty'] || row['qty'];
          const qty = parseInt(String(rawQty || 0));
          
          if (rawCraft && rawCraft !== '--') {
            const reqStart = parseExcelDate(row['Requirement Start'] || row['requirementStart']) || projectMap[cleanCodeToken].startDate;
            const reqEnd = parseExcelDate(row['Requirement End'] || row['requirementEnd']) || projectMap[cleanCodeToken].endDate;
            projectMap[cleanCodeToken].requirements.push({
              id: `REQ-${Date.now()}-${Math.random().toString(16).substring(2, 10)}`,
              craft: String(rawCraft),
              phase: (row['Phase'] || row['phase'] || 'TA') as ProjectPhase,
              qty: qty,
              startDate: reqStart,
              endDate: reqEnd,
            });
          }
        });

        const imported = Object.values(projectMap);

        if (imported.length === 0) {
          alert('No valid project records found in Excel. Verify columns "Project Title", "Project Code".');
          return;
        }

        const map = new Map(projects.map(p => [p.code.toUpperCase().trim(), p]));
        imported.forEach(item => {
          map.set(item.code.toUpperCase().trim(), item);
        });
        setProjects(Array.from(map.values()));
        alert(`Successfully imported/synchronized ${imported.length} projects from Google Drive file: ${fileName}`);
        setIsDriveModalOpen(false);
      }
    } catch (err: any) {
      console.error(err);
      alert(`Could not import selected Google Drive file: ${err.message || err}`);
    } finally {
      setIsImportLoading(false);
    }
  };
  
  const [collapsedProjects, setCollapsedProjects] = useState<Record<string, boolean>>({});
  const [collapsedPhases, setCollapsedPhases] = useState<Record<string, boolean>>({});
  
  // Edit Project state
  const [editingProject, setEditingProject] = useState<Project | null>(null);

  const handleSaveEdit = () => {
    if (!editingProject) return;
    if (!editingProject.name || !editingProject.code) {
      alert('Project name and code are required.');
      return;
    }
    const cleanCode = String(editingProject.code).trim();
    const tokenCode = cleanCode.toUpperCase();
    
    // Check duplication with other projects
    const codeExists = projects.some(p => p.id !== editingProject.id && p.code.toUpperCase().trim() === tokenCode);
    if (codeExists) {
      alert(`Error: A project with Code "${cleanCode}" already exists.`);
      return;
    }

    setProjects(projects.map(p => {
      if (p.id === editingProject.id) {
        return {
          ...p,
          name: editingProject.name,
          code: cleanCode,
          location: editingProject.location,
          department: editingProject.department,
          startDate: editingProject.startDate,
          endDate: editingProject.endDate,
          status: editingProject.status
        };
      }
      return p;
    }));
    setEditingProject(null);
  };

  const toggleProject = (projId: string) => {
    setCollapsedProjects(prev => ({ ...prev, [projId]: !prev[projId] }));
  };

  const togglePhase = (projIdAndPhase: string) => {
    setCollapsedPhases(prev => ({ ...prev, [projIdAndPhase]: !prev[projIdAndPhase] }));
  };

  const expandAllProjects = () => {
    setCollapsedProjects({});
  };

  const collapseAllProjects = () => {
    const collapsed: Record<string, boolean> = {};
    projects.forEach(p => {
      collapsed[p.id] = true;
    });
    setCollapsedProjects(collapsed);
  };

  const [newProject, setNewProject] = useState<Partial<Project>>({
    name: '',
    code: '',
    location: 'Main Site',
    department: '',
    startDate: dayjs().format('YYYY-MM-DD'),
    endDate: dayjs().add(6, 'month').format('YYYY-MM-DD')
  });

  const projectCodeCounts = projects.reduce((acc, p) => {
    acc[p.code] = (acc[p.code] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const handleCreateProject = () => {
    if (!newProject.name || !newProject.code) {
      alert('Missing project name or code');
      return;
    }
    const cleanCode = String(newProject.code).trim();
    const tokenCode = cleanCode.toUpperCase();
    if (projects.some(p => p.code.toUpperCase().trim() === tokenCode)) {
      alert(`Error: A project with Code "${cleanCode}" already exists.`);
      return;
    }

    const project: Project = {
      id: `P-${tokenCode}`,
      name: newProject.name!,
      code: cleanCode,
      location: newProject.location || 'Site Area A',
      department: newProject.department || '',
      startDate: newProject.startDate || dayjs().format('YYYY-MM-DD'),
      endDate: newProject.endDate || dayjs().add(6, 'month').format('YYYY-MM-DD'),
      requirements: []
    };
    setProjects([...projects, project]);
    setNewProject({ 
      name: '', 
      code: '', 
      location: 'Main Site',
      department: '',
      startDate: dayjs().format('YYYY-MM-DD'),
      endDate: dayjs().add(6, 'month').format('YYYY-MM-DD')
    });
    // Do not close automatically to allow multiple entries
  };

  const handleExport = () => {
    const flatProjects = projects.flatMap(p => 
      p.requirements.length > 0 ? p.requirements.map(r => ({
        'Project Title': p.name,
        'Project Code': p.code,
        'Location': p.location,
        'Department': p.department,
        'Start Date': formatToExcelDate(p.startDate),
        'Finish Date': formatToExcelDate(p.endDate),
        'Required Craft': r.craft,
        'Phase': r.phase,
        'Required Qty': r.qty,
        'Requirement Start': formatToExcelDate(r.startDate),
        'Requirement End': formatToExcelDate(r.endDate)
      })) : [{
        'Project Title': p.name,
        'Project Code': p.code,
        'Location': p.location,
        'Department': p.department,
        'Start Date': formatToExcelDate(p.startDate),
        'Finish Date': formatToExcelDate(p.endDate),
        'Required Craft': '--',
        'Phase': '--',
        'Required Qty': 0,
        'Requirement Start': '--',
        'Requirement End': '--'
      }]
    );
    const ws = XLSX.utils.json_to_sheet(flatProjects);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Projects");
    XLSX.writeFile(wb, `Kanooz_Projects_Requirements_${formatToExcelDate(new Date())}.xlsx`);
  };

  const handleExportToDrive = async () => {
    if (!googleUser) {
      alert("Please connect Google Drive first (available in the import modal or Reports tab) to export.");
      setIsDriveModalOpen(true);
      return;
    }

    const flatProjects = projects.flatMap(p => 
      p.requirements.length > 0 ? p.requirements.map(r => ({
        'Project Title': p.name,
        'Project Code': p.code,
        'Location': p.location,
        'Department': p.department,
        'Start Date': formatToExcelDate(p.startDate),
        'Finish Date': formatToExcelDate(p.endDate),
        'Required Craft': r.craft,
        'Phase': r.phase,
        'Required Qty': r.qty,
        'Requirement Start': formatToExcelDate(r.startDate),
        'Requirement End': formatToExcelDate(r.endDate)
      })) : [{
        'Project Title': p.name,
        'Project Code': p.code,
        'Location': p.location,
        'Department': p.department,
        'Start Date': formatToExcelDate(p.startDate),
        'Finish Date': formatToExcelDate(p.endDate),
        'Required Craft': '--',
        'Phase': '--',
        'Required Qty': 0,
        'Requirement Start': '--',
        'Requirement End': '--'
      }]
    );
    const ws = XLSX.utils.json_to_sheet(flatProjects);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Projects");

    try {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const fileName = `Kanooz_Projects_Requirements_${formatToExcelDate(new Date())}.xlsx`;
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

        // Group by Project Code
        const projectMap: Record<string, Project> = {};

        data.forEach((row: any) => {
          const rawCode = String(row['Project Code'] || row['code'] || '').trim();
          const cleanCodeToken = rawCode.toUpperCase();
          const name = row['Project Title'] || row['name'] || '';
          if (!rawCode || !name) return;

          if (!projectMap[cleanCodeToken]) {
            projectMap[cleanCodeToken] = {
              id: `P-${cleanCodeToken}`,
              name: name,
              code: rawCode,
              location: row['Location'] || row['location'] || 'Main Site',
              department: row['Department'] || row['department'] || '',
              startDate: parseExcelDate(row['Start Date'] || row['startDate']) || dayjs().format('YYYY-MM-DD'),
              endDate: parseExcelDate(row['Finish Date'] || row['endDate']) || dayjs().add(6, 'month').format('YYYY-MM-DD'),
              requirements: []
            };
          }

          const rawCraft = row['Required Craft'] || row['craft'];
          const rawQty = row['Required Qty'] || row['qty'];
          const qty = parseInt(String(rawQty || 0));
          
          // If there's a valid craft (not the placeholder '--'), add it as a requirement
          if (rawCraft && rawCraft !== '--') {
            const reqStart = parseExcelDate(row['Requirement Start'] || row['requirementStart']) || projectMap[cleanCodeToken].startDate;
            const reqEnd = parseExcelDate(row['Requirement End'] || row['requirementEnd']) || projectMap[cleanCodeToken].endDate;
            projectMap[cleanCodeToken].requirements.push({
              id: `REQ-${Date.now()}-${Math.random().toString(16).substring(2, 10)}`,
              craft: String(rawCraft),
              phase: (row['Phase'] || row['phase'] || 'TA') as ProjectPhase,
              qty: qty,
              startDate: reqStart,
              endDate: reqEnd,
            });
          }
        });

        const imported = Object.values(projectMap);

        if (imported.length === 0) {
          alert('No valid project records found in Excel. Verify columns "Project Title", "Project Code".');
          return;
        }

        // Merge: overwrite existing project indices by code so assignments stay perfectly functional
        const map = new Map(projects.map(p => [p.code.toUpperCase().trim(), p]));
        imported.forEach(item => {
          map.set(item.code.toUpperCase().trim(), item);
        });
        setProjects(Array.from(map.values()));

        if (fileInputRef.current) fileInputRef.current.value = '';
        alert(`Successfully synchronized ${imported.length} projects with their site requisitions.`);
      } catch (err) {
        console.error(err);
        alert('Failed to parse Projects Excel. Please ensure correct headers (Project Title, Project Code).');
      }
    };
    reader.readAsBinaryString(file);
  };


  const phases: ProjectPhase[] = ['Pre-TA', 'TA', 'Post-TA'];

  const addRequirement = (projectId: string, phase: ProjectPhase = 'TA') => {
    setProjects(projects.map(p => {
      if (p.id !== projectId) return p;
      return {
        ...p,
        requirements: [...p.requirements, {
          id: `REQ-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
          craft: 'Pipe Fitter',
          phase: phase,
          qty: 1,
          startDate: p.startDate,
          endDate: p.endDate
        }]
      };
    }));
  };

  const updateReqField = (projectId: string, reqId: string, field: keyof ProjectRequirement, value: any) => {
    setProjects(projects.map(p => {
      if (p.id !== projectId) return p;
      return {
        ...p,
        requirements: p.requirements.map(r => r.id === reqId ? { ...r, [field]: value } : r)
      };
    }));
  };

  const sortedAndFilteredProjects = (() => {
    const filtered = projects.filter(p => 
      p.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
      p.code.toLowerCase().includes(searchTerm.toLowerCase())
    );
    
    if (sortBy === 'status') {
      return [...filtered].sort((a, b) => {
        const statusA = getProjectActualStatus(a);
        const statusB = getProjectActualStatus(b);
        if (statusA !== statusB) {
          return statusA.localeCompare(statusB);
        }
        return dayjs(a.startDate).valueOf() - dayjs(b.startDate).valueOf();
      });
    } else {
      return [...filtered].sort((a, b) => {
        return dayjs(a.startDate).valueOf() - dayjs(b.startDate).valueOf();
      });
    }
  })();

  return (
    <div className="space-y-6">
      {/* ... keeping registration UI as is (starts at line 137) ... */}
      {isAdding && (
        <div className="bg-white p-6 rounded-2xl border border-indigo-200 shadow-lg animate-in fade-in slide-in-from-top-4 duration-300">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-lg font-bold text-indigo-900 flex items-center gap-2">
              <Layers className="w-5 h-5" /> 
              Register New Project
            </h3>
            <button onClick={onCloseAdd} className="p-2 hover:bg-gray-100 rounded-full text-gray-400">
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Project Title</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. Shutdown 2026 Phase 1"
                value={newProject.name}
                onChange={e => setNewProject({...newProject, name: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Project Code</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. PR-882"
                value={newProject.code}
                onChange={e => setNewProject({...newProject, code: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Location</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. Ras Tanura"
                value={newProject.location}
                onChange={e => setNewProject({...newProject, location: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Department</label>
              <input 
                type="text" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                placeholder="e.g. Maintenance"
                value={newProject.department}
                onChange={e => setNewProject({...newProject, department: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Start Date</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newProject.startDate}
                onChange={e => setNewProject({...newProject, startDate: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Finish Date</label>
              <input 
                type="date" 
                className="w-full px-4 py-2 border border-gray-250 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
                value={newProject.endDate}
                onChange={e => setNewProject({...newProject, endDate: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase">Project Status</label>
              <select
                className="w-full px-4 py-2 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none bg-white font-semibold text-slate-800"
                value={newProject.status || 'Auto'}
                onChange={e => setNewProject({...newProject, status: e.target.value as any})}
              >
                <option value="Auto">Auto (Date-based)</option>
                <option value="In Progress">In Progress (Manual)</option>
                <option value="Completed">Completed (Manual)</option>
                <option value="Hold">Hold</option>
                <option value="Rescheduled">Rescheduled</option>
                <option value="Cancelled">Cancelled</option>
              </select>
            </div>
            <div className="space-y-2 lg:col-span-2 flex justify-end">
              <button 
                onClick={handleCreateProject}
                className="px-8 bg-indigo-600 text-white py-2 rounded-xl font-bold flex items-center justify-center gap-2 hover:bg-indigo-700 transition-all shadow-md mt-6"
              >
                <Plus className="w-4 h-4" />
                Initialize Project
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="flex flex-col md:flex-row items-center justify-between gap-4 mb-6">
        <div className="w-full md:w-auto flex flex-col sm:flex-row items-stretch sm:items-center gap-3 flex-1">
          <div className="relative w-full md:w-80">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Search projects..." 
              className="w-full pl-10 pr-4 py-2 bg-white border border-[#E5E5E5] rounded-xl text-sm focus:ring-2 focus:ring-indigo-100 outline-none"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest leading-none shrink-0">Sort By:</span>
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as any)}
              className="px-3 py-1.5 bg-white border border-[#E5E5E5] rounded-xl text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 cursor-pointer shadow-3xs"
            >
              <option value="startDate">Start Date (Asc)</option>
              <option value="status">Status, then Start Date (Asc)</option>
            </select>
          </div>
        </div>
        
        <div className="flex flex-wrap items-center gap-2">
          {projects.length > 0 && (
            <div className="flex items-center gap-1.5 border-r border-[#E5E5E5] pr-3 mr-1">
              <button 
                onClick={expandAllProjects}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-slate-705 border border-slate-200 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                title="Expand All Projects"
              >
                <FolderPlus className="w-3.5 h-3.5 text-slate-500" />
                Expand All
              </button>
              <button 
                onClick={collapseAllProjects}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-slate-705 border border-slate-200 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                title="Collapse All Projects"
              >
                <FolderMinus className="w-3.5 h-3.5 text-slate-500" />
                Collapse All
              </button>
            </div>
          )}

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
            onClick={handleOpenDriveImport}
            className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-indigo-700 border border-indigo-100 bg-indigo-50 rounded-lg hover:bg-indigo-100 transition-colors"
          >
            <Cloud className="w-3.5 h-3.5" />
            Import from Drive
          </button>
          <button 
            onClick={handleExport}
            className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-[#666] border border-[#E5E5E5] rounded-lg hover:bg-gray-50 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Export Excel
          </button>
          <button 
            onClick={handleExportToDrive}
            className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-[#666] border border-[#E5E5E5] rounded-lg hover:bg-gray-50 transition-colors"
          >
            <Cloud className="w-3.5 h-3.5" />
            Export to Drive
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-6">
        {sortedAndFilteredProjects.map(project => {
          const isDuplicate = projectCodeCounts[project.code] > 1;
          const isProjCollapsed = !!collapsedProjects[project.id];
          
          return (
          <div key={project.id} className={cn(
            "bg-white rounded-xl border border-[#E5E5E5] shadow-sm overflow-hidden border-l-4 transition-all hover:shadow-md",
            isDuplicate ? "border-l-red-500" : "border-l-indigo-600"
          )}>
            <div 
              className="p-4 bg-[#FAFAFB]/50 border-b border-[#E5E5E5] flex items-center justify-between cursor-pointer hover:bg-slate-50/60 transition-colors select-none"
              onClick={() => toggleProject(project.id)}
            >
              <div className="flex items-center gap-3 min-w-0">
                <div className="p-1 hover:bg-gray-250/50 rounded-lg text-slate-500 shrink-0">
                  {isProjCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                </div>
                <div className="min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <h3 className="text-base font-bold text-[#1A1A1A]">{project.name}</h3>
                    {(() => {
                      const status = getProjectActualStatus(project);
                      let badgeColor = '';
                      switch (status) {
                        case 'In Progress':
                          badgeColor = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                          break;
                        case 'Completed':
                          badgeColor = 'bg-blue-50 text-blue-700 border-blue-200';
                          break;
                        case 'Hold':
                          badgeColor = 'bg-amber-50 text-amber-700 border-amber-200';
                          break;
                        case 'Rescheduled':
                          badgeColor = 'bg-indigo-50 text-indigo-750 border-indigo-200';
                          break;
                        case 'Cancelled':
                          badgeColor = 'bg-rose-50 text-rose-700 border-rose-200';
                          break;
                      }
                      return (
                        <span className={cn(
                          "px-2 py-0.5 text-[9px] font-bold uppercase tracking-wider rounded-full border shadow-3xs",
                          badgeColor
                        )}>
                          {status}
                        </span>
                      );
                    })()}
                    {isDuplicate && <AlertCircle className="w-4 h-4 text-red-500" title="Duplicate Project Code" />}
                  </div>
                  <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-1.5">
                    <span className={cn(
                      "flex items-center gap-1 text-[10px] font-bold uppercase tracking-wide px-1.5 py-0.5 rounded",
                      isDuplicate ? "bg-red-50 text-red-650" : "text-indigo-600"
                    )}>
                      <Hash className="w-3 h-3" /> {project.code}
                    </span>
                    <span className="flex items-center gap-1 text-[10px] font-semibold text-[#666]">
                      <MapPin className="w-3 h-3 text-[#999]" /> {project.location}
                    </span>
                    <span className="flex items-center gap-1 text-[10px] font-semibold text-[#666]">
                      <Building2 className="w-3 h-3 text-[#999]" /> {project.department || '--'}
                    </span>
                    <span className="flex items-center gap-1 text-[10px] font-semibold text-[#666]">
                      <Calendar className="w-3 h-3 text-[#999]" /> {dayjs(project.startDate).format('DD MMM YY')} — {dayjs(project.endDate).format('DD MMM YY')}
                    </span>
                  </div>
                </div>
              </div>
              
              <div className="flex items-center gap-3 shrink-0" onClick={(e) => e.stopPropagation()}>
                <button 
                  onClick={() => setEditingProject(project)}
                  className="flex items-center gap-1.5 text-[10px] font-bold text-slate-700 hover:text-indigo-750 bg-slate-50 border border-slate-200 px-2.5 py-1.5 rounded-lg transition-colors"
                >
                  <Edit2 className="w-3.5 h-3.5 text-slate-500" />
                  Edit Project
                </button>

                <button 
                  onClick={() => addRequirement(project.id)}
                  className="flex items-center gap-1.5 text-[10px] font-bold text-indigo-700 hover:text-indigo-800 bg-indigo-50 px-2.5 py-1.5 rounded-lg transition-colors"
                >
                  <Calculator className="w-3.5 h-3.5" />
                  Define Requirements
                </button>
                
                {confirmDeleteProjectId === project.id ? (
                  <div className="flex items-center gap-2 animate-in fade-in zoom-in duration-200">
                    <button 
                      onClick={() => {
                        setProjects(projects.filter(p => p.id !== project.id));
                        setConfirmDeleteProjectId(null);
                      }}
                      className="px-3 py-2 bg-red-500 text-white text-[10px] font-bold rounded-lg hover:bg-red-600 transition-colors shadow-sm"
                    >
                      Confirm Delete
                    </button>
                    <button 
                      onClick={() => setConfirmDeleteProjectId(null)}
                      className="px-3 py-2 bg-gray-100 text-gray-600 text-[10px] font-bold rounded-lg hover:bg-gray-200 transition-colors"
                    >
                      Cancel
                    </button>
                  </div>
                ) : (
                  <button 
                    onClick={() => setConfirmDeleteProjectId(project.id)}
                    className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                    title="Delete Project"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            </div>

            {/* WBS Phase-wise Project Requirements container */}
            {!isProjCollapsed && (
              <div className="p-4 border-t border-[#E5E5E5] bg-[#FAFAFB]/20 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-slate-50/70 p-3 rounded-xl border border-slate-100">
                <div>
                  <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wide">Work Breakdown Structure (WBS) Requirements</h4>
                  <p className="text-[10px] text-gray-400 font-medium">Define skilled craft personnel needed for each key event phase.</p>
                </div>
                <div className="flex items-center gap-1.5 flex-wrap">
                  <button 
                    onClick={() => addRequirement(project.id, 'Pre-TA')}
                    className="px-2.5 py-1.5 text-[10px] font-extrabold text-amber-800 bg-amber-50 rounded-lg border border-amber-200 hover:bg-amber-100/70 transition-colors"
                  >
                    + Add Pre-TA
                  </button>
                  <button 
                    onClick={() => addRequirement(project.id, 'TA')}
                    className="px-2.5 py-1.5 text-[10px] font-extrabold text-rose-800 bg-rose-50 rounded-lg border border-rose-200 hover:bg-rose-100/70 transition-colors"
                  >
                    + Add TA
                  </button>
                  <button 
                    onClick={() => addRequirement(project.id, 'Post-TA')}
                    className="px-2.5 py-1.5 text-[10px] font-extrabold text-emerald-800 bg-emerald-50 rounded-lg border border-emerald-200 hover:bg-emerald-100/70 transition-colors"
                  >
                    + Add Post-TA
                  </button>
                </div>
              </div>

              <div className="space-y-3">
                {phases.map(phase => {
                  const phaseKey = `${project.id}-${phase}`;
                  const isPhaseCollapsed = !!collapsedPhases[phaseKey];
                  const phaseReqs = project.requirements.filter(r => r.phase === phase);

                  return (
                    <div key={phase} className="border border-slate-200/60 rounded-xl overflow-hidden shadow-3xs bg-white">
                      {/* Level 2: WBS Phase Header */}
                      <div 
                        className="px-4 py-2.5 bg-slate-50/50 border-b border-slate-200/40 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors select-none"
                        onClick={() => togglePhase(phaseKey)}
                      >
                        <div className="flex items-center gap-2">
                          <div className="p-0.5 text-slate-500 hover:bg-slate-250/10 rounded">
                            {isPhaseCollapsed ? <ChevronRight className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                          </div>
                          <span className={cn(
                            "px-2 py-0.5 text-[9px] font-extrabold rounded uppercase tracking-wider",
                            phase === 'TA' ? 'bg-rose-100 text-rose-800 border border-rose-200' :
                            phase === 'Pre-TA' ? 'bg-amber-100 text-amber-800 border border-amber-200' :
                            'bg-emerald-100 text-emerald-800 border border-emerald-200'
                          )}>
                            {phase} Phase Requisitions
                          </span>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="text-[10px] text-slate-500 font-extrabold bg-slate-100 px-2 py-0.5 rounded-full">
                            {phaseReqs.length} Crafts
                          </span>
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              addRequirement(project.id, phase);
                            }}
                            className="px-2 py-0.5 text-[9px] bg-white border border-slate-200 hover:bg-indigo-50 hover:text-indigo-700 hover:border-indigo-200 font-extrabold text-slate-650 rounded transition shadow-3xs"
                          >
                            + Quick Add
                          </button>
                        </div>
                      </div>

                      {/* Level 2 Content */}
                      {!isPhaseCollapsed && (
                        <div>
                          {phaseReqs.length === 0 ? (
                            <div className="p-5 text-center text-xs text-slate-400 italic bg-white flex flex-col items-center justify-center gap-1.5">
                              <span>No personnel requirements defined for {phase} phase yet.</span>
                              <button
                                onClick={() => addRequirement(project.id, phase)}
                                className="text-[10px] font-bold text-indigo-600 hover:underline inline-flex items-center gap-1 mt-1"
                              >
                                + Define first {phase} slot requirement
                              </button>
                            </div>
                          ) : (
                            <div className="overflow-x-auto bg-white">
                              <table className="w-full text-left border-collapse text-xs">
                                <thead>
                                  <tr className="bg-slate-50/30 border-b border-slate-150">
                                    <th className="px-5 py-2.5 text-[9px] font-extrabold text-[#888] uppercase tracking-wider">Required Craft Name</th>
                                    <th className="px-5 py-2.5 text-[9px] font-extrabold text-[#888] uppercase tracking-wider text-center w-28">Required Qty</th>
                                    <th className="px-5 py-2.5 text-[9px] font-extrabold text-[#888] uppercase tracking-wider">Schedule Period</th>
                                    <th className="px-5 py-2.5 text-[9px] font-extrabold text-[#888] uppercase tracking-wider text-right w-24">Actions</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-[#F3F4F6]">
                                  {phaseReqs.map(req => (
                                    <tr 
                                      key={req.id} 
                                      className="hover:bg-indigo-50/5 transition-colors group"
                                    >
                                      <td className="px-5 py-2.5">
                                        <input 
                                          type="text"
                                          list="common-crafts"
                                          className="text-xs font-bold text-[#1A1A1A] bg-transparent outline-none w-full border-b border-transparent focus:border-indigo-300 py-0.5"
                                          placeholder="Enter Craft..."
                                          value={req.craft}
                                          onChange={(e) => updateReqField(project.id, req.id, 'craft', e.target.value)}
                                        />
                                      </td>
                                      <td className="px-5 py-2.5 text-center">
                                        <input 
                                          type="number" 
                                          placeholder="Quantity" 
                                          className="w-16 px-1.5 py-1 border border-[#E5E5E5] rounded text-xs font-bold text-center focus:ring-1 focus:ring-indigo-500 outline-none bg-white"
                                          value={req.qty} 
                                          onChange={(e) => updateReqField(project.id, req.id, 'qty', parseInt(e.target.value) || 0)}
                                        />
                                      </td>
                                      <td className="px-5 py-2.5">
                                        <div className="flex items-center gap-1.5">
                                          <input 
                                            type="date" 
                                            className="text-[10px] p-1 border border-gray-100 rounded bg-white outline-none font-medium text-slate-700"
                                            value={req.startDate}
                                            onChange={(e) => updateReqField(project.id, req.id, 'startDate', e.target.value)}
                                          />
                                          <span className="text-gray-400 text-[10px]">—</span>
                                          <input 
                                            type="date" 
                                            className="text-[10px] p-1 border border-gray-100 rounded bg-white outline-none font-medium text-slate-700"
                                            value={req.endDate}
                                            onChange={(e) => updateReqField(project.id, req.id, 'endDate', e.target.value)}
                                          />
                                        </div>
                                      </td>
                                      <td className="px-5 py-2.5 text-right">
                                        {confirmDeleteReqId === req.id ? (
                                          <div className="flex items-center justify-end gap-1.5">
                                            <button 
                                              onClick={() => {
                                                const updatedReqs = project.requirements.filter(r => r.id !== req.id);
                                                setProjects(projects.map(p => p.id === project.id ? { ...p, requirements: updatedReqs } : p));
                                                setConfirmDeleteReqId(null);
                                              }}
                                              className="px-2 py-0.5 bg-red-500 text-white text-[9px] font-bold rounded hover:bg-red-600 transition"
                                            >
                                              Confirm
                                            </button>
                                            <button 
                                              onClick={() => setConfirmDeleteReqId(null)}
                                              className="px-2 py-0.5 bg-white text-gray-500 text-[9px] font-bold rounded border border-gray-200 hover:bg-gray-100 transition"
                                            >
                                              Exit
                                            </button>
                                          </div>
                                        ) : (
                                          <button 
                                            onClick={() => setConfirmDeleteReqId(req.id)}
                                            className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                                            title="Remove requirement"
                                          >
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        )}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
            )}
          </div>
        );
      })}
      </div>

      {/* Edit Project Modal Overlay */}
      {editingProject && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-[24px] border border-slate-200 shadow-[0_20px_50px_rgba(0,0,0,0.15)] max-w-lg w-full overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
              <div>
                <b className="text-base text-slate-900">Edit Project Settings</b>
                <p className="text-[11px] text-slate-500 mt-0.5">Modify core project timelines and area parameters</p>
              </div>
              <button 
                onClick={() => setEditingProject(null)} 
                className="p-1 px-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100/80 rounded-lg text-sm transition-all"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Form Content */}
            <div className="p-6 space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                
                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Project Title</span>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.name}
                    onChange={e => setEditingProject({...editingProject, name: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Project Code</span>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.code}
                    onChange={e => setEditingProject({...editingProject, code: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Location</span>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.location}
                    onChange={e => setEditingProject({...editingProject, location: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Department</span>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-850 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.department || ''}
                    onChange={e => setEditingProject({...editingProject, department: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Start Date</span>
                  <input 
                    type="date" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-850 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.startDate}
                    onChange={e => setEditingProject({...editingProject, startDate: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Finish Date</span>
                  <input 
                    type="date" 
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-850 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.endDate}
                    onChange={e => setEditingProject({...editingProject, endDate: e.target.value})}
                  />
                </div>

                <div className="space-y-1.5">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest block">Project Status</span>
                  <select
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-850 focus:border-indigo-500 outline-none bg-white shadow-3xs"
                    value={editingProject.status || 'Auto'}
                    onChange={e => setEditingProject({...editingProject, status: e.target.value as any})}
                  >
                    <option value="Auto">Auto (Date-based)</option>
                    <option value="In Progress">In Progress (Manual)</option>
                    <option value="Completed">Completed (Manual)</option>
                    <option value="Hold">Hold</option>
                    <option value="Rescheduled">Rescheduled</option>
                    <option value="Cancelled">Cancelled</option>
                  </select>
                </div>

              </div>
            </div>

            {/* Modal Actions */}
            <div className="px-6 py-4 bg-slate-50/50 border-t border-slate-100 flex items-center justify-end gap-2.5">
              <button 
                onClick={() => setEditingProject(null)}
                className="px-4 py-2 hover:bg-slate-100 text-xs font-bold text-slate-600 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button 
                onClick={handleSaveEdit}
                className="px-5 py-2 bg-indigo-600 text-white hover:bg-indigo-700 text-xs font-bold rounded-lg transition-colors shadow-xs"
              >
                Save Changes
              </button>
            </div>

          </div>
        </div>
      )}

      {/* Google Drive Import Modal */}
      {isDriveModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-[2px] animate-in fade-in duration-200 no-print">
          <div className="bg-white rounded-2xl border border-gray-100 shadow-2xl p-6 w-full max-w-lg mx-4 flex flex-col max-h-[85vh] animate-in zoom-in-95 duration-200 text-left">
            <div className="flex items-center justify-between border-b pb-4 mb-4">
              <div className="flex items-center gap-2">
                <Cloud className="w-5 h-5 text-indigo-600" />
                <h3 className="text-base font-bold text-gray-900">Import Projects from Google Drive</h3>
              </div>
              <button 
                onClick={() => setIsDriveModalOpen(false)}
                className="p-1.5 hover:bg-gray-100 rounded-full text-gray-400 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {!googleUser ? (
              <div className="py-8 text-center space-y-4">
                <div className="p-3 bg-indigo-50 text-indigo-600 rounded-full w-fit mx-auto">
                  <Cloud className="w-8 h-8" />
                </div>
                <div className="space-y-1">
                  <h4 className="font-bold text-slate-800">Google Drive Connection Required</h4>
                  <p className="text-xs text-slate-500 max-w-sm mx-auto leading-normal">
                    Please connect your Google Drive account first to securely search and import project requirements or backup packs.
                  </p>
                </div>
                
                <button
                  onClick={onGoogleConnect}
                  className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-2 px-4 rounded-xl shadow-md transition-colors cursor-pointer"
                >
                  <span>Connect Google Drive</span>
                </button>

                <p className="text-[10px] text-slate-500 font-medium text-center bg-amber-50 rounded-lg p-2.5 border border-amber-100 leading-normal max-w-xs mx-auto">
                  💡 <b>Cookie/Popup Workaround:</b> If popups are blocked by your browser, navigate to the <b>Backup & Sync</b> tab and click <b>Use Redirect Flow</b>. You can also click "New Tab  ↗" in the Reports tab to load the page standalone.
                </p>
              </div>
            ) : (
              <>
                {/* File search/list view */}
                <div className="mb-3">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                    <input 
                      type="text"
                      placeholder="Search spreadsheets / backups in Drive..."
                      className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-indigo-150"
                      value={driveSearchQuery}
                      onChange={e => setDriveSearchQuery(e.target.value)}
                    />
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto space-y-2 pr-1 min-h-[250px] max-h-[400px]">
                  {isFetchLoading ? (
                    <div className="py-20 text-center text-slate-500 text-xs flex flex-col items-center justify-center gap-2">
                      <RefreshCw className="w-5 h-5 text-indigo-500 animate-spin" />
                      <span>Scanning Google Drive files...</span>
                    </div>
                  ) : driveFiles.length === 0 ? (
                    <div className="py-20 text-center text-slate-400 text-xs">
                      No spreadsheets (.xlsx/.xls) or backup files found in Drive.
                      <button 
                        onClick={fetchDriveFiles} 
                        className="block mx-auto mt-2 text-indigo-650 hover:underline font-bold"
                      >
                        Reload List
                      </button>
                    </div>
                  ) : (
                    (() => {
                      const filteredFiles = driveFiles.filter(f => 
                        f.name.toLowerCase().includes(driveSearchQuery.toLowerCase())
                      );
                      
                      if (filteredFiles.length === 0) {
                        return <p className="text-center text-slate-400 text-xs py-10">No matching files found.</p>;
                      }

                      return filteredFiles.map(file => {
                        const isSelected = selectedDriveFileId === file.id;
                        const isJson = file.name.endsWith('.json');
                        return (
                          <div 
                            key={file.id}
                            onClick={() => !isImportLoading && setSelectedDriveFileId(file.id)}
                            className={cn(
                              "p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between text-left",
                              isSelected 
                                ? "border-indigo-500 bg-indigo-50/40 shadow-xs" 
                                : "border-slate-100 hover:bg-slate-50"
                            )}
                          >
                            <div className="flex items-center gap-3 truncate">
                              <div className={cn(
                                "p-2 rounded-lg text-xs font-bold shrink-0",
                                isJson ? "bg-amber-50 text-amber-700" : "bg-emerald-50 text-emerald-700"
                              )}>
                                {isJson ? 'JSON' : 'XLSX'}
                              </div>
                              <div className="truncate">
                                <p className="text-xs font-bold text-slate-850 truncate">{file.name}</p>
                                <p className="text-[9px] text-slate-500">
                                  Modified: {file.modifiedTime ? dayjs(file.modifiedTime).format('DD MMM YYYY, hh:mm A') : 'Unknown'}
                                </p>
                              </div>
                            </div>
                            <input 
                              type="radio"
                              checked={isSelected}
                              onChange={() => {}}
                              className="accent-indigo-600 shrink-0"
                            />
                          </div>
                        );
                      });
                    })()
                  )}
                </div>

                <div className="border-t pt-4 mt-4 flex items-center justify-between">
                  <button
                    onClick={fetchDriveFiles}
                    disabled={isFetchLoading || isImportLoading}
                    className="text-xs font-bold text-slate-650 hover:text-indigo-650 p-2 rounded-lg hover:bg-slate-50 flex items-center gap-1.5 cursor-pointer"
                  >
                    <RefreshCw className={cn("w-3.5 h-3.5", isFetchLoading && "animate-spin")} />
                    Reload
                  </button>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setIsDriveModalOpen(false)}
                      disabled={isImportLoading}
                      className="px-4 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-500 hover:bg-slate-50 cursor-pointer"
                    >
                      Cancel
                    </button>
                    <button
                      onClick={() => {
                        const file = driveFiles.find(f => f.id === selectedDriveFileId);
                        if (file) handleImportFromSelectedDriveFile(file.id, file.name);
                      }}
                      disabled={!selectedDriveFileId || isImportLoading}
                      className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold shadow-md disabled:opacity-50 inline-flex items-center gap-1.5 cursor-pointer"
                    >
                      {isImportLoading ? (
                        <>
                          <RefreshCw className="w-3 h-3 animate-spin" />
                          Importing...
                        </>
                      ) : (
                        'Load & Import'
                      )}
                    </button>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      <datalist id="common-crafts">
        {DEFAULT_CRAFTS.map(c => <option key={c} value={c} />)}
      </datalist>
    </div>
  );
}
