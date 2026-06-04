import React, { useState } from 'react';
import { 
  ChevronDown, 
  ChevronRight, 
  UserPlus, 
  UserMinus, 
  RefreshCw, 
  CheckCircle2, 
  Sparkles, 
  Download, 
  Search, 
  HelpCircle, 
  User, 
  SlidersHorizontal,
  FolderOpen,
  Calendar,
  X,
  Plus,
  HelpCircle as QuestionIcon
} from 'lucide-react';
import { Assignment, Manpower, Project, ProjectPhase, Craft } from '../types';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import { formatToExcelDate, getProjectActualStatus } from '../lib/dateUtils';
import { cn } from '../lib/utils';

interface Props {
  assignments: Assignment[];
  manpower: Manpower[];
  projects: Project[];
  onAutoAssign: () => void;
  setAssignments: (a: Assignment[]) => void;
}

const normalizeCraft = (c: string) => {
  if (!c) return '';
  return c.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
};

export default function AssignmentList({ assignments, manpower, projects, onAutoAssign, setAssignments }: Props) {
  // Collapsed status for Projects & Phases
  const [collapsedProjects, setCollapsedProjects] = useState<Record<string, boolean>>({});
  const [collapsedPhases, setCollapsedPhases] = useState<Record<string, boolean>>({});
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState<'startDate' | 'status'>('startDate');

  // Worker Selection Modal State
  const [modalOpen, setModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState<'assign' | 'replace'>('assign');
  const [targetProject, setTargetProject] = useState<Project | null>(null);
  const [targetPhase, setTargetPhase] = useState<ProjectPhase>('TA');
  const [targetCraft, setTargetCraft] = useState<Craft>('');
  const [targetStartDate, setTargetStartDate] = useState('');
  const [targetEndDate, setTargetEndDate] = useState('');
  const [assignmentToReplace, setAssignmentToReplace] = useState<Assignment | null>(null);
  const [searchWorkerQuery, setSearchWorkerQuery] = useState('');

  const getWorker = (id: string) => manpower.find(m => m.id === id);
  const getProject = (id: string) => projects.find(p => p.id === id);

  // Toggle Collapse States
  const toggleProject = (projId: string) => {
    setCollapsedProjects(prev => ({ ...prev, [projId]: !prev[projId] }));
  };

  const togglePhase = (projIdAndPhase: string) => {
    setCollapsedPhases(prev => ({ ...prev, [projIdAndPhase]: !prev[projIdAndPhase] }));
  };

  const expandAll = () => {
    setCollapsedProjects({});
    setCollapsedPhases({});
  };

  const collapseAll = () => {
    const projCollapsed: Record<string, boolean> = {};
    const phaseCollapsed: Record<string, boolean> = {};
    projects.forEach(p => {
      projCollapsed[p.id] = true;
      p.requirements.forEach(r => {
        phaseCollapsed[`${p.id}-${r.phase}`] = true;
      });
    });
    setCollapsedProjects(projCollapsed);
    setCollapsedPhases(phaseCollapsed);
  };

  // Check worker availability for period
  const getWorkerAvailabilityStatus = (
    worker: Manpower,
    startDate: string,
    endDate: string,
    ignoreAssignmentId?: string
  ) => {
    const workerJoin = dayjs(worker.joinDate);
    const workerRel = dayjs(worker.releaseDate);
    const reqS = dayjs(startDate);
    const reqE = dayjs(endDate);

    if (!workerJoin.isValid() || !workerRel.isValid() || !reqS.isValid() || !reqE.isValid()) {
      return { available: false, reason: 'Invalid dates' };
    }

    // Check service contract dates
    if (!workerJoin.isSameOrBefore(reqS, 'day') || !workerRel.isSameOrAfter(reqE, 'day')) {
      return { 
        available: false, 
        reason: `Contract limit (${dayjs(worker.joinDate).format('DD MMM YY')} - ${dayjs(worker.releaseDate).format('DD MMM YY')})` 
      };
    }

    // Check Vacation
    if (worker.vacationStart && worker.vacationEnd) {
      const vacS = dayjs(worker.vacationStart);
      const vacE = dayjs(worker.vacationEnd);
      if (vacS.isValid() && vacE.isValid()) {
        if (reqS.isBetween(vacS, vacE, 'day', '[]') || reqE.isBetween(vacS, vacE, 'day', '[]') ||
            vacS.isBetween(reqS, reqE, 'day', '[]')) {
          return { available: false, reason: 'On scheduled vacation' };
        }
      }
    }

    // Check other assignments
    const conflicts = assignments.filter(a => {
      if (a.id === ignoreAssignmentId || a.workerId !== worker.id) return false;
      const otherP = getProject(a.projectId);
      if (otherP) {
        const actStatus = getProjectActualStatus(otherP);
        if (actStatus === 'Hold' || actStatus === 'Cancelled') {
          return false;
        }
      }
      return true;
    });
    for (const conf of conflicts) {
      const existingS = dayjs(conf.startDate);
      const existingE = dayjs(conf.endDate);
      if (reqS.isBetween(existingS, existingE, 'day', '[]') || reqE.isBetween(existingS, existingE, 'day', '[]') ||
          existingS.isBetween(reqS, reqE, 'day', '[]')) {
        const otherProj = getProject(conf.projectId);
        return { 
          available: false, 
          reason: `Assigned to ${otherProj?.name || 'another project'} (${existingS.format('DD MMM')} - ${existingE.format('DD MMM')})` 
        };
      }
    }

    return { available: true };
  };

  // Open replace / assign modal
  const openModalForSlot = (
    mode: 'assign' | 'replace',
    project: Project,
    phase: ProjectPhase,
    craft: Craft,
    startDate: string,
    endDate: string,
    existingAssignment?: Assignment
  ) => {
    setModalMode(mode);
    setTargetProject(project);
    setTargetPhase(phase);
    setTargetCraft(craft);
    setTargetStartDate(startDate);
    setTargetEndDate(endDate);
    setAssignmentToReplace(existingAssignment || null);
    setSearchWorkerQuery('');
    setModalOpen(true);
  };

  // Confirm worker assignment / substitution
  const chooseWorker = (workerId: string) => {
    if (!targetProject) return;

    if (modalMode === 'replace' && assignmentToReplace) {
      // Edit existing assignment ID to target the new worker details
      setAssignments(assignments.map(a => 
        a.id === assignmentToReplace.id ? { ...a, workerId } : a
      ));
    } else {
      // Standard assign
      const newAssign: Assignment = {
        id: `ASSIGN-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
        workerId,
        projectId: targetProject.id,
        craft: targetCraft,
        phase: targetPhase,
        startDate: targetStartDate,
        endDate: targetEndDate
      };
      setAssignments([...assignments, newAssign]);
    }

    setModalOpen(false);
  };

  // Clear Assignment Spot - "To be Announced"
  const clearAssignmentSpot = (assignId: string) => {
    if (confirm("Are you sure you want to clear this worker and mark it as 'To be Announced'?")) {
      setAssignments(assignments.filter(a => a.id !== assignId));
    }
  };

  // Export
  const exportToExcel = () => {
    const data = assignments.map(a => {
      const worker = getWorker(a.workerId);
      const project = getProject(a.projectId);
      return {
        'Worker Name': worker?.name || 'Unknown',
        'Staff ID': worker?.badgeNo || '',
        'Badge Number': worker?.badgeNo || '',
        'Passport/Iqama': worker?.passportIqama || '',
        'Craft': a.craft,
        'Strength Rating': worker?.strength || 'Good',
        'Project': project?.name || 'Unknown',
        'Phase': a.phase,
        'Start Date': formatToExcelDate(a.startDate),
        'End Date': formatToExcelDate(a.endDate)
      };
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Assignments');
    XLSX.writeFile(wb, `assignments_${formatToExcelDate(new Date())}.xlsx`);
  };

  // Filter matching candidates for our modal - return all candidates so they can assign any craft/category
  const getModalCandidates = () => {
    return manpower;
  };

  const modalCandidates = getModalCandidates();
  
  // Search state query
  const filteredModalCandidates = modalCandidates.filter(m => 
    m.name.toLowerCase().includes(searchWorkerQuery.toLowerCase()) ||
    m.badgeNo.toLowerCase().includes(searchWorkerQuery.toLowerCase()) ||
    (m.craft && m.craft.toLowerCase().includes(searchWorkerQuery.toLowerCase()))
  ).map(worker => {
    const status = getWorkerAvailabilityStatus(
      worker, 
      targetStartDate, 
      targetEndDate, 
      assignmentToReplace?.id
    );
    return { worker, status };
  }).sort((a, b) => {
    // 1. Sort available workers first
    if (a.status.available !== b.status.available) {
      return a.status.available ? -1 : 1;
    }
    
    // 2. Sort those with matching target craft first (perfect fits on top)
    const normTarget = normalizeCraft(targetCraft);
    const isMatchingA = normalizeCraft(a.worker.craft) === normTarget;
    const isMatchingB = normalizeCraft(b.worker.craft) === normTarget;
    if (isMatchingA !== isMatchingB) {
      return isMatchingA ? -1 : 1;
    }

    // 3. Sort by strength: Excellent -> Good -> Average
    const strengths = { Excellent: 3, Good: 2, Average: 1 };
    const strA = strengths[a.worker.strength || 'Good'] || 2;
    const strB = strengths[b.worker.strength || 'Good'] || 2;
    if (strA !== strB) {
      return strB - strA;
    }

    // 4. Sort alphabetically by name
    return a.worker.name.localeCompare(b.worker.name);
  });

  const sortedAndFilteredProjects = (() => {
    const list = projects.filter(p => 
      p.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
      p.code.toLowerCase().includes(searchTerm.toLowerCase())
    );

    if (sortBy === 'status') {
      return [...list].sort((a, b) => {
        const statusA = getProjectActualStatus(a);
        const statusB = getProjectActualStatus(b);
        if (statusA !== statusB) {
          return statusA.localeCompare(statusB);
        }
        return dayjs(a.startDate).valueOf() - dayjs(b.startDate).valueOf();
      });
    } else {
      return [...list].sort((a, b) => {
        return dayjs(a.startDate).valueOf() - dayjs(b.startDate).valueOf();
      });
    }
  })();

  return (
    <div className="space-y-6 font-sans">
      {/* Search Header / Config bar */}
      <div className="bg-white rounded-2xl border border-[#E5E5E5] p-5 shadow-sm flex flex-col gap-4 mb-6">
        <div>
          <h3 className="text-base font-bold text-[#1A1A1A]">Assignments</h3>
          <p className="text-xs text-gray-500 mt-0.5">Dynamic job mapping, workforce deployments, and manual gap reconciliation</p>
        </div>
        
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 border-t border-slate-100 pt-3">
          <div className="relative flex-1 min-w-[180px] max-w-xs">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Search WBS by project..." 
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
                <option value="startDate">Start Date (Asc)</option>
                <option value="status">Project Status (Asc) then Start Date</option>
              </select>
            </div>
            
            <div className="flex items-center gap-1.5 shrink-0">
              {projects.length > 0 && (
                <div className="flex items-center gap-1 border-r border-[#E5E5E5] pr-2.5 mr-0.5 shrink-0">
                  <button 
                    onClick={expandAll}
                    className="flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-bold text-slate-655 border border-slate-250 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                    title="Expand All"
                  >
                    <FolderOpen className="w-3.5 h-3.5 text-slate-550" />
                    Expand All
                  </button>
                  <button 
                    onClick={collapseAll}
                    className="flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-bold text-slate-655 border border-slate-250 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                    title="Collapse All"
                  >
                    <FolderOpen className="w-3.5 h-3.5 text-slate-400" />
                    Collapse All
                  </button>
                </div>
              )}

              {assignments.length > 0 && (
                <button 
                  onClick={exportToExcel}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-655 border border-slate-200 rounded-lg bg-white hover:bg-slate-50 transition-colors shrink-0"
                >
                  <Download className="w-3.5 h-3.5 text-slate-550" />
                  Export Excel
                </button>
              )}

              <button 
                onClick={onAutoAssign}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-indigo-700 border border-indigo-150 bg-indigo-50 rounded-lg hover:bg-indigo-100 transition-colors shrink-0"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                Auto-Assign All
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Main List */}
      <div className="space-y-4">
        {sortedAndFilteredProjects.map(project => {
          const isProjCollapsed = collapsedProjects[project.id];
          
          // Find requirements and phase grouping
          const validReqs = project.requirements.filter(r => r.qty > 0);
          const phasesInProject = Array.from(new Set(validReqs.map(r => r.phase)));

          // Calculate counts
          const totalRequiredSlots = validReqs.reduce((sum, r) => sum + r.qty, 0);
          const projectAssignedCount = assignments.filter(a => a.projectId === project.id).length;
          const openSlotsCount = Math.max(0, totalRequiredSlots - projectAssignedCount);

          return (
            <div key={project.id} className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden transition-all duration-300">
              
              {/* Level 1: Project Header */}
              <div 
                className="p-4 bg-slate-50 border-b border-slate-200/60 flex items-center justify-between cursor-pointer hover:bg-slate-100/50 transition-colors select-none"
                onClick={() => toggleProject(project.id)}
              >
                <div className="flex items-center gap-3 min-w-0">
                  <div className="p-1 hover:bg-gray-200/80 rounded-lg transition-colors text-slate-500">
                    {isProjCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                  </div>
                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="bg-indigo-50 text-indigo-700 px-2 py-0.5 text-[10px] font-extrabold uppercase rounded border border-indigo-100 tracking-wider">
                        {project.code}
                      </span>
                      <h4 className="text-sm font-bold text-[#1a1a1a] truncate">{project.name}</h4>
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
                            "px-1.5 py-0.5 text-[8px] font-bold uppercase tracking-wider rounded border shadow-3xs",
                            badgeColor
                          )}>
                            {status}
                          </span>
                        );
                      })()}
                    </div>
                    <p className="text-[10px] text-gray-500 font-medium mt-0.5">{project.location} • {dayjs(project.startDate).format('DD MMM')} — {dayjs(project.endDate).format('DD MMM YY')}</p>
                  </div>
                </div>

                <div className="flex items-center gap-3 shrink-0">
                  <span className="text-[10px] bg-slate-200/60 px-2 py-0.5 rounded-full font-bold text-slate-650 flex items-center gap-1">
                    Req Slots: <b>{totalRequiredSlots}</b>
                  </span>
                  <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-100 px-2 py-0.5 rounded-full font-bold flex items-center gap-1">
                    Assigned: <b>{projectAssignedCount}</b>
                  </span>
                  {openSlotsCount > 0 ? (
                    <span className="text-[10px] bg-amber-50 text-amber-700 border border-amber-100 px-2 py-0.5 rounded-full font-bold shrink-0">
                      TBA: <b>{openSlotsCount}</b>
                    </span>
                  ) : (
                    <span className="text-[10px] bg-green-100 text-green-800 px-2 py-0.5 rounded-full font-extrabold shrink-0 uppercase tracking-tight">
                      Fulfillment Met
                    </span>
                  )}
                </div>
              </div>

              {!isProjCollapsed && (
                <div className="p-4 space-y-4">
                  
                  {phasesInProject.length === 0 && (
                    <div className="p-4 text-center text-xs text-gray-450 italic">
                      No personnel requirements defined for this project. Switch to the "Projects" tab to configure requirements.
                    </div>
                  )}

                  {phasesInProject.map(phase => {
                    const phaseKey = `${project.id}-${phase}`;
                    const isPhaseCollapsed = collapsedPhases[phaseKey];
                    const phaseRequirements = validReqs.filter(r => r.phase === phase);

                    return (
                      <div key={phase} className="border border-slate-200/90 rounded-xl overflow-hidden shadow-2xs">
                        
                        {/* Level 2: Phase Header */}
                        <div 
                          className="px-4 py-2 bg-slate-100/35 border-b border-indigo-100/50 flex items-center justify-between cursor-pointer hover:bg-slate-100/70 transition-colors select-none"
                          onClick={() => togglePhase(phaseKey)}
                        >
                          <div className="flex items-center gap-2">
                            {isPhaseCollapsed ? <ChevronRight className="w-3.5 h-3.5 text-gray-500" /> : <ChevronDown className="w-3.5 h-3.5 text-gray-500" />}
                            <span className={cn(
                              "px-1.5 py-0.5 text-[9px] font-bold rounded uppercase tracking-wider",
                              phase === 'TA' ? 'bg-rose-100 text-rose-800' :
                              phase === 'Pre-TA' ? 'bg-amber-100 text-amber-800' :
                              'bg-green-100 text-green-800'
                            )}>
                              {phase} Phase Requirements
                            </span>
                          </div>
                          
                          <div className="text-[10px] font-bold text-gray-600">
                            {phaseRequirements.length} crafts required
                          </div>
                        </div>

                        {!isPhaseCollapsed && (
                          <div className="divide-y divide-slate-150 bg-white">
                            
                            {/* Level 3: Craft Grouping */}
                            {phaseRequirements.map(req => {
                              // Find all active assignments that fulfill this exact project requirement
                              const normRequiredCraft = normalizeCraft(req.craft);
                              const matchingAssignments = assignments.filter(a => 
                                a.projectId === project.id && 
                                a.phase === phase && 
                                normalizeCraft(a.craft) === normRequiredCraft
                              );

                              return (
                                <div key={req.id} className="p-4 space-y-3">
                                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-dashed border-slate-200 pb-2">
                                    <div className="flex items-center gap-2 flex-wrap text-xs font-bold text-[#1A1A1A]">
                                      <span className="px-2 py-0.5 bg-indigo-50 border border-indigo-150 text-indigo-700 font-extrabold uppercase rounded text-[10px]">
                                        {req.craft}
                                      </span>
                                      <span className="text-[10px] text-gray-500 font-medium">
                                        Required Interval: {dayjs(req.startDate).format('DD MMM YY')} — {dayjs(req.endDate).format('DD MMM YY')}
                                      </span>
                                    </div>
                                    <div className="text-[10px] font-bold text-slate-650 shrink-0">
                                      Demand: <b className="text-indigo-600">{matchingAssignments.length}</b> allocated of <b className="text-gray-800">{req.qty}</b> planned
                                    </div>
                                  </div>

                                  {/* Level 4: The Slot Assignments list (1 slot per qty) */}
                                  <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-2xs">
                                    <table className="w-full text-left border-collapse text-xs">
                                      <thead>
                                        <tr className="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                                          <th className="px-4 py-3 w-12 text-center">Slot</th>
                                          <th className="px-4 py-3">Deployed Worker / Status</th>
                                          <th className="px-4 py-3 w-32">Badge ID</th>
                                          <th className="px-4 py-3 w-40">Passport / Iqama</th>
                                          <th className="px-4 py-3 w-28">Strength</th>
                                          <th className="px-4 py-3 w-44 text-right">Actions</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {Array.from({ length: req.qty }).map((_, slotIdx) => {
                                          const activeAssignment = matchingAssignments[slotIdx];
                                          
                                          if (activeAssignment) {
                                            const worker = getWorker(activeAssignment.workerId);
                                            return (
                                              <tr key={slotIdx} className="hover:bg-indigo-50/10 transition-colors font-sans">
                                                <td className="px-4 py-3 text-center font-bold text-slate-400">
                                                  #{slotIdx + 1}
                                                </td>
                                                <td className="px-4 py-3">
                                                  <div className="flex items-center gap-2">
                                                    <User className="w-4 h-4 text-indigo-500 shrink-0" />
                                                    <span className="font-bold text-slate-900">{worker?.name || 'Unregistered Worker'}</span>
                                                  </div>
                                                </td>
                                                <td className="px-4 py-3 font-mono text-[11px] font-semibold text-slate-600">
                                                  {worker?.badgeNo || 'N/A'}
                                                </td>
                                                <td className="px-4 py-3 text-slate-655 font-semibold">
                                                  {worker?.passportIqama || '--'}
                                                </td>
                                                <td className="px-4 py-3">
                                                  <span className={cn(
                                                    "px-2 py-0.5 text-[9px] font-extrabold rounded border uppercase tracking-wider",
                                                    (worker?.strength || 'Good') === 'Excellent' ? 'bg-purple-50 text-purple-700 border-purple-200' :
                                                    (worker?.strength || 'Good') === 'Average' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                                                    'bg-emerald-50 text-emerald-700 border-emerald-200'
                                                  )}>
                                                    {worker?.strength || 'Good'}
                                                  </span>
                                                </td>
                                                <td className="px-4 py-3 text-right">
                                                  <div className="flex items-center justify-end gap-1.5">
                                                    <button 
                                                      onClick={() => openModalForSlot('replace', project, phase, req.craft, req.startDate, req.endDate, activeAssignment)}
                                                      className="px-2 py-1 text-[10px] font-bold text-indigo-700 hover:bg-indigo-50 rounded-lg border border-indigo-150 transition-colors cursor-pointer"
                                                      title="Replace worker with same craft"
                                                    >
                                                      🔄 Substitute
                                                    </button>
                                                    <button 
                                                      onClick={() => clearAssignmentSpot(activeAssignment.id)}
                                                      className="px-2 py-1 text-[10px] font-bold text-rose-700 hover:bg-rose-50 rounded-lg border border-rose-100 transition-colors cursor-pointer"
                                                      title="Mark as To Be Announced"
                                                    >
                                                      🚫 Empty
                                                    </button>
                                                  </div>
                                                </td>
                                              </tr>
                                            );
                                          } else {
                                            // Empty Slot
                                            return (
                                              <tr key={slotIdx} className="bg-slate-50/30 hover:bg-slate-50/70 transition-colors font-sans italic">
                                                <td className="px-4 py-3 text-center font-bold text-slate-350">
                                                  #{slotIdx + 1}
                                                </td>
                                                <td className="px-4 py-3 text-slate-400">
                                                  <div className="flex items-center gap-2">
                                                    <QuestionIcon className="w-4 h-4 text-slate-350 shrink-0" />
                                                    <span className="font-semibold text-slate-400">To be Announced</span>
                                                    <span className="px-1.5 py-0.2 bg-slate-100 border border-slate-200 text-slate-400 text-[8px] font-bold rounded uppercase not-italic tracking-wide ml-1">
                                                      Open Slot
                                                    </span>
                                                  </div>
                                                </td>
                                                <td className="px-4 py-3 text-slate-350 font-mono text-[11px] not-italic">
                                                  --
                                                </td>
                                                <td className="px-4 py-3 text-slate-350 not-italic">
                                                  --
                                                </td>
                                                <td className="px-4 py-3 not-italic text-slate-350">
                                                  --
                                                </td>
                                                <td className="px-4 py-3 text-right">
                                                  <button 
                                                    onClick={() => openModalForSlot('assign', project, phase, req.craft, req.startDate, req.endDate)}
                                                    className="inline-flex py-1 px-2.5 bg-white hover:bg-indigo-50 border border-indigo-200 hover:border-indigo-455 rounded-lg text-[10px] font-bold text-indigo-600 items-center justify-center gap-1 transition-all not-italic cursor-pointer"
                                                  >
                                                    <Plus className="w-3.5 h-3.5" />
                                                    Assign Worker
                                                  </button>
                                                </td>
                                              </tr>
                                            );
                                          }
                                        })}
                                      </tbody>
                                    </table>
                                  </div>
                                </div>
                              );
                            })}

                          </div>
                        )}

                      </div>
                    );
                  })}

                </div>
              )}

            </div>
          );
        })}
      </div>

      {assignments.length === 0 && (
        <div className="bg-white rounded-2xl border border-[#E5E5E5] p-12 text-center shadow-sm">
          <div className="w-16 h-16 bg-indigo-50 rounded-full flex items-center justify-center mb-4 mx-auto border border-indigo-100">
            <Sparkles className="w-8 h-8 text-indigo-500" />
          </div>
          <h4 className="text-base font-bold mb-1">WBS Matrix Empty</h4>
          <p className="text-xs text-slate-500 max-w-sm mx-auto mb-6">No personnel assigned to projects. You can launch auto-assignment to populate the matrix dynamically or assign manually slot-by-slot.</p>
          <button 
            onClick={onAutoAssign}
            className="px-6 py-2.5 bg-indigo-600 text-white rounded-xl text-xs font-bold hover:bg-indigo-700 shadow-md transition-all inline-flex items-center gap-2"
          >
            <RefreshCw className="w-4 h-4 animate-spin-slow" />
            Run Auto-Placement Engine
          </button>
        </div>
      )}

      {/* Choose/Replace Worker Modal Overlay */}
      {modalOpen && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[90vh] animate-in fade-in zoom-in-95 duration-200">
            
            {/* Modal Header */}
            <div className="p-4 bg-slate-50 border-b border-slate-150 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  {modalMode === 'replace' ? 'Substitute Craft Personnel' : 'Deploy Qualified Worker'}
                </h3>
                <p className="text-[10.5px] text-gray-500 mt-0.5 font-medium leading-tight">
                  Finding eligible <span className="text-indigo-600 uppercase font-extrabold">{targetCraft}s</span> • {dayjs(targetStartDate).format('DD MMM')} to {dayjs(targetEndDate).format('DD MMM YY')}
                </p>
              </div>
              <button 
                onClick={() => setModalOpen(false)}
                className="p-1.5 hover:bg-slate-200 rounded-lg text-slate-400"
              >
                <X className="w-4.5 h-4.5" />
              </button>
            </div>

            {/* Modal search bar */}
            <div className="p-3 border-b border-slate-100 flex items-center gap-2 bg-white">
              <Search className="w-4 h-4 text-slate-450 shrink-0 ml-1" />
              <input 
                type="text" 
                placeholder="Search candidates by name, badge..."
                className="w-full text-xs outline-none bg-transparent font-medium py-1 placeholder-slate-400"
                value={searchWorkerQuery}
                onChange={e => setSearchWorkerQuery(e.target.value)}
                autoFocus
              />
            </div>

            {/* Candidates List with real-time Availability checking and Strength balancing indicator */}
            <div className="p-2 h-[420px] overflow-y-auto bg-slate-50/70 divide-y divide-slate-150">
              
              {filteredModalCandidates.length === 0 && (
                <div className="p-8 text-center text-xs text-slate-400 italic">
                  No registered {targetCraft || 'workers'} matched your search query. Verify primary crafts in the Manpower Pool.
                </div>
              )}

              {filteredModalCandidates.map(({ worker, status }) => (
                <div 
                  key={worker.id}
                  className={cn(
                    "p-3 rounded-xl flex items-center justify-between gap-3 text-left transition-all mb-2",
                    status.available 
                      ? "hover:bg-indigo-50/40 cursor-pointer bg-white border border-slate-200 shadow-2xs hover:border-indigo-300" 
                      : "opacity-72 bg-slate-100/40 cursor-not-allowed border border-slate-150"
                  )}
                  onClick={() => status.available && chooseWorker(worker.id)}
                >
                  <div className="min-w-0">
                    <div className="flex items-center gap-2 flex-wrap">
                      <p className="text-xs font-bold text-slate-900 truncate">{worker.name}</p>
                      
                      {/* Craft Designation Badge */}
                      <span className={cn(
                        "px-1.5 py-0.2 text-[8px] font-extrabold rounded uppercase tracking-wider border",
                        normalizeCraft(worker.craft) === normalizeCraft(targetCraft)
                          ? "bg-indigo-50 text-indigo-700 border-indigo-200"
                          : "bg-slate-50 text-slate-600 border-slate-200"
                      )}>
                        {worker.craft}
                      </span>

                      {/* Strength level badges */}
                      <span className={cn(
                        "px-1.5 py-0.2 text-[8px] font-extrabold rounded uppercase tracking-wider",
                        worker.strength === 'Excellent' ? 'bg-purple-100 text-purple-800' :
                        worker.strength === 'Average' ? 'bg-amber-100 text-amber-800' :
                        'bg-emerald-100 text-emerald-800'
                      )}>
                        {worker.strength || 'Good'}
                      </span>
                    </div>

                    <p className="text-[10px] text-slate-500 font-semibold mt-1">
                      Badge: <span className="text-slate-850 font-bold">{worker.badgeNo}</span> • Contract End: {dayjs(worker.releaseDate).format('DD MMM YY')}
                    </p>

                    {/* Conflict context indicator */}
                    {!status.available && (
                      <p className="text-[9px] text-rose-600 font-bold mt-1 flex items-center gap-1">
                        ⚠️ Limit: {status.reason}
                      </p>
                    )}
                  </div>

                  {status.available ? (
                    <button 
                      onClick={(e) => {
                        e.stopPropagation();
                        chooseWorker(worker.id);
                      }}
                      className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-[9px] font-bold shadow-xs flex items-center gap-1 cursor-pointer"
                    >
                      Deploy
                    </button>
                  ) : (
                    <span className="text-[9px] text-slate-400 font-extrabold uppercase bg-slate-200 px-1.5 py-0.5 rounded shrink-0">
                      Unavailable
                    </span>
                  )}
                </div>
              ))}
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
