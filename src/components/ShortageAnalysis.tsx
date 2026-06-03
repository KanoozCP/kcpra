import { useState, useMemo } from 'react';
import { 
  AlertCircle, 
  ArrowRight, 
  UserMinus, 
  ShieldAlert, 
  CheckCircle2, 
  ChevronDown, 
  ChevronRight, 
  FolderMinus, 
  FolderPlus,
  Compass,
  Calendar,
  Layers,
  Wrench
} from 'lucide-react';
import { Manpower, Project, Assignment } from '../types';
import { calculateShortages } from '../lib/assignment-engine';
import { cn } from '../lib/utils';
import dayjs from 'dayjs';

interface Props {
  manpower: Manpower[];
  projects: Project[];
  assignments: Assignment[];
}

export default function ShortageAnalysis({ manpower, projects, assignments }: Props) {
  // Compute raw shortages through the analysis engine
  const shortages = useMemo(() => calculateShortages(manpower, projects, assignments), [manpower, projects, assignments]);

  const criticalCount = shortages.filter(s => s.assigned === 0).length;
  const totalGapHeads = shortages.reduce((sum, s) => sum + s.gap, 0);

  // WBS Collapse States
  const [collapsedProjects, setCollapsedProjects] = useState<Record<string, boolean>>({});
  const [collapsedPhases, setCollapsedPhases] = useState<Record<string, boolean>>({});

  // Grouping logic for WBS representation
  const wbsShortages = useMemo(() => {
    // Group shortages by projects first
    const projectsMap: Record<string, {
      project: Project;
      totalGap: number;
      phases: Record<string, {
        phase: string;
        totalGap: number;
        items: typeof shortages;
      }>;
    }> = {};

    shortages.forEach(s => {
      // Find matching project reference
      const proj = projects.find(p => p.id === s.projectId);
      if (!proj) return;

      if (!projectsMap[proj.id]) {
        projectsMap[proj.id] = {
          project: proj,
          totalGap: 0,
          phases: {}
        };
      }

      const pGroup = projectsMap[proj.id];
      pGroup.totalGap += s.gap;

      if (!pGroup.phases[s.phase]) {
        pGroup.phases[s.phase] = {
          phase: s.phase,
          totalGap: 0,
          items: []
        };
      }

      const phaseGroup = pGroup.phases[s.phase];
      phaseGroup.totalGap += s.gap;
      phaseGroup.items.push(s);
    });

    return Object.values(projectsMap).sort((a, b) => {
      // Prioritize projects with highest gap totals at the top
      return b.totalGap - a.totalGap;
    });
  }, [shortages, projects]);

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
    wbsShortages.forEach(pGroup => {
      projCollapsed[pGroup.project.id] = true;
      Object.keys(pGroup.phases).forEach(phKey => {
        phaseCollapsed[`${pGroup.project.id}-${phKey}`] = true;
      });
    });
    setCollapsedProjects(projCollapsed);
    setCollapsedPhases(phaseCollapsed);
  };

  return (
    <div className="space-y-6 font-sans">
      {/* Shortage Top Dashboard Cards */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5">
          <div className="p-3 bg-red-50 rounded-xl shrink-0">
            <AlertCircle className="text-red-500 w-5 h-5 animate-pulse" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Total Deficit Categories</p>
            <h4 className="text-lg font-bold text-slate-900 mt-0.5">{shortages.length} Gaps Found</h4>
          </div>
        </div>

        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5 border-l-4 border-l-red-500">
          <div className="p-3 bg-rose-50 rounded-xl shrink-0">
            <ShieldAlert className="text-rose-600 w-5 h-5" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Critical Unstaffed Roles</p>
            <h4 className="text-lg font-bold text-rose-600 mt-0.5">{criticalCount} Zero Allocations</h4>
          </div>
        </div>

        <div className="bg-white p-4.5 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-4.5">
          <div className="p-3 bg-indigo-50 rounded-xl shrink-0">
            <UserMinus className="text-indigo-600 w-5 h-5" />
          </div>
          <div>
            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Accumulated Head Gap</p>
            <h4 className="text-lg font-bold text-indigo-700 mt-0.5">{totalGapHeads} Missing Personnel</h4>
          </div>
        </div>
      </div>

      {/* Main WBS Container */}
      <div className="bg-white rounded-2xl border border-[#E5E5E5] shadow-xs overflow-hidden min-h-[420px]">
        {/* WBS Heading and Control Utility Bar */}
        <div className="p-5 border-b border-[#E5E5E5] flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-[#FAFAFB]">
          <div>
            <h3 className="text-sm font-bold text-[#1A1A1A]">Work Breakdown Structure (WBS) Shortage Map</h3>
            <p className="text-[10.5px] text-gray-500 font-medium mt-0.5">Hierarchical evaluation of personnel deficits categorized by project scope, phase schedules, and necessary specialties.</p>
          </div>
          {wbsShortages.length > 0 && (
            <div className="flex items-center gap-2 self-start sm:self-center">
              <button 
                onClick={expandAll}
                className="px-3 py-1.5 border border-slate-250 bg-white hover:bg-slate-100/50 rounded-xl text-[10px] font-bold text-slate-700 transition flex items-center gap-1.5"
              >
                <FolderPlus className="w-3.5 h-3.5" />
                Expand All
              </button>
              <button 
                onClick={collapseAll}
                className="px-3 py-1.5 border border-slate-250 bg-white hover:bg-slate-100/50 rounded-xl text-[10px] font-bold text-slate-700 transition flex items-center gap-1.5"
              >
                <FolderMinus className="w-3.5 h-3.5" />
                Collapse All
              </button>
            </div>
          )}
        </div>

        {/* WBS Hierarchical Render */}
        <div className="p-5 space-y-4">
          {wbsShortages.map(pGroup => {
            const isProjCollapsed = !!collapsedProjects[pGroup.project.id];
            const phasesKeys = Object.keys(pGroup.phases);

            return (
              <div 
                key={pGroup.project.id} 
                className="border border-slate-200/85 rounded-2xl overflow-hidden transition-all duration-300 shadow-3xs"
              >
                {/* LEVEL 1: Project Node */}
                <div 
                  className="p-3.5 bg-slate-50 border-b border-slate-200/60 flex items-center justify-between cursor-pointer hover:bg-slate-100/60 transition-colors select-none"
                  onClick={() => toggleProject(pGroup.project.id)}
                >
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="p-1 hover:bg-gray-250/50 rounded-lg text-slate-500 shrink-0">
                      {isProjCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                    </div>
                    <div className="min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="bg-indigo-50 border border-indigo-150 text-indigo-700 px-2 py-0.5 text-[10px] font-extrabold uppercase rounded tracking-wider">
                          {pGroup.project.code}
                        </span>
                        <h4 className="text-xs font-bold text-[#1a1a1a] truncate">{pGroup.project.name}</h4>
                      </div>
                      <p className="text-[10px] text-gray-450 font-medium mt-0.5">
                        Location: {pGroup.project.location} • Period: {dayjs(pGroup.project.startDate).format('DD MMM')} to {dayjs(pGroup.project.endDate).format('DD MMM YY')}
                      </p>
                    </div>
                  </div>

                  <div className="shrink-0 pl-2">
                    <span className="text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-150 rounded-full px-2.5 py-1">
                      Total Deficit: <b>{pGroup.totalGap} Heads</b>
                    </span>
                  </div>
                </div>

                {/* Level 1 Content */}
                {!isProjCollapsed && (
                  <div className="p-4 space-y-3 bg-white">
                    {phasesKeys.map(phKey => {
                      const phaseGroup = pGroup.phases[phKey];
                      const uniquePhaseKey = `${pGroup.project.id}-${phKey}`;
                      const isPhaseCollapsed = !!collapsedPhases[uniquePhaseKey];

                      return (
                        <div key={phKey} className="border border-slate-200/80 rounded-xl overflow-hidden">
                          {/* LEVEL 2: Phase Node */}
                          <div 
                            className="px-4 py-2 bg-slate-100/35 border-b border-slate-200/50 flex items-center justify-between cursor-pointer hover:bg-slate-100/70 transition-colors select-none"
                            onClick={() => togglePhase(uniquePhaseKey)}
                          >
                            <div className="flex items-center gap-2.5">
                              {isPhaseCollapsed ? <ChevronRight className="w-3.5 h-3.5 text-gray-500" /> : <ChevronDown className="w-3.5 h-3.5 text-gray-500" />}
                              <span className={cn(
                                "px-1.5 py-0.5 text-[9px] font-bold rounded uppercase tracking-wider",
                                phKey === 'TA' ? 'bg-rose-100 text-rose-800' :
                                phKey === 'Pre-TA' ? 'bg-amber-100 text-amber-800' :
                                'bg-green-100 text-green-800'
                              )}>
                                {phKey} Phase Deficiencies
                              </span>
                            </div>
                            <div className="text-[10px] font-bold text-rose-700 bg-rose-50/55 border border-rose-100/80 px-2 py-0.5 rounded-full">
                              Phase Deficit: <b>-{phaseGroup.totalGap}</b>
                            </div>
                          </div>

                          {/* Level 2 Content */}
                          {!isPhaseCollapsed && (
                            <div className="divide-y divide-slate-100 bg-white">
                              {/* LEVEL 3: Craft and Requirement shortages */}
                              {phaseGroup.items.map((shortage, sIdx) => (
                                <div key={sIdx} className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4.5 hover:bg-red-50/5 transition-all">
                                  
                                  {/* Craft Description & Schedule Frame */}
                                  <div className="min-w-0">
                                    <div className="flex items-center gap-2 flex-wrap">
                                      <span className="px-2.5 py-0.5 bg-indigo-50 border border-indigo-150 text-indigo-700 font-extrabold uppercase rounded-lg text-[10px] tracking-wide">
                                        {shortage.craft}
                                      </span>
                                      
                                      {shortage.assigned === 0 && (
                                        <span className="px-1.5 py-0.5 bg-rose-100 text-rose-800 text-[8px] font-bold uppercase rounded border border-rose-200 tracking-tighter shrink-0 animate-pulse">
                                          ⚠️ Critical Zero-Staffing
                                        </span>
                                      )}
                                    </div>
                                    
                                    <div className="flex items-center gap-2 mt-2 text-gray-550 font-medium">
                                      <Calendar className="w-3 h-3 text-slate-400 shrink-0" />
                                      <span className="text-[10px] leading-none">
                                        Required Span: <b>{dayjs(shortage.start).format('DD MMM YY')}</b> <ArrowRight className="inline w-3 h-3 mx-0.5" /> <b>{dayjs(shortage.end).format('DD MMM YY')}</b>
                                      </span>
                                    </div>
                                  </div>

                                  {/* Visual Allocation Scale and Gap indicator */}
                                  <div className="flex items-center gap-6.5 shrink-0 select-none">
                                    {/* Visual scale */}
                                    <div className="flex flex-col items-end">
                                      <div className="flex items-center gap-2">
                                        <span className="text-[10.5px] font-semibold text-slate-500">Allocation Metric:</span>
                                        <span className="text-[10.5px] font-bold text-slate-800">{shortage.assigned} / {shortage.required}</span>
                                      </div>
                                      <div className="w-24 h-1.5 bg-slate-100 rounded-full mt-1.5 overflow-hidden border border-slate-200/50">
                                        <div 
                                          className={cn(
                                            "h-full rounded-full transition-all",
                                            shortage.assigned === 0 ? "bg-rose-450" : "bg-indigo-500"
                                          )}
                                          style={{ width: `${(shortage.assigned / shortage.required) * 100}%` }}
                                        />
                                      </div>
                                    </div>

                                    {/* Precise Gap Indicator */}
                                    <div className="bg-rose-50 border border-rose-150 px-3 py-1.5 rounded-xl flex items-center gap-1 text-center shrink-0">
                                      <ShieldAlert className="w-4 h-4 text-rose-600" />
                                      <span className="text-xs font-extrabold text-rose-700">
                                        -{shortage.gap} Staff Needed
                                      </span>
                                    </div>
                                  </div>

                                </div>
                              ))}
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

        {/* Full Fulfillment Welcome Card if shortages list is empty */}
        {shortages.length === 0 && (
          <div className="py-24 px-6 text-center">
            <div className="w-16 h-16 bg-emerald-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-emerald-100">
              <CheckCircle2 className="w-10 h-10 text-emerald-500" />
            </div>
            <h4 className="text-base font-extrabold text-[#1A1A1A] mb-1">Excellent: 100% Demand Satisfaction</h4>
            <p className="text-xs text-slate-500 max-w-sm mx-auto leading-relaxed">All specified project craft schedules are currently fully staffed by personnel in the active assignment schedule. No talent gaps detected.</p>
          </div>
        )}
      </div>
    </div>
  );
}
