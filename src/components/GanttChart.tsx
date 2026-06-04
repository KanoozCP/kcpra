import { useState, useMemo } from 'react';
import dayjs from 'dayjs';
import minMax from 'dayjs/plugin/minMax';
import * as XLSX from 'xlsx';
import { Manpower, Project, Assignment } from '../types';
import { cn } from '../lib/utils';
import { getProjectActualStatus } from '../lib/dateUtils';
import { 
  ZoomIn, 
  ZoomOut, 
  Download, 
  ChevronDown, 
  ChevronRight, 
  Briefcase, 
  User, 
  Calendar,
  Layers
} from 'lucide-react';

dayjs.extend(minMax);

interface Props {
  manpower: Manpower[];
  projects: Project[];
  assignments: Assignment[];
}

export default function GanttView({ manpower, projects, assignments }: Props) {
  const [zoomScale, setZoomScale] = useState(1);
  const [timelineMode, setTimelineMode] = useState<'daily' | 'weekly' | 'monthly'>('weekly');
  const [expandedProjects, setExpandedProjects] = useState<Record<string, boolean>>({});
  const [sortBy, setSortBy] = useState<'startDate' | 'status' | 'craft' | 'type'>('startDate');

  // Default collapse/expand states: default all to expanded for easy discoverability
  const toggleProject = (projectId: string) => {
    setExpandedProjects(prev => ({
      ...prev,
      [projectId]: !prev[projectId]
    }));
  };

  const expandAll = () => {
    const expanded: Record<string, boolean> = {};
    projects.forEach(p => { expanded[p.id] = true; });
    setExpandedProjects(expanded);
  };

  const collapseAll = () => {
    setExpandedProjects({});
  };

  const handleExport = () => {
    const data = assignments.map(a => {
      const worker = manpower.find(m => m.id === a.workerId);
      const project = projects.find(p => p.id === a.projectId);
      return {
        'Project Code': project?.code || '',
        'Project Name': project?.name || 'Unknown',
        'Worker Code': worker?.badgeNo || '',
        'Worker Name': worker?.name || 'Unknown',
        'Craft': a.craft,
        'Rating': worker?.strength || 'Good',
        'Phase': a.phase,
        'Start Date': a.startDate,
        'End Date': a.endDate
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Gantt Planning");
    XLSX.writeFile(wb, `KCP_Gantt_Timeline_${dayjs().format('YYYY-MM-DD')}.xlsx`);
  };

  // Determine date range for timeline
  const { startDate, endDate, days, weeks, months } = useMemo(() => {
    // Standard default calendar range
    let start = dayjs().startOf('month');
    let end = dayjs().add(2, 'month').endOf('month');

    // Expand boundaries dynamically based on projects or assignments
    const projectDates = projects.flatMap(p => [dayjs(p.startDate), dayjs(p.endDate)]);
    const assignmentDates = assignments.flatMap(a => [dayjs(a.startDate), dayjs(a.endDate)]);
    const allDates = [...projectDates, ...assignmentDates].filter(d => d.isValid());

    if (allDates.length > 0) {
      const actualStart = dayjs.min(allDates)!;
      const actualEnd = dayjs.max(allDates)!;
      if (actualStart.isBefore(start)) start = actualStart.startOf('month');
      if (actualEnd.isAfter(end)) end = actualEnd.endOf('month').add(7, 'days'); // subtle breathing padding
    }

    const diffDays = end.diff(start, 'day');
    const dayList = Array.from({ length: diffDays + 1 }, (_, i) => start.add(i, 'day'));
    
    // Groups for different modes
    const weekList: dayjs.Dayjs[] = [];
    let curW = start.startOf('week');
    while (curW.isBefore(end) || curW.isSame(end, 'day')) {
      weekList.push(curW);
      curW = curW.add(1, 'week');
    }

    const monthList: dayjs.Dayjs[] = [];
    let curM = start.startOf('month');
    while (curM.isBefore(end) || curM.isSame(end, 'day')) {
      monthList.push(curM);
      curM = curM.add(1, 'month');
    }

    return { 
      startDate: start, 
      endDate: end, 
      days: dayList,
      weeks: weekList,
      months: monthList
    };
  }, [projects, assignments]);

  // Construct hierarchal WBS projectwise visible rows
  const visibleRows = useMemo(() => {
    const list: any[] = [];
    
    // 1. Sort projects
    const sortedProjects = [...projects].sort((a, b) => {
      if (sortBy === 'status') {
        const statusA = getProjectActualStatus(a);
        const statusB = getProjectActualStatus(b);
        if (statusA !== statusB) {
          return statusA.localeCompare(statusB);
        }
      }
      return dayjs(a.startDate).valueOf() - dayjs(b.startDate).valueOf();
    });

    sortedProjects.forEach(project => {
      // Find all assignments for this specific project and sort them
      let projectAssignments = assignments.filter(a => a.projectId === project.id);
      
      if (sortBy === 'craft') {
        projectAssignments = [...projectAssignments].sort((a, b) => {
          return a.craft.localeCompare(b.craft);
        });
      } else if (sortBy === 'type') {
        projectAssignments = [...projectAssignments].sort((a, b) => {
          const workerA = manpower.find(m => m.id === a.workerId);
          const workerB = manpower.find(m => m.id === b.workerId);
          const typeA = workerA?.employmentType || '';
          const typeB = workerB?.employmentType || '';
          return typeA.localeCompare(typeB);
        });
      }

      const isExpanded = !!expandedProjects[project.id];

      list.push({
        type: 'project',
        id: project.id,
        projectId: project.id,
        label: project.name,
        code: project.code,
        startDate: project.startDate,
        endDate: project.endDate,
        workerCount: projectAssignments.length,
        isExpanded
      });

      if (isExpanded) {
        projectAssignments.forEach(a => {
          const worker = manpower.find(m => m.id === a.workerId);
          list.push({
            type: 'worker',
            id: `WROW-${project.id}-${a.id}`,
            projectId: project.id,
            assignment: a,
            worker,
            label: worker?.name || 'To Be Announced',
            code: worker ? `${worker.badgeNo} • ${a.craft}` : `TBA • ${a.craft}`,
            startDate: a.startDate,
            endDate: a.endDate,
            phase: a.phase,
            craft: a.craft
          });
        });
      }
    });

    return list;
  }, [projects, assignments, manpower, expandedProjects, sortBy]);

  // View cell widths
  const baseCellWidth = timelineMode === 'daily' ? 44 : timelineMode === 'weekly' ? 100 : 200;
  const cellWidth = baseCellWidth * zoomScale;
  const labelWidth = 260;

  const currentTimeline = timelineMode === 'daily' ? days : timelineMode === 'weekly' ? weeks : months;

  // Helpers to calculate relative Gantt bar layouts
  const getGanttBarLayout = (startStr: string, endStr: string) => {
    const start = dayjs(startStr);
    const end = dayjs(endStr);
    
    // Check out of bounds
    if (end.isBefore(startDate) || start.isAfter(endDate)) {
      return { show: false, left: 0, width: 0 };
    }

    const activeStart = start.isBefore(startDate) ? startDate : start;
    const activeEnd = end.isAfter(endDate) ? endDate : end;

    const diffStart = activeStart.diff(startDate, 'day');
    const diffSpan = activeEnd.diff(activeStart, 'day') + 1;
    
    // Pixel sizing
    const pixelPerDay = timelineMode === 'daily' ? cellWidth : timelineMode === 'weekly' ? cellWidth / 7 : cellWidth / 30;

    return {
      show: true,
      left: diffStart * pixelPerDay,
      width: diffSpan * pixelPerDay
    };
  };

  return (
    <div className="space-y-6 font-sans">
      <div className="bg-white rounded-2xl border border-[#E5E5E5] p-5 shadow-sm flex flex-col gap-4 mb-6">
        <div>
          <h3 className="text-base font-bold text-[#1A1A1A]">Gantt Timeline</h3>
          <p className="text-xs text-gray-500 mt-0.5">Visual scheduling, deployment durations, and crew coordination timeline</p>
        </div>
        
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 border-t border-slate-100 pt-3">
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Collapse Actions */}
            <div className="flex items-center gap-1 border-r border-[#E5E5E5] pr-2.5 mr-0.5 shrink-0">
              <button 
                onClick={expandAll}
                className="flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-bold text-slate-655 border border-slate-250 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                title="Expand All"
              >
                Expand All
              </button>
              <button 
                onClick={collapseAll}
                className="flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-bold text-slate-655 border border-slate-250 bg-white rounded-lg hover:bg-slate-50 transition-colors shadow-3xs"
                title="Collapse All"
              >
                Collapse All
              </button>
            </div>

            <div className="flex items-center gap-1.5 shrink-0">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest block leading-none">Resolution:</span>
              <div className="flex p-0.5 bg-gray-100 border border-[#E5E5E5] rounded-xl shrink-0">
                {(['daily', 'weekly', 'monthly'] as const).map(mode => (
                  <button 
                    key={mode}
                    onClick={() => setTimelineMode(mode)}
                    className={cn(
                      "px-2.5 py-1 rounded-lg text-[9px] uppercase font-bold transition-all",
                      timelineMode === mode ? "bg-white shadow-3xs text-indigo-700 font-extrabold" : "text-[#555]"
                    )}
                  >
                    {mode}
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <div className="flex items-center gap-1.5 shrink-0">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest block leading-none">Sort:</span>
              <select
                value={sortBy}
                onChange={(e) => setSortBy(e.target.value as any)}
                className="px-2.5 py-1.5 bg-white border border-[#E5E5E5] rounded-xl text-[11px] font-bold text-[#333] outline-none cursor-pointer hover:bg-slate-50"
              >
                <option value="startDate">Project Start (Asc)</option>
                <option value="status">Project Status (Asc)</option>
                <option value="craft">Crew Craft (A-Z)</option>
                <option value="type">Employment Type</option>
              </select>
            </div>

            <button 
              onClick={handleExport}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-[#444] border border-[#E5E5E5] rounded-lg bg-white hover:bg-slate-50 transition-colors shrink-0"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              Export Excel
            </button>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-[#E5E5E5] shadow-sm overflow-hidden flex flex-col h-[650px] font-sans">
        {/* Legends & Zoom Controls */}
        <div className="p-3 border-b border-[#E5E5E5] flex items-center justify-between shrink-0 bg-slate-50/50">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded shadow-3xs bg-amber-100 border border-amber-300" />
              <span className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">Pre-TA</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded shadow-3xs bg-rose-100 border border-rose-300" />
              <span className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">TA Phase</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded shadow-3xs bg-emerald-100 border border-emerald-300" />
              <span className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">Post-TA</span>
            </div>
          </div>

          <div className="flex items-center gap-1">
            <button 
              onClick={() => setZoomScale(s => Math.min(s + 0.1, 2.5))}
              className="p-1.5 hover:bg-gray-150 rounded-lg transition-colors border border-[#E5E5E5] shrink-0 bg-white"
              title="Zoom In"
            >
              <ZoomIn className="w-3.5 h-3.5 text-gray-650" />
            </button>
            <button 
              onClick={() => setZoomScale(s => Math.max(s - 0.1, 0.6))}
              className="p-1.5 hover:bg-gray-150 rounded-lg transition-colors border border-[#E5E5E5] shrink-0 bg-white"
              title="Zoom Out"
            >
              <ZoomOut className="w-3.5 h-3.5 text-gray-650" />
            </button>
          </div>
        </div>

        {/* Gantt Timeline Container */}
        <div className="flex-1 overflow-auto relative">
        <div className="min-w-max">
          
          {/* Timeline Grid Header (Sticky Top) */}
          <div className="flex border-b border-[#F0F0F0] sticky top-0 bg-white z-20">
            <div 
              className="shrink-0 border-r border-[#F0F0F0] p-3 bg-slate-50 flex items-center shadow-xs"
              style={{ width: labelWidth }}
            >
              <span className="text-[9px] font-extrabold text-[#666] uppercase tracking-wider">PROJECT-BASED DEPLOYMENT HIERARCHY</span>
            </div>
            <div className="flex bg-white">
              {currentTimeline.map((item, i) => {
                const isSunday = item.day() === 0;
                let label = '';
                let subLabel = '';
                
                if (timelineMode === 'daily') {
                  label = item.format('ddd').toUpperCase();
                  subLabel = item.format('D');
                } else if (timelineMode === 'weekly') {
                  label = 'WK ' + item.format('WW');
                  subLabel = item.format('DD MMM');
                } else {
                  label = item.format('MMM').toUpperCase();
                  subLabel = item.format('YYYY');
                }

                return (
                  <div 
                    key={i} 
                    className={cn(
                      "flex flex-col items-center justify-center py-2 border-r border-[#F0F0F0] shrink-0 relative",
                      timelineMode === 'daily' && isSunday && "bg-gray-150/40"
                    )}
                    style={{ width: cellWidth }}
                  >
                    {timelineMode === 'daily' && item.date() === 1 && (
                      <span className="absolute -top-px left-0 px-1 py-0.5 bg-indigo-600 text-[8px] font-bold text-white uppercase rounded-b-sm z-10 whitespace-nowrap">
                        {item.format('MMM YY')}
                      </span>
                    )}
                    <span className={cn("text-[8px] font-bold", timelineMode === 'daily' && isSunday ? "text-red-500" : "text-[#888]")}>
                      {label}
                    </span>
                    <span className="text-[10px] font-extrabold text-[#1A1A1A] mt-0.5">{subLabel}</span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Timeline Gantt Grid rows */}
          <div className="relative divide-y divide-slate-100">
            {visibleRows.map((row) => {
              const isProj = row.type === 'project';
              
              return (
                <div key={row.id} className={cn(
                  "flex transition-colors group", 
                  isProj ? "bg-[#FAFAFB]/40 hover:bg-[#F3F4F6]/50" : "bg-white hover:bg-slate-50/70"
                )}>
                  
                  {/* Left Column: Label */}
                  <div 
                    className="shrink-0 border-r border-[#F0F0F0] p-3 flex flex-col justify-center select-none"
                    style={{ width: labelWidth }}
                  >
                    {isProj ? (
                      // Project Row Title
                      <div 
                        className="flex items-center gap-1.5 cursor-pointer"
                        onClick={() => toggleProject(row.id)}
                      >
                        <div className="p-0.5 hover:bg-slate-200 rounded text-slate-500 transition-colors shrink-0">
                          {row.isExpanded ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
                        </div>
                        <div className="min-w-0">
                          <h4 className="text-xs font-bold text-slate-900 truncate leading-tight flex items-center gap-1.5">
                            <Briefcase className="w-3 h-3 text-indigo-500 shrink-0" />
                            <span className="truncate">{row.label}</span>
                            {(() => {
                              const proj = projects.find(p => p.id === row.projectId);
                              if (!proj) return null;
                              const status = getProjectActualStatus(proj);
                              let dotColor = '';
                              switch (status) {
                                case 'In Progress': dotColor = 'bg-emerald-500'; break;
                                case 'Completed': dotColor = 'bg-blue-500'; break;
                                case 'Hold': dotColor = 'bg-amber-500'; break;
                                case 'Rescheduled': dotColor = 'bg-indigo-500'; break;
                                case 'Cancelled': dotColor = 'bg-rose-500'; break;
                              }
                              return (
                                <span className="flex items-center gap-1 text-[7.5px] font-extrabold uppercase tracking-wide text-slate-500 shrink-0 ml-1" title={status}>
                                  <span className={cn("w-1.5 h-1.5 rounded-full inline-block shrink-0", dotColor)} />
                                  {status.replace(/ \(Manual\)/gi, '')}
                                </span>
                              );
                            })()}
                          </h4>
                          <div className="flex items-center gap-1.5 mt-1">
                            <span className="text-[8.5px] font-bold bg-indigo-50 border border-indigo-100/80 text-indigo-700 px-1 rounded uppercase tracking-tighter">
                              {row.code}
                            </span>
                            <span className="text-[9px] text-[#888] font-bold">
                              ({row.workerCount} deployed)
                            </span>
                          </div>
                        </div>
                      </div>
                    ) : (
                      // Nested Assigned Worker Row
                      <div className="pl-6 flex items-start gap-1.5 min-w-0">
                        <span className="text-gray-400 font-mono text-xs select-none mt-0.5">↳</span>
                        <div className="min-w-0 leading-tight">
                          <p className="text-xs font-semibold text-slate-800 truncate flex items-center gap-1">
                            <User className="w-3 h-3 text-slate-400 shrink-0" />
                            {row.label}
                          </p>
                          <p className="text-[9px] font-bold text-indigo-700 uppercase tracking-tight truncate mt-0.5">
                            {row.code}
                          </p>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Right Column: Timeline visualization grid */}
                  <div className="flex relative h-11">
                    
                    {/* Background Grid partition columns */}
                    {currentTimeline.map((item, i) => (
                      <div 
                        key={i} 
                        className={cn(
                          "h-full border-r border-[#F3F4F6] shrink-0",
                          timelineMode === 'daily' && item.day() === 0 && "bg-gray-150/15"
                        )} 
                        style={{ width: cellWidth }} 
                      />
                    ))}
                    
                    {/* Timeline Data Bars */}
                    {(() => {
                      if (isProj) {
                        // For Project Parent: Render a subtle, elegant project milestone tracker bar
                        const layout = getGanttBarLayout(row.startDate, row.endDate);
                        if (!layout.show) return null;

                        return (
                          <div 
                            className="absolute top-3.5 h-4.5 rounded bg-slate-200/50 border border-slate-300 flex items-center justify-between px-2 overflow-hidden pointer-events-none select-none"
                            style={{ 
                              left: layout.left, 
                              width: layout.width,
                              zIndex: 5 
                            }}
                            title={`Project: ${row.startDate} to ${row.endDate}`}
                          >
                            <span className="text-[7.5px] font-extrabold text-slate-500 tracking-wider uppercase truncate">
                              OVERALL PROJECT SPAN
                            </span>
                            <span className="text-[7.5px] font-bold text-slate-400 shrink-0">
                              {dayjs(row.startDate).format('DD MMM')} - {dayjs(row.endDate).format('DD MMM')}
                            </span>
                          </div>
                        );
                      } else {
                        // For Assigned Worker Child: Render their specific deployment interval colored by Phase
                        const layout = getGanttBarLayout(row.startDate, row.endDate);
                        if (!layout.show) return null;

                        return (
                          <div 
                            className={cn(
                              "absolute top-2 h-7 rounded-lg shadow-xs border flex items-center justify-between px-3 overflow-hidden cursor-pointer hover:brightness-105 hover:shadow-sm transition-all group/bar",
                              row.phase === 'TA' ? 'bg-rose-100 text-rose-800 border-rose-250' : 
                              row.phase === 'Pre-TA' ? 'bg-amber-100 text-amber-800 border-amber-250' : 
                              'bg-emerald-100/90 text-emerald-800 border-emerald-300'
                            )}
                            style={{ 
                              left: layout.left, 
                              width: layout.width,
                              zIndex: 10 
                            }}
                          >
                            <span className="text-[9.5px] font-extrabold whitespace-nowrap truncate uppercase tracking-tighter">
                              {row.craft} ({row.phase})
                            </span>
                            
                            {/* Hover info overlay summary tooltip */}
                            <div className="hidden group-hover/bar:flex absolute left-0 bottom-full mb-1.5 p-2 bg-[#1E293B] text-white text-[9px] rounded-lg pointer-events-none z-50 flex-col gap-0.5 shadow-lg font-sans border border-slate-700 w-max leading-none">
                              <p className="font-extrabold text-indigo-400">{row.label}</p>
                              <p className="font-bold">{row.craft} • {row.phase}</p>
                              <p className="font-semibold text-slate-300 mt-1">{dayjs(row.startDate).format('DD MMM YY')} — {dayjs(row.endDate).format('DD MMM YY')}</p>
                            </div>
                          </div>
                        );
                      }
                    })()}

                  </div>
                </div>
              );
            })}
          </div>

        </div>
      </div>
      
    </div>
  </div>
  );
}
