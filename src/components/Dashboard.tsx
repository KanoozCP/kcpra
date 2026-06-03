import React from 'react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  PieChart, 
  Pie, 
  Cell, 
  Legend,
  AreaChart,
  Area,
  ComposedChart,
  Line,
  ReferenceLine
} from 'recharts';
import { 
  Users, 
  Briefcase, 
  Calendar, 
  AlertTriangle,
  Award,
  CircleDot,
  CheckCircle2,
  Percent,
  TrendingUp,
  Activity,
  UserCheck,
  Plane
} from 'lucide-react';
import { Manpower, Project, Assignment } from '../types';
import dayjs from 'dayjs';

interface Props {
  manpower: Manpower[];
  projects: Project[];
  assignments: Assignment[];
}

export default function Dashboard({ manpower, projects, assignments }: Props) {
  const todayStr = '2026-05-21';
  const today = dayjs(todayStr);

  // 1. CALCULATE CORE METRICS AND STATS
  const totalPool = manpower.length;
  
  // Vacation count: workers on vacation today
  const vacationCount = manpower.filter(m => {
    if (m.vacationStart && m.vacationEnd) {
      const vStart = dayjs(m.vacationStart);
      const vEnd = dayjs(m.vacationEnd);
      return vStart.isValid() && vEnd.isValid() && today.isBetween(vStart, vEnd, 'day', '[]');
    }
    return false;
  }).length;

  const deployedCount = assignments.length;
  const idleCount = Math.max(0, totalPool - deployedCount - vacationCount);
  const utilizationRate = totalPool > 0 ? Math.round((deployedCount / totalPool) * 100) : 0;

  // Calculate Urgent Staffing Deficits/Gaps across all projects
  const totalDemandQty = projects.reduce((acc, p) => {
    return acc + p.requirements.reduce((pAcc, req) => pAcc + req.qty, 0);
  }, 0);

  const totalGaps = projects.reduce((acc, p) => {
    const projectGaps = p.requirements.reduce((pAcc, req) => {
      const assignedToReq = assignments.filter(a => 
        a.projectId === p.id && 
        a.craft === req.craft && 
        a.phase === req.phase
      ).length;
      return pAcc + Math.max(0, req.qty - assignedToReq);
    }, 0);
    return acc + projectGaps;
  }, 0);

  const overallRequirementsFulfilled = totalDemandQty > 0 
    ? Math.round(((totalDemandQty - totalGaps) / totalDemandQty) * 100) 
    : 100;

  // Stats definition for rendering
  const stats = [
    { 
      label: 'Total workforce size', 
      value: totalPool, 
      sub: `${deployedCount} assigned, ${idleCount} available`,
      icon: Users, 
      color: 'bg-indigo-50 text-indigo-700 border border-indigo-100',
      subColor: 'text-indigo-600'
    },
    { 
      label: 'Active deployments', 
      value: deployedCount, 
      sub: `${utilizationRate}% deployment utilization`,
      icon: UserCheck, 
      color: 'bg-emerald-50 text-emerald-700 border border-emerald-100',
      subColor: 'text-emerald-600'
    },
    { 
      label: 'Deficits / Missing Personnel', 
      value: totalGaps, 
      sub: totalGaps > 0 ? `${totalGaps} positions unfilled` : 'All sites fully staffed',
      icon: AlertTriangle, 
      color: totalGaps > 0 
        ? 'bg-rose-50 text-rose-700 border border-rose-100' 
        : 'bg-emerald-50 text-emerald-700 border border-emerald-100',
      subColor: totalGaps > 0 ? 'text-rose-600' : 'text-emerald-600'
    },
    { 
      label: 'Fulfilling rate', 
      value: `${overallRequirementsFulfilled}%`, 
      sub: `${totalDemandQty - totalGaps} of ${totalDemandQty} requested placed`,
      icon: Activity, 
      color: 'bg-purple-50 text-purple-700 border border-purple-100',
      subColor: 'text-purple-600'
    },
  ];

  // 2. PREPARE GORGEOUS AND HIGHLY VALUE-ADD CHART DATASETS
  
  // A. Top Projects Demand vs Assigned headcount
  const projectStaffingData = projects.map(p => {
    const required = p.requirements.reduce((sum, r) => sum + r.qty, 0);
    const assigned = assignments.filter(a => a.projectId === p.id).length;
    const gap = Math.max(0, required - assigned);
    const rate = required > 0 ? Math.round((assigned / required) * 100) : 100;

    return {
      name: p.name,
      code: p.code,
      Required: required,
      Assigned: assigned,
      Gap: gap,
      Rate: rate
    };
  }).sort((a, b) => b.Required - a.Required).slice(0, 6); // Look at top 6 projects

  // B. Staffing Deficit / Missing headcount by projects (Highly Actionable!)
  const projectGapsData = projectStaffingData
    .filter(p => p.Gap > 0)
    .sort((a, b) => b.Gap - a.Gap);

  // C. Bench Availability and Assigned Breakdown by Craft
  const allCraftsSet = new Set<string>();
  manpower.forEach(m => { if (m.craft) allCraftsSet.add(m.craft); });
  projects.forEach(p => { p.requirements.forEach(req => { if (req.craft) allCraftsSet.add(req.craft); }); });

  const craftChartData = Array.from(allCraftsSet).map(craft => {
    const totalSelected = manpower.filter(m => m.craft === craft);
    const supply = totalSelected.length;
    
    // Deployed for this craft
    const deployed = assignments.filter(a => a.craft === craft).length;
    
    // Vacation count for this craft
    const vacation = totalSelected.filter(m => {
      if (m.vacationStart && m.vacationEnd) {
        const vStart = dayjs(m.vacationStart);
        const vEnd = dayjs(m.vacationEnd);
        return vStart.isValid() && vEnd.isValid() && today.isBetween(vStart, vEnd, 'day', '[]');
      }
      return false;
    }).length;

    const idle = Math.max(0, supply - deployed - vacation);

    // Total Project Demand for this craft
    const demand = projects.reduce((acc, p) => {
      const match = p.requirements.filter(r => r.craft === craft).reduce((s, r) => s + r.qty, 0);
      return acc + match;
    }, 0);

    return {
      craft,
      Supply: supply,
      Deployed: deployed,
      IdleBench: idle,
      Vacation: vacation,
      Demand: demand,
      Deficit: Math.max(0, demand - deployed)
    };
  })
  .sort((a, b) => b.Demand - a.Demand || b.Supply - a.Supply)
  .slice(0, 8); // top 8 crafts

  // D. Workforce Employment sponsorship types
  const directCount = manpower.filter(m => m.employmentType === 'Direct').length;
  const qiwaCount = manpower.filter(m => m.employmentType === 'Qiwa').length;
  const localHireCount = manpower.filter(m => m.employmentType === 'Local Hire').length;

  const sponsorshipData = [
    { name: 'Direct', value: directCount, color: '#4F46E5' },
    { name: 'Qiwa', value: qiwaCount, color: '#3B82F6' },
    { name: 'Local Hire', value: localHireCount, color: '#10B981' },
  ].filter(c => c.value > 0);

  // E. Talent Quality Rating Distribution
  const ratingCounts = manpower.reduce((acc: Record<string, number>, m) => {
    const rating = m.strength || 'Good';
    acc[rating] = (acc[rating] || 0) + 1;
    return acc;
  }, {});

  const qualityData = [
    { name: 'Excellent Performance', value: ratingCounts['Excellent'] || 0, color: '#8B5CF6' },
    { name: 'Good Performance', value: ratingCounts['Good'] || 0, color: '#EC4899' },
    { name: 'Average Performance', value: ratingCounts['Average'] || 0, color: '#F59E0B' },
  ].filter(q => q.value > 0);

  // F. Projects Staffing completion percentages
  const staffingProgressData = projectStaffingData.map(p => ({
    name: p.name,
    Completion: p.Rate
  })).sort((a, b) => b.Completion - a.Completion);

  // 3. CHART STYLING CONSTANTS (WHITE BACKGROUND DEFAULTS)
  const gridColor = '#F3F4F6';
  const labelColor = '#4B5563';
  const tooltipStyle = {
    backgroundColor: '#FFFFFF',
    borderColor: '#E2E8F0',
    color: '#1F2937',
    borderRadius: '12px',
    boxShadow: '0 10px 30px rgba(0, 0, 0, 0.04)',
    fontSize: '12px',
    padding: '10px'
  };

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-12">
      
      {/* 1. Header Information */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white p-6 rounded-[24px] border border-slate-200/80 shadow-[0_8px_30px_rgb(0,0,0,0.02)] relative overflow-hidden">
        <div>
          <h2 className="text-lg font-bold text-slate-900">Workforce Analytics & Resource Insights</h2>
          <p className="text-xs text-slate-500 mt-1">
            Real-time status updates, critical staffing gaps, and interactive allocation forecasts based on registered project requirements.
          </p>
        </div>
        <div className="flex items-center gap-3 self-start sm:self-center shrink-0">
          <span className="flex items-center gap-1.5 text-xs text-indigo-700 bg-indigo-50 font-bold px-3 py-1.5 rounded-lg border border-indigo-100">
            <Calendar className="w-3.5 h-3.5" />
            As of May 21, 2026
          </span>
          <span className="flex items-center gap-1 text-xs text-emerald-700 bg-emerald-50 font-bold px-3 py-1.5 rounded-lg border border-emerald-100">
            <CheckCircle2 className="w-3.5 h-3.5" />
            Active Sync
          </span>
        </div>
      </div>

      {/* 2. Stats KPI Ribbon */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 no-print">
        {stats.map((s, idx) => (
          <div key={idx} className="bg-white p-5 rounded-[24px] border border-slate-200/80 flex items-center justify-between shadow-[0_8px_30px_rgb(0,0,0,0.02)] transition-transform hover:translate-y-[-2px] duration-150">
            <div className="space-y-1">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{s.label}</p>
              <h4 className="text-2xl font-black text-slate-900 tracking-tight">{s.value}</h4>
              <p className={`text-[11px] font-bold ${s.subColor}`}>{s.sub}</p>
            </div>
            <div className={`w-12 h-12 rounded-xl flex items-center justify-center ${s.color}`}>
              <s.icon className="w-5 h-5 animate-pulse-slow" />
            </div>
          </div>
        ))}
      </div>

      {/* 3. Primary row: Project staffing review & Gap Deficits */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        {/* Chart 1: Project Allocation Details */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Demand</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Demanded quantities compared directly to current active headcount</p>
              </div>
              <span className="text-[10px] bg-slate-100 text-slate-600 font-bold px-2 py-0.5 rounded uppercase">Quantity Metrics</span>
            </div>
          </div>
          
          <div className="h-[280px] w-full mt-2">
            {projectStaffingData.length === 0 ? (
              <div className="h-full flex items-center justify-center text-xs text-slate-400 italic">
                No active requirements loaded
              </div>
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart 
                  data={projectStaffingData} 
                  barGap={8}
                  margin={{ left: -10, right: 10, top: 10, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={gridColor} />
                  <XAxis 
                    dataKey="name" 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tick={{ fill: labelColor }}
                    tickFormatter={(val) => val.length > 12 ? `${val.slice(0, 10)}...` : val}
                  />
                  <YAxis 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tick={{ fill: labelColor }}
                  />
                  <Tooltip 
                    contentStyle={tooltipStyle}
                    cursor={{ fill: 'rgba(79, 70, 229, 0.02)' }}
                  />
                  <Legend 
                    verticalAlign="top" 
                    height={32} 
                    iconSize={10} 
                    iconType="circle"
                    wrapperStyle={{ fontSize: '11px', color: labelColor }} 
                  />
                  <Bar dataKey="Required" name="Staff Requested" fill="#E2E8F0" radius={[4, 4, 0, 0]} barSize={16} />
                  <Bar dataKey="Assigned" name="Staff Allocated" fill="#4F46E5" radius={[4, 4, 0, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>

        {/* Chart 2: Urgent Deficits / Gaps by Project */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Shortage</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Missing headcount placements grouped by site (Action Required)</p>
              </div>
              <span className="text-[10px] bg-rose-50 text-rose-700 font-bold px-2 py-0.5 rounded uppercase">Crisis Tracker</span>
            </div>
          </div>

          <div className="h-[280px] w-full mt-2">
            {projectGapsData.length === 0 ? (
              <div className="h-full flex flex-col items-center justify-center text-slate-400 italic space-y-2">
                <CheckCircle2 className="w-8 h-8 text-emerald-500" />
                <span className="text-xs">Zero Staffing Deficits! All project requests are fully fulfilled.</span>
              </div>
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart 
                  data={projectGapsData}
                  margin={{ left: -10, right: 10, top: 10, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={gridColor} />
                  <XAxis 
                    dataKey="name" 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tick={{ fill: labelColor }}
                    tickFormatter={(val) => val.length > 12 ? `${val.slice(0, 10)}...` : val}
                  />
                  <YAxis 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tick={{ fill: labelColor }}
                    allowDecimals={false}
                  />
                  <Tooltip 
                    contentStyle={tooltipStyle}
                    cursor={{ fill: 'rgba(239, 68, 68, 0.02)' }}
                  />
                  <Bar dataKey="Gap" name="Missing Headcount" fill="#EF4444" radius={[4, 4, 0, 0]} barSize={24} />
                </BarChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>

      </div>

      {/* 4. Secondary row: Bench Status by Craft & Project Success Rate */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        {/* Chart 3: Active vs Bench pool by Craft */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Deployment vs Available MP</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Assigned workers, available workers, and vacationers</p>
              </div>
              <span className="text-[10px] bg-emerald-50 text-emerald-700 font-bold px-2 py-0.5 rounded uppercase">Available Level</span>
            </div>
          </div>

          <div className="h-[280px] w-full mt-2">
            {craftChartData.length === 0 ? (
              <div className="h-full flex items-center justify-center text-xs text-slate-400 italic">
                No craft resources to display
              </div>
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart 
                  data={craftChartData} 
                  layout="vertical"
                  margin={{ left: 10, right: 15, top: 10, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke={gridColor} />
                  <XAxis type="number" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: labelColor }} />
                  <YAxis 
                    dataKey="craft" 
                    type="category" 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false} 
                    width={100}
                    tick={{ fill: labelColor }}
                  />
                  <Tooltip contentStyle={tooltipStyle} />
                  <Legend 
                    verticalAlign="top" 
                    height={32} 
                    iconSize={8} 
                    iconType="circle"
                    wrapperStyle={{ fontSize: '11px', color: labelColor }} 
                  />
                  <Bar dataKey="Deployed" name="Assigned & Deployed" fill="#4F46E5" stackId="a" barSize={14} />
                  <Bar dataKey="IdleBench" name="Available" fill="#10B981" stackId="a" barSize={14} />
                  <Bar dataKey="Vacation" name="On Vacation" fill="#F59E0B" stackId="a" barSize={14} />
                </BarChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>

        {/* Chart 4: Project Staffing Progress Meter */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Project Allocation Rate</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Percentage of project staffing demands successfully met</p>
              </div>
              <span className="text-[10px] bg-indigo-50 text-indigo-700 font-bold px-2 py-0.5 rounded uppercase">KPI Ratios</span>
            </div>
          </div>

          <div className="h-[280px] w-full mt-2 flex flex-col justify-between">
            {staffingProgressData.length === 0 ? (
              <div className="h-full flex items-center justify-center text-xs text-slate-400 italic">
                No loaded projects
              </div>
            ) : (
              <div className="flex flex-col justify-center h-full space-y-4">
                {staffingProgressData.map((proj, idx) => {
                  let barColor = 'bg-emerald-500';
                  if (proj.Completion < 50) barColor = 'bg-rose-500';
                  else if (proj.Completion < 85) barColor = 'bg-amber-500';

                  return (
                    <div key={idx} className="space-y-1.5">
                      <div className="flex justify-between text-xs font-semibold text-slate-700">
                        <span className="truncate max-w-[70%]">{proj.name}</span>
                        <span className="font-bold">{proj.Completion}% Fulfilled</span>
                      </div>
                      <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden relative">
                        <div 
                          className={`h-full ${barColor} rounded-full transition-all duration-500`} 
                          style={{ width: `${Math.min(100, proj.Completion)}%` }} 
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>

      </div>

      {/* 5. Tertiary row: Sponsorship & Talent Quality Distr */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        {/* Chart 5: Sponsorship Models */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Workforce Composition (By Sponsorship Model)</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Distribution between Qiwa sponsorship, Direct hires, and Local Contractors</p>
              </div>
              <span className="text-[10px] bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded uppercase">Dispatch Composition</span>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-12 gap-6 items-center flex-1 mt-2">
            <div className="sm:col-span-6 h-[200px] relative flex justify-center items-center">
              {sponsorshipData.length === 0 ? (
                <div className="text-xs text-slate-400 italic">No hires recorded</div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={sponsorshipData}
                      innerRadius={55}
                      outerRadius={75}
                      paddingAngle={4}
                      dataKey="value"
                    >
                      {sponsorshipData.map((entry, idx) => (
                        <Cell key={`cell-${idx}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={tooltipStyle} />
                  </PieChart>
                </ResponsiveContainer>
              )}
              {/* Central Utility percentage badge */}
              <div className="absolute flex flex-col items-center justify-center">
                <span className="text-xs font-bold text-slate-400 uppercase tracking-widest leading-none">Total MP</span>
                <span className="text-xl font-black text-slate-800 leading-tight mt-1">{totalPool}</span>
              </div>
            </div>

            <div className="sm:col-span-6 space-y-3">
              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-2">Sponsorship breakdown</p>
              <div className="space-y-2">
                {sponsorshipData.map((item, idx) => (
                  <div key={idx} className="flex items-center justify-between text-xs p-2 rounded hover:bg-slate-50 transition-colors">
                    <div className="flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: item.color }} />
                      <span className="font-semibold text-slate-600">{item.name}</span>
                    </div>
                    <div>
                      <b className="text-slate-900">{item.value}</b>
                      <span className="text-[10px] text-slate-400 ml-1">({Math.round((item.value / (totalPool || 1)) * 105)}%)</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Chart 6: Performance Matrix */}
        <div className="bg-white p-6 rounded-[24px] border border-slate-200/85 shadow-[0_8px_30px_rgb(0,0,0,0.02)] hover:shadow-[0_12px_40px_rgb(0,0,0,0.04)] transition-all duration-300 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Workforce Talent Quality Distribution</h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Strength ratings breakdown according to supervisory evaluations</p>
              </div>
              <span className="text-[10px] bg-purple-50 text-purple-700 font-bold px-2 py-0.5 rounded uppercase">Talent Matrix</span>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-12 gap-6 items-center flex-1 mt-2">
            <div className="sm:col-span-6 h-[200px] relative flex justify-center items-center">
              {qualityData.length === 0 ? (
                <div className="text-xs text-slate-400 italic">No ratings provided</div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={qualityData}
                      innerRadius={0}
                      outerRadius={75}
                      paddingAngle={2}
                      dataKey="value"
                    >
                      {qualityData.map((entry, idx) => (
                        <Cell key={`cell-${idx}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={tooltipStyle} />
                  </PieChart>
                </ResponsiveContainer>
              )}
            </div>

            <div className="sm:col-span-6 space-y-3">
              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Supervisory Grades</p>
              <div className="space-y-2">
                {qualityData.map((item, idx) => (
                  <div key={idx} className="flex items-center justify-between text-xs p-2 rounded hover:bg-slate-50 transition-colors">
                    <div className="flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: item.color }} />
                      <span className="font-semibold text-slate-600">{item.name}</span>
                    </div>
                    <div>
                      <b className="text-slate-900">{item.value}</b>
                      <span className="text-[9px] text-slate-450 ml-1">({Math.round((item.value / (totalPool || 1)) * 100)}%)</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

      </div>

    </div>
  );
}
