import dayjs from 'dayjs';
import isBetween from 'dayjs/plugin/isBetween';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore';
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter';
import { Manpower, Project, Assignment, ProjectRequirement } from '../types';

dayjs.extend(isBetween);
dayjs.extend(isSameOrBefore);
dayjs.extend(isSameOrAfter);

const normalizeCraft = (c: string) => {
  if (!c) return '';
  return c.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
};

function interlaceByStrength(candidates: Manpower[]): Manpower[] {
  const excellent = candidates.filter(w => (w.strength || 'Good') === 'Excellent');
  const good = candidates.filter(w => (w.strength || 'Good') === 'Good');
  const average = candidates.filter(w => (w.strength || 'Good') === 'Average');
  
  const interlaced: Manpower[] = [];
  const maxLen = Math.max(excellent.length, good.length, average.length);
  for (let j = 0; j < maxLen; j++) {
    if (j < excellent.length) interlaced.push(excellent[j]);
    if (j < good.length) interlaced.push(good[j]);
    if (j < average.length) interlaced.push(average[j]);
  }
  return interlaced;
}

/**
 * The core assignment engine logic.
 * Tries to fulfill project requirements based on availability and craft.
 */
export function runAutoAssignment(
  manpower: Manpower[],
  projects: Project[]
): Assignment[] {
  const newAssignments: Assignment[] = [];
  let assignCounter = 0;
  
  // Sort projects by earliest start date
  const sortedProjects = [...projects].sort((a, b) => {
    const getEarliest = (p: Project) => {
      const validReqs = (p.requirements || []).filter(r => dayjs(r.startDate).isValid());
      if (validReqs.length > 0) {
        return Math.min(...validReqs.map(r => dayjs(r.startDate).valueOf()));
      }
      return dayjs(p.startDate).isValid() ? dayjs(p.startDate).valueOf() : Infinity;
    };
    return getEarliest(a) - getEarliest(b);
  });

  const workerSchedule: Record<string, { start: string; end: string }[]> = {};

  const isWorkerAvailable = (workerId: string, start: string, end: string) => {
    const s = dayjs(start);
    const e = dayjs(end);
    if (!s.isValid() || !e.isValid()) return false;

    const schedules = workerSchedule[workerId] || [];
    
    for (const schedule of schedules) {
      const existingS = dayjs(schedule.start);
      const existingE = dayjs(schedule.end);
      if (s.isBetween(existingS, existingE, 'day', '[]') || e.isBetween(existingS, existingE, 'day', '[]') || 
          existingS.isBetween(s, e, 'day', '[]')) {
        return false;
      }
    }
    return true;
  };

  const markWorkerBusy = (workerId: string, start: string, end: string) => {
    if (!workerSchedule[workerId]) workerSchedule[workerId] = [];
    workerSchedule[workerId].push({ start, end });
  };

  sortedProjects.forEach(project => {
    project.requirements.forEach(req => {
      const { qty, startDate: start, endDate: end, phase, craft } = req;

      if (qty <= 0 || !start || !end) return;

      // Filter available workers for this specific requirement
      const candidates = manpower.filter(w => {
        if (!w.craft || !craft) return false;
        if (normalizeCraft(w.craft) !== normalizeCraft(craft)) return false;
        
        const workerJoin = dayjs(w.joinDate);
        const workerRel = dayjs(w.releaseDate);
        const reqS = dayjs(start);
        const reqE = dayjs(end);

        // Date validity check
        if (!workerJoin.isValid() || !workerRel.isValid() || !reqS.isValid() || !reqE.isValid()) return false;

        // Worker must be within their service period
        if (!workerJoin.isSameOrBefore(reqS, 'day') || !workerRel.isSameOrAfter(reqE, 'day')) return false;

        // Check current session's timeline
        if (!isWorkerAvailable(w.id, start, end)) return false;

        // Check vacation
        if (w.vacationStart && w.vacationEnd) {
          const vacS = dayjs(w.vacationStart);
          const vacE = dayjs(w.vacationEnd);
          if (vacS.isValid() && vacE.isValid()) {
            // If request spans any part of vacation, skip
            if (reqS.isBetween(vacS, vacE, 'day', '[]') || reqE.isBetween(vacS, vacE, 'day', '[]') ||
                vacS.isBetween(reqS, reqE, 'day', '[]')) {
              return false;
            }
          }
        }

        return true;
      });

      const balancedCandidates = interlaceByStrength(candidates);

      // Assign workers up to required qty
      for (let i = 0; i < Math.min(qty, balancedCandidates.length); i++) {
        const worker = balancedCandidates[i];
        newAssignments.push({
          id: `ASSIGN-${Date.now()}-${assignCounter++}-${Math.random().toString(36).substring(2, 7)}`,
          workerId: worker.id,
          projectId: project.id,
          craft,
          phase,
          startDate: start,
          endDate: end,
        });
        markWorkerBusy(worker.id, start, end);
      }
    });
  });

  return newAssignments;
}

export function calculateShortages(manpower: Manpower[], projects: Project[], assignments: Assignment[]) {
  const shortages: any[] = [];

  projects.forEach(project => {
    project.requirements.forEach(req => {
      const { qty: requiredQty, startDate: start, endDate: end, phase, craft } = req;

      if (requiredQty <= 0) return;

      const assignedCount = assignments.filter(a => 
        a.projectId === project.id && 
        a.phase === phase && 
        normalizeCraft(a.craft) === normalizeCraft(craft)
      ).length;

      if (assignedCount < requiredQty) {
        shortages.push({
          projectId: project.id,
          projectCode: project.code,
          project: project.name,
          craft,
          phase,
          required: requiredQty,
          assigned: assignedCount,
          gap: requiredQty - assignedCount,
          start: start,
          end: end
        });
      }
    });
  });

  return shortages;
}
