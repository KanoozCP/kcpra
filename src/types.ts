export type Craft = string;

export type EmploymentType = 'Direct' | 'Qiwa' | 'Local Hire';

export interface Manpower {
  id: string;
  name: string;
  badgeNo: string;
  craft: Craft;
  joinDate: string;
  releaseDate: string;
  employmentType: EmploymentType;
  passportIqama?: string;
  vacationStart?: string;
  vacationEnd?: string;
  strength?: 'Good' | 'Average' | 'Excellent';
}

export type ProjectPhase = 'Pre-TA' | 'TA' | 'Post-TA';

export interface ProjectRequirement {
  id: string;
  craft: Craft;
  phase: ProjectPhase;
  qty: number;
  startDate: string;
  endDate: string;
}

export type ProjectStatus = 'Auto' | 'In Progress' | 'Completed' | 'Hold' | 'Rescheduled' | 'Cancelled';

export interface Project {
  id: string;
  name: string;
  code: string;
  location: string;
  department: string;
  startDate: string;
  endDate: string;
  requirements: ProjectRequirement[];
  status?: ProjectStatus;
}

export interface Assignment {
  id: string;
  workerId: string;
  projectId: string;
  craft: Craft;
  phase: ProjectPhase;
  startDate: string;
  endDate: string;
}

export enum Tab {
  DASHBOARD = 'dashboard',
  MANPOWER = 'manpower',
  PROJECTS = 'projects',
  ASSIGNMENTS = 'assignments',
  GANTT = 'gantt',
  SHORTAGE = 'shortage',
  UNASSIGNED = 'unassigned',
  REPORTS = 'reports'
}
