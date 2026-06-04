import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import { Project } from '../types';

dayjs.extend(customParseFormat);

/**
 * Resolves the actual status of a project. If manual override is selected (Hold, Rescheduled, Cancelled, In Progress, Completed),
 * it returns that. Otherwise, it dynamically determines if the project is "In Progress" or "Completed" based on current date.
 */
export function getProjectActualStatus(project: Project): 'In Progress' | 'Completed' | 'Hold' | 'Rescheduled' | 'Cancelled' {
  if (project.status && project.status !== 'Auto') {
    return project.status as 'In Progress' | 'Completed' | 'Hold' | 'Rescheduled' | 'Cancelled';
  }
  const today = dayjs();
  const end = dayjs(project.endDate);
  
  if (today.isAfter(end, 'day')) {
    return 'Completed';
  } else {
    return 'In Progress';
  }
}

/**
 * Parses any incoming excel date value (serial number or formatted string)
 * into a browser-standard YYYY-MM-DD string.
 */
export function parseExcelDate(val: any): string {
  if (val === undefined || val === null || val === '') {
    return '';
  }

  // Handle excel date serial numbers (like 44670)
  if (typeof val === 'number') {
    // Excel dates start on Dec 30 1899 due to 1900 leap year bug
    const date = new Date((val - 25569) * 86400 * 1000);
    return dayjs(date).format('YYYY-MM-DD');
  }

  const str = String(val).trim();
  if (!str || str === '--') return '';

  // Try parsing popular formats
  const formats = [
    'DD-MMM-YY',
    'DD-MMM-YYYY',
    'YYYY-MM-DD',
    'DD/MM/YYYY',
    'MM/DD/YYYY',
    'DD-MM-YYYY',
    'YYYY/MM/DD',
    'D-M-YY',
    'D/M/YY'
  ];

  for (const fmt of formats) {
    const parsed = dayjs(str, fmt, true);
    if (parsed.isValid()) {
      return parsed.format('YYYY-MM-DD');
    }
  }

  // Fallback direct parsing of native formats
  const nativeParsed = dayjs(str);
  if (nativeParsed.isValid()) {
    return nativeParsed.format('YYYY-MM-DD');
  }

  return '';
}

/**
 * Standard utility to format any YYYY-MM-DD (or valid dayjs date)
 * to the user's requested 'DD-MMM-YY' (e.g. 20-May-26) format.
 */
export function formatToExcelDate(dateVal: any): string {
  if (!dateVal) return '';
  const d = dayjs(dateVal);
  if (!d.isValid()) return '';
  return d.format('DD-MMM-YY');
}
