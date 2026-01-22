import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell, Sector, LabelList
} from 'recharts';
import { DashboardData } from './types';
import { parseCSV } from './utils/dataParser';

// --- CONFIGURATION ---
const PUBLISHED_ID = '2PACX-1vSrx7lqwi5bjj99rYho8jYGBYH47sYw2a5d62uPGrKS-HvSgiz6o-Rx_opsCMGNhVNRjJNx2bi6OTfK';
const BASE_URL = `https://docs.google.com/spreadsheets/d/e/${PUBLISHED_ID}/pub?output=csv`;
const HTML_URL = `https://docs.google.com/spreadsheets/d/e/${PUBLISHED_ID}/pubhtml`;
const REFRESH_INTERVAL = 120000;

const defaultPageTitle = 'Ifocus RC Build Reports';

// Tabs configuration
const TABS_CONFIG = {
  SUMMARY: { id: 'summary', label: 'Report Summary', gid: '0', icon: '📊' },
  NEW_ISSUES: { 
    id: 'new_issues', 
    label: 'New Issues', 
    gid: '476295067', 
    icon: '🐛',
    overrideUrl: 'https://docs.google.com/spreadsheets/d/1yjf5kI6WPNwi_WhH3dFTgiIUONztM54I50sJtnr_PxY/export?format=csv&gid=476295067'
  },
  VALIDATION: { 
    id: 'validation', 
    label: 'Ticket Validation', 
    gid: '2057375142', 
    icon: '✅',
    overrideUrl: 'https://docs.google.com/spreadsheets/d/1yjf5kI6WPNwi_WhH3dFTgiIUONztM54I50sJtnr_PxY/export?format=csv&gid=2057375142'
  },
};

const EXECUTION_COLORS = {
  pass: '#10B981',
  fail: '#F43F5E',
  notConsidered: '#94A3B8',
  automation: '#8B5CF6',
  manual: '#EC4899',
  critical: '#EF4444',
  major: '#F59E0B',
  minor: '#3B82F6',
};

const BUILD_COL_ALIASES = ['RC Build', 'Build Version', 'Build', 'Version', 'Build Number', 'Release Build', 'Build details'];
const PLATFORM_COL_ALIASES = ['Platform', 'OS', 'Environment', 'Device'];
const DATE_COL_ALIASES = ['Build Date', 'Date', 'Reported Date', 'Created At', 'Start Date'];
const SEVERITY_COL_ALIASES = ['Severity', 'Issue Severity', 'Priority', 'Status Severity'];
const STATUS_COL_ALIASES = ['Status', 'Overall Status', 'Result', 'Results', 'Execution Status', 'Build Status'];
const TYPE_COL_ALIASES = ['Build Type', 'Type', 'Deployment Type', 'Category'];
const AUTO_COL_ALIASES = ['Automation executed', 'Automation', 'Auto Executed', 'Automation Test Cases'];
const MANUAL_COL_ALIASES = ['Manual executed', 'Manual', 'Manual Executed', 'Manual Test Cases'];
const RELEASE_STORE_ALIASES = ['Released to store', 'Store Released', 'Store Status', 'App Store Status', 'Play Store Status'];

// Data extraction aliases for metrics
const TOTAL_COL_ALIASES = ['Total Test Cases', 'Total', 'Total Cases', 'Total Count'];
const EXECUTED_COL_ALIASES = ['Executed', 'Execution Count', 'Run'];
const PASSED_COL_ALIASES = ['Passed', 'Pass', 'Success', 'Passed Cases'];
const FAILED_COL_ALIASES = ['Failed', 'Fail', 'Failure', 'Failed Cases'];
const NOT_CONSIDERED_COL_ALIASES = ['Not considered', 'Not Considered', 'N/A', 'Skipped', 'Not Run', 'Pending', 'Not Executed'];
const CRITICAL_ALIASES = ['Critical Issues', 'Critical'];
const MAJOR_ALIASES = ['Major Issues', 'Major'];
const MINOR_ALIASES = ['Minor Issues', 'Minor'];

// --- HELPERS ---

// Robust column finder (case-insensitive)
const findCol = (headers: string[] = [], aliases: string[] = []) => {
  if (!headers || headers.length === 0) return undefined;
  return headers.find(h => aliases.some(a => a.toLowerCase() === h.toLowerCase().trim()));
};

const getVal = (row: any, header: string | undefined) => {
  if (!header || row[header] === undefined) return 0;
  return Number(row[header]) || 0;
};

const getValByAliases = (row: any, headers: string[], aliases: string[]) => {
  const col = findCol(headers, aliases);
  return getVal(row, col);
};

// Robust comparison helper
const smartCompare = (a: any, b: any) => {
  const sa = String(a || '').trim();
  const sb = String(b || '').trim();
  
  if (!sa || !sb) return false;
  if (sa === sb) return true;
  if (sa.toLowerCase() === sb.toLowerCase()) return true;

  // Cleanup function: remove prefixes and standardize separators
  const clean = (s: string) => s
    .replace(/^(rc|build|v|ver|version)[\s:.-]*/i, '') // Remove prefix words and immediate separators
    .trim();

  const ca = clean(sa);
  const cb = clean(sb);

  if (ca === cb) return true;
  if (ca.toLowerCase() === cb.toLowerCase()) return true;

  // Version segment comparison (handles 1.0 vs 1.0.0, ignores trailing zeros)
  const isVersion = (s: string) => /^[\d.]+$/.test(s);
  
  if (isVersion(ca) && isVersion(cb)) {
    const splitA = ca.split('.').map(Number);
    const splitB = cb.split('.').map(Number);
    
    // Remove trailing zeros (e.g., 1.0.0 becomes 1)
    while (splitA.length > 0 && splitA[splitA.length - 1] === 0) splitA.pop();
    while (splitB.length > 0 && splitB[splitB.length - 1] === 0) splitB.pop();
    
    if (splitA.length !== splitB.length) return false;
    return splitA.every((val, i) => val === splitB[i]);
  }
  
  return false;
};

/**
 * Reusable helper to determine status styles based on semantic value
 */
const getStatusStyles = (value: any) => {
  const s = String(value || '').toLowerCase();
  
  if (s.includes('not implemented') || s.includes('not fixed') || s.includes('fail') || s.includes('error')) {
    return 'bg-rose-100 text-rose-700 dark:bg-rose-900/30 dark:text-rose-400';
  }
  if (s.includes('implemented') || s.includes('fixed') || s.includes('pass') || s.includes('success') || s.includes('completed')) {
    return 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400';
  }
  if (s.includes('in progress') || s.includes('pending') || s.includes('started')) {
    return 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400';
  }
  if (s.includes('cnv')) {
    return 'bg-yellow-100 text-yellow-700 dark:bg-yellow-900/30 dark:text-yellow-400';
  }
  if (s.includes('not considered')) {
    return 'bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-400';
  }
  return 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400';
};

/**
 * Helper to determine Build Type colors
 */
const getBuildTypeStyles = (type: any) => {
  const t = String(type || '').toLowerCase();
  if (t.includes('hotfix')) return 'bg-rose-100 text-rose-700 dark:bg-rose-900/30 dark:text-rose-400';
  if (t.includes('planned')) return 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400';
  if (t.includes('adhoc') || t.includes('ad-hoc')) return 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400';
  if (t.includes('release')) return 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400';
  return 'bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-400';
};

// --- INTERFACES ---
interface MetricCardProps {
  title: string;
  value: string | number;
  icon: string;
}

interface CardProps {
  title: string;
  children?: React.ReactNode;
  loading?: boolean;
  error?: string | null;
  onRetry?: () => void;
  discoveredTabs?: Record<string, string>;
  fullWidth?: boolean;
}

interface BadgeProps {
  value: any;
  color: string;
  label: string;
  size?: 'sm' | 'md';
  details?: string[];
}

// --- SUB-COMPONENTS ---

const CustomTooltip = ({ active, payload, label }: any) => {
  if (!active || !payload?.length) return null;
  const title = label || payload[0].name;
  return (
    <div className="bg-white dark:bg-slate-800 p-4 border border-slate-200 dark:border-slate-700 shadow-2xl rounded-2xl animate-in fade-in duration-150 min-w-[160px] pointer-events-none z-[200]">
      <p className="text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2.5 border-b border-slate-100 dark:border-slate-700 pb-2.5 truncate max-w-[200px]">{title}</p>
      <div className="space-y-2">
        {payload.map((entry: any, index: number) => {
          // Robust percentage retrieval for Pie charts
          const preCalcPercent = entry.payload?.percent;
          let percentDisplay = null;

          if (preCalcPercent !== undefined && !isNaN(preCalcPercent)) {
             percentDisplay = (preCalcPercent * 100).toFixed(1);
          }
          
          return (
            <div key={index} className="flex items-center justify-between gap-6">
              <div className="flex items-center gap-2.5">
                <div className="w-2.5 h-2.5 rounded-full ring-2 ring-white shadow-sm" style={{ backgroundColor: entry.color || entry.fill }} />
                <span className="text-[11px] font-bold text-slate-700 dark:text-slate-200">{entry.name}</span>
              </div>
              <div className="flex items-center gap-1.5 text-right">
                <span className="text-[11px] font-extrabold text-slate-900 dark:text-white">{entry.value}</span>
                {percentDisplay && (
                  <span className="text-[10px] font-black text-primary-600 dark:text-primary-400">
                    ({percentDisplay}%)
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

function Calendar({ selectedDate, onSelect, onClose }: { selectedDate: string, onSelect: (date: string) => void, onClose: () => void }) {
  const [viewDate, setViewDate] = useState(() => selectedDate ? new Date(selectedDate) : new Date());
  const month = viewDate.getMonth();
  const year = viewDate.getFullYear();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const firstDayOfMonth = new Date(year, month, 1).getDay();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const prevMonth = () => setViewDate(new Date(year, month - 1, 1));
  const nextMonth = () => setViewDate(new Date(year, month + 1, 1));
  const isSelected = (d: number) => {
    if (!selectedDate) return false;
    const s = new Date(selectedDate);
    return s.getDate() === d && s.getMonth() === month && s.getFullYear() === year;
  };
  const isToday = (d: number) => {
    const today = new Date();
    return today.getDate() === d && today.getMonth() === month && today.getFullYear() === year;
  };
  const handleDateClick = (d: number) => {
    const date = new Date(year, month, d);
    const offset = date.getTimezoneOffset();
    const adjusted = new Date(date.getTime() - (offset * 60 * 1000));
    onSelect(adjusted.toISOString().split('T')[0]);
    onClose();
  };
  return (
    <div className="p-4 w-72 select-none">
      <div className="flex items-center justify-between mb-4">
        <button onClick={prevMonth} className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-xl transition-colors text-slate-400 hover:text-primary-600">
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M15 19l-7-7 7-7" /></svg>
        </button>
        <div className="text-[11px] font-black uppercase tracking-widest text-slate-600 dark:text-slate-300">
          {monthNames[month]} {year}
        </div>
        <button onClick={nextMonth} className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-xl transition-colors text-slate-400 hover:text-primary-600">
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M9 5l7 7-7 7" /></svg>
        </button>
      </div>
      <div className="grid grid-cols-7 gap-1 mb-2">
        {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map(d => (
          <div key={d} className="text-center text-[9px] font-black text-slate-400 uppercase">{d}</div>
        ))}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {Array.from({ length: firstDayOfMonth }).map((_, i) => <div key={`empty-${i}`} />)}
        {Array.from({ length: daysInMonth }).map((_, i) => {
          const d = i + 1;
          const active = isSelected(d);
          const current = isToday(d);
          return (
            <button
              key={d}
              onClick={() => handleDateClick(d)}
              className={`
                aspect-square rounded-lg text-[10px] font-bold flex items-center justify-center transition-all
                ${active ? 'bg-primary-600 text-white shadow-lg shadow-primary-500/30' : 
                  current ? 'text-primary-600 bg-primary-50 dark:bg-primary-900/20' : 
                  'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}
              `}
            >
              {d}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function DateSelector({ value, onChange, placeholder }: { value: string, onChange: (v: string) => void, placeholder: string }) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);
  
  const formatDate = (val: string) => {
    if (!val) return placeholder;
    return new Date(val).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  return (
    <div className="relative flex-1 min-w-0 h-11" ref={containerRef}>
      <div 
        onClick={() => setIsOpen(!isOpen)}
        className={`w-full h-full bg-slate-50 dark:bg-slate-800 px-4 rounded-2xl border transition-all flex items-center justify-between cursor-pointer group z-0 ${isOpen ? 'border-primary-500 ring-2 ring-primary-500/10' : 'border-transparent hover:border-slate-200 dark:hover:border-slate-700'}`}
      >
        <span className={`text-xs font-bold truncate mr-2 ${value ? 'text-slate-900 dark:text-white' : 'text-slate-400'}`}>
          {formatDate(value)}
        </span>
        <svg className={`w-4 h-4 transition-colors ${isOpen ? 'text-primary-500' : 'text-slate-300 group-hover:text-primary-500'}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2-2v12a2 2 0 002 2z" />
        </svg>
      </div>
      {isOpen && (
        <div className="absolute top-full left-0 mt-2 z-[60] bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-2xl glass animate-in zoom-in-95 duration-200">
          <Calendar selectedDate={value} onSelect={onChange} onClose={() => setIsOpen(false)} />
        </div>
      )}
      {value && !isOpen && (
        <button 
          onClick={(e) => { e.preventDefault(); e.stopPropagation(); onChange(''); }} 
          className="absolute right-10 top-1/2 -translate-y-1/2 p-2 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-full transition-all z-20"
          type="button"
        >
          <svg className="w-3 h-3 text-slate-400 hover:text-rose-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      )}
    </div>
  );
}

// --- MAIN APP COMPONENT ---

export default function App() {
  const [theme, setTheme] = useState<'light' | 'dark'>(() => (localStorage.getItem('dashboard-theme') as 'light' | 'dark') || 'light');
  const [dataMap, setDataMap] = useState<Record<string, DashboardData>>({});
  const [activeTab, setActiveTab] = useState<string>(TABS_CONFIG.SUMMARY.id);
  const [loadingMap, setLoadingMap] = useState<Record<string, boolean>>({});
  const [errorMap, setErrorMap] = useState<Record<string, string | null>>({});
  const [refreshProgress, setRefreshProgress] = useState(0);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [selectedPlatform, setSelectedPlatform] = useState<string>('All');
  const [selectedBuild, setSelectedBuild] = useState<string>('All');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  
  const [dynamicGidMap, setDynamicGidMap] = useState<Record<string, string>>({});
  const [discoveredTabs, setDiscoveredTabs] = useState<Record<string, string>>({});

  const isDark = theme === 'dark';

  useEffect(() => {
    localStorage.setItem('dashboard-theme', theme);
    document.documentElement.classList.toggle('dark', isDark);
  }, [theme, isDark]);

  const discoverGids = useCallback(async () => {
    try {
      const resp = await fetch(`${HTML_URL}&t=${Date.now()}`);
      if (!resp.ok) return null;
      const html = await resp.text();
      const tabMap: Record<string, string> = {};
      const liRegex = /id="sheet-button-([^"]+)">.*?>(.*?)<\/a>/g;
      let match;
      while ((match = liRegex.exec(html)) !== null) {
        const gid = match[1];
        const name = match[2].trim();
        tabMap[name] = gid;
      }
      if (Object.keys(tabMap).length === 0) {
        const scriptRegex = /"([^"]+)",\d+,"([^"]+)",\d+,\d+,"[^"]*",\d+/g;
        while ((match = scriptRegex.exec(html)) !== null) {
          const gid = match[2];
          const name = match[1];
          if (/^\d+$/.test(gid) && name.length < 50) {
             tabMap[name] = gid;
          }
        }
      }
      setDiscoveredTabs(tabMap);
      return tabMap;
    } catch (e) {
      return null;
    }
  }, []);

  const fetchData = useCallback(async (tabId: string, silent = false, retryDiscovery = true) => {
    if (!silent) { 
      setLoadingMap(p => ({ ...p, [tabId]: true })); 
      setErrorMap(p => ({ ...p, [tabId]: null })); 
    }
    const config = (TABS_CONFIG as any)[tabId.toUpperCase()] || Object.values(TABS_CONFIG).find(t => t.id === tabId);
    if (!config) return;
    let url: string;
    if (config.overrideUrl) {
      url = `${config.overrideUrl}&t=${Date.now()}`;
    } else {
      const currentGid = dynamicGidMap[tabId] || config.gid;
      url = `${BASE_URL}&gid=${currentGid}&t=${Date.now()}`;
    }
    try {
      const resp = await fetch(url);
      if (!resp.ok) {
        if ((resp.status === 400 || resp.status === 404) && retryDiscovery && !config.overrideUrl) {
          const tabs = await discoverGids();
          if (tabs) {
            const labelLower = config.label.toLowerCase().replace(/\s/g, '');
            const foundName = Object.keys(tabs).find(name => {
              const nameLower = name.toLowerCase().replace(/\s/g, '');
              return nameLower === labelLower || nameLower.includes(labelLower) || labelLower.includes(nameLower);
            });
            if (foundName) {
              const newGid = tabs[foundName];
              setDynamicGidMap(prev => ({ ...prev, [tabId]: newGid }));
              return fetchData(tabId, silent, false);
            }
          }
        }
        throw new Error(`Invalid source configuration. Tab '${config.label}' was not found.`);
      }
      const text = await resp.text();
      const parsed = parseCSV(text);
      setDataMap(p => ({ ...p, [tabId]: parsed }));
    } catch (e: any) {
      if (!silent) setErrorMap(p => ({ ...p, [tabId]: e.message }));
    } finally { 
      if (!silent) setLoadingMap(p => ({ ...p, [tabId]: false })); 
    }
  }, [dynamicGidMap, discoverGids]);

  const syncAll = useCallback(async (isAuto = false) => {
    await Promise.all(Object.values(TABS_CONFIG).map(t => fetchData(t.id, isAuto)));
    setRefreshProgress(0);
  }, [fetchData]);

  const handleResetFilters = useCallback(() => {
    setSelectedPlatform('All');
    setSelectedBuild('All');
    setStartDate('');
    setEndDate('');
  }, []);

  useEffect(() => { syncAll(); }, []);

  useEffect(() => {
    const start = Date.now();
    const interval = setInterval(() => {
      const prog = Math.min(((Date.now() - start) / REFRESH_INTERVAL) * 100, 100);
      setRefreshProgress(prog);
      if (prog >= 100) syncAll(true);
    }, 1000);
    return () => clearInterval(interval);
  }, [syncAll, activeTab]);

  const platforms = useMemo(() => {
    const all = new Set<string>();
    Object.values(dataMap).forEach((d: DashboardData) => {
      const col = findCol(d.headers, PLATFORM_COL_ALIASES);
      if (col) d.rows.forEach(r => r[col] && all.add(String(r[col]).trim()));
    });
    return Array.from(all).sort();
  }, [dataMap]);

  /**
   * Refined builds useMemo for dynamic filtering based on Platform
   */
  const builds = useMemo(() => {
    const all = new Set<string>();
    const summaryData = dataMap[TABS_CONFIG.SUMMARY.id];
    if (summaryData) {
      const pCol = findCol(summaryData.headers, PLATFORM_COL_ALIASES);
      const bCol = findCol(summaryData.headers, BUILD_COL_ALIASES);
      if (bCol) {
        summaryData.rows.forEach(r => {
          const pVal = pCol ? String(r[pCol] || '').trim() : '';
          const bVal = String(r[bCol] || '').trim();
          if (bVal && (selectedPlatform === 'All' || smartCompare(pVal, selectedPlatform))) {
            all.add(bVal);
          }
        });
      }
    }
    return Array.from(all).filter(Boolean).sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
  }, [dataMap, selectedPlatform]);

  /**
   * Selection Sync and Current Status Data Extraction
   */
  const currentBuildInfo = useMemo(() => {
    if (selectedBuild === 'All') return null;
    const summaryData = dataMap[TABS_CONFIG.SUMMARY.id];
    if (!summaryData) return null;

    const bCol = findCol(summaryData.headers, BUILD_COL_ALIASES);
    const pCol = findCol(summaryData.headers, PLATFORM_COL_ALIASES);
    const tCol = findCol(summaryData.headers, TYPE_COL_ALIASES);
    const dCol = findCol(summaryData.headers, DATE_COL_ALIASES);
    const sCol = findCol(summaryData.headers, STATUS_COL_ALIASES);
    const rCol = findCol(summaryData.headers, RELEASE_STORE_ALIASES);

    if (!bCol) return null;

    const matchedRow = summaryData.rows.find(r => 
      smartCompare(r[bCol], selectedBuild) &&
      (selectedPlatform === 'All' || (pCol && smartCompare(r[pCol], selectedPlatform)))
    );

    return matchedRow ? {
      build: selectedBuild,
      platform: pCol ? matchedRow[pCol] : selectedPlatform,
      type: tCol ? matchedRow[tCol] : null,
      startDate: dCol ? matchedRow[dCol] : null,
      status: sCol ? matchedRow[sCol] : null,
      releasedToStore: rCol ? matchedRow[rCol] : null,
    } : null;
  }, [selectedBuild, selectedPlatform, dataMap]);

  /**
   * Dynamic Page Title Logic
   */
  const dynamicPageTitle = useMemo(() => {
    if (selectedBuild !== 'All') {
      return `Ifocus RC build report for ${selectedPlatform} - ${selectedBuild}`;
    }
    return defaultPageTitle;
  }, [selectedPlatform, selectedBuild]);

  const handleBuildChange = (newBuild: string) => {
    setSelectedBuild(newBuild);
    if (newBuild !== 'All' && selectedPlatform === 'All') {
      const summaryData = dataMap[TABS_CONFIG.SUMMARY.id];
      if (summaryData) {
        const bCol = findCol(summaryData.headers, BUILD_COL_ALIASES);
        const pCol = findCol(summaryData.headers, PLATFORM_COL_ALIASES);
        if (bCol && pCol) {
          const row = summaryData.rows.find(r => smartCompare(r[bCol], newBuild));
          if (row && row[pCol]) {
            setSelectedPlatform(String(row[pCol]).trim());
          }
        }
      }
    }
  };

  useEffect(() => {
    // Auto-select build if only one is available for the selected platform
    if (selectedPlatform !== 'All' && builds.length === 1 && selectedBuild !== builds[0]) {
      setSelectedBuild(builds[0]);
    } else if (selectedBuild !== 'All' && !builds.includes(selectedBuild)) {
      // Reset if current selection is no longer valid
      setSelectedBuild('All');
    }
  }, [selectedPlatform, builds, selectedBuild]);

  const filteredRows = useMemo(() => {
    const data = dataMap[activeTab];
    if (!data) return [];
    let rows = [...data.rows];
    
    // Use helpers to find columns based on aliases
    const pCol = findCol(data.headers, PLATFORM_COL_ALIASES);
    const bCol = findCol(data.headers, BUILD_COL_ALIASES);
    const dCol = findCol(data.headers, DATE_COL_ALIASES);

    // Filter Logic:
    // 1. If a specific Build is selected, it takes precedence. 
    //    We filter strictly by that build. This assumes the Build dropdown 
    //    only shows valid builds for the current context/platform.
    if (selectedBuild !== 'All' && bCol) {
       rows = rows.filter(r => smartCompare(r[bCol], selectedBuild));
    } 
    // 2. If Build is 'All', but Platform is specific
    else if (selectedPlatform !== 'All') {
       if (pCol) {
         // If table has explicit platform column, use it
         rows = rows.filter(r => smartCompare(r[pCol], selectedPlatform));
       } else if (bCol && activeTab !== TABS_CONFIG.SUMMARY.id) {
         // Fallback: Use the inferred list of builds for this platform
         // (Only applied to non-Summary tabs that lack platform column, like New Issues)
         rows = rows.filter(r => builds.some(b => smartCompare(r[bCol], b)));
       }
    }

    if (startDate || endDate) {
      if (dCol) {
        rows = rows.filter(r => {
          if (!r[dCol]) return false;
          const bd = new Date(r[dCol]);
          return (!startDate || bd >= new Date(startDate)) && (!endDate || bd <= new Date(endDate));
        });
      }
    }
    return rows;
  }, [dataMap, activeTab, selectedPlatform, selectedBuild, startDate, endDate, builds]);

  const displayHeaders = useMemo(() => {
    const data = dataMap[activeTab];
    if (!data) return [];
    
    // Filter out "S no" variations
    const sNoAliases = ['s no', 's.no', 's.no.', 's. no', 'serial no', 'no.', 'no'];
    let headers = data.headers.filter(h => !sNoAliases.includes(h.toLowerCase().trim()));

    if (activeTab === TABS_CONFIG.VALIDATION.id) {
      const pCol = findCol(data.headers, PLATFORM_COL_ALIASES);
      const bCol = findCol(data.headers, BUILD_COL_ALIASES);
      return headers.filter(h => h !== pCol && h !== bCol);
    }
    return headers;
  }, [dataMap, activeTab]);

  const summaryStats = useMemo(() => {
    const data = dataMap[TABS_CONFIG.SUMMARY.id];
    if (!data) return null;
    let rows = [...data.rows];
    const pCol = findCol(data.headers, PLATFORM_COL_ALIASES);
    const bCol = findCol(data.headers, BUILD_COL_ALIASES);
    const dCol = findCol(data.headers, DATE_COL_ALIASES);

    if (selectedPlatform !== 'All' && pCol) rows = rows.filter(r => smartCompare(r[pCol], selectedPlatform));
    if (selectedBuild !== 'All' && bCol) rows = rows.filter(r => smartCompare(r[bCol], selectedBuild));
    if (startDate || endDate) rows = rows.filter(r => {
      if (!r[dCol || '']) return false;
      const bd = new Date(r[dCol || '']);
      return (!startDate || bd >= new Date(startDate)) && (!endDate || bd <= new Date(endDate));
    });
    return rows.reduce((acc, r) => ({
      total: acc.total + getValByAliases(r, data.headers, TOTAL_COL_ALIASES),
      executed: acc.executed + getValByAliases(r, data.headers, EXECUTED_COL_ALIASES),
      passed: acc.passed + getValByAliases(r, data.headers, PASSED_COL_ALIASES),
      failed: acc.failed + getValByAliases(r, data.headers, FAILED_COL_ALIASES),
      critical: acc.critical + getValByAliases(r, data.headers, CRITICAL_ALIASES),
      major: acc.major + getValByAliases(r, data.headers, MAJOR_ALIASES),
      minor: acc.minor + getValByAliases(r, data.headers, MINOR_ALIASES),
    }), { total: 0, executed: 0, passed: 0, failed: 0, critical: 0, major: 0, minor: 0 });
  }, [dataMap, selectedPlatform, selectedBuild, startDate, endDate]);

  const pieData = useMemo(() => {
    const data = dataMap[TABS_CONFIG.SUMMARY.id];
    if (!data) return [];
    let rows = [...data.rows];
    const pCol = findCol(data.headers, PLATFORM_COL_ALIASES);
    const bCol = findCol(data.headers, BUILD_COL_ALIASES);
    const dCol = findCol(data.headers, DATE_COL_ALIASES);

    if (selectedPlatform !== 'All' && pCol) rows = rows.filter(r => smartCompare(r[pCol], selectedPlatform));
    if (selectedBuild !== 'All' && bCol) rows = rows.filter(r => smartCompare(r[bCol], selectedBuild));
    if (startDate || endDate) rows = rows.filter(r => {
      if (!r[dCol || '']) return false;
      const bd = new Date(r[dCol || '']);
      return (!startDate || bd >= new Date(startDate)) && (!endDate || bd <= new Date(endDate));
    });
    const totals = rows.reduce((acc, r) => ({
      passed: acc.passed + getValByAliases(r, data.headers, PASSED_COL_ALIASES),
      failed: acc.failed + getValByAliases(r, data.headers, FAILED_COL_ALIASES),
      notConsidered: acc.notConsidered + getValByAliases(r, data.headers, NOT_CONSIDERED_COL_ALIASES),
    }), { passed: 0, failed: 0, notConsidered: 0 });
    const sum = totals.passed + totals.failed + totals.notConsidered;
    return [
      { name: 'Pass', value: totals.passed, color: EXECUTION_COLORS.pass, percent: sum ? totals.passed / sum : 0 },
      { name: 'Fail', value: totals.failed, color: EXECUTION_COLORS.fail, percent: sum ? totals.failed / sum : 0 },
      { name: 'N/A', value: totals.notConsidered, color: EXECUTION_COLORS.notConsidered, percent: sum ? totals.notConsidered / sum : 0 },
    ].filter(d => d.value > 0);
  }, [dataMap, selectedPlatform, selectedBuild, startDate, endDate]);

  const trendData = useMemo(() => {
    const summaryData = dataMap[TABS_CONFIG.SUMMARY.id];
    const issuesData = dataMap[TABS_CONFIG.NEW_ISSUES.id];
    if (!summaryData) return [];
    let sumRows = [...summaryData.rows];
    const pColSum = findCol(summaryData.headers, PLATFORM_COL_ALIASES);
    const bColSum = findCol(summaryData.headers, BUILD_COL_ALIASES);
    const dColSum = findCol(summaryData.headers, DATE_COL_ALIASES);
    
    if (selectedPlatform !== 'All' && pColSum) sumRows = sumRows.filter(r => smartCompare(r[pColSum], selectedPlatform));
    if (selectedBuild !== 'All' && bColSum) sumRows = sumRows.filter(r => smartCompare(r[bColSum], selectedBuild));

    if (startDate || endDate) sumRows = sumRows.filter(r => {
      if (!r[dColSum || '']) return false;
      const bd = new Date(r[dColSum || '']);
      return (!startDate || bd >= new Date(startDate)) && (!endDate || bd <= new Date(endDate));
    });
    const recentSumRows = sumRows.slice(0, 10).reverse();
    return recentSumRows.map(r => {
      // Use raw access instead of getValByAliases to preserve string nature of build names (e.g. avoid 0 for "RC 1")
      const rawBuildName = String(r[bColSum || ''] || 'Unknown').trim();
      const platformName = pColSum ? String(r[pColSum] || '').trim() : '';

      const autoVal = getValByAliases(r, summaryData.headers, AUTO_COL_ALIASES);
      const manualVal = getValByAliases(r, summaryData.headers, MANUAL_COL_ALIASES);
      let critical = 0, major = 0, minor = 0;
      if (issuesData) {
        const bColIss = findCol(issuesData.headers, BUILD_COL_ALIASES);
        const sColIss = findCol(issuesData.headers, SEVERITY_COL_ALIASES);
        if (bColIss && sColIss) {
          issuesData.rows.forEach(ir => {
            const iBuild = String(ir[bColIss] || '').trim();
            if (smartCompare(iBuild, rawBuildName)) {
              const sev = String(ir[sColIss] || '').toLowerCase();
              if (sev.includes('crit')) critical++;
              else if (sev.includes('maj')) major++;
              else if (sev.includes('min')) minor++;
            }
          });
        }
      }
      if (critical === 0 && major === 0 && minor === 0) {
        critical = getValByAliases(r, summaryData.headers, CRITICAL_ALIASES);
        major = getValByAliases(r, summaryData.headers, MAJOR_ALIASES);
        minor = getValByAliases(r, summaryData.headers, MINOR_ALIASES);
      }
      return {
        name: rawBuildName,
        platform: platformName,
        fullName: rawBuildName,
        Passed: getValByAliases(r, summaryData.headers, PASSED_COL_ALIASES),
        Failed: getValByAliases(r, summaryData.headers, FAILED_COL_ALIASES),
        Critical: critical,
        Major: major,
        Minor: minor,
        Automation: autoVal,
        Manual: manualVal,
      };
    });
  }, [dataMap, selectedPlatform, selectedBuild, startDate, endDate]);

  /**
   * Determine if any issues exist in the current filtered context for the trend chart
   */
  const trendHasNoIssues = useMemo(() => {
    if (selectedBuild !== 'All') {
      // If a specific build is selected, we only check that build's summary stats
      return (summaryStats?.critical === 0 && summaryStats?.major === 0 && summaryStats?.minor === 0);
    }
    // Otherwise check if all points in the trend are zero
    return trendData.length === 0 || trendData.every(d => (d.Critical + d.Major + d.Minor) === 0);
  }, [trendData, summaryStats, selectedBuild]);

  const renderActiveShape = (props: any) => {
    const { cx, cy, innerRadius, outerRadius, startAngle, endAngle, fill } = props;
    return (
      <g>
        <Sector cx={cx} cy={cy} innerRadius={innerRadius} outerRadius={outerRadius + 8} startAngle={startAngle} endAngle={endAngle} fill={fill} />
        <Sector cx={cx} cy={cy} startAngle={startAngle} endAngle={endAngle} innerRadius={outerRadius + 10} outerRadius={outerRadius + 12} fill={fill} opacity={0.3} />
      </g>
    );
  };

  /**
   * Refined Pie Label for external positioning and readability
   */
  const renderCustomizedPieLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent, value, name }: any) => {
    const RADIAN = Math.PI / 180;
    // Push labels slightly further out for clarity
    const radius = outerRadius + 30; 
    const x = cx + radius * Math.cos(-midAngle * RADIAN);
    const y = cy + radius * Math.sin(-midAngle * RADIAN);
    const labelPercent = (percent * 100).toFixed(1);
    
    // Hide label if slice is too small to avoid clutter (threshold lowered to 2%)
    if (percent < 0.02) return null;

    return (
      <g>
        {/* Draw a connecting line manually to ensure it points to the right place */}
        <path d={`M${cx + outerRadius * Math.cos(-midAngle * RADIAN)},${cy + outerRadius * Math.sin(-midAngle * RADIAN)}L${x},${y}`} stroke={isDark ? '#475569' : '#cbd5e1'} fill="none" />
        <text 
          x={x} 
          y={y - 7} 
          fill={isDark ? '#94a3b8' : '#64748b'} 
          textAnchor={x > cx ? 'start' : 'end'} 
          dominantBaseline="central" 
          className="text-[10px] md:text-[11px] font-black uppercase tracking-widest pointer-events-none"
        >
          {name}
        </text>
        <text 
          x={x} 
          y={y + 7} 
          fill={isDark ? '#f8fafc' : '#0f172a'} 
          textAnchor={x > cx ? 'start' : 'end'} 
          dominantBaseline="central" 
          className="text-[11px] md:text-[13px] font-black pointer-events-none"
        >
          {`${value} (${labelPercent}%)`}
        </text>
      </g>
    );
  };

  const CustomXAxisTick = (props: any) => {
    const { x, y, payload } = props;
    const dataItem = trendData[payload.index];
    if (!dataItem) return null;
    
    return (
      <g transform={`translate(${x},${y})`}>
        <text 
          x={0} 
          y={0} 
          dy={16} 
          textAnchor="end" 
          fill={isDark ? '#cbd5e1' : '#1e293b'} 
          fontSize={10} 
          fontWeight={800}
          transform="rotate(-35)"
        >
          {payload.value}
          <tspan x="0" dy="12" fill={isDark ? '#64748b' : '#94a3b8'} fontSize={9} fontWeight={600} style={{ textTransform: 'uppercase' }}>
            {dataItem.platform}
          </tspan>
        </text>
      </g>
    );
  };

  const isAnyFilterActive = useMemo(() => {
    return selectedPlatform !== 'All' || selectedBuild !== 'All' || startDate !== '' || endDate !== '';
  }, [selectedPlatform, selectedBuild, startDate, endDate]);

  return (
    <div className="min-h-screen bg-[#F8FAFC] dark:bg-[#020617] pb-12 transition-all">
      <div className="fixed top-0 left-0 right-0 h-1 bg-slate-100 dark:bg-slate-900 z-[100] overflow-hidden">
        <div className="h-full bg-primary-600 transition-all duration-1000 ease-linear" style={{ width: `${refreshProgress}%` }} />
      </div>
      <nav className="sticky top-0 z-[90] glass border-b border-slate-200 dark:border-slate-800 px-4 md:px-6 py-4 shadow-sm">
        <div className="max-w-7xl mx-auto flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-primary-600 rounded-xl flex items-center justify-center text-white font-black text-xl shadow-lg shrink-0">i</div>
            <h1 className="text-sm md:text-xl font-black uppercase tracking-tight text-primary-600 truncate">{dynamicPageTitle}</h1>
          </div>
          <div className="flex gap-2 shrink-0">
            <button onClick={() => setTheme(t => t === 'light' ? 'dark' : 'light')} className="p-2.5 rounded-xl bg-slate-100 dark:bg-slate-800 hover:scale-105 transition-all">{isDark ? '☀️' : '🌙'}</button>
            <button onClick={() => syncAll()} className="px-5 py-2.5 bg-primary-600 text-white rounded-xl text-[10px] font-black uppercase active:scale-95 shadow-lg">Sync</button>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 md:px-8 py-8 space-y-8">
        {/* Filter Section */}
        <section className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 border border-slate-200 dark:border-slate-800 shadow-sm relative">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-[10px] font-black uppercase tracking-widest text-slate-400">Dashboard Filters</h2>
            {isAnyFilterActive && (
              <button onClick={handleResetFilters} className="text-[10px] font-black uppercase text-rose-500 hover:text-rose-600 transition-all flex items-center gap-1.5 group animate-in fade-in slide-in-from-right-2 duration-300">
                <span className="p-1.5 rounded-full bg-rose-50 dark:bg-rose-900/20 group-hover:bg-rose-100 dark:group-hover:bg-rose-900/40 transition-colors">
                  <svg className="w-3.5 h-3.5 group-hover:rotate-[-45deg] transition-transform duration-300" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg>
                </span>
                Clear All Filters
              </button>
            )}
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
            <div><label className="text-[10px] font-black uppercase text-slate-400 ml-1">Platform</label><select value={selectedPlatform} onChange={e => { const val = e.target.value; setSelectedPlatform(val); if (val === 'All') setSelectedBuild('All'); }} className="w-full bg-slate-50 dark:bg-slate-800 p-3.5 mt-1 rounded-2xl text-xs font-bold border border-transparent focus:border-primary-500/30 transition-all appearance-none cursor-pointer"><option value="All">All Platforms</option>{platforms.map(p => <option key={p} value={p}>{p}</option>)}</select></div>
            <div><label className="text-[10px] font-black uppercase text-slate-400 ml-1">Build Version</label><select value={selectedBuild} onChange={e => handleBuildChange(e.target.value)} className="w-full bg-slate-50 dark:bg-slate-800 p-3.5 mt-1 rounded-2xl text-xs font-bold border border-transparent focus:border-primary-500/30 transition-all appearance-none cursor-pointer"><option value="All">All Builds</option>{builds.map(b => <option key={b} value={b}>{b}</option>)}</select></div>
            <div className="md:col-span-2"><label className="text-[10px] font-black uppercase text-slate-400 ml-1">Date Range</label><div className="flex items-center gap-2 mt-1"><DateSelector value={startDate} onChange={setStartDate} placeholder="Start Date" /><span className="text-slate-300">~</span><DateSelector value={endDate} onChange={setEndDate} placeholder="End Date" /></div></div>
          </div>
        </section>

        {/* Current Status Section */}
        {currentBuildInfo && (
          <section className="animate-in fade-in slide-in-from-top-2 duration-500">
            <div className="bg-white dark:bg-slate-900 rounded-[2rem] p-8 border border-slate-200 dark:border-slate-800 shadow-sm">
              <div className="flex items-center gap-3 mb-6">
                <div className="w-2 h-2 rounded-full bg-primary-500 animate-pulse" />
                <h2 className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Current Status</h2>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-8">
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Build Number</p>
                  <p className="text-sm font-black text-primary-600 dark:text-primary-400">#{currentBuildInfo.build}</p>
                </div>
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Platform</p>
                  <p className="text-sm font-black text-slate-900 dark:text-white uppercase">{currentBuildInfo.platform}</p>
                </div>
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Build Type</p>
                  {currentBuildInfo.type ? (
                    <span className={`px-2.5 py-0.5 rounded-lg text-[10px] font-black uppercase tracking-wider ${getBuildTypeStyles(currentBuildInfo.type)}`}>
                      {currentBuildInfo.type}
                    </span>
                  ) : <p className="text-sm font-black text-slate-300">—</p>}
                </div>
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Start Date</p>
                  <p className="text-sm font-black text-slate-900 dark:text-white">{currentBuildInfo.startDate || '—'}</p>
                </div>
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Overall Status</p>
                  {currentBuildInfo.status ? (
                    <span className={`px-2.5 py-0.5 rounded-lg text-[10px] font-black uppercase tracking-wider ${getStatusStyles(currentBuildInfo.status)}`}>
                      {currentBuildInfo.status}
                    </span>
                  ) : <p className="text-sm font-black text-slate-300">—</p>}
                </div>
                <div>
                  <p className="text-[9px] font-black uppercase text-slate-400 tracking-wider mb-1.5">Released to store</p>
                  {currentBuildInfo.releasedToStore ? (
                    <span className={`px-2.5 py-0.5 rounded-lg text-[10px] font-black uppercase tracking-wider ${getStatusStyles(currentBuildInfo.releasedToStore)}`}>
                      {currentBuildInfo.releasedToStore}
                    </span>
                  ) : <p className="text-sm font-black text-slate-300">—</p>}
                </div>
              </div>
            </div>
          </section>
        )}

        <div className="flex bg-slate-100 dark:bg-slate-900/50 p-1 rounded-[1.8rem] w-full md:w-auto overflow-x-auto gap-1 shadow-inner no-scrollbar">
          {Object.values(TABS_CONFIG).map(tab => (
            <button key={tab.id} onClick={() => setActiveTab(tab.id)} className={`px-8 py-3.5 rounded-2xl text-xs font-black uppercase tracking-wider transition-all flex items-center gap-2 whitespace-nowrap ${activeTab === tab.id ? 'bg-white dark:bg-slate-800 text-primary-600 shadow-md' : 'text-slate-500 hover:text-slate-800'}`}><span>{tab.icon}</span>{tab.label}</button>
          ))}
        </div>

        {activeTab === 'summary' ? (
          <div className="space-y-8 animate-in fade-in duration-500">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <MetricCard title="Total Cases" value={summaryStats?.total || 0} icon="🎯" />
              <MetricCard title="Executed" value={summaryStats?.executed || 0} icon="⚡" />
              <MetricCard title="Pass Rate" value={`${summaryStats?.executed ? ((summaryStats.passed / summaryStats.executed) * 100).toFixed(1) : 0}%`} icon="✅" />
              <MetricCard title="Critical" value={summaryStats?.critical || 0} icon="🌋" />
            </div>
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              <Card title="Distribution" loading={loadingMap[activeTab]} error={errorMap[activeTab]} onRetry={() => fetchData(activeTab)}>
                <div className="h-[400px] md:h-[360px] relative">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie 
                        {...({ activeIndex, activeShape: renderActiveShape } as any)} 
                        data={pieData} 
                        cx="50%" 
                        cy="50%" 
                        startAngle={90}
                        endAngle={-270}
                        innerRadius="50%" 
                        outerRadius="70%" 
                        paddingAngle={4} 
                        dataKey="value" 
                        label={renderCustomizedPieLabel} 
                        labelLine={false}
                        onMouseEnter={(_, i) => setActiveIndex(i)} 
                        onMouseLeave={() => setActiveIndex(-1)}
                      >
                        {pieData.map((e, i) => <Cell key={i} fill={e.color} stroke="none" />)}
                      </Pie>
                      <Tooltip content={<CustomTooltip />} cursor={false} wrapperStyle={{ outline: 'none' }} />
                      <Legend verticalAlign="bottom" height={36} iconType="circle" wrapperStyle={{ fontSize: '11px', fontWeight: 'bold', paddingTop: '10px' }} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-center pointer-events-none select-none">
                    <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Executed</div>
                    <div className="text-2xl md:text-3xl font-black text-slate-900 dark:text-white leading-tight">{summaryStats?.executed || 0}</div>
                  </div>
                </div>
              </Card>
              <div className="lg:col-span-2">
                <Card title="Methodology Trend" loading={loadingMap[activeTab]} error={errorMap[activeTab]}>
                  <div className="h-[300px]"><ResponsiveContainer width="100%" height="100%">
                    <BarChart data={trendData} margin={{ top: 30, right: 10, left: -20, bottom: 20 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={isDark ? '#1e293b' : '#f1f5f9'} />
                      <XAxis dataKey="name" tick={<CustomXAxisTick />} axisLine={false} tickLine={false} interval={0} height={70} />
                      <YAxis tick={{ fontSize: 10, fontWeight: 800 }} axisLine={false} tickLine={false} />
                      <Tooltip content={<CustomTooltip />} />
                      <Legend verticalAlign="top" align="right" wrapperStyle={{ fontSize: '11px', fontWeight: 'bold' }} />
                      <Bar name="Automation" dataKey="Automation" fill={EXECUTION_COLORS.automation} radius={[6, 6, 0, 0]} barSize={16}>
                        <LabelList position="top" fontSize={11} fontWeight="900" fill={isDark ? '#cbd5e1' : '#1e293b'} offset={10} formatter={(val: number) => val > 0 ? val : ''} />
                      </Bar>
                      <Bar name="Manual" dataKey="Manual" fill={EXECUTION_COLORS.manual} radius={[6, 6, 0, 0]} barSize={16}>
                        <LabelList position="top" fontSize={11} fontWeight="900" fill={isDark ? '#cbd5e1' : '#1e293b'} offset={10} formatter={(val: number) => val > 0 ? val : ''} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer></div>
                </Card>
              </div>
            </div>
            <Card title="Issue Severity Trend" fullWidth>
              <div className="h-[360px] relative flex flex-col items-center justify-center">
                {trendHasNoIssues ? (
                  <div className="flex flex-col items-center justify-center text-center p-12 space-y-3 animate-in fade-in zoom-in-95 duration-500">
                    <div className="text-5xl opacity-30 grayscale mb-2">🎉</div>
                    <h4 className="text-base font-black text-slate-900 dark:text-white uppercase tracking-tight">Perfect Score!</h4>
                    <p className="text-sm font-bold text-slate-400 max-w-sm">No issues reported in this build.</p>
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={trendData} margin={{ top: 40, right: 20, left: -20, bottom: 20 }} stackOffset="none">
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={isDark ? '#1e293b' : '#f1f5f9'} />
                      <XAxis dataKey="name" tick={<CustomXAxisTick />} axisLine={false} tickLine={false} interval={0} height={70} />
                      <YAxis tick={{ fontSize: 10, fontWeight: 800 }} axisLine={false} tickLine={false} />
                      <Tooltip content={<CustomTooltip />} />
                      <Legend verticalAlign="top" align="right" wrapperStyle={{ fontSize: '11px', fontWeight: 'bold', paddingBottom: '30px' }} />
                      <Bar name="Critical" dataKey="Critical" stackId="severity" fill={EXECUTION_COLORS.critical} barSize={32}>
                        <LabelList dataKey="Critical" position="center" fill="#fff" fontSize={11} fontWeight="900" formatter={(val: any) => val > 0 ? val : ''} />
                      </Bar>
                      <Bar name="Major" dataKey="Major" stackId="severity" fill={EXECUTION_COLORS.major} barSize={32}>
                        <LabelList dataKey="Major" position="center" fill="#fff" fontSize={11} fontWeight="900" formatter={(val: any) => val > 0 ? val : ''} />
                      </Bar>
                      <Bar name="Minor" dataKey="Minor" stackId="severity" fill={EXECUTION_COLORS.minor} radius={[6, 6, 0, 0]} barSize={32}>
                        <LabelList dataKey="Minor" position="center" fill="#fff" fontSize={11} fontWeight="900" formatter={(val: any) => val > 0 ? val : ''} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </Card>
            <Card title="Execution Matrix & Build Details" fullWidth>
              <div className="max-h-[600px] overflow-auto custom-scrollbar border border-slate-100 dark:border-slate-800 rounded-2xl shadow-inner bg-white dark:bg-slate-900 relative">
                <table className="w-full text-left min-w-[1400px] border-separate border-spacing-0">
                  <thead className="bg-slate-50 dark:bg-slate-900 shadow-sm">
                    <tr className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 sticky left-0 top-0 z-50 min-w-[200px] text-left whitespace-nowrap">Build Version</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 sticky top-0 z-40 text-left whitespace-nowrap">Start Date</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-center sticky top-0 z-40 whitespace-nowrap">Overall Status</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 sticky top-0 z-40 text-left whitespace-nowrap">Build Type</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-right sticky top-0 z-40 whitespace-nowrap">Total Test Cases</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-right sticky top-0 z-40 whitespace-nowrap">Passed Test Cases</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-right sticky top-0 z-40 whitespace-nowrap">Failed Test Cases</th>
                      <th className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-right sticky top-0 z-40 whitespace-nowrap">Issue Severities</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                    {filteredRows.map((row, idx) => {
                      const bCol = findCol(dataMap[activeTab].headers, BUILD_COL_ALIASES);
                      const pCol = findCol(dataMap[activeTab].headers, PLATFORM_COL_ALIASES);
                      const dCol = findCol(dataMap[activeTab].headers, DATE_COL_ALIASES);
                      const sCol = findCol(dataMap[activeTab].headers, STATUS_COL_ALIASES);
                      const tCol = findCol(dataMap[activeTab].headers, TYPE_COL_ALIASES);
                      
                      const total = getValByAliases(row, dataMap[activeTab].headers, TOTAL_COL_ALIASES);
                      const passed = getValByAliases(row, dataMap[activeTab].headers, PASSED_COL_ALIASES);
                      const failed = getValByAliases(row, dataMap[activeTab].headers, FAILED_COL_ALIASES);
                      const critical = getValByAliases(row, dataMap[activeTab].headers, CRITICAL_ALIASES);
                      const major = getValByAliases(row, dataMap[activeTab].headers, MAJOR_ALIASES);
                      const minor = getValByAliases(row, dataMap[activeTab].headers, MINOR_ALIASES);

                      return (
                        <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/30 transition-colors">
                          <td className="px-6 py-5 sticky left-0 z-30 bg-white dark:bg-slate-900 border-r border-slate-50 dark:border-slate-800 whitespace-nowrap min-w-[200px] text-left">
                            <div className="font-extrabold text-sm text-slate-900 dark:text-white">{row[bCol || ''] || '-'}</div>
                            <div className="text-[10px] font-bold text-slate-500 uppercase mt-0.5 tracking-wide">{row[pCol || ''] || 'N/A'}</div>
                          </td>
                          <td className="px-6 py-5 text-[11px] font-bold text-slate-500 whitespace-nowrap text-left">{row[dCol || ''] || '-'}</td>
                          <td className="px-6 py-5 text-center"><span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest ${getStatusStyles(row[sCol || ''])} whitespace-nowrap`}>{row[sCol || ''] || 'N/A'}</span></td>
                          <td className="px-6 py-5 text-left"><span className={`px-2.5 py-0.5 rounded-lg text-[10px] font-black uppercase tracking-wider ${getBuildTypeStyles(row[tCol || ''])} whitespace-nowrap`}>{row[tCol || ''] || 'N/A'}</span></td>
                          <td className="px-6 py-5 text-xs font-black text-slate-500 text-right">{total}</td>
                          <td className="px-6 py-5 text-xs font-black text-emerald-600 text-right">{passed}</td>
                          <td className="px-6 py-5 text-xs font-black text-rose-500 text-right">{failed}</td>
                          <td className="px-6 py-5 text-right">
                            <div className="flex gap-1 justify-end">
                              <Badge value={critical} color="bg-rose-500" label="Critical" size="sm" />
                              <Badge value={major} color="bg-amber-500" label="Major" size="sm" />
                              <Badge value={minor} color="bg-blue-500" label="Minor" size="sm" />
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {filteredRows.length === 0 && (
                      <tr><td colSpan={8} className="px-6 py-12 text-center text-slate-400 font-bold text-xs uppercase tracking-widest italic">No matching build details found</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </Card>
          </div>
        ) : (
          <Card title={activeTab === 'new_issues' ? 'Issue Backlog (External Data)' : 'Validation Queue'} loading={loadingMap[activeTab]} error={errorMap[activeTab]} onRetry={() => fetchData(activeTab)} discoveredTabs={discoveredTabs} fullWidth>
            <div className="max-h-[600px] overflow-auto custom-scrollbar border border-slate-100 dark:border-slate-800 rounded-2xl shadow-inner bg-white dark:bg-slate-900 relative">
              <table className="w-full text-left min-w-[1000px] border-separate border-spacing-0">
                <thead className="bg-slate-50 dark:bg-slate-900 shadow-sm">
                  <tr className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                    {displayHeaders.map((h, i) => <th key={i} className="px-6 py-4 border-b border-slate-200 dark:border-slate-700 whitespace-nowrap bg-slate-50 dark:bg-slate-900 sticky top-0 z-40">{h}</th>)}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {filteredRows.map((row, i) => (
                    <tr key={i} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/30 transition-colors">
                      {displayHeaders.map((h, j) => {
                        const isStatus = STATUS_COL_ALIASES.some(a => a.toLowerCase() === h.toLowerCase());
                        const val = row[h];
                        if (isStatus && val) {
                          return <td key={j} className="px-6 py-5 min-w-[120px] whitespace-nowrap"><span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest ${getStatusStyles(val)}`}>{val}</span></td>;
                        }
                        return <td key={j} className="px-6 py-5 text-[12px] font-bold text-slate-600 dark:text-slate-300 leading-relaxed min-w-[120px] whitespace-nowrap">{val || '-'}</td>;
                      })}
                    </tr>
                  ))}
                  {filteredRows.length === 0 && !loadingMap[activeTab] && (
                    <tr><td colSpan={displayHeaders.length || 1} className="px-6 py-12 text-center text-slate-400 font-bold text-xs uppercase tracking-widest italic">No records match the current filters</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>
        )}
      </main>
    </div>
  );
}

function MetricCard({ title, value, icon }: MetricCardProps) {
  return (
    <div className="bg-white dark:bg-slate-900 p-8 rounded-[2.5rem] border border-slate-200 dark:border-slate-800 shadow-sm relative group hover:-translate-y-1 transition-all duration-300">
      <div className="absolute top-0 right-0 p-4 opacity-5 group-hover:opacity-10 transition-all text-6xl pointer-events-none rotate-6">{icon}</div>
      <span className="text-[11px] font-black text-slate-400 uppercase mb-2 block tracking-widest">{title}</span>
      <div className="text-3xl font-black text-slate-900 dark:text-white tracking-tight">{value}</div>
    </div>
  );
}

function Card({ title, children, loading, error, onRetry, discoveredTabs, fullWidth }: CardProps) {
  const tabsList = discoveredTabs ? Object.entries(discoveredTabs) : [];
  return (
    <div className={`bg-white dark:bg-slate-900 rounded-[2.5rem] border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col min-h-[400px] ${fullWidth ? 'lg:col-span-3' : ''}`}>
      <div className="px-8 py-6 border-b border-slate-50 dark:border-slate-800/50 flex justify-between items-center"><h3 className="text-[12px] font-black uppercase tracking-widest text-slate-400">{title}</h3></div>
      <div className="p-8 flex-1 relative flex flex-col">
        {loading && <div className="absolute inset-0 z-50 flex items-center justify-center bg-white/60 dark:bg-slate-900/60 backdrop-blur-sm rounded-[2.5rem]"><div className="w-8 h-8 border-3 border-primary-600 border-t-transparent rounded-full animate-spin" /></div>}
        {error ? (
          <div className="flex-1 flex flex-col items-center justify-center text-center space-y-4">
            <div className="text-4xl opacity-40">⚠️</div>
            <p className="text-[11px] font-bold text-rose-500 max-w-xs">{error}</p>
            {tabsList.length > 0 && (
              <div className="bg-slate-50 dark:bg-slate-800 p-4 rounded-2xl text-left w-full max-sm border border-slate-100 dark:border-slate-700">
                <p className="text-[9px] font-black uppercase text-slate-400 mb-2 tracking-widest">Available Sheets (Discovery):</p>
                <div className="space-y-1">
                  {tabsList.map(([name, gid]) => (
                    <div key={gid} className="flex justify-between items-center text-[10px] font-bold text-slate-600 dark:text-slate-400"><span>{name}</span><span className="bg-slate-100 dark:bg-slate-700 px-1 rounded text-[8px]">GID: {gid}</span></div>
                  ))}
                </div>
              </div>
            )}
            <button onClick={onRetry} className="px-8 py-3 bg-slate-900 dark:bg-slate-800 text-white rounded-xl text-[10px] font-black uppercase tracking-widest active:scale-95 shadow-lg">Retry Sync</button>
          </div>
        ) : children}
      </div>
    </div>
  );
}

function Badge({ value, color, label, size = 'md', details }: BadgeProps) {
  const vNum = Number(value);
  const isZero = isNaN(vNum) || vNum === 0;
  const displayVal = isZero ? '-' : value;
  const s = size === 'sm' ? 'w-8 h-8 text-[10px]' : 'w-10 h-10 text-[12px]';
  const hasDetails = details && details.length > 0;
  
  // Show tooltip if not zero
  const canShowTooltip = !isZero;

  return (
    <div className="relative group/badge inline-block">
      <div className={`${s} rounded-lg flex items-center justify-center font-black transition-all ${isZero ? 'bg-slate-100 text-slate-300 dark:bg-slate-800 dark:text-slate-700' : `${color} text-white shadow-sm ${canShowTooltip ? 'cursor-help' : ''} hover:scale-105 active:scale-95`}`}>
        {displayVal}
      </div>
      
      {canShowTooltip && (
        <div className={`absolute bottom-full left-1/2 -translate-x-1/2 mb-3 p-3 bg-white text-slate-900 rounded-xl opacity-0 scale-95 translate-y-2 group-hover/badge:opacity-100 group-hover/badge:scale-100 group-hover/badge:translate-y-0 pointer-events-none transition-all duration-200 ease-out z-[100] shadow-2xl border border-slate-200 ${hasDetails ? 'min-w-[200px] max-w-[340px]' : 'whitespace-nowrap'}`}>
          <div className={`flex items-center ${hasDetails ? 'justify-between mb-2 border-b border-slate-100 pb-2' : 'justify-center'}`}>
             <span className="text-[11px] font-black uppercase tracking-wider">{displayVal} {label} {vNum === 1 ? 'Issue' : 'Issues'}</span>
          </div>
          {hasDetails && (
            <ul className="space-y-1.5 max-h-[160px] overflow-y-auto no-scrollbar pr-1 text-left">
              {details.slice(0, 10).map((d, i) => (
                <li key={i} className="text-[10px] font-medium leading-tight text-slate-600 list-disc list-inside break-words">{d}</li>
              ))}
              {details.length > 10 && <li className="text-[9px] font-black text-primary-600 uppercase tracking-tighter pt-1 sticky bottom-0 bg-white">+ {details.length - 10} more issues</li>}
            </ul>
          )}
          <div className="absolute top-full left-1/2 -translate-x-1/2 w-0 h-0 border-l-[6px] border-l-transparent border-r-[6px] border-r-transparent border-t-[6px] border-t-white" />
        </div>
      )}
    </div>
  );
}
