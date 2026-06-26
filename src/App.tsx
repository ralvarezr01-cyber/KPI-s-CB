/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { 
  FileUp, 
  History,
  FileSpreadsheet,
  CalendarDays,
  CheckCircle2,
  BarChart3,
  Download,
  Percent
} from 'lucide-react';
import { 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  AreaChart, 
  Area
} from 'recharts';
import * as XLSX from 'xlsx';
import { motion } from 'motion/react';
import { parseExcelDashboard } from './services/excelService';
import { DashboardData, RawRecord } from './types';
import { cn } from './lib/utils';

// Final Evaluation Naming and Level Helpers
const getEvaluation = (efficiencyPct: number) => {
  if (efficiencyPct >= 100) return { note: 5, label: 'Extraordinario', color: 'text-emerald-700' };
  if (efficiencyPct >= 80) return { note: 4, label: 'Excelente', color: 'text-emerald-500' };
  if (efficiencyPct >= 50) return { note: 3, label: 'Satisfactorio', color: 'text-blue-500' };
  if (efficiencyPct >= 0) return { note: 2, label: 'Regular', color: 'text-orange-500' };
  return { note: 1, label: 'Insuficiente', color: 'text-rose-600' };
};

// Optimized Helper for Aging Detection
const getIsOld = (label: string) => {
  const l = label.toUpperCase();
  return l.includes('91') || l.includes('120') || l.includes('MAS DE') || l.includes('150') || l.includes('180');
};

// Optimized Helper for Target Aging buckets (4, 5, 6)
const getIsTargetAging = (label: string) => {
  const l = label.toUpperCase();
  return l.includes('4.') || l.includes('5.') || l.includes('6.') ||
         l.includes('MENOR A 120') || l.includes('MENOR A 180') || l.includes('MAYOR A 180') ||
         l.includes('0-30') || l.includes('31-60') || l.includes('61-90') || l.includes('91-120') || 
         l.includes('121-150') || l.includes('151-180') || l.includes('180') || l.includes('MAS DE');
};

// Target categorization for 5-view logic:
// Categorías 1, 2, 3 (Sin desfase: 0-30 días, 31-60 días, 61-90 días)
const getIsSinDesfase = (label: string) => {
  const l = label.toUpperCase();
  return l.includes('1.') || l.includes('2.') || l.includes('3.') ||
         l.includes('0-30') || l.includes('31-60') || l.includes('61-90') ||
         l.includes('MENOR A 30') || l.includes('MENOR A 60') || l.includes('MENOR A 90') ||
         l.includes('MAYOR A 30');
};

// Categorías 4, 5, 6 (Desfasado: Menor a 120 días, Menor a 180 días, Mayor a 180 días, etc.)
const getIsDesfasado = (label: string) => {
  const l = label.toUpperCase();
  return l.includes('4.') || l.includes('5.') || l.includes('6.') ||
         l.includes('MENOR A 120') || l.includes('MENOR A 180') || l.includes('MAYOR A 180') ||
         l.includes('91-120') || l.includes('121-150') || l.includes('151-180') ||
         l.includes('180') || l.includes('MAS DE');
};

const getRecordController = (r: any): string => {
  return r['Contralor Corporativo'] || r['Contralor Responsable'] || r['Contralor'] || r['Responsable'] || r['contralor'] || 'N/A';
};

const getKPI1Status = (pct: number) => {
  if (pct >= 4 && pct <= 6) {
    return { text: "Cumple", color: "bg-emerald-100 text-emerald-800 border-emerald-200" };
  } else {
    return { text: "No cumple", color: "bg-rose-100 text-rose-800 border-rose-200" };
  }
};

const getKPI2Status = (pct: number) => {
  if (pct >= 1.5 && pct <= 3.5) {
    return { text: "Cumple", color: "bg-emerald-100 text-emerald-800 border-emerald-200" };
  } else {
    return { text: "No cumple", color: "bg-rose-100 text-rose-800 border-rose-200" };
  }
};

const formatToTitleCase = (str: string) => {
  if (!str) return 'N/A';
  if (str === 'ALL' || str === 'N/A') return str;
  return str.toLowerCase().replace(/(^|\s)\S/g, l => l.toUpperCase());
};

export default function App() {
  const [data, setData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(false);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [vistas, setVistas] = useState<'acumulado' | 'sin_desfase' | 'desfasado' | 'ejecutiva'>('acumulado');
  const [selectedSociedad, setSelectedSociedad] = useState<string>('ALL');
  const [selectedPeriods, setSelectedPeriods] = useState<string[]>([]);
  const [groupBy, setGroupBy] = useState<string>('Sociedad');

  // Initial Data Load from Server
  useEffect(() => {
    const fetchPersistentData = async () => {
      try {
        const response = await fetch('/api/data');
        const savedData = await response.json();
        if (savedData) {
          setData(savedData);
          setSelectedPeriods(savedData.periods.map(p => p.periodLabel));
          if (savedData.availableColumns.includes('Soc.')) setGroupBy('Soc.');
        }
      } catch (err) {
        console.error("Persistence Load Error:", err);
      } finally {
        setIsInitialLoad(false);
      }
    };
    fetchPersistentData();
  }, []);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    try {
      const result = await parseExcelDashboard(file);
      setData(result);
      setSelectedPeriods(result.periods.map(p => p.periodLabel)); // Select all by default for trend
      
      // Try to find a good default groupBy besides Sociedad
      if (result.availableColumns.includes('Soc.')) setGroupBy('Soc.');

      // Persist to Server
      fetch('/api/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result)
      }).catch(err => console.error("Persistence Save Error:", err));

    } catch (err) {
      console.error(err);
      setError('Error al procesar el archivo Excel. Verifique el formato.');
    } finally {
      setLoading(false);
    }
  };

  const togglePeriod = (period: string) => {
    setSelectedPeriods(prev => 
      prev.includes(period) 
        ? prev.filter(p => p !== period) 
        : [...prev, period]
    );
  };

  // Aggregated data based on selected periods, sociedad, and dynamic GROUP BY
  // 1. Data Aggregation for Selected Periods
  const aggregateData = useMemo(() => {
    if (!data || selectedPeriods.length === 0) return null;

    const filteredSnapshots = data.periods.filter(p => selectedPeriods.includes(p.periodLabel));
    
    let totalVolume = 0;
    let totalAmount = 0;
    const groupByMap = new Map<string, { count: number; amount: number }>();

    // Optimization: Single pass through records
    for (const s of filteredSnapshots) {
      const recordsToProcess = selectedSociedad === 'ALL' 
        ? s.records 
        : (s.recordsBySociedad?.[selectedSociedad] || []);

      for (const r of recordsToProcess) {
        const label = String(r['Antigüedad Label'] || '');
        if (vistas === 'sin_desfase' && !getIsSinDesfase(label)) continue;
        if (vistas === 'desfasado' && !getIsDesfasado(label)) continue;

        totalVolume++;
        const amt = (Number(r['Importe en MD']) || 0);
        totalAmount += amt;

        const groupVal = String(r[groupBy] || 'N/A').trim();
        const current = groupByMap.get(groupVal) || { count: 0, amount: 0 };
        current.count++;
        current.amount += amt;
        groupByMap.set(groupVal, current);
      }
    }

    const groupData = Array.from(groupByMap.entries())
      .map(([label, d]) => ({ label, ...d }))
      .sort((a, b) => b.amount - a.amount);

    return {
      totalVolume,
      totalAmount,
      groupData
    };
  }, [data, selectedPeriods, selectedSociedad, groupBy, vistas]);

  // 2. Specialized Ranking Logic for the ABSOLUTE LATEST period
  const rankingData = useMemo(() => {
    if (!data || data.periods.length === 0) return { socs: [], movsHigh: [], movsLow: [] };

    // Always take the last period (absolute latest)
    const period = data.periods[data.periods.length - 1];
    
    const socMap = new Map<string, { volOld: number; volNew: number; amtOld: number; amtNew: number }>();
    const posOld: RawRecord[] = [];
    const posNew: RawRecord[] = [];
    const negOld: RawRecord[] = [];
    const negNew: RawRecord[] = [];
    
    period.records.forEach(r => {
      const label = String(r['Antigüedad Label'] || '');
      if (vistas === 'sin_desfase' && !getIsSinDesfase(label)) return;
      if (vistas === 'desfasado' && !getIsDesfasado(label)) return;

      const soc = String(r['Soc.'] || r['Sociedad'] || 'N/A');
      const amt = Number(r['Importe en MD']) || 0;
      const curr = socMap.get(soc) || { volOld: 0, volNew: 0, amtOld: 0, amtNew: 0 };
      
      const isOldRec = getIsOld(label);
      if (isOldRec) {
        curr.volOld++;
        curr.amtOld += amt;
        if (amt > 0) posOld.push(r);
        else if (amt < 0) negOld.push(r);
      } else {
        curr.volNew++;
        curr.amtNew += amt;
        if (amt > 0) posNew.push(r);
        else if (amt < 0) negNew.push(r);
      }
      socMap.set(soc, curr);
    });

    const socArray = Array.from(socMap.entries()).map(([soc, d]) => ({ soc, ...d }));

    const movsHigh = [
      ...posOld.sort((a, b) => (Number(b['Importe en MD']) || 0) - (Number(a['Importe en MD']) || 0)).slice(0, 10),
      ...posNew.sort((a, b) => (Number(b['Importe en MD']) || 0) - (Number(a['Importe en MD']) || 0)).slice(0, 10)
    ].map(r => ({
      soc: String(r['Soc.'] || r['Sociedad'] || 'N/A'),
      amt: Number(r['Importe en MD']) || 0,
      label: String(r['Antigüedad Label'] || ''),
      isOld: getIsOld(String(r['Antigüedad Label'] || ''))
    }));

    const movsLow = [
      ...negOld.sort((a, b) => (Number(a['Importe en MD']) || 0) - (Number(b['Importe en MD']) || 0)).slice(0, 10),
      ...negNew.sort((a, b) => (Number(a['Importe en MD']) || 0) - (Number(b['Importe en MD']) || 0)).slice(0, 10)
    ].map(r => ({
      soc: String(r['Soc.'] || r['Sociedad'] || 'N/A'),
      amt: Number(r['Importe en MD']) || 0,
      label: String(r['Antigüedad Label'] || ''),
      isOld: getIsOld(String(r['Antigüedad Label'] || ''))
    }));

    return {
      socs: socArray,
      movsHigh,
      movsLow,
      latestLabel: period.periodLabel
    };
  }, [data, vistas]);


  // Helper to parse month names and years from period labels
  const parsePeriodLabel = (label: string): { month: string; year: number; original: string } => {
    const clean = label.toUpperCase().trim();
    let month = "DESCONOCIDO";
    let year = 2026; 
    
    const spanishMonths = [
      { name: "ENERO", keys: ["ENERO", "ENE"] },
      { name: "FEBRERO", keys: ["FEBRERO", "FEB"] },
      { name: "MARZO", keys: ["MARZO", "MAR"] },
      { name: "ABRIL", keys: ["ABRIL", "ABR"] },
      { name: "MAYO", keys: ["MAYO", "MAY"] },
      { name: "JUNIO", keys: ["JUNIO", "JUN"] },
      { name: "JULIO", keys: ["JULIO", "JUL"] },
      { name: "AGOSTO", keys: ["AGOSTO", "AGO"] },
      { name: "SEPTIEMBRE", keys: ["SEPTIEMBRE", "SEP"] },
      { name: "OCTUBRE", keys: ["OCTUBRE", "OCT"] },
      { name: "NOVIEMBRE", keys: ["NOVIEMBRE", "NOV"] },
      { name: "DICIEMBRE", keys: ["DICIEMBRE", "DIC"] }
    ];

    for (const m of spanishMonths) {
      if (m.keys.some(k => clean.includes(k))) {
        month = m.name;
        break;
      }
    }

    const yearMatch = clean.match(/\b(20\d{2}|\d{2})\b/);
    if (yearMatch) {
      const yr = parseInt(yearMatch[1], 10);
      year = yr < 100 ? 2000 + yr : yr;
    }

    return { month, year, original: label };
  };

  const getMonthIndex = (monthStr: string): number => {
    const clean = monthStr.toUpperCase().trim();
    if (clean.includes('ENE')) return 0;
    if (clean.includes('FEB')) return 1;
    if (clean.includes('MAR')) return 2;
    if (clean.includes('ABR')) return 3;
    if (clean.includes('MAY')) return 4;
    if (clean.includes('JUN')) return 5;
    if (clean.includes('JUL')) return 6;
    if (clean.includes('AGO')) return 7;
    if (clean.includes('SEP')) return 8;
    if (clean.includes('OCT')) return 9;
    if (clean.includes('NOV')) return 10;
    if (clean.includes('DIC')) return 11;
    return -1;
  };

  // Automated Year-over-Year Comparative Finder
  const monthComparisons = useMemo(() => {
    if (!data || data.periods.length === 0) return [];

    const grouped = new Map<string, typeof data.periods>();
    data.periods.forEach(p => {
      const parsed = parsePeriodLabel(p.periodLabel);
      const m = parsed.month;
      const list = grouped.get(m) || [];
      list.push(p);
      grouped.set(m, list);
    });

    const pairs: {
      monthName: string;
      basePeriod: typeof data.periods[0];
      compPeriod: typeof data.periods[0];
    }[] = [];

    grouped.forEach((list, monthName) => {
      const sorted = [...list].sort((a, b) => {
        const yearA = parsePeriodLabel(a.periodLabel).year;
        const yearB = parsePeriodLabel(b.periodLabel).year;
        return yearA - yearB;
      });

      if (sorted.length >= 2) {
        pairs.push({
          monthName,
          basePeriod: sorted[sorted.length - 2],
          compPeriod: sorted[sorted.length - 1]
        });
      }
    });

    return pairs.sort((a, b) => getMonthIndex(a.monthName) - getMonthIndex(b.monthName));
  }, [data]);

  // Comparative analysis calculations for all matching months
  const comparativeAnalysis = useMemo(() => {
    if (!data || monthComparisons.length === 0) return [];

    return monthComparisons.map(({ monthName, basePeriod, compPeriod }) => {
      const recordsBase = selectedSociedad === 'ALL'
        ? basePeriod.records
        : (basePeriod.recordsBySociedad?.[selectedSociedad] || []);

      const mapBase = new Map<string, { count: number; amount: number }>();
      recordsBase.forEach(r => {
        const label = String(r['Antigüedad Label'] || 'N/A');
        if (vistas === 'sin_desfase' && !getIsSinDesfase(label)) return;
        if (vistas === 'desfasado' && !getIsDesfasado(label)) return;

        const current = mapBase.get(label) || { count: 0, amount: 0 };
        current.count++;
        current.amount += (Number(r['Importe en MD']) || 0);
        mapBase.set(label, current);
      });

      const agingA = Array.from(mapBase.entries())
        .map(([label, d]) => ({ label, ...d }))
        .sort((a, b) => a.label.localeCompare(b.label));

      const recordsComp = selectedSociedad === 'ALL'
        ? compPeriod.records
        : (compPeriod.recordsBySociedad?.[selectedSociedad] || []);

      const mapComp = new Map<string, { count: number; amount: number }>();
      recordsComp.forEach(r => {
        const label = String(r['Antigüedad Label'] || 'N/A');
        if (vistas === 'sin_desfase' && !getIsSinDesfase(label)) return;
        if (vistas === 'desfasado' && !getIsDesfasado(label)) return;

        const current = mapComp.get(label) || { count: 0, amount: 0 };
        current.count++;
        current.amount += (Number(r['Importe en MD']) || 0);
        mapComp.set(label, current);
      });

      const agingB = Array.from(mapComp.entries())
        .map(([label, d]) => ({ label, ...d }))
        .sort((a, b) => a.label.localeCompare(b.label));

      return {
        monthName,
        baseLabel: basePeriod.periodLabel,
        compLabel: compPeriod.periodLabel,
        agingA,
        agingB
      };
    });
  }, [data, monthComparisons, selectedSociedad, vistas]);

  // Find the chronologically latest period matching a Yo-Yo month
  const latestMonthComp = useMemo(() => {
    if (monthComparisons.length === 0) return null;
    return monthComparisons[monthComparisons.length - 1];
  }, [monthComparisons]);

  // Extract society name with custom heuristics
  const getSocietyNameFromRecord = (r: RawRecord, socId: string): string => {
    const possibleKeys = [
      'Texto de sociedad', 'Texto de la sociedad', 'Texto sociedad', 
      'Texto de la Soc.', 'Texto de Soc.', 'Texto Soc.',
      'Denominación de la sociedad', 'Denominación sociedad', 'Denominación',
      'Nombre de sociedad', 'Nombre de la sociedad', 'Nombre sociedad', 'Nombre 1', 'Nombre',
      'Sociedad', 'Soc.'
    ];
    for (const k of possibleKeys) {
      if (r[k]) {
        const val = String(r[k]).trim();
        if (val && val !== socId && isNaN(Number(val))) {
          return val;
        }
      }
    }
    
    for (const key of Object.keys(r)) {
      const kLower = key.toLowerCase();
      if (kLower.includes('sociedad') || kLower.includes('nombre') || kLower.includes('texto') || kLower.includes('denomin')) {
        const val = String(r[key]).trim();
        if (val && val !== socId && isNaN(Number(val))) {
          return val;
        }
      }
    }
    
    const knownSocs: Record<string, string> = {
      '110': 'Distribuidora Liverpool',
      '111': 'LPC',
      '112': 'Suburbia',
      '0110': 'Distribuidora Liverpool',
      '0111': 'LPC',
      '0112': 'Suburbia',
    };
    return knownSocs[socId] || '';
  };

  const getSocDetailsForController = (controllerName: string, allRecords: RawRecord[]) => {
    const socIds = new Set<string>();
    const socNames = new Map<string, string>();

    const controllerRecords = allRecords.filter(r => {
      const rawName = getRecordController(r) || 'N/A';
      return rawName.trim() === controllerName;
    });

    controllerRecords.forEach(r => {
      const socId = String(r['Soc.'] || r['Sociedad'] || 'N/A').trim();
      if (socId && socId !== 'N/A') {
        socIds.add(socId);
        const name = getSocietyNameFromRecord(r, socId);
        if (name && !socNames.has(socId)) {
          socNames.set(socId, name);
        }
      }
    });

    return Array.from(socIds).sort().map(id => ({
      id,
      name: socNames.get(id) || getSocietyNameFromRecord({} as any, id) || 'Sociedad'
    }));
  };

  // Generic function to calculate KPIs for any period snapshot
  const calculateKPIForPeriod = (period: typeof data.periods[0], socFilter: string) => {
    const records = socFilter === 'ALL'
      ? period.records
      : (period.recordsBySociedad?.[socFilter] || []);

    const controllerMap = new Map<string, {
      name: string;
      totalVol: number;
      totalAmt: number;
      oldVol: number;
      oldAmt: number;
      sociedades: Set<string>;
    }>();

    records.forEach(r => {
      const rawName = getRecordController(r) || 'N/A';
      if (!rawName || rawName === 'null' || rawName === 'undefined') return;
      const normalizedName = rawName.trim();
      const soc = String(r['Soc.'] || r['Sociedad'] || 'N/A').trim();

      const current = controllerMap.get(normalizedName) || {
        name: normalizedName,
        totalVol: 0,
        totalAmt: 0,
        oldVol: 0,
        oldAmt: 0,
        sociedades: new Set<string>()
      };

      const amt = Number(r['Importe en MD']) || 0;
      const label = String(r['Antigüedad Label'] || '');

      current.totalVol++;
      current.totalAmt += amt;
      if (soc && soc !== 'N/A') {
        current.sociedades.add(soc);
      }

      if (getIsDesfasado(label)) {
        current.oldVol++;
        current.oldAmt += amt;
      }

      controllerMap.set(normalizedName, current);
    });

    const list = Array.from(controllerMap.values()).map(c => {
      const kpi1Val = c.totalVol > 0 ? (c.oldVol / c.totalVol) * 100 : 0;
      const kpi2Val = Math.abs(c.totalAmt) > 0 ? (Math.abs(c.oldAmt) / Math.abs(c.totalAmt)) * 100 : 0;

      return {
        name: c.name,
        totalVol: c.totalVol,
        totalAmt: c.totalAmt,
        oldVol: c.oldVol,
        oldAmt: c.oldAmt,
        sociedades: Array.from(c.sociedades).sort(),
        kpi1Val,
        kpi2Val
      };
    });

    return {
      periodLabel: period.periodLabel,
      byController: new Map(list.map(item => [item.name, item])),
      list
    };
  };

  // Structured KPIs data for the listed view (del mes actual más reciente tanto de año anterior como año actual)
  const kpiListData = useMemo(() => {
    if (!data || !latestMonthComp) return null;

    const { basePeriod, compPeriod, monthName } = latestMonthComp;

    const priorKPIResult = calculateKPIForPeriod(basePeriod, 'ALL');
    const currentKPIResult = calculateKPIForPeriod(compPeriod, 'ALL');

    const allControllers = new Set<string>([
      ...priorKPIResult.list.map(c => c.name),
      ...currentKPIResult.list.map(c => c.name)
    ]);

    const combinedRecords = [
      ...basePeriod.records,
      ...compPeriod.records
    ];

    const list = Array.from(allControllers).map(name => {
      const prior = priorKPIResult.byController.get(name);
      const current = currentKPIResult.byController.get(name);
      const societies = getSocDetailsForController(name, combinedRecords);

      return {
        name,
        societies,
        prior: prior ? {
          totalVol: prior.totalVol,
          totalAmt: prior.totalAmt,
          oldVol: prior.oldVol,
          oldAmt: prior.oldAmt,
          kpi1Val: prior.kpi1Val,
          kpi2Val: prior.kpi2Val,
          status1: getKPI1Status(prior.kpi1Val),
          status2: getKPI2Status(prior.kpi2Val)
        } : null,
        current: current ? {
          totalVol: current.totalVol,
          totalAmt: current.totalAmt,
          oldVol: current.oldVol,
          oldAmt: current.oldAmt,
          kpi1Val: current.kpi1Val,
          kpi2Val: current.kpi2Val,
          status1: getKPI1Status(current.kpi1Val),
          status2_fixed: getKPI2Status(current.kpi2Val)
        } : null
      };
    }).sort((a, b) => {
      const volA = a.current?.totalVol ?? 0;
      const volB = b.current?.totalVol ?? 0;
      return volB - volA;
    });

    return {
      monthName,
      baseLabel: basePeriod.periodLabel,
      currentLabel: compPeriod.periodLabel,
      list
    };
  }, [data, latestMonthComp]);



  const timeSeriesData = useMemo(() => {
    if (!data) return [];
    
    return data.periods.map(p => {
      let volOld = 0;
      let volNew = 0;
      let totalVol = 0;
      let totalAmount = 0;
      p.records.forEach(r => {
        const label = String(r['Antigüedad Label'] || '');
        if (vistas === 'sin_desfase' && !getIsSinDesfase(label)) return;
        if (vistas === 'desfasado' && !getIsDesfasado(label)) return;

        totalVol++;
        totalAmount += (Number(r['Importe en MD']) || 0);
        if (getIsOld(label)) volOld++;
        else volNew++;
      });
      return {
        name: p.periodLabel,
        volume: totalVol,
        volOld,
        volNew,
        amount: totalAmount
      };
    });
  }, [data, vistas]);


  const prevYearSnapshot = useMemo(() => {
    if (!data) return null;
    return data.periods.find(p => p.periodLabel.toUpperCase().includes('AÑO ANT')) || data.periods[0];
  }, [data]);

  const showInThousands = vistas === 'ejecutiva';

  const displayCurrency = (val: number) => {
    if (showInThousands) {
      const reducedVal = val / 1000;
      return new Intl.NumberFormat('es-MX', {
        style: 'currency',
        currency: 'MXN',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      }).format(reducedVal) + ' k';
    }
    return formatCurrency(val);
  };

  if (isInitialLoad) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-8 text-center font-sans">
        <motion.div
           initial={{ opacity: 0, scale: 0.95 }}
           animate={{ opacity: 1, scale: 1 }}
           className="space-y-6"
        >
          <div className="w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin mx-auto ring-4 ring-blue-50"></div>
          <div>
            <h1 className="text-2xl font-black text-slate-900 tracking-tight">KPI´s CB</h1>
            <p className="text-slate-500 font-medium">Recuperando información guardada...</p>
          </div>
        </motion.div>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center justify-center p-6 bg-[radial-gradient(circle_at_top_right,rgba(37,99,235,0.05),transparent),radial-gradient(circle_at_bottom_left,rgba(37,99,235,0.05),transparent)]">
        <motion.div 
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-2xl w-full text-center space-y-16"
        >
          <div className="space-y-6">
            <div className="inline-flex items-center gap-3 bg-blue-50 px-6 py-2 rounded-full border border-blue-100 mb-4">
              <div className="w-2 h-2 bg-blue-600 rounded-full animate-pulse" />
              <span className="text-[10px] font-black text-blue-700 uppercase tracking-[0.3em]">Banking Intelligence System</span>
            </div>
            <h1 className="text-7xl font-black text-slate-900 tracking-tighter uppercase italic leading-none">
              RECON <span className="text-blue-700 not-italic font-extralight italic">PRO</span>
            </h1>
            <p className="text-slate-500 text-lg font-medium max-w-md mx-auto">Consolida, analiza y visualiza el comportamiento de tus partidas conciliatorias con precisión ejecutiva.</p>
          </div>

          <label className="block group">
            <div className="relative overflow-hidden bg-white border-2 border-slate-100 rounded-[3rem] p-16 cursor-pointer transition-all hover:border-blue-500 hover:shadow-[0_40px_80px_-20px_rgba(37,99,235,0.15)] group-active:scale-95">
              <div className="absolute top-0 right-0 p-8 opacity-5 group-hover:opacity-10 transition-opacity">
                <FileSpreadsheet size={160} />
              </div>
              <div className="relative flex flex-col items-center space-y-8">
                <div className="w-20 h-20 bg-slate-900 text-white rounded-3xl flex items-center justify-center shadow-2xl group-hover:bg-blue-700 transition-colors">
                  <FileUp size={36} />
                </div>
                <div className="space-y-2">
                  <p className="text-2xl font-black text-slate-900 uppercase">Cargar Dashboard Maestro</p>
                  <p className="text-sm text-slate-400 font-medium">Hojas Multi-Periodo • Auto-Aging (Col N) • Pivot General</p>
                </div>
              </div>
              <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} disabled={loading} />
            </div>
          </label>

          {loading && (
            <div className="fixed inset-0 z-50 bg-white/90 backdrop-blur-2xl flex flex-col items-center justify-center">
              <motion.div 
                animate={{ rotate: 360 }}
                transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
                className="w-16 h-16 border-4 border-blue-700 border-t-transparent rounded-full mb-8"
              />
              <p className="text-xl font-black text-slate-900 uppercase tracking-tighter italic">Procesando Inteligencia de Datos...</p>
              <p className="text-slate-400 text-xs font-bold uppercase tracking-[0.4em] mt-4">Mapeando Nodos y Jerarquías</p>
            </div>
          )}
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-[#f8fafc] font-sans text-slate-900 overflow-hidden">
      {/* Sidebar - Nano Design */}
      <aside className="w-64 bg-white border-r border-slate-100 flex flex-col p-5 shrink-0 z-30 shadow-sm overflow-hidden">
        <div className="mb-8">
          <div className="flex items-center gap-2">
             <div className="w-8 h-8 bg-slate-900 rounded-lg flex items-center justify-center shadow-lg">
                <BarChart3 size={16} className="text-white" />
             </div>
             <div>
                <h1 className="text-sm font-black tracking-tighter italic uppercase">Recon <span className="text-blue-700 font-extralight not-italic">Pro</span></h1>
             </div>
          </div>
        </div>

        <div className="space-y-8 flex-1 overflow-y-auto no-scrollbar pb-6">
          {/* Multi-Period Selector */}
          <div>
            <div className="flex items-center justify-between mb-4">
              <label className="text-[8px] uppercase text-slate-400 font-black tracking-[0.2em]">Periodos</label>
            </div>
            <div className="space-y-1">
              {data.periods.map(p => (
                <button
                  key={p.periodLabel}
                  onClick={() => togglePeriod(p.periodLabel)}
                  className={cn(
                    "w-full flex items-center justify-between px-3 py-2.5 rounded-xl border transition-all text-left",
                    selectedPeriods.includes(p.periodLabel) 
                      ? "bg-slate-900 border-slate-900 text-white shadow-md text-[10px]" 
                      : "bg-white border-slate-100 hover:border-blue-200 text-slate-500 text-[10px]"
                  )}
                >
                  <span className="font-bold truncate max-w-[120px]">{p.periodLabel}</span>
                  <span className={cn("font-mono text-[8px]", selectedPeriods.includes(p.periodLabel) ? "text-blue-400" : "text-slate-300")}>{p.records.length}</span>
                </button>
              ))}
            </div>
          </div>

          {/* Group By Selector */}
          <div>
            <label className="text-[8px] uppercase text-slate-400 font-black tracking-[0.2em] block mb-2">Eje Análisis</label>
            <select 
              value={groupBy}
              onChange={(e) => setGroupBy(e.target.value)}
              className="w-full bg-slate-50 border-slate-100 text-slate-900 rounded-xl px-3 py-2 text-[10px] font-bold uppercase focus:ring-2 focus:ring-blue-100 outline-none transition-all shadow-xs"
            >
              {data.availableColumns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>

          {/* Segment Selector */}
          <div>
            <label className="text-[8px] uppercase text-slate-400 font-black tracking-[0.2em] block mb-2">Sociedad</label>
            <select 
              value={selectedSociedad}
              onChange={(e) => setSelectedSociedad(e.target.value)}
              className="w-full bg-slate-50 border-slate-100 text-slate-900 rounded-xl px-3 py-2 text-[10px] font-bold uppercase focus:ring-2 focus:ring-blue-100 outline-none transition-all shadow-xs"
            >
              <option value="ALL">Consolidado</option>
              {data.allSociedades.map(soc => (
                <option key={soc} value={soc}>{soc}</option>
              ))}
            </select>
          </div>
        </div>

        <label className="mt-4 flex items-center justify-center gap-2 text-[8px] font-black uppercase tracking-widest text-slate-400 hover:text-blue-700 transition-all py-4 border-t border-slate-50 cursor-pointer group">
          <FileUp size={12} className="group-hover:scale-110 transition-transform" /> 
          Reemplazar Datos Excel
          <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
        </label>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col p-4 overflow-y-auto no-scrollbar scroll-smooth">
        <header className="flex justify-between items-center mb-4 shrink-0">
          <div className="space-y-0.5">
            <div className="flex items-center gap-2">
              <span className="text-[7px] font-black text-blue-700 uppercase tracking-[0.2em] bg-blue-50 px-2 py-0.5 rounded-full">Executive Analytics</span>
              <span className="text-[7px] font-bold text-slate-300 uppercase tracking-widest">{selectedPeriods.length} Períodos</span>
            </div>
            <h2 className="text-xl font-black text-slate-900 tracking-tighter uppercase italic leading-none">Cifras Reconciliadas</h2>
          </div>

          <div className="flex items-center gap-4 bg-white px-4 py-2 rounded-xl border border-slate-100 shadow-xs">
             <div className="text-right">
               <p className="text-[7px] font-black text-slate-400 uppercase tracking-widest mb-0">Partidas</p>
               <p className="text-sm font-black text-slate-900 leading-none">{aggregateData ? aggregateData.totalVolume.toLocaleString() : '0'}</p>
             </div>
             <div className="h-4 w-px bg-slate-100" />
             <div className="text-right">
               <p className="text-[7px] font-black text-blue-700 uppercase tracking-widest mb-0">Importe</p>
               <p className="text-sm font-black text-blue-700 leading-none">{aggregateData ? displayCurrency(aggregateData.totalAmount) : '$0.00'}</p>
             </div>
          </div>
        </header>

        {/* Dynamic Vistas Tabs Selector */}
        <div className="flex flex-wrap gap-2 mb-6 bg-slate-100/90 p-1.5 rounded-2xl w-fit border border-slate-200/50 backdrop-blur-md">
          <button
            onClick={() => setVistas('acumulado')}
            className={cn(
              "flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-black uppercase tracking-tight transition-all",
              vistas === 'acumulado'
                ? "bg-slate-900 text-white shadow-lg scale-[1.02]"
                : "text-slate-500 hover:text-slate-900/80 hover:bg-white/40"
            )}
          >
            <BarChart3 size={14} />
            Partidas Abiertas
          </button>
          <button
            onClick={() => setVistas('sin_desfase')}
            className={cn(
              "flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-black uppercase tracking-tight transition-all",
              vistas === 'sin_desfase'
                ? "bg-emerald-600 text-white shadow-lg scale-[1.02]"
                : "text-slate-500 hover:text-emerald-700 hover:bg-emerald-50/50"
            )}
          >
            <CheckCircle2 size={14} className="text-emerald-500 bg-white rounded-full p-px" />
            Menores a 90 días
          </button>
          <button
            onClick={() => setVistas('desfasado')}
            className={cn(
              "flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-black uppercase tracking-tight transition-all",
              vistas === 'desfasado'
                ? "bg-rose-600 text-white shadow-lg scale-[1.02]"
                : "text-slate-500 hover:text-rose-600 hover:bg-rose-50/50"
            )}
          >
            <History size={14} className="text-rose-400" />
            Mayores a 91 días
          </button>
          <button
            onClick={() => setVistas('ejecutiva')}
            className={cn(
              "flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-black uppercase tracking-tight transition-all",
              vistas === 'ejecutiva'
                ? "bg-blue-600 text-white shadow-lg scale-[1.02]"
                : "text-slate-500 hover:text-blue-600 hover:bg-blue-50/50"
            )}
          >
            <FileSpreadsheet size={14} className="text-blue-400" />
            KPI´S
          </button>
        </div>

        {/* Subtitle Indicator for Thousands unit under KPI's */}
        {showInThousands && (
          <div className="mb-6 px-4 py-2.5 bg-blue-50 text-blue-700 text-[10px] font-black uppercase tracking-[0.12em] rounded-xl border border-blue-100 flex items-center gap-2 w-fit animate-pulse">
            <div className="w-1.5 h-1.5 rounded-full bg-blue-600" />
            <span>* KPI´S: Cifras de Importes representadas en Miles de Pesos ($ k). Gráficas ocultas por claridad ejecutiva.</span>
          </div>
        )}

        {aggregateData && (
          <div className="space-y-6 animate-in fade-in duration-500">
            
            {/* 1. LARGE SPLIT TREND CHARTS */}
            {vistas !== 'ejecutiva' && (
            <div className={`grid grid-cols-1 gap-4 ${vistas === 'acumulado' ? 'md:grid-cols-2' : ''}`}>
               {/* Trend Chart: Age < 90 Days */}
               {(vistas === 'acumulado' || vistas === 'sin_desfase') && (
                <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm h-[280px] flex flex-col relative overflow-hidden group">
                  <div className="flex justify-between items-center mb-4 relative z-10">
                    <div className="space-y-0.5">
                      <h3 className="font-bold text-sm text-slate-900 leading-none">Tendencia Reciente</h3>
                      <p className="text-slate-400 text-[8px] font-bold">Partidas Menores a 90 Días</p>
                    </div>
                    <div className="flex items-center gap-2 px-3 py-1 bg-emerald-50 rounded-full">
                      <div className="w-1.5 h-1.5 rounded-full bg-emerald-600" />
                      <span className="text-[8px] font-bold text-emerald-700">V. Reciente</span>
                    </div>
                  </div>
                  <div className="flex-1 min-h-0 relative z-10">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={timeSeriesData.filter(d => selectedPeriods.includes(d.name))}>
                        <defs>
                          <linearGradient id="recentGradient" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#10b981" stopOpacity={0.15}/>
                            <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="5 5" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="name" stroke="#cbd5e1" fontSize={8} tickLine={false} axisLine={false} tick={{ fontWeight: 800 }} dy={10} />
                        <YAxis stroke="#cbd5e1" fontSize={8} tickLine={false} axisLine={false} tickFormatter={(v) => v.toLocaleString()} tick={{ fontWeight: 800 }} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 30px -5px rgb(0 0 0 / 0.1)', padding: '10px' }} />
                        <Area type="monotone" dataKey="volNew" stroke="#10b981" strokeWidth={3} fillOpacity={1} fill="url(#recentGradient)" activeDot={{ r: 4 }} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
               </div>
               )}

               {/* Trend Chart: Age > 91 Days */}
               {(vistas === 'acumulado' || vistas === 'desfasado') && (
                <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm h-[280px] flex flex-col relative overflow-hidden group">
                  <div className="flex justify-between items-center mb-4 relative z-10">
                    <div className="space-y-0.5">
                      <h3 className="font-bold text-sm text-slate-900 leading-none">Tendencia Antigua</h3>
                      <p className="text-slate-400 text-[8px] font-bold">Partidas Mayores a 91 Días</p>
                    </div>
                    <div className="flex items-center gap-2 px-3 py-1 bg-rose-50 rounded-full">
                      <div className="w-1.5 h-1.5 rounded-full bg-rose-600" />
                      <span className="text-[8px] font-bold text-rose-700">V. Antiguo</span>
                    </div>
                  </div>
                  <div className="flex-1 min-h-0 relative z-10">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={timeSeriesData.filter(d => selectedPeriods.includes(d.name))}>
                        <defs>
                          <linearGradient id="oldGradient" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#f43f5e" stopOpacity={0.15}/>
                            <stop offset="95%" stopColor="#f43f5e" stopOpacity={0}/>
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="5 5" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="name" stroke="#cbd5e1" fontSize={8} tickLine={false} axisLine={false} tick={{ fontWeight: 800 }} dy={10} />
                        <YAxis stroke="#cbd5e1" fontSize={8} tickLine={false} axisLine={false} tickFormatter={(v) => v.toLocaleString()} tick={{ fontWeight: 800 }} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 30px -5px rgb(0 0 0 / 0.1)', padding: '10px' }} />
                        <Area type="monotone" dataKey="volOld" stroke="#f43f5e" strokeWidth={3} fillOpacity={1} fill="url(#oldGradient)" activeDot={{ r: 4 }} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
               </div>
               )}
            </div>
            )}

            {/* 2. AUTOMATED COMPARISON ENGINE - Consolidated Volume and Amount per Month */}
            {vistas !== 'ejecutiva' && comparativeAnalysis.length > 0 && (
              <div className="space-y-8 animate-in fade-in duration-500 font-sans">
                {comparativeAnalysis.map((analysis) => (
                  <div key={analysis.monthName} className="space-y-4 border-l-4 border-blue-600 pl-4 py-2 bg-slate-50/50 rounded-r-3xl p-4">
                    <div className="flex flex-col sm:flex-row gap-4 bg-slate-900 p-4 rounded-3xl text-white items-start sm:items-center justify-between shadow-lg">
                      <div className="flex items-center gap-4">
                        <CalendarDays size={18} className="text-blue-400" />
                        <div>
                          <h3 className="text-xs font-black uppercase tracking-tight italic font-sans">
                            Comparativo Mensual: {analysis.monthName}
                          </h3>
                          <p className="text-[9px] text-slate-400 uppercase font-black font-sans">
                            {analysis.baseLabel} vs {analysis.compLabel}
                          </p>
                        </div>
                      </div>
                      <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest font-sans">
                        Eje de Análisis por {groupBy}
                      </div>
                    </div>

                    <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                      <ComparisonTable 
                        title={`Comparativo de Volumen de Transacciones (${analysis.monthName})`} 
                        baseLabel={analysis.baseLabel}
                        compLabel={analysis.compLabel}
                        dataA={analysis.agingA} 
                        dataB={analysis.agingB} 
                        type="count" 
                        color="bg-slate-800" 
                        showInThousands={showInThousands} 
                      />
                      <ComparisonTable 
                        title={`Comparativo de Importe de Transacciones (${analysis.monthName})`} 
                        baseLabel={analysis.baseLabel}
                        compLabel={analysis.compLabel}
                        dataA={analysis.agingA} 
                        dataB={analysis.agingB} 
                        type="amount" 
                        color="bg-blue-800" 
                        showInThousands={showInThousands} 
                      />
                    </div>
                  </div>
                ))}
              </div>
            )}


            {/* 3. LATEST PERIOD RANKINGS (Not Variance) */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
               <RankingTable 
                  title={`Top 10 Sociedades x Volumen (${rankingData.latestLabel})`} 
                  data={rankingData.socs} showInThousands={showInThousands} 
                  metric="vol" 
                  color="bg-slate-900"
               />
               <MovementRankingTable 
                  title={`Top 20 Importes (+) (${rankingData.latestLabel})`} 
                  data={rankingData.movsHigh} showInThousands={showInThousands} 
                  color="bg-blue-700"
               />
               <MovementRankingTable 
                  title={`Top 20 Importes (-) (${rankingData.latestLabel})`} 
                  data={rankingData.movsLow} showInThousands={showInThousands} 
                  color="bg-rose-700"
                />
             </div>

            {/* 4. EXECUTIVE KPI LIST VIEW (Replaces efficiency monitor and old KPI table) */}
            {vistas === 'ejecutiva' && kpiListData && (
              <div className="space-y-6 pt-4 font-sans max-w-7xl mx-auto">
                <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 mb-2">
                  <div className="flex items-center gap-3">
                    <div className="p-3 bg-slate-900 text-white rounded-2xl shadow-xl border border-slate-800">
                      <Percent size={22} className="text-blue-400" />
                    </div>
                    <div>
                      <h3 className="text-base font-black text-slate-900 tracking-tight uppercase italic leading-tight">
                        Evaluación de KPI´S por Contralor Responsable
                      </h3>
                      <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mt-0.5">
                        Comparativo Default de Mes Más Reciente: {kpiListData.monthName} ({kpiListData.baseLabel} vs {kpiListData.currentLabel})
                      </p>
                    </div>
                  </div>
                  <div className="px-3 py-1.5 bg-slate-100 border border-slate-200 rounded-full text-[9px] font-black text-slate-700 uppercase tracking-wider">
                    Unidad de Importe: Miles ($ k)
                  </div>
                </div>

                {/* Tabular List structure */}
                <div className="space-y-4">
                  {kpiListData.list.map((row) => (
                    <div 
                      key={row.name} 
                      className="bg-white rounded-[2rem] border border-slate-100 shadow-xl overflow-hidden hover:border-blue-200 transition-all duration-300 grid grid-cols-1 lg:grid-cols-12 gap-6 p-6 group animate-in fade-in duration-300"
                    >
                      {/* Left Block: Controller & Assigned Societies (Col Span 4) */}
                      <div className="lg:col-span-4 lg:border-r border-slate-100 lg:pr-6 flex flex-col justify-between space-y-4">
                        <div>
                          <div className="flex items-center gap-2 mb-2">
                            <span className="w-2.5 h-2.5 bg-blue-600 rounded-full animate-pulse" />
                            <h4 className="text-sm font-black text-slate-900 tracking-tight uppercase italic">
                              {row.name || 'Sin Contralor Asignado'}
                            </h4>
                          </div>
                          <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-3">
                            Contralor Corporativo
                          </p>
                        </div>

                        {/* Society listing with styled badges */}
                        <div className="space-y-1">
                          <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest mb-1.5 font-sans">
                            Sociedades Incluidas
                          </p>
                          <div className="flex flex-wrap gap-1.5 max-h-[140px] overflow-y-auto no-scrollbar py-0.5">
                            {row.societies.map((soc) => (
                              <span 
                                key={soc.id} 
                                className="px-2.5 py-1 bg-slate-50 hover:bg-slate-100 text-slate-705 rounded-xl text-[9px] font-bold tracking-tight border border-slate-200/60 transition-colors flex items-center gap-1.5"
                              >
                                <span className="text-[8px] px-1 bg-slate-200 text-slate-800 rounded font-black font-mono">
                                  Soc {soc.id}
                                </span>
                                <span className="uppercase text-slate-600 font-semibold truncate max-w-[120px] font-sans">
                                  {soc.name}
                                </span>
                              </span>
                            ))}
                          </div>
                        </div>
                      </div>

                      {/* Right Block: Dynamic Year-over-Year Comparisons & Semaphores (Col Span 8) */}
                      <div className="lg:col-span-8 grid grid-cols-1 md:grid-cols-2 gap-6 font-sans">
                        
                        {/* Column 1: Prior Month / Year */}
                        {row.prior ? (
                          <div className="bg-slate-50/50 rounded-2xl p-4 border border-slate-100 flex flex-col justify-between font-sans">
                            <div>
                              <div className="flex items-center justify-between border-b border-slate-100 pb-2 mb-3 font-sans">
                                <span className="text-[10px] uppercase font-black text-slate-400 tracking-widest">
                                  Periodo Anterior
                                </span>
                                <span className="text-[11px] font-black text-slate-900 bg-slate-200/50 px-2 py-0.5 rounded-lg uppercase">
                                  {kpiListData.baseLabel}
                                </span>
                              </div>

                              <div className="grid grid-cols-2 gap-4 mb-4 font-sans">
                                <div>
                                  <p className="text-[8px] font-black text-slate-400 uppercase mb-0.5">Partidas Abiertas</p>
                                  <p className="text-sm font-black text-slate-900">{row.prior.totalVol.toLocaleString()}</p>
                                </div>
                                <div>
                                  <p className="text-[8px] font-black text-slate-400 uppercase mb-0.5">Importe Total</p>
                                  <p className="text-sm font-black text-blue-700 font-mono">
                                    ${(row.prior.totalAmt / 1000).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })} k
                                  </p>
                                </div>
                              </div>
                            </div>

                            {/* Semaphores */}
                            <div className="space-y-2 border-t border-slate-100 pt-3">
                              <div className="flex items-center justify-between text-[10px]">
                                <span className="text-slate-500 font-bold uppercase tracking-tight">KPI 1 (Volumen):</span>
                                <div className="flex items-center gap-1.5">
                                  <span className="font-mono font-black text-slate-800">{row.prior.kpi1Val.toFixed(2)}%</span>
                                  <span className={cn(
                                    "px-2 py-0.5 rounded-full text-[8px] font-black uppercase tracking-wider flex items-center gap-1",
                                    row.prior.status1.color
                                  )}>
                                    <span className="w-1 h-1 rounded-full bg-current animate-pulse" />
                                    {row.prior.status1.text}
                                  </span>
                                </div>
                              </div>

                              <div className="flex items-center justify-between text-[10px]">
                                <span className="text-slate-500 font-bold uppercase tracking-tight">KPI 2 (Importe):</span>
                                <div className="flex items-center gap-1.5">
                                  <span className="font-mono font-black text-slate-800">{row.prior.kpi2Val.toFixed(2)}%</span>
                                  <span className={cn(
                                    "px-2 py-0.5 rounded-full text-[8px] font-black uppercase tracking-wider flex items-center gap-1",
                                    row.prior.status2.color
                                  )}>
                                    <span className="w-1 h-1 rounded-full bg-current animate-pulse" />
                                    {row.prior.status2.text}
                                  </span>
                                </div>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div className="bg-slate-50/20 rounded-2xl p-4 border border-slate-100 flex items-center justify-center text-slate-400 text-xs italic font-medium font-sans">
                            No disponible
                          </div>
                        )}

                        {/* Column 2: Current Month / Year */}
                        {row.current ? (
                          <div className="bg-blue-50/10 rounded-2xl p-4 border border-blue-100/40 flex flex-col justify-between font-sans">
                            <div>
                              <div className="flex items-center justify-between border-b border-blue-50 pb-2 mb-3">
                                <span className="text-[10px] uppercase font-black text-blue-700/80 tracking-widest">
                                  Periodo Evaluado
                                </span>
                                <span className="text-[11px] font-black text-blue-700 bg-blue-100/55 px-2 py-0.5 rounded-lg uppercase">
                                  {kpiListData.currentLabel}
                                </span>
                              </div>

                              <div className="grid grid-cols-2 gap-4 mb-4">
                                <div>
                                  <p className="text-[8px] font-black text-slate-400 uppercase mb-0.5">Partidas Abiertas</p>
                                  <p className="text-sm font-black text-slate-900">{row.current.totalVol.toLocaleString()}</p>
                                </div>
                                <div>
                                  <p className="text-[8px] font-black text-slate-400 uppercase mb-0.5">Importe Total</p>
                                  <p className="text-sm font-black text-blue-700 font-mono">
                                    ${(row.current.totalAmt / 1000).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })} k
                                  </p>
                                </div>
                              </div>
                            </div>

                            {/* Semaphores */}
                            <div className="space-y-2 border-t border-blue-50 pt-3">
                              <div className="flex items-center justify-between text-[10px]">
                                <span className="text-slate-500 font-bold uppercase tracking-tight">KPI 1 (Volumen):</span>
                                <div className="flex items-center gap-1.5">
                                  <span className="font-mono font-black text-slate-800">{row.current.kpi1Val.toFixed(2)}%</span>
                                  <span className={cn(
                                    "px-2 py-0.5 rounded-full text-[8px] font-black uppercase tracking-wider flex items-center gap-1",
                                    row.current.status1.color
                                  )}>
                                    <span className="w-1 h-1 rounded-full bg-current animate-pulse" />
                                    {row.current.status1.text}
                                  </span>
                                </div>
                              </div>

                              <div className="flex items-center justify-between text-[10px]">
                                <span className="text-slate-500 font-bold uppercase tracking-tight">KPI 2 (Importe):</span>
                                <div className="flex items-center gap-1.5">
                                  <span className="font-mono font-black text-slate-800">{row.current.kpi2Val.toFixed(2)}%</span>
                                  <span className={cn(
                                    "px-2 py-0.5 rounded-full text-[8px] font-black uppercase tracking-wider flex items-center gap-1",
                                    getKPI2Status(row.current.kpi2Val).color
                                  )}>
                                    <span className="w-1 h-1 rounded-full bg-current animate-pulse" />
                                    {getKPI2Status(row.current.kpi2Val).text}
                                  </span>
                                </div>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div className="bg-slate-50/20 rounded-2xl p-4 border border-none flex items-center justify-center text-slate-400 text-xs italic font-medium font-sans">
                            No disponible
                          </div>
                        )}

                      </div>
                    </div>
                  ))}
                </div>

                {/* Elegant KPI Methodology note */}
                <div className="bg-slate-50 rounded-2xl p-4 border border-slate-100">
                  <h4 className="text-[10px] font-black text-slate-900 uppercase tracking-widest mb-3 flex items-center gap-2">
                    <History size={12} className="text-blue-700 font-sans" />
                    Criterios y Reglas de Evaluación de KPI´s
                  </h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-balance text-[10px] text-slate-600 leading-relaxed font-sans">
                    <div className="space-y-1 p-2 bg-white rounded-lg border border-slate-100/50">
                      <p className="text-[8px] font-black text-slate-400 uppercase tracking-wider"><span className="text-blue-600">KPI 1:</span> Partidas Antiguas (Volumen)</p>
                      <p className="text-slate-600 leading-tight">Porcentaje de partidas cuya antigüedad es mayor a 91 Días respecto al total asignado a cada Contralor. <strong className="text-slate-800 font-black">Cumple:</strong> si se sitúa en un rango del <strong className="text-blue-700 font-black">4% al 6%</strong> inclusive.</p>
                    </div>
                    <div className="space-y-1 p-2 bg-white rounded-lg border border-slate-100/50">
                      <p className="text-[8px] font-black text-slate-400 uppercase tracking-wider"><span className="text-blue-600">KPI 2:</span> Partidas Antiguas (Importe)</p>
                      <p className="text-slate-600 leading-tight">Mide el peso del importe financiero de partidas con antigüedad de más de 91 Días contra el total de su cartera. <strong className="text-slate-800 font-black">Cumple:</strong> de <strong className="text-blue-700 font-black">1.5% a 3.5%</strong> inclusive.</p>
                    </div>
                  </div>
                </div>
              </div>
            )}



            {/* FOOTER */}
            <footer className="pt-6 pb-4 border-t border-slate-100 flex justify-between items-center text-[8px] font-black uppercase text-slate-300">
               <div className="flex gap-4">
                  <span>Recon Pro v3.2 Nano</span>
                  <span>© 2026 analytical services</span>
               </div>
               <button className="bg-slate-900 text-white px-6 py-2 rounded-full hover:scale-105 transition-all text-[8px]">Exportar Reporte</button>
            </footer>
          </div>
        )}
      </main>
    </div>
  );
}

// Consolidated Comparison Table Component
const ComparisonTable = ({ 
  title, 
  baseLabel,
  compLabel,
  dataA, 
  dataB, 
  type, 
  color, 
  showInThousands 
}: { 
  title: string, 
  baseLabel: string,
  compLabel: string,
  dataA: any[], 
  dataB: any[], 
  type: 'count' | 'amount', 
  color: string, 
  showInThousands?: boolean 
}) => {
  const formatVal = (v: number) => {
    if (type === 'count') return formatNumber(v);
    if (showInThousands) {
      return new Intl.NumberFormat('es-MX', {
        style: 'currency',
        currency: 'MXN',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      }).format(v / 1000) + ' k';
    }
    return formatCurrency(v).split('.')[0];
  };

  const combinedRows = useMemo(() => {
    const labels = Array.from(new Set([
      ...dataA.map(x => x.label),
      ...dataB.map(x => x.label)
    ])).sort((a, b) => a.localeCompare(b.label));

    return labels.map(label => {
      const aRow = dataA.find(x => x.label === label) || { count: 0, amount: 0 };
      const bRow = dataB.find(x => x.label === label) || { count: 0, amount: 0 };
      const baseVal = type === 'count' ? aRow.count : aRow.amount;
      const compVal = type === 'count' ? bRow.count : bRow.amount;
      return {
        label,
        baseVal,
        compVal,
        diffVal: baseVal - compVal
      };
    });
  }, [dataA, dataB, type]);

  const totalBase = combinedRows.reduce((s, r) => s + r.baseVal, 0);
  const totalComp = combinedRows.reduce((s, r) => s + r.compVal, 0);
  const totalDiff = totalBase - totalComp;

  return (
    <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-full font-sans">
      <div className={cn("px-5 py-4 text-white flex justify-between items-center", color)}>
        <div>
          <h4 className="text-xs font-black tracking-tight">{title}</h4>
          <p className="text-[8px] opacity-75 font-semibold uppercase tracking-wider">
            Cifras en {type === 'count' ? 'Unidades' : showInThousands ? 'Miles de Pesos ($ k)' : 'Pesos'}
          </p>
        </div>
        <div className="text-right">
          <p className="text-xs font-black">{formatVal(totalBase)}</p>
          <p className="text-[8px] opacity-75 font-medium">Total Base</p>
        </div>
      </div>
      <div className="flex-1 overflow-x-auto no-scrollbar">
        <table className="w-full text-left text-xs">
          <thead className="bg-slate-50 font-bold text-slate-400 text-[9px] tracking-wider border-b border-slate-100">
            <tr>
              <th className="px-5 py-2.5">Antigüedad</th>
              <th className="px-5 py-2.5 text-right">{baseLabel}</th>
              <th className="px-5 py-2.5 text-right">{compLabel}</th>
              <th className="px-5 py-2.5 text-right">Diferencia (Δ)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {combinedRows.map((row) => (
              <tr key={row.label} className="hover:bg-slate-50/50 transition-colors">
                <td className="px-5 py-2.5 font-semibold text-slate-600 truncate max-w-[150px]">{row.label}</td>
                <td className="px-5 py-2.5 text-right font-medium text-slate-950">{formatVal(row.baseVal)}</td>
                <td className="px-5 py-2.5 text-right font-medium text-slate-500">{formatVal(row.compVal)}</td>
                <td className={cn(
                  "px-5 py-2.5 text-right font-bold",
                  row.diffVal > 0 ? "text-rose-600" : row.diffVal < 0 ? "text-emerald-600" : "text-slate-400"
                )}>
                  {row.diffVal > 0 ? `+${formatVal(row.diffVal)}` : formatVal(row.diffVal)}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot className="bg-slate-50/50 font-black border-t border-slate-100 text-[11px]">
            <tr>
              <td className="px-5 py-2.5 text-slate-900">Total</td>
              <td className="px-5 py-2.5 text-right text-slate-950">{formatVal(totalBase)}</td>
              <td className="px-5 py-2.5 text-right text-slate-500">{formatVal(totalComp)}</td>
              <td className={cn(
                "px-5 py-2.5 text-right",
                totalDiff > 0 ? "text-rose-700" : totalDiff < 0 ? "text-emerald-700" : "text-slate-400"
              )}>
                {totalDiff > 0 ? `+${formatVal(totalDiff)}` : formatVal(totalDiff)}
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
};

// Simplified Ranking Component
const RankingTable = ({ title, data, metric, color, showInThousands }: { title: string, data: any[], metric: 'vol' | 'amtHigh' | 'amtLow', color: string, showInThousands?: boolean }) => {
  const sorted = useMemo(() => {
    return [...data].sort((a, b) => {
      if (metric === 'vol') return (b.volOld + b.volNew) - (a.volOld + a.volNew);
      if (metric === 'amtHigh') return (b.amtOld + b.amtNew) - (a.amtOld + a.amtNew);
      if (metric === 'amtLow') return (a.amtOld + a.amtNew) - (b.amtOld + b.amtNew);
      return 0;
    }).slice(0, 10);
  }, [data, metric]);

  const formatVal = (v: number, forceVol?: boolean) => {
    if (forceVol || metric === 'vol') return formatNumber(v);
    if (showInThousands) {
      return new Intl.NumberFormat('es-MX', {
        style: 'currency',
        currency: 'MXN',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      }).format(v / 1000) + ' k';
    }
    return formatCurrency(v).split('.')[0];
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-full font-sans">
      <div className={cn("px-4 py-3 text-white shadow-xs", color)}>
        <h4 className="text-[10px] font-black">{title}</h4>
      </div>
      <div className="flex-1 max-h-[400px] overflow-y-auto no-scrollbar">
        <table className="w-full text-left text-[9px] relative">
          <thead className="bg-slate-50 font-bold text-slate-400 border-b border-slate-100 sticky top-0 z-10 shadow-xs">
            <tr>
              <th className="px-4 py-2">Soc.</th>
              <th className="px-4 py-2 text-right">{">"}91d</th>
              <th className="px-4 py-2 text-right">{"<"}90d</th>
              <th className="px-4 py-2 text-right bg-slate-100/50">Total</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {sorted.map((row) => {
              const valOld = metric === 'vol' ? row.volOld : row.amtOld;
              const valNew = metric === 'vol' ? row.volNew : row.amtNew;
              const total = valOld + valNew;
              return (
                <tr key={row.soc} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-2 font-bold text-slate-900 truncate max-w-[80px]">{row.soc}</td>
                  <td className="px-4 py-2 text-right text-slate-500 font-medium">
                    {formatVal(valOld)}
                  </td>
                  <td className="px-4 py-2 text-right text-slate-500 font-medium">
                    {formatVal(valNew)}
                  </td>
                  <td className="px-4 py-2 text-right font-black text-slate-900 bg-slate-50/20">
                    {formatVal(total)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// Ranking for individual movements
const MovementRankingTable = ({ title, data, color, showInThousands }: { title: string, data: any[], color: string, showInThousands?: boolean }) => {
  const formatVal = (v: number) => {
    if (showInThousands) {
      return new Intl.NumberFormat('es-MX', {
        style: 'currency',
        currency: 'MXN',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      }).format(v / 1000) + ' k';
    }
    return formatCurrency(v).split('.')[0];
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-full font-sans">
      <div className={cn("px-4 py-3 text-white shadow-xs", color)}>
        <h4 className="text-[10px] font-black">{title}</h4>
      </div>
      <div className="flex-1 max-h-[400px] overflow-y-auto no-scrollbar">
        <table className="w-full text-left text-[9px] relative">
          <thead className="bg-slate-50 font-bold text-slate-400 border-b border-slate-100 sticky top-0 z-10 shadow-xs">
            <tr>
              <th className="px-4 py-2">Soc.</th>
              <th className="px-4 py-2">Antigüedad</th>
              <th className="px-4 py-2 text-right">Importe</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {data.map((row, idx) => (
              <tr key={idx} className="hover:bg-slate-50 transition-colors">
                <td className="px-4 py-2 font-black text-slate-900 truncate max-w-[80px]">
                  <span className="text-[7px] text-slate-300 mr-1">#{idx + 1}</span>
                  {row.soc}
                </td>
                <td className={cn("px-4 py-2 font-bold text-[8px]", row.isOld ? "text-rose-600" : "text-emerald-600")}>
                  {row.label || 'N/A'}
                </td>
                <td className={cn("px-4 py-2 text-right font-black", row.amt > 0 ? "text-emerald-700" : "text-rose-600")}>
                  {formatVal(row.amt)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// Global Formatter Helper
const formatCurrency = (val: number) => {
  return new Intl.NumberFormat('es-MX', {
    style: 'currency',
    currency: 'MXN',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(val);
};

const formatNumber = (val: number) => {
  return new Intl.NumberFormat('es-MX').format(val);
};

