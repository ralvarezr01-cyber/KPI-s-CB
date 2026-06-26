export interface RawRecord {
  'Número de documento': string | number;
  'Importe en MD': number;
  'Sociedad': string;
  'Fecha de documento'?: any;
  'Referencia'?: string;
  'Texto'?: string;
  [key: string]: any;
}

export interface AgingBucket {
  label: string;
  count: number;
  amount: number;
  order: number;
}

export interface PeriodSnapshot {
  periodLabel: string; // Sheet name (e.g., "AÑO ANT", "ENE 2026")
  totalVolume: number;
  totalAmount: number;
  records: RawRecord[];
  agingBuckets: AgingBucket[];
  sociedades: string[];
  recordsBySociedad?: Record<string, RawRecord[]>;
}

export interface DashboardData {
  periods: PeriodSnapshot[];
  allSociedades: string[];
  availableColumns: string[];
}
