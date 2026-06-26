import * as XLSX from 'xlsx';
import { RawRecord, PeriodSnapshot, DashboardData, AgingBucket } from '../types';

export async function parseExcelDashboard(file: File): Promise<DashboardData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const periods: PeriodSnapshot[] = [];
        const sociedadesSet = new Set<string>();
        const allAvailableColumns = new Set<string>();

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];

          if (rows.length < 1) return;

          // Find the header row by searching for specific column names
          let headerRowIdx = -1;
          for (let i = 0; i < Math.min(rows.length, 20); i++) {
            const row = rows[i];
            if (row && row.some(cell => {
              const val = String(cell || '').toLowerCase();
              return val.includes('nº doc.') || val.includes('importe en md') || val.includes('soc.');
            })) {
              headerRowIdx = i;
              break;
            }
          }

          if (headerRowIdx === -1) headerRowIdx = 0;

          const rawHeaders = rows[headerRowIdx].map(h => String(h || '').trim());
          const headers = rawHeaders.map(h => h.toLowerCase());
          const dataRows = rows.slice(headerRowIdx + 1);

          // Find critical column indices using user-provided headers
          const docNumIdx = headers.findIndex(h => h.includes('nº doc.') || h.includes('nº doc'));
          const amountIdx = headers.findIndex(h => h.includes('importe en md'));
          const socIdx = headers.findIndex(h => h === 'soc.' || h === 'sociedad');
          const agingColIdx = headers.findIndex(h => h.includes('antigüedad'));
          const contralorIdx = headers.findIndex(h => h.includes('contralor corporativo'));
          const responsableIdx = headers.findIndex(h => h.includes('responsable'));

          const rawRecords: RawRecord[] = [];
          const recordsBySociedad: Record<string, RawRecord[]> = {};
          let totalVolume = 0;
          let totalAmount = 0;
          const agingMap = new Map<string, { count: number; amount: number; order: number }>();

          // Collect all column names for the UI dropdown
          rawHeaders.forEach(h => { if(h) allAvailableColumns.add(h); });

          dataRows.forEach((row) => {
            if (!row || row.length === 0) return;
            
            const docNum = docNumIdx !== -1 ? row[docNumIdx] : null;
            if (docNum === null || docNum === undefined || String(docNum).trim() === '') return;

            const rawAmount = amountIdx !== -1 ? row[amountIdx] : 0;
            const amount = typeof rawAmount === 'number' ? rawAmount : parseFloat(String(rawAmount).replace(/[^0-9.-]+/g, "")) || 0;
            
            const soc = socIdx !== -1 ? String(row[socIdx] || 'N/A') : 'N/A';
            
            // Dynamic Aging Detection
            let agingLabel = 'Sin Clasificar';
            if (agingColIdx !== -1) {
              const rawAging = row[agingColIdx];
              if (rawAging !== null && rawAging !== undefined) {
                agingLabel = String(rawAging).trim();
                if (!isNaN(Number(agingLabel)) && agingLabel !== '') {
                  const days = Number(agingLabel);
                  if (days < 30) agingLabel = '0-30 días';
                  else if (days < 60) agingLabel = '31-60 días';
                  else if (days < 90) agingLabel = '61-90 días';
                  else if (days < 120) agingLabel = '91-120 días';
                  else if (days < 150) agingLabel = '121-150 días';
                  else if (days < 180) agingLabel = '151-180 días';
                  else agingLabel = 'MAS DE 180 días';
                }
              }
            }

            const contralor = contralorIdx !== -1 ? String(row[contralorIdx] || 'N/A') : 'N/A';

            sociedadesSet.add(soc);
            totalVolume++;
            totalAmount += amount;

            const currentBucket = agingMap.get(agingLabel) || { count: 0, amount: 0, order: 0 };
            currentBucket.count++;
            currentBucket.amount += amount;
            agingMap.set(agingLabel, currentBucket);

            // Create a record with ALL columns mapped by their header name
            const record: RawRecord = {
              'Número de documento': String(docNum),
              'Importe en MD': amount,
              'Sociedad': soc,
              'Antigüedad Label': agingLabel,
              'Contralor Corporativo': contralor
            };

            for (let i = 0; i < rawHeaders.length; i++) {
              if (rawHeaders[i]) {
                record[rawHeaders[i]] = row[i];
              }
            }

            rawRecords.push(record);
            if (!recordsBySociedad[soc]) recordsBySociedad[soc] = [];
            recordsBySociedad[soc].push(record);
          });

          if (totalVolume === 0) return;

          const agingBuckets: AgingBucket[] = Array.from(agingMap.entries())
            .map(([label, data]) => ({ label, ...data }))
            .sort((a, b) => a.label.localeCompare(b.label)); // Natural order is better if user provides text

          periods.push({
            periodLabel: sheetName,
            totalVolume,
            totalAmount,
            records: rawRecords,
            recordsBySociedad,
            agingBuckets,
            sociedades: Object.keys(recordsBySociedad)
          });
        });

        resolve({
          periods,
          allSociedades: Array.from(sociedadesSet).sort(),
          availableColumns: Array.from(allAvailableColumns).sort()
        } as any); // Type cast to include availableColumns temporarily
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
