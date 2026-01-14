import type { ProcessResult, TextValueInfo, MissingValueInfo, DescriptionMismatchInfo, SkippedRowInfo, UnreconciledRow, FlaggedRowInfo } from '../types';
import { translations, Language } from '../translations';

declare const XLSX: any; // Using XLSX from a CDN script

// --- UTILITY FUNCTIONS ---

const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => resolve(event.target?.result as ArrayBuffer);
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

const normalizeHeader = (header: any): string => {
    if (header === null || header === undefined) return "";
    return String(header).trim().toLowerCase().replace(/\s+/g, ' ');
}

const normalizeCodeHeader = (header: any): string => {
    if (header === null || header === undefined) return "";
    return String(header).trim().toLowerCase().replace(/[\s\[\]\(\)]+/g, '');
}

const normalizeCode = (code: any): string => {
    return String(code ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
};

const normalizeDescription = (desc: any): string => {
  if (desc === null || desc === undefined) return "";
  return String(desc)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/&/g, 'and')
    .replace(/[.,]$/, '');
};

const constructA3Code = (segments: (string | number)[]): string | null => {
    const [l1_raw, l2_raw, l3_raw, l4_raw] = segments;

    const l1 = String(l1_raw ?? '').trim().toUpperCase();
    const l2_num = String(l2_raw ?? '').trim().replace(/[^0-9]/g, '');
    const l3_num = String(l3_raw ?? '').trim().replace(/[^0-9]/g, '');
    const l4_num = String(l4_raw ?? '').trim().replace(/[^0-9]/g, '');

    if (!l1 || !l2_num) return null;

    const l2 = l2_num.padStart(3, '0');
    let code = '';
    if (l3_num && !l4_num) {
        const l3 = l3_num.padStart(2, '0');
        code = `${l1}${l2}${l3}`;
    } else if (l3_num && l4_num) {
        const l3 = l3_num.padStart(2, '0');
        const l4 = l4_num.padStart(2, '0');
        code = `${l1}${l2}${l3}${l4}`;
    } else {
        return null;
    }

    return (code.length === 6 || code.length === 8) ? code : null;
};

const parseNumericValue = (value: any): number => {
    if (value === null || value === undefined) return 0;
    let strValue = String(value).trim();
    if (strValue === '' || strValue === '9999999' || strValue === '9999999999') return 0;
    
    const lower = strValue.toLowerCase();
    const excluded = /\b(included|incl\.?|includi|incluso)\b/i;
    if (excluded.test(lower)) return 0;
    
    const lastComma = strValue.lastIndexOf(',');
    const lastDot = strValue.lastIndexOf('.');
    if (lastComma > lastDot) {
      strValue = strValue.replace(/\./g, '').replace(',', '.');
    } else {
      strValue = strValue.replace(/,/g, '');
    }
    
    const match = strValue.match(/-?\d+(\.\d+)?/);
    const num = match ? parseFloat(match[0]) : NaN;
    return isNaN(num) ? 0 : num;
};

const parseQuantity = (value: any): number | null => {
    if (value === null || value === undefined) return null;
    let strValue = String(value).trim().replace(/\u00A0/g, ' ');
    if (strValue === '') return null;
    
    const clean = strValue.replace(/[a-zA-Z()]/g, '').trim();
    if (clean === '') {
        const match = strValue.match(/-?[\d.,]+/);
        if (!match) return null;
        strValue = match[0];
    } else {
        strValue = clean;
    }

    const lastComma = strValue.lastIndexOf(',');
    const lastDot = strValue.lastIndexOf('.');
    if (lastComma > lastDot) {
      strValue = strValue.replace(/\./g, '').replace(',', '.');
    } else {
      strValue = strValue.replace(/,/g, '');
    }
    
    const num = parseFloat(strValue);
    return isNaN(num) ? null : num;
};

interface A3ColIndices {
  qty: number; unitPrice: number; discountedUnitPrice: number; description: number;
  totalPrice: number; l1: number; l2: number; l3: number; l4: number;
}

const findA3Headers = (a3Data: any[][]): { colIndices: A3ColIndices; headerRowIndex: number } => {
  let headerRowIndex = -1;
  const colIndices: A3ColIndices = { qty: -1, unitPrice: -1, discountedUnitPrice: -1, description: -1, totalPrice: -1, l1: -1, l2: -1, l3: -1, l4: -1 };

  for (let i = 0; i < Math.min(a3Data.length, 20); i++) {
    const row = a3Data[i];
    if (!row) continue;
    row.forEach((cell, index) => {
      const norm = normalizeHeader(cell);
      const code = normalizeCodeHeader(cell);
      if (colIndices.qty === -1 && norm.includes('q.ty')) { colIndices.qty = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.unitPrice === -1 && norm.includes('unit pric') && !norm.includes('discounted')) { colIndices.unitPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.discountedUnitPrice === -1 && norm.includes('discounted unit price')) { colIndices.discountedUnitPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.description === -1 && norm.includes('description')) { colIndices.description = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.totalPrice === -1 && (norm.includes('total pr') || norm.includes('total price'))) { colIndices.totalPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l1 === -1 && code === 'l1') { colIndices.l1 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l2 === -1 && code === 'l2') { colIndices.l2 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l3 === -1 && code === 'l3') { colIndices.l3 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l4 === -1 && code === 'l4') { colIndices.l4 = index; headerRowIndex = Math.max(headerRowIndex, i); }
    });
  }
  return { colIndices, headerRowIndex };
};

export const scanForInvalidCodes = async (a3File: File): Promise<UnreconciledRow[]> => {
    const buffer = await readFileAsArrayBuffer(a3File);
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const { colIndices, headerRowIndex } = findA3Headers(data);
    const unreconciled: UnreconciledRow[] = [];

    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        const segments = [colIndices.l1, colIndices.l2, colIndices.l3, colIndices.l4].map(idx => row[idx]);
        if (segments.every(s => s === null || s === undefined || String(s).trim() === '')) continue;
        const code = constructA3Code(segments);
        if (!code && row[colIndices.description]) {
            unreconciled.push({
                rowIndex: i, excelRow: i + 1, segments: segments.map(s => s ?? ''),
                description: String(row[colIndices.description] ?? '')
            });
        }
    }
    return unreconciled;
};

export const processFiles = async (
  a3File: File, 
  dbFile: File,
  corrections?: Map<number, string>,
  commessa?: string,
  numeroOrdine?: string,
  language: Language = 'it'
): Promise<ProcessResult> => {
  const [a3Buffer, dbBuffer] = await Promise.all([readFileAsArrayBuffer(a3File), readFileAsArrayBuffer(dbFile)]);
  const a3Wb = XLSX.read(a3Buffer, { type: 'array', cellStyles: true });
  const dbWb = XLSX.read(dbBuffer, { type: 'array' });
  const a3Sheet = a3Wb.Sheets[a3Wb.SheetNames[0]];
  const dbSheet = dbWb.Sheets[dbWb.SheetNames[0]];
  const a3Data: any[][] = XLSX.utils.sheet_to_json(a3Sheet, { header: 1 });
  const dbData: any[][] = XLSX.utils.sheet_to_json(dbSheet, { header: 1 });
  const { colIndices: a3Cols, headerRowIndex: a3Start } = findA3Headers(a3Data);

  // DB logic
  const dbHeaders = dbData[0].map(normalizeHeader);
  const dbCodeIdx = dbHeaders.findIndex(h => h.includes('articolo'));
  const dbDescIdx = dbHeaders.findIndex(h => h.includes('descriz'));
  const dbMap = new Map<string, any[]>();
  for(let i=1; i<dbData.length; i++) {
      const code = normalizeCode(dbData[i][dbCodeIdx]);
      if(code && !dbMap.has(code)) dbMap.set(code, dbData[i]);
  }

  // "Numero Ordine" column removed from output
  const outputAoa: any[][] = [['Articolo', 'Descrizione', 'Descrizione supp', 'QUANTITA', 'Prezzo', 'Sconto', 'Prezzo Totale', 'Commessa']];
  const foundCodes: string[] = [];
  const notFound: string[] = [];
  const mismatches: DescriptionMismatchInfo[] = [];
  const lumpsumRows: FlaggedRowInfo[] = [];
  const includedRows: FlaggedRowInfo[] = [];
  const skippedStruck: SkippedRowInfo[] = [];

  for (let i = a3Start + 1; i < a3Data.length; i++) {
    const row = a3Data[i];
    if (!row || row.length === 0) continue;

    const keyIndices = [a3Cols.l1, a3Cols.l2, a3Cols.l3, a3Cols.l4, a3Cols.description];
    let isStruck = false;
    for (const c of keyIndices) {
        if (c === -1) continue;
        const cell = a3Sheet[XLSX.utils.encode_cell({r: i, c})];
        if (cell?.s?.font?.strike) { isStruck = true; break; }
    }

    const segments = [a3Cols.l1, a3Cols.l2, a3Cols.l3, a3Cols.l4].map(idx => row[idx]);
    let code = corrections?.get(i) || constructA3Code(segments);
    
    if (isStruck) {
        skippedStruck.push({ row_excel: i + 1, article: code || 'N/A', reason: 'strikethrough' });
        continue;
    }

    if (!code) continue;
    const dbRow = dbMap.get(normalizeCode(code));
    if (!dbRow) { notFound.push(code); continue; }

    const uPrice = parseNumericValue(row[a3Cols.unitPrice]);
    const dPrice = parseNumericValue(row[a3Cols.discountedUnitPrice]);
    const qtyParsed = parseQuantity(row[a3Cols.qty]);
    const a3Desc = String(row[a3Cols.description] || '').trim();
    const dbDesc = String(dbRow[dbDescIdx] || '').trim();

    let priceFlag = 'NORMAL';
    let qty = qtyParsed || 1;
    let listPrice = uPrice;
    let discount = 0;

    if (uPrice === 0 && dPrice === 0) {
        priceFlag = 'INCLUDED';
        includedRows.push({ codice: code, cella: 'N/A' });
        continue; 
    } else {
        if (qtyParsed === null || qtyParsed === 0) {
            priceFlag = 'LUMPSUM';
            qty = 1;
            lumpsumRows.push({ codice: code, cella: 'N/A' });
        }
        if (listPrice > 0 && dPrice > 0 && dPrice < listPrice) {
            discount = -((1 - (dPrice / listPrice)) * 100);
        }
    }

    const mismatch = normalizeDescription(a3Desc) !== normalizeDescription(dbDesc);
    if(mismatch) mismatches.push({ codice: code, db_description: dbDesc, a3_description: a3Desc });

    const suppParts = [];
    if (mismatch) suppParts.push(a3Desc);
    if (priceFlag === 'LUMPSUM') suppParts.push('LUMPSUM');
    
    const excelRowIndex = outputAoa.length + 1;
    const qtyCell = `D${excelRowIndex}`;
    const priceCell = `E${excelRowIndex}`;
    const discCell = `F${excelRowIndex}`;
    
    const totalFormula = `ROUND((${qtyCell}*${priceCell})*(1+IF(ISBLANK(${discCell}),0,${discCell})/100),2)`;
    const calculatedTotal = Number((qty * listPrice * (1 + discount / 100)).toFixed(2));

    outputAoa.push([
        dbRow[dbCodeIdx], dbDesc, suppParts.join(' | ') || null,
        qty, listPrice, discount !== 0 ? discount : null,
        { f: totalFormula, v: calculatedTotal, t: 'n' }, commessa || null
    ]);
    foundCodes.push(code);
  }

  const newWs = XLSX.utils.aoa_to_sheet(outputAoa);
  
  const range = XLSX.utils.decode_range(newWs['!ref']!);
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({c: C, r: R});
      const cell = newWs[addr];
      if (!cell) continue;

      if (R === 0) {
          cell.t = 's';
          cell.z = '@';
      } else {
          const isText = [0, 1, 2, 7].includes(C);
          if (isText) {
              cell.t = 's';
              cell.z = '@';
              if (cell.v !== null && cell.v !== undefined) {
                  cell.v = String(cell.v);
              }
          } else {
              cell.t = 'n';
              cell.z = (C === 3) ? '0' : '0.00';
              if (cell.f) {
                cell.t = 'n';
                if (cell.v === undefined) cell.v = 0;
              } else if (typeof cell.v !== 'number') {
                  const rawVal = String(cell.v || '0').replace(',', '.');
                  const parsed = parseFloat(rawVal);
                  cell.v = isNaN(parsed) ? 0 : parsed;
              }
          }
      }
    }
  }

  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, newWs, "Foglio1");
  
  const buffer = XLSX.write(newWb, { 
      type: 'array', 
      bookType: 'xls',
      cellDates: true,
      cellStyles: true 
  });

  return {
    summary: {
      updated_rows: foundCodes.length,
      found_and_updated: foundCodes,
      not_found_in_db: notFound.sort(),
      duplicates_in_db: [],
      text_values_detected: [],
      missing_values_replaced: [],
      description_mismatches: mismatches,
      unprocessed_db_rows: { headers: dbHeaders, rows: [] },
      skipped_strikethrough_rows: skippedStruck,
      lumpsum_rows: lumpsumRows,
      included_rows: includedRows,
      fallback_mode: 'disabled',
      assunzioni: translations[language].assumptionsList,
      output_file: `${commessa || 'COMMESSA'}-${numeroOrdine || 'ORDINE'}-A3.xls`
    },
    updatedFileBuffer: new Uint8Array(buffer)
  };
};
