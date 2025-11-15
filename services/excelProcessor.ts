import type { ProcessResult, TextValueInfo, MissingValueInfo, DescriptionMismatchInfo, SkippedRowInfo, UnreconciledRow, FlaggedRowInfo } from '../types';

declare const XLSX: any; // Using XLSX from a CDN script

// --- UTILITY FUNCTIONS ---

const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      resolve(event.target?.result as ArrayBuffer);
    };
    reader.onerror = (error) => {
      reject(error);
    };
    reader.readAsArrayBuffer(file);
  });
};

const normalizeHeader = (header: any): string => {
    if (header === null || header === undefined) return "";
    return String(header).trim().toLowerCase().replace(/\s+/g, ' ');
}

const normalizeCodeHeader = (header: any): string => {
    if (header === null || header === undefined) return "";
    // Rimuove spazi e parentesi quadre/tonde per un match affidabile (es. "[L1]" -> "l1")
    return String(header).trim().toLowerCase().replace(/[\s\[\]\(\)]+/g, '');
}

const normalizeCode = (code: any): string => {
    // Aggressively remove all non-alphanumeric characters and convert to uppercase
    // for robust matching. Handles various separators ('-', '.', '/', ' ', etc.) and case differences.
    return String(code ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
};

const normalizeDescription = (desc: any): string => {
  if (desc === null || desc === undefined) return "";
  return String(desc)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ') // Collapse multiple spaces
    .replace(/&/g, 'and') // Replace & with 'and'
    .replace(/[.,]$/, ''); // Remove trailing punctuation (dot or comma)
};

const constructA3Code = (segments: (string | number)[]): string | null => {
    const [l1_raw, l2_raw, l3_raw, l4_raw] = segments;

    const l1 = String(l1_raw ?? '').trim().toUpperCase();
    const l2_num = String(l2_raw ?? '').trim().replace(/[^0-9]/g, '');
    const l3_num = String(l3_raw ?? '').trim().replace(/[^0-9]/g, '');
    const l4_num = String(l4_raw ?? '').trim().replace(/[^0-9]/g, '');

    if (!l1 || !l2_num) {
        return null; // Invalid code parts
    }

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
        return null; // Does not fit pattern
    }

    // Return the code only if it has the correct length (6 or 8)
    if (code.length === 6 || code.length === 8) {
        return code;
    }
    
    return null;
};

const parseQuantity = (value: any): number | null => {
    if (value === null || value === undefined) return null;

    // 1. Initial cleanup: trim spaces, incl. non-breaking space (U+00A0)
    let strValue = String(value).trim().replace(/\u00A0/g, ' ');
    if (strValue === '') return null;

    const originalStrForFallback = strValue;

    // 2. Remove text wrappers and any letters. Keep digits, ,, . and -.
    strValue = strValue.replace(/[a-zA-Z()]/g, '').trim();
    
    // If cleaning left nothing, it's not a parsable number.
    if (strValue === '') {
        // Fallback: try to find a number in the original string
        const fallbackMatch = originalStrForFallback.match(/-?[\d.,]+/);
        if (fallbackMatch) {
            strValue = fallbackMatch[0];
        } else {
            return null;
        }
    }

    // 3 & 4. Normalize based on last separator to handle EU/US formats
    const lastComma = strValue.lastIndexOf(',');
    const lastDot = strValue.lastIndexOf('.');

    let standardizedValue = strValue;
    if (lastComma > lastDot) {
      // European format (e.g., "1.234,50"). Remove dots, replace comma with dot.
      standardizedValue = standardizedValue.replace(/\./g, '').replace(',', '.');
    } else {
      // US/default format (e.g., "1,234.50"). Just remove commas.
      standardizedValue = standardizedValue.replace(/,/g, '');
    }
    
    // 5. Attempt to parse the standardized value
    const num = parseFloat(standardizedValue);

    if (!isNaN(num)) {
        return num;
    }

    // 6. Final fallback if primary method failed.
    const fallbackMatch = originalStrForFallback.match(/-?\d+(\.\d+)?/);
    if (fallbackMatch) {
        const fallbackNum = parseFloat(fallbackMatch[0]);
        if (!isNaN(fallbackNum)) return fallbackNum;
    }
    
    return null;
};

const parseNumericValue = (value: any): number => {
    if (value === null || value === undefined) return 0;
  
    const strValue = String(value).trim();
    if (strValue === '') return 0;
  
    // Handle placeholders
    if (strValue === '9999999' || strValue === '9999999999' || (typeof value === 'number' && (value === 9999999 || value === 9999999999))) {
        return 0;
    }
    
    // Handle text that means zero price but is not a number
    const lowerStrValue = strValue.toLowerCase();
    const includedRegex = /\b(included|incl\.?|includi|incluso)\b/i;
    // Also handle some lumpsum keywords if they appear alone, as they imply no numeric value
    const lumpsumKeywords = ['lumpsum', 'lump sum', 'l.s.', 'ls', 'lump-sum', 'forfait'];
    
    if (includedRegex.test(lowerStrValue) || lumpsumKeywords.some(kw => lowerStrValue === kw)) {
        return 0;
    }
    
    let standardizedValue = strValue;
    const lastComma = standardizedValue.lastIndexOf(',');
    const lastDot = standardizedValue.lastIndexOf('.');
  
    // Normalize number string based on detected decimal separator
    if (lastComma > lastDot) {
      // European format (e.g., "1.234,56"). Remove dots, replace comma.
      standardizedValue = standardizedValue.replace(/\./g, '').replace(',', '.');
    } else {
      // US/default format (e.g., "1,234.56"). Just remove commas.
      standardizedValue = standardizedValue.replace(/,/g, '');
    }
    
    // Extract the first valid floating-point number from the string.
    // This handles cases where numbers are mixed with text or currency symbols.
    const match = standardizedValue.match(/-?\d+(\.\d+)?/);
    const num = match ? parseFloat(match[0]) : NaN;
    
    if (isNaN(num)) {
      return 0;
    }
  
    return num;
};


// --- TYPE DEFINITIONS FOR INDICES ---
interface A3ColIndices {
  qty: number;
  unitPrice: number;
  discountedUnitPrice: number;
  description: number;
  totalPrice: number;
  priceFlag: number; // -1 if not found
  length: number; // -1 if not found
  l1: number;
  l2: number;
  l3: number;
  l4: number;
}

interface DBColIndices {
  code: number;
  description: number;
}
const findA3Headers = (a3Data: any[][]): { colIndices: A3ColIndices; headerRowIndex: number } => {
  let headerRowIndex = -1;
  const colIndices: A3ColIndices = { qty: -1, unitPrice: -1, discountedUnitPrice: -1, description: -1, totalPrice: -1, priceFlag: -1, length: -1, l1: -1, l2: -1, l3: -1, l4: -1 };

  for (let i = 0; i < Math.min(a3Data.length, 20); i++) {
    const row = a3Data[i];
    if (!row || !Array.isArray(row)) continue;

    row.forEach((cell, index) => {
      const normalizedStd = normalizeHeader(cell);
      const normalizedCode = normalizeCodeHeader(cell);

      if (colIndices.qty === -1 && normalizedStd.includes('q.ty')) { colIndices.qty = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.unitPrice === -1 && normalizedStd.includes('unit pric') && !normalizedStd.includes('discounted')) { colIndices.unitPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.discountedUnitPrice === -1 && normalizedStd.includes('discounted unit price')) { colIndices.discountedUnitPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.description === -1 && normalizedStd.includes('description')) { colIndices.description = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.totalPrice === -1 && (normalizedStd.includes('total pr') || normalizedStd.includes('total price'))) { colIndices.totalPrice = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.priceFlag === -1 && normalizedStd.includes('priceflag')) { colIndices.priceFlag = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.length === -1 && normalizedStd.includes('length')) { colIndices.length = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l1 === -1 && normalizedCode === 'l1') { colIndices.l1 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l2 === -1 && normalizedCode === 'l2') { colIndices.l2 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l3 === -1 && normalizedCode === 'l3') { colIndices.l3 = index; headerRowIndex = Math.max(headerRowIndex, i); }
      if (colIndices.l4 === -1 && normalizedCode === 'l4') { colIndices.l4 = index; headerRowIndex = Math.max(headerRowIndex, i); }
    });
  }
  
  const mandatoryHeaders: (keyof Omit<A3ColIndices, 'priceFlag' | 'length'>)[] = ['qty', 'unitPrice', 'discountedUnitPrice', 'description', 'totalPrice', 'l1', 'l2', 'l3', 'l4'];
  const missingHeaders = mandatoryHeaders.filter(key => colIndices[key] === -1);

  if (missingHeaders.length > 0) {
    throw new Error(`Impossibile trovare le seguenti intestazioni obbligatorie nel file A3: ${missingHeaders.join(', ')}. Assicurati che le colonne necessarie siano presenti nelle prime 20 righe.`);
  }

  return { colIndices, headerRowIndex };
};
const extractA3Updates = (
  a3Data: any[][], 
  a3Sheet: any, 
  colIndices: A3ColIndices, 
  headerRowIndex: number,
  corrections?: Map<number, string>
) => {
  type PriceFlag = 'LUMPSUM' | 'INCLUDED' | 'NORMAL';
  const updates: { 
    code: string | null; 
    qty: number | null; 
    listPrice: number;
    discount: number;
    description: string; 
    priceFlag: PriceFlag;
    originalTotalPrice: number;
    flagSourceCell: string | null;
  }[] = [];

  const skippedStrikethroughRows: SkippedRowInfo[] = [];
  const merges = a3Sheet['!merges'] || [];
  
  for (let i = headerRowIndex + 1; i < a3Data.length; i++) {
    const row = a3Data[i];
    if (!row || row.length === 0) continue;
    
    // --- STRIKETHROUGH CHECK (KEY-CELL FOCUSED, MERGE & RICH-TEXT AWARE) ---
    const keyCellIndices = [
        colIndices.l1, colIndices.l2, colIndices.l3, colIndices.l4,
        colIndices.description, colIndices.qty, colIndices.unitPrice, colIndices.discountedUnitPrice
    ].filter(idx => idx > -1);

    let isStrikethrough = false;
    for (const c of keyCellIndices) {
        const r = i; // Current row index (0-based)
        let effectiveCell = null;
        
        // Find if the current cell (r, c) is part of a merged range.
        const mergeRange = merges.find((range: {s: {r: number, c: number}, e: {r: number, c: number}}) => 
            r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c
        );

        if (mergeRange) {
            // If merged, the style is determined by the top-left cell of the range.
            const topLeftAddress = XLSX.utils.encode_cell({ r: mergeRange.s.r, c: mergeRange.s.c });
            effectiveCell = a3Sheet[topLeftAddress];
        } else {
            // Otherwise, check the cell itself.
            const cellAddress = XLSX.utils.encode_cell({ r, c });
            effectiveCell = a3Sheet[cellAddress];
        }
        
        let cellIsStruck = false;
        // Check 1: Full cell strikethrough
        if (effectiveCell?.s?.font?.strike) {
            cellIsStruck = true;
        } 
        // Check 2: Rich text strikethrough (if full cell isn't struck)
        else if (effectiveCell?.r && Array.isArray(effectiveCell.r)) {
            for (const run of effectiveCell.r) {
                if (run.s?.font?.strike) {
                    cellIsStruck = true;
                    break; // A single struck run is sufficient
                }
            }
        }

        if (cellIsStruck) {
            isStrikethrough = true;
            break; // One struck key cell is enough to skip the whole row.
        }
    }

    let code: string | null = null;
    const codeColIndices = [colIndices.l1, colIndices.l2, colIndices.l3, colIndices.l4];
    const codeSegments = codeColIndices.map(idx => row[idx]);

    if (corrections && corrections.has(i)) {
      code = corrections.get(i)!;
    } else {
      code = constructA3Code(codeSegments);
    }
    
    if (isStrikethrough) {
        const articleCode = code || 'N/A';
        skippedStrikethroughRows.push({
            row_excel: i + 1, // Use 1-based index for user readability
            article: articleCode,
            reason: "strikethrough"
        });
        continue;
    }

    // --- ROW PARSING ---
    const discountedUnitPrice = parseNumericValue(row[colIndices.discountedUnitPrice]);
    const unitPrice = parseNumericValue(row[colIndices.unitPrice]);
    const originalTotalPrice = parseNumericValue(row[colIndices.totalPrice]);
    const descriptionText = String(row[colIndices.description] ?? '');
    const parsedQty = parseQuantity(row[colIndices.qty]);
    
    if (code === null && originalTotalPrice === 0 && descriptionText.trim() === '') {
        continue;
    }

    // --- VALUE & FLAG DETERMINATION (NEW LOGIC) ---
    let priceFlag: PriceFlag = 'NORMAL';
    let finalQty: number | null = parsedQty;
    let listPrice = 0;
    let discount = 0;
    const priceToConsider = discountedUnitPrice > 0 ? discountedUnitPrice : unitPrice;

    if (priceToConsider === 0) {
        // Scenario B: Included (price is not present)
        priceFlag = 'INCLUDED';
        finalQty = 0;
        listPrice = 0;
        discount = 0;
    } else {
        // Scenarios A (Lumpsum) and C (Normal) share price/discount logic
        if (parsedQty === null || parsedQty === 0) {
            // Scenario A: Lumpsum
            priceFlag = 'LUMPSUM';
            finalQty = 1;
        } else {
            // Scenario C: Normal
            priceFlag = 'NORMAL';
            // finalQty is already parsedQty
        }

        listPrice = unitPrice;
        if (listPrice > 0 && discountedUnitPrice > 0 && discountedUnitPrice < listPrice) {
            discount = -((1 - (discountedUnitPrice / listPrice)) * 100);
        }
    }

    updates.push({ 
        code, 
        qty: finalQty,
        listPrice,
        discount,
        description: descriptionText,
        priceFlag,
        originalTotalPrice: originalTotalPrice,
        flagSourceCell: priceFlag !== 'NORMAL' ? 'Logica Prezzo/Q.tà' : null
    });
  }

  return { updates, textValuesDetected: [], missingValuesReplaced: [], skippedStrikethroughRows };
};
const findDbHeaders = (dbDataAoA: any[][]): DBColIndices => {
  if (dbDataAoA.length === 0) throw new Error("Il file Database è vuoto.");
  
  const dbHeadersRaw = dbDataAoA[0].map(cell => normalizeHeader(cell));
  const dbColIndices: DBColIndices = {
      code: dbHeadersRaw.findIndex(h => h.includes('articolo')),
      description: dbHeadersRaw.findIndex(h => h.includes('descrizione') || h.includes('description')),
  };

  if (dbColIndices.code === -1) throw new Error("Colonna contenente 'Articolo' non trovata nel file Database.");
  if (dbColIndices.description === -1) throw new Error("Colonna contenente 'Descrizione' o 'Description' non trovata nel file Database.");

  return dbColIndices;
};
const reconcileAndGenerateOutput = (
    dbDataAoA: any[][],
    dbColIndices: DBColIndices,
    a3Updates: { code: string | null; qty: number | null; listPrice: number; discount: number; description: string; priceFlag: string, originalTotalPrice: number, flagSourceCell: string | null }[],
    commessa?: string
) => {
    const finalOutputHeaders = [
        'Articolo', 
        'Descrizione', 
        'Descrizione secondaria', 
        'Quantità', 
        'Prezzo', 
        'Sconto', 
        'Prezzo Totale', 
        'Commessa'
    ];

    const outputDataAoA: any[][] = [finalOutputHeaders];
    
    const dbCodeMap6 = new Map<string, { rowData: any[] }>();
    const dbCodeMap8 = new Map<string, { rowData: any[] }>();
    const duplicatesInDb = new Set<string>();
    
    for (let i = 1; i < dbDataAoA.length; i++) {
        const row = dbDataAoA[i];
        if (!row || !row[dbColIndices.code]) continue;
        const code = normalizeCode(row[dbColIndices.code]);
        if (!code) continue;

        let targetMap: Map<string, { rowData: any[] }> | undefined;
        if (code.length === 6) targetMap = dbCodeMap6;
        else if (code.length === 8) targetMap = dbCodeMap8;
        
        if (targetMap) {
            if (targetMap.has(code)) {
                duplicatesInDb.add(code);
            } else {
                targetMap.set(code, { rowData: row });
            }
        }
    }

    const foundAndUpdated: string[] = [];
    const descriptionMismatches: DescriptionMismatchInfo[] = [];
    const notFoundInDb: string[] = [];
    const lumpsumRows: FlaggedRowInfo[] = [];
    const includedRows: FlaggedRowInfo[] = [];
    
    const a3CodeSet = new Set(a3Updates.map(u => u.code).filter(Boolean));
    const unprocessedDbRows: any[][] = [];

    // Identify unprocessed DB rows
    for (let i = 1; i < dbDataAoA.length; i++) {
        const row = dbDataAoA[i];
        const code = row[dbColIndices.code] ? normalizeCode(row[dbColIndices.code]) : '';
        if (code && !a3CodeSet.has(code)) {
            unprocessedDbRows.push(row);
        }
    }

    let excelRowIndex = 2; // Data rows start from the second row in Excel
    for (const a3Update of a3Updates) {
        const { code: a3Code } = a3Update;
        
        if (!a3Code) {
            notFoundInDb.push('[Codice Invalido/Mancante in A3]');
            continue;
        }

        let dbMatch: { rowData: any[] } | undefined;
        if (a3Code.length === 6) dbMatch = dbCodeMap6.get(a3Code);
        else if (a3Code.length === 8) dbMatch = dbCodeMap8.get(a3Code);

        if (!dbMatch) {
            notFoundInDb.push(a3Code);
            continue;
        }

        if (a3Update.priceFlag === 'LUMPSUM') {
          lumpsumRows.push({ codice: a3Code, cella: a3Update.flagSourceCell || 'N/A' });
        } else if (a3Update.priceFlag === 'INCLUDED') {
          includedRows.push({ codice: a3Code, cella: a3Update.flagSourceCell || 'N/A' });
          continue; // Do not add "INCLUDED" rows to the output file
        }
        
        const { rowData: originalDbRow } = dbMatch;

        const dbDescription = String(originalDbRow[dbColIndices.description] ?? '').trim();
        const a3Description = String(a3Update.description ?? '').trim();
        const isMismatch = normalizeDescription(dbDescription) !== normalizeDescription(a3Description);
        
        // --- FORMULA GENERATION ---
        const qtyColLetter = 'D';
        const priceColLetter = 'E';
        const discountColLetter = 'F'; // Was 'G', shifted due to PriceFlag removal

        const qtyCell = `${qtyColLetter}${excelRowIndex}`;
        const priceCell = `${priceColLetter}${excelRowIndex}`;
        const discountCell = `${discountColLetter}${excelRowIndex}`;
        const rowTotalFormula = `ROUND(${qtyCell}*${priceCell}*(1+IF(ISBLANK(${discountCell}),0,${discountCell})/100),2)`;

        const descriptionPart = isMismatch ? a3Description : null;
        const flagPart = (a3Update.priceFlag === 'LUMPSUM' || a3Update.priceFlag === 'INCLUDED') ? a3Update.priceFlag : null;
        const secondaryDescription = [descriptionPart, flagPart].filter(Boolean).join(' | ');

        const newRow: any[] = [
            originalDbRow[dbColIndices.code], // Articolo
            dbDescription,                     // Descrizione
            secondaryDescription || null,      // Descrizione secondaria
            a3Update.qty,                      // Quantità
            a3Update.listPrice,                // Prezzo
            a3Update.discount !== 0 ? a3Update.discount : null, // Sconto
            { f: rowTotalFormula },            // Prezzo Totale
            commessa ?? null                   // Commessa
        ];
        
        outputDataAoA.push(newRow);
        foundAndUpdated.push(a3Code);
        excelRowIndex++;
        
        if (isMismatch) {
            descriptionMismatches.push({
                codice: a3Code,
                db_description: dbDescription,
                a3_description: a3Description,
            });
        }
    }
    
    return {
        outputDataAoA,
        updatedRowsCount: foundAndUpdated.length,
        notFoundInDb,
        duplicatesInDb: Array.from(duplicatesInDb),
        foundAndUpdated,
        descriptionMismatches,
        unprocessedDbRows,
        lumpsumRows,
        includedRows,
    };
};

export const scanForInvalidCodes = async (a3File: File): Promise<UnreconciledRow[]> => {
    const a3Buffer = await readFileAsArrayBuffer(a3File);
    const a3Workbook = XLSX.read(a3Buffer, { type: 'array', cellStyles: true });
    
    if (!a3Workbook.SheetNames?.length) {
        throw new Error("Il file A3 non è valido o non contiene fogli.");
    }
    const a3SheetName = a3Workbook.SheetNames.includes("A3") ? "A3" : a3Workbook.SheetNames[0];
    const a3Sheet = a3Workbook.Sheets[a3SheetName];
    if (!a3Sheet) {
        throw new Error("Impossibile trovare il foglio di lavoro richiesto nel file A3.");
    }

    const a3Data: any[][] = XLSX.utils.sheet_to_json(a3Sheet, { header: 1 });
    const { colIndices, headerRowIndex } = findA3Headers(a3Data);
    
    const unreconciledRows: UnreconciledRow[] = [];

    for (let i = headerRowIndex + 1; i < a3Data.length; i++) {
        const row = a3Data[i];
        if (!row || row.length === 0) continue;

        const codeColIndices = [colIndices.l1, colIndices.l2, colIndices.l3, colIndices.l4];
        const codeSegments = codeColIndices.map(idx => row[idx]);

        if (codeSegments.every(s => s === null || s === undefined || String(s).trim() === '')) continue;
        
        const code = constructA3Code(codeSegments);
        
        if (!code && row[colIndices.description]) {
            unreconciledRows.push({
                rowIndex: i,
                excelRow: i + 1,
                segments: codeSegments.map(s => s ?? ''),
                description: String(row[colIndices.description] ?? '')
            });
        }
    }
    
    return unreconciledRows;
};

export const processFiles = async (
  a3File: File, 
  dbFile: File,
  corrections?: Map<number, string>,
  commessa?: string
): Promise<ProcessResult> => {
  // --- 1. Read files and load workbooks ---
  const [a3Buffer, dbBuffer] = await Promise.all([
    readFileAsArrayBuffer(a3File),
    readFileAsArrayBuffer(dbFile),
  ]);

  const a3Workbook = XLSX.read(a3Buffer, { type: 'array', cellStyles: true });
  const dbWorkbook = XLSX.read(dbBuffer, { type: 'array' });

  // --- 2. Get sheets and convert to Array of Arrays ---
  if (!a3Workbook.SheetNames?.length) {
    throw new Error("Il file A3 (Ordine Fornitore) non è valido o non contiene fogli di lavoro.");
  }
  if (!dbWorkbook.SheetNames?.length) {
    throw new Error("Il file Database non è valido o non contiene fogli di lavoro.");
  }
  
  const a3SheetName = a3Workbook.SheetNames.includes("A3") ? "A3" : a3Workbook.SheetNames[0];
  const dbSheetName = dbWorkbook.SheetNames.includes("Elenco articoli") ? "Elenco articoli" : dbWorkbook.SheetNames[0];
  
  const a3Sheet = a3Workbook.Sheets[a3SheetName];
  const dbSheet = dbWorkbook.Sheets[dbSheetName];
  if (!a3Sheet || !dbSheet) {
    throw new Error("Impossibile trovare i fogli di lavoro richiesti nei file Excel.");
  }
  
  const a3Data: any[][] = XLSX.utils.sheet_to_json(a3Sheet, { header: 1 });
  const dbDataAoA: any[][] = XLSX.utils.sheet_to_json(dbSheet, { header: 1 });

  // --- 3. Process A3 data to get updates ---
  const { colIndices: a3ColIndices, headerRowIndex } = findA3Headers(a3Data);
  const { updates: a3Updates, textValuesDetected, missingValuesReplaced, skippedStrikethroughRows } = extractA3Updates(a3Data, a3Sheet, a3ColIndices, headerRowIndex, corrections);

  // --- 4. Reconcile data and generate new output structure ---
  const dbColIndices = findDbHeaders(dbDataAoA);
  const { outputDataAoA, unprocessedDbRows, ...updateStats } = reconcileAndGenerateOutput(dbDataAoA, dbColIndices, a3Updates, commessa);
  
  // --- 5. Generate Output File ---
  const newSheet = XLSX.utils.aoa_to_sheet(outputDataAoA);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, dbSheetName);
  const updatedFileBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
  
  // --- 6. Build Summary ---
  const fallbackMode: 'enabled' | 'disabled' = 'disabled';
  
  const assunzioni = [
      "Fedeltà ai Dati: Il 'Prezzo' (di listino), la 'Quantità' e lo 'Sconto' dell'output sono presi o calcolati direttamente dal file A3.",
      "Identificazione Flag (LUMPSUM/INCLUDED): La logica non si basa più sulla ricerca di parole chiave, ma sulla relazione tra prezzo e quantità:",
      "  - Una riga è 'LUMPSUM' se ha un prezzo unitario ma nessuna quantità. In questo caso, la quantità in output è 1, e il prezzo di listino e lo sconto sono calcolati come per una riga normale.",
      "  - Una riga è 'INCLUDED' se non ha un prezzo unitario. Queste righe vengono conteggiate nel riepilogo ma sono escluse dal file di output finale.",
      "Riconoscimento Strikethrough Avanzato: Una riga viene ignorata se una delle sue celle chiave (Codice, Descrizione, Q.tà, Prezzo) è barrata. La detezione supporta: barrato su cella intera, barrato parziale (rich text), e celle unite (merged cells). La formattazione condizionale che applica il barrato potrebbe non essere rilevata.",
      "Calcolo Sconto: La colonna 'Sconto' è calcolata dalla differenza tra 'Unit Price' e 'Discounted Unit Price' del file A3 e espressa come percentuale negativa (es. -10 per 10%).",
      "Struttura Colonne Output: Il file di output generato ha una struttura fissa: Articolo, Descrizione, Descrizione secondaria, Quantità, Prezzo, Sconto, Prezzo Totale, Commessa.",
      "Calcolo Prezzo Totale con Formula: La colonna 'Prezzo Totale' in ogni riga è calcolata tramite una formula Excel (=ROUND(D_ * E_ * (1 + F_/100), 2)), rendendola dinamica a eventuali modifiche.",
      "Descrizione Secondaria: La colonna 'Descrizione secondaria' viene usata per indicare se una riga è stata trattata come 'LUMPSUM', e per riportare la descrizione A3 se differisce da quella del database.",
      "I valori segnaposto (es. 9999999) e le celle vuote o testuali in campi numerici sono stati convertiti a 0.",
      "Output Filtrato: L'output contiene solo le righe con codici presenti in entrambi i file, escludendo quelle identificate come 'INCLUDED'.",
    ];

  const summary = {
    updated_rows: updateStats.updatedRowsCount,
    found_and_updated: updateStats.foundAndUpdated,
    not_found_in_db: updateStats.notFoundInDb.sort(),
    duplicates_in_db: updateStats.duplicatesInDb.sort(),
    text_values_detected: textValuesDetected,
    missing_values_replaced: missingValuesReplaced,
    description_mismatches: updateStats.descriptionMismatches,
    unprocessed_db_rows: {
        headers: dbDataAoA[0],
        rows: unprocessedDbRows,
    },
    skipped_strikethrough_rows: skippedStrikethroughRows,
    lumpsum_rows: (updateStats.lumpsumRows || []).sort((a,b) => a.codice.localeCompare(b.codice)),
    included_rows: (updateStats.includedRows || []).sort((a,b) => a.codice.localeCompare(b.codice)),
    fallback_mode: fallbackMode,
    assunzioni,
    output_file: `(scaricabile dall'interfaccia)`,
  };

  return {
    summary,
    updatedFileBuffer: new Uint8Array(updatedFileBuffer),
  };
};