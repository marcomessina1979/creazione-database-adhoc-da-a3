
import React, { useCallback } from 'react';
import type { ProcessResult, TextValueInfo, MissingValueInfo, DescriptionMismatchInfo, UnprocessedRowInfo, SkippedRowInfo, FlaggedRowInfo } from '../types';

interface ResultsDisplayProps {
  results: ProcessResult;
  a3FileName: string;
}

const StatCard: React.FC<{ title: string; value: number | string; icon: React.ReactNode; colorClass: string }> = ({ title, value, icon, colorClass }) => (
    <div className="bg-white p-4 rounded-xl shadow-lg border border-gray-200/80 flex items-center space-x-4">
        <div className={`rounded-full p-3 ${colorClass}`}>
            {icon}
        </div>
        <div>
            <h3 className="text-sm font-medium text-gray-500 truncate">{title}</h3>
            <p className="mt-1 text-2xl font-bold text-gray-900">{value}</p>
        </div>
    </div>
);


const Details: React.FC<{ title: string; count: number; children: React.ReactNode; defaultOpen?: boolean; }> = ({ title, count, children, defaultOpen = false }) => (
    <details className="bg-white p-4 rounded-lg shadow-md border border-gray-200/80 group" open={defaultOpen}>
        <summary className="font-semibold text-gray-800 cursor-pointer list-none flex items-center justify-between">
            <span>{title} <span className="text-gray-500 font-normal">({count})</span></span>
            <svg className="w-5 h-5 transform transition-transform duration-200 group-open:rotate-180" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
            </svg>
        </summary>
        <div className="mt-4">
            {children}
        </div>
    </details>
);


const DetailsList: React.FC<{ title: string; items: string[] | TextValueInfo[] | MissingValueInfo[] | DescriptionMismatchInfo[] | SkippedRowInfo[] | FlaggedRowInfo[]; emptyText: string }> = ({ title, items, emptyText }) => {
    if (items.length === 0) {
        return null;
    }

    return (
        <Details title={title} count={items.length}>
            <div className="max-h-60 overflow-y-auto pr-2">
                <pre className="bg-gray-100 p-3 rounded-md text-sm font-mono whitespace-pre-wrap">
                    <code>{JSON.stringify(items, null, 2)}</code>
                </pre>
            </div>
        </Details>
    );
};

const TablePreview: React.FC<{ title: string; data: UnprocessedRowInfo; emptyText: string }> = ({ title, data, emptyText }) => {
    const { headers, rows } = data;
    const MAX_ROWS = 50;

    if (rows.length === 0) {
        return null;
    }

    return (
        <Details title={title} count={rows.length}>
            <div className="mt-4 max-h-96 overflow-y-auto pr-2 relative border border-gray-200 rounded-lg">
                <table className="w-full text-sm text-left text-gray-500">
                    <thead className="text-xs text-gray-700 uppercase bg-gray-100 sticky top-0">
                        <tr>
                            {headers.map((h, index) => (
                                <th key={index} scope="col" className="px-4 py-2 font-semibold">
                                    {String(h)}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {rows.slice(0, MAX_ROWS).map((row, rowIndex) => (
                            <tr key={rowIndex} className="bg-white border-b hover:bg-gray-50">
                                {row.map((cell, cellIndex) => (
                                    <td key={cellIndex} className="px-4 py-2 whitespace-nowrap">
                                        {String(cell ?? '')}
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
                 {rows.length > MAX_ROWS && 
                    <p className="text-center text-sm text-gray-600 mt-2 p-2 bg-gray-100 rounded-b-md">
                        ... e altre {rows.length - MAX_ROWS} righe non mostrate.
                    </p>
                }
            </div>
        </Details>
    );
};


export const ResultsDisplay: React.FC<ResultsDisplayProps> = ({ results, a3FileName }) => {
  const { summary } = results;
  
  const handleDownload = useCallback(() => {
    // MIME type for Excel 97-2003 (.xls)
    const blob = new Blob([results.updatedFileBuffer], {
      type: 'application/vnd.ms-excel',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    
    const nameWithoutExt = a3FileName.substring(0, a3FileName.lastIndexOf('.')) || a3FileName;
    // Download as .xls file
    a.download = `${nameWithoutExt}_output_AdHoc.xls`;
    
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [results, a3FileName]);

  const ICONS = {
    updated: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" /></svg>,
    found: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" /></svg>,
    lumpsum: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M5 8h14M5 8a2 2 0 110-4h14a2 2 0 110 4M5 8v10a2 2 0 002 2h10a2 2 0 002-2V8m-9 4h4" /></svg>,
    included: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3m0 0v3m0-3h3m-3 0H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>,
    mismatch: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>,
    notFound: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M8.228 9c.549-1.165 2.03-2 3.772-2 2.21 0 4 1.343 4 3 0 1.4-1.278 2.575-3.006 2.907-.542.104-.994.54-.994 1.093m0 3h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>,
    duplicate: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" /></svg>,
    skipped: <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M15 12H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
  };

  return (
    <div className="mt-10 pt-8 border-t border-gray-200/80">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6">
        <h2 className="text-3xl font-bold text-gray-900 mb-2 sm:mb-0">Risultati Elaborazione</h2>
        <button
            onClick={handleDownload}
            className="px-6 py-3 bg-green-600 text-white font-bold rounded-lg shadow-md hover:bg-green-700 transition-all duration-300 ease-in-out focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500 flex items-center space-x-2"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414l-3 3a1 1 0 01-1.414 0l-3-3a1 1 0 010-1.414z" clipRule="evenodd" />
            </svg>
            <span>Scarica File .XLS (Excel 97-2003)</span>
          </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <StatCard title="Righe Aggiornate" value={summary.updated_rows} icon={ICONS.updated} colorClass="bg-blue-100 text-blue-600" />
        <StatCard title="Righe LUMPSUM" value={summary.lumpsum_rows.length} icon={ICONS.lumpsum} colorClass="bg-cyan-100 text-cyan-600" />
        <StatCard title="Discordanze Desc." value={summary.description_mismatches.length} icon={ICONS.mismatch} colorClass="bg-purple-100 text-purple-600" />
        <StatCard title="Non Trovati nel DB" value={summary.not_found_in_db.length} icon={ICONS.notFound} colorClass="bg-yellow-100 text-yellow-600" />
      </div>
      <div className="space-y-4">
        <Details title="Assunzioni Applicate" count={summary.assunzioni.length + 1} defaultOpen>
            <div className="text-sm text-gray-600 space-y-2 p-2 bg-gray-50 rounded-md">
                <p className="font-semibold text-gray-800 flex items-center">
                    Modalità Esportazione:
                    <span className={`ml-2 text-xs font-bold px-2 py-0.5 rounded-full text-green-800 bg-green-100`}>
                        EXCEL 97-2003 (BIFF8)
                    </span>
                </p>
                <ul className="list-disc list-inside pt-1">
                    {summary.assunzioni.map((item, index) => <li key={index} className="mt-1">{item}</li>)}
                </ul>
            </div>
        </Details>
        <TablePreview 
          title="Anteprima Righe del DB non Copiate" 
          data={summary.unprocessed_db_rows}
          emptyText="Tutte le righe del DB con un codice corrispondente nel file A3 sono state incluse nell'output."
        />
        <DetailsList title="Codici trovati e aggiornati" items={summary.found_and_updated} emptyText="Nessun codice è stato trovato e aggiornato." />
        <DetailsList title="Righe contrassegnate come LUMPSUM" items={summary.lumpsum_rows} emptyText="Nessuna riga è stata identificata come LUMPSUM." />
        <DetailsList title="Righe contrassegnate come INCLUDED" items={summary.included_rows} emptyText="Nessuna riga è stata identificata come INCLUDED." />
        <DetailsList title="Codici non trovati nel Database" items={summary.not_found_in_db} emptyText="Tutti i codici del file A3 sono stati trovati nel database." />
        <DetailsList title="Codici duplicati nel Database (prima occorrenza usata)" items={summary.duplicates_in_db} emptyText="Nessun codice duplicato trovato nel database." />
        <DetailsList title="Righe Barrate Ignorate" items={summary.skipped_strikethrough_rows} emptyText="Nessuna riga barrata (strikethrough) è stata rilevata." />
        <DetailsList title="Discordanze Descrizione Rilevate" items={summary.description_mismatches} emptyText="Nessuna discordanza nelle descrizioni è stata rilevata." />
        <DetailsList title="Valori testuali rilevati (sostituiti con 0)" items={summary.text_values_detected} emptyText="Nessun valore testuale problematico rilevato." />
        <DetailsList title="Valori mancanti rilevati (sostituiti con 0)" items={summary.missing_values_replaced} emptyText="Nessun valore mancante rilevato." />
      </div>
    </div>
  );
};
