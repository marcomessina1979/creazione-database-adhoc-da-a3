
import React, { useState, useCallback, useEffect } from 'react';
import { FileUpload } from './components/FileUpload';
import { ResultsDisplay } from './components/ResultsDisplay';
import { Spinner } from './components/Spinner';
import { Reconciliation } from './components/Reconciliation';
import { processFiles, scanForInvalidCodes } from './services/excelProcessor';
import { getDbFile, saveDbFile, clearDbFile } from './services/dbCache';
import type { ProcessResult, UnreconciledRow } from './types';

const App: React.FC = () => {
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [a3File, setA3File] = useState<File | null>(null);
  const [dbFile, setDbFile] = useState<File | null>(null);
  const [dbFileNameFromCache, setDbFileNameFromCache] = useState<string | null>(null);
  const [showDbSuccess, setShowDbSuccess] = useState<boolean>(false);
  const [commessa, setCommessa] = useState<string>('');
  const [a3Results, setA3Results] = useState<ProcessResult | null>(null);
  const [unreconciledRows, setUnreconciledRows] = useState<UnreconciledRow[] | null>(null);

  useEffect(() => {
    // On component mount, try to load the DB file from cache
    const loadCachedDb = async () => {
      try {
        const cachedFile = await getDbFile();
        if (cachedFile) {
          setDbFile(cachedFile);
          setDbFileNameFromCache(cachedFile.name);
        }
      } catch (err) {
        console.error("Failed to load database from cache:", err);
      }
    };
    loadCachedDb();
  }, []);
  
  useEffect(() => {
    if (showDbSuccess) {
      const timer = setTimeout(() => setShowDbSuccess(false), 2500);
      return () => clearTimeout(timer);
    }
  }, [showDbSuccess]);

  const handleDbFileSelect = useCallback(async (file: File | null) => {
    if (file) {
      try {
        await saveDbFile(file);
        setDbFile(file);
        setDbFileNameFromCache(file.name);
        setShowDbSuccess(true);
      } catch (err) {
        console.error("Failed to save database to cache:", err);
        setError("Impossibile salvare il database in memoria.");
        setDbFile(file); // Set file even if caching fails, for current session use
      }
    }
  }, []);

  const handleClearCachedDb = useCallback(async () => {
    try {
      await clearDbFile();
      setDbFile(null);
      setDbFileNameFromCache(null);
      setShowDbSuccess(false);
    } catch (err) {
      console.error("Failed to clear cached database:", err);
      setError("Impossibile rimuovere il database dalla memoria.");
    }
  }, []);


  const handleA3ProcessClick = useCallback(async () => {
    if (!a3File || !dbFile) {
      setError("Per favore, carica entrambi i file prima di procedere.");
      return;
    }

    setIsLoading(true);
    setError(null);
    setA3Results(null);
    setUnreconciledRows(null);

    try {
      const invalidRows = await scanForInvalidCodes(a3File);

      if (invalidRows.length > 0) {
        setUnreconciledRows(invalidRows);
        setIsLoading(false);
      } else {
        const result = await processFiles(a3File, dbFile, undefined, commessa);
        setA3Results(result);
        setIsLoading(false);
      }
    } catch (err) {
      if (err instanceof Error) {
        setError(`Errore durante l'elaborazione: ${err.message}`);
      } else {
        setError("Si è verificato un errore sconosciuto.");
      }
      console.error(err);
      setIsLoading(false);
    }
  }, [a3File, dbFile, commessa]);

  const handleReconciliationConfirm = useCallback(async (corrections: Map<number, string>) => {
    if (!a3File || !dbFile) return;

    setIsLoading(true);
    setError(null);
    setA3Results(null);
    setUnreconciledRows(null);

    try {
        const result = await processFiles(a3File, dbFile, corrections, commessa);
        setA3Results(result);
    } catch (err) {
        if (err instanceof Error) {
            setError(`Errore durante l'elaborazione: ${err.message}`);
        } else {
            setError("Si è verificato un errore sconosciuto.");
        }
        console.error(err);
    } finally {
        setIsLoading(false);
    }
  }, [a3File, dbFile, commessa]);

  const handleReconciliationCancel = useCallback(() => {
    // Process without any corrections, effectively skipping the invalid rows.
    handleReconciliationConfirm(new Map());
  }, [handleReconciliationConfirm]);
  

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col items-center p-4 sm:p-6 lg:p-8">
       {unreconciledRows && (
        <Reconciliation 
          rows={unreconciledRows}
          onConfirm={handleReconciliationConfirm}
          onCancel={handleReconciliationCancel}
        />
      )}
      <div className="w-full max-w-4xl mx-auto">
        <header className="text-center mb-10">
            <div className="inline-block bg-blue-100 text-blue-600 p-3 rounded-full mb-4 ring-8 ring-blue-50">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
            </div>
            <h1 className="text-3xl sm:text-5xl font-extrabold text-gray-900 tracking-tight">Creazione Database AdHoc da A3</h1>
            <p className="mt-4 max-w-2xl mx-auto text-lg text-gray-500">
                Genera un file database per AdHoc partendo da un Ordine Fornitore (A3).
            </p>
        </header>


        <main className="bg-white rounded-2xl shadow-xl border border-gray-200/80 p-6 sm:p-10 space-y-8 w-full">
            <div className="space-y-10 animate-fade-in">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-10">
                    {/* Step 1 */}
                    <div className="flex flex-col space-y-3">
                        <label className="text-base font-semibold text-gray-900 flex items-center">
                            <span className="flex items-center justify-center text-white bg-blue-600 rounded-full w-7 h-7 text-sm font-bold mr-3">1</span>
                            Carica il file A3
                            <span className="text-gray-500 font-normal ml-1.5">(Ordine Fornitore)</span>
                        </label>
                        <FileUpload
                          id="a3-file"
                          onFileSelect={setA3File}
                          acceptedFileType=".xlsx"
                        />
                    </div>
                    
                    {/* Step 2 */}
                    <div className="flex flex-col space-y-3">
                        <label className="text-base font-semibold text-gray-900 flex items-center">
                            <span className="flex items-center justify-center text-white bg-blue-600 rounded-full w-7 h-7 text-sm font-bold mr-3">2</span>
                            {dbFileNameFromCache ? 'File Database in memoria' : 'Carica il file Database'}
                        </label>
                        {dbFileNameFromCache ? (
                            <div className={`relative flex flex-col items-center justify-center w-full h-32 px-4 transition-all duration-300 border rounded-lg ${showDbSuccess ? 'bg-green-50 border-green-300' : 'bg-slate-50 border-slate-200'}`}>
                                {showDbSuccess ? (
                                   <div className="text-center animate-fade-in">
                                      <svg className="w-10 h-10 text-green-500 mx-auto" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
                                        <path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                                      </svg>
                                      <span className="font-medium text-green-700 mt-2 block">Database caricato!</span>
                                   </div>
                                ) : (
                                  <>
                                    <div className="bg-green-100 text-green-600 rounded-full p-2">
                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                                            <path strokeLinecap="round" strokeLinejoin="round" d="M5 8h14M5 8a2 2 0 110-4h14a2 2 0 110 4M5 8v10a2 2 0 002 2h10a2 2 0 002-2V8m-9 4h4" />
                                        </svg>
                                    </div>
                                    <span className="font-semibold text-gray-800 mt-2 text-center break-all">{dbFileNameFromCache}</span>
                                    <button 
                                      onClick={handleClearCachedDb}
                                      className="mt-3 px-3 py-1 text-xs bg-white text-gray-700 font-semibold rounded-md border border-gray-300 hover:bg-gray-100 transition-colors"
                                    >
                                      Sostituisci
                                    </button>
                                  </>
                                )}
                            </div>
                        ) : (
                            <FileUpload
                                id="db-file"
                                onFileSelect={handleDbFileSelect}
                                acceptedFileType=".xlsx"
                            />
                        )}
                    </div>
                </div>

                {/* Step 3 */}
                <div className="flex flex-col space-y-3">
                    <label htmlFor="commessa-input" className="text-base font-semibold text-gray-900 flex items-center">
                        <span className="flex items-center justify-center text-white bg-blue-600 rounded-full w-7 h-7 text-sm font-bold mr-3">3</span>
                        Inserisci Numero Commessa
                        <span className="text-gray-500 font-normal ml-1.5">(Opzionale)</span>
                    </label>
                    <input
                        id="commessa-input"
                        type="text"
                        value={commessa}
                        onChange={(e) => setCommessa(e.target.value)}
                        placeholder="Es. C23-005 o C993"
                        className="w-full px-4 py-2 text-gray-700 bg-white border border-gray-300 rounded-lg shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"
                    />
                </div>
              
                <div className="border-t border-gray-200/80 pt-8 text-center">
                    <button
                      onClick={handleA3ProcessClick}
                      disabled={!a3File || !dbFile || isLoading}
                      className="px-10 py-4 bg-blue-600 text-white font-bold text-lg rounded-lg shadow-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-all duration-300 ease-in-out focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transform hover:scale-105"
                    >
                      {isLoading ? (
                        <div className="flex items-center justify-center">
                          <Spinner />
                          <span className="ml-3">Elaborazione...</span>
                        </div>
                      ) : (
                        'Elabora File'
                      )}
                    </button>
                </div>
            </div>

          {error && (
            <div className="mt-8 bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-md" role="alert">
              <p className="font-bold">Errore</p>
              <p>{error}</p>
            </div>
          )}
          
          {a3Results && <ResultsDisplay results={a3Results} a3FileName={a3File?.name ?? 'A3_file.xlsx'} />}
        </main>

        <footer className="text-center mt-12 text-sm text-gray-500">
          <p>&copy; {new Date().getFullYear()} Excel Processor App. Tutti i diritti riservati.</p>
        </footer>
      </div>
    </div>
  );
};

export default App;
