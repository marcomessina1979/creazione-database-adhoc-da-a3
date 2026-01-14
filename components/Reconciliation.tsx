import React, { useState } from 'react';
import type { UnreconciledRow } from '../types';
import { translations, Language } from '../translations';

interface ReconciliationProps {
  rows: UnreconciledRow[];
  onConfirm: (corrections: Map<number, string>) => void;
  onCancel: () => void;
  language: Language;
}

export const Reconciliation: React.FC<ReconciliationProps> = ({ rows, onConfirm, onCancel, language }) => {
    const [corrections, setCorrections] = useState<Map<number, string>>(new Map());
    const t = translations[language];

    const handleCodeChange = (rowIndex: number, code: string) => {
        const newCorrections = new Map(corrections);
        const normalizedCode = code.toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (normalizedCode) {
            newCorrections.set(rowIndex, normalizedCode);
        } else {
            newCorrections.delete(rowIndex);
        }
        setCorrections(newCorrections);
    };

    const handleSubmit = () => {
        onConfirm(corrections);
    };

    return (
        <div className="fixed inset-0 bg-gray-900 bg-opacity-75 flex items-center justify-center p-4 z-50 animate-fade-in">
            <div className="bg-white rounded-xl shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col">
                <header className="p-6 border-b border-gray-200">
                    <h2 className="text-2xl font-bold text-gray-800">{t.reconciliationTitle}</h2>
                    <p className="mt-1 text-gray-600">
                        {t.reconciliationSub(rows.length)}
                    </p>
                </header>

                <main className="flex-grow p-6 overflow-y-auto">
                    <div className="overflow-x-auto border border-gray-200 rounded-lg">
                        <table className="w-full text-sm text-left text-gray-700">
                            <thead className="text-xs text-gray-800 uppercase bg-gray-100 sticky top-0">
                                <tr>
                                    <th scope="col" className="px-4 py-3 font-semibold">{t.excelRow}</th>
                                    <th scope="col" className="px-4 py-3 font-semibold">{t.originalSegments}</th>
                                    <th scope="col" className="px-4 py-3 font-semibold">{t.a3Desc}</th>
                                    <th scope="col" className="px-4 py-3 font-semibold w-1/4">{t.correctCode}</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((row) => (
                                    <tr key={row.rowIndex} className="bg-white border-b last:border-b-0 hover:bg-gray-50">
                                        <td className="px-4 py-2 font-medium">{row.excelRow}</td>
                                        <td className="px-4 py-2 font-mono text-gray-600">{row.segments.join(' - ')}</td>
                                        <td className="px-4 py-2 max-w-sm truncate" title={row.description}>{row.description}</td>
                                        <td className="px-4 py-2">
                                            <input
                                                type="text"
                                                className="w-full px-2 py-1.5 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500"
                                                placeholder="Es. ABC123DE"
                                                maxLength={12} 
                                                onChange={(e) => handleCodeChange(row.rowIndex, e.target.value)}
                                            />
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </main>

                <footer className="p-6 border-t border-gray-200 bg-gray-50 flex justify-end space-x-4 rounded-b-xl">
                    <button
                        onClick={onCancel}
                        className="px-6 py-2 bg-white text-gray-800 font-semibold rounded-lg border border-gray-300 hover:bg-gray-100 transition-colors"
                    >
                        {t.cancelIgnore}
                    </button>
                    <button
                        onClick={handleSubmit}
                        className="px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 transition-colors"
                    >
                        {t.confirmProceed}
                    </button>
                </footer>
            </div>
        </div>
    );
};