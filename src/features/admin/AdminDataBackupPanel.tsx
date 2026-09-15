import React, { useRef, useState } from 'react';
import { Download, Upload, Database, X, ShieldCheck } from 'lucide-react';
import { encodeBinaryBackup, exportBudgetControlBackup, decodeBackupFile, restoreBudgetControlBackup } from '../../adminDataBackup';

export default function AdminDataBackupPanel() {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  const runExport = async (binary: boolean) => {
    setBusy(true); setMessage('');
    try {
      const backup = await exportBudgetControlBackup();
      const payload = binary ? encodeBinaryBackup(backup) : backup;
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = binary ? `budget-control-backup-${new Date().toISOString().slice(0,10)}.bin` : `budget-control-backup-${new Date().toISOString().slice(0,10)}.json`;
      a.click();
      URL.revokeObjectURL(url);
      setMessage(`Exported ${backup.collections.reduce((n, c) => n + c.documents.length, 0)} records.`);
    } catch (e) { setMessage(e instanceof Error ? e.message : 'Export failed.'); }
    finally { setBusy(false); }
  };

  const runImport = async (file: File) => {
    setBusy(true); setMessage('');
    try {
      const backup = await decodeBackupFile(file);
      const result = await restoreBudgetControlBackup(backup);
      setMessage(`Imported ${result.written} records. Reload the page to refresh data.`);
    } catch (e) { setMessage(e instanceof Error ? e.message : 'Import failed.'); }
    finally { setBusy(false); if (inputRef.current) inputRef.current.value = ''; }
  };

  return (
    <>
      <button onClick={() => setOpen(true)} className="fixed bottom-4 right-4 z-[100] inline-flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-lg hover:bg-slate-800" title="Admin database backup and restore">
        <Database className="h-4 w-4" /> Data Backup
      </button>

      {open && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center bg-black/40 p-4">
          <div className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-2xl">
            <div className="mb-5 flex items-center justify-between">
              <div>
                <div className="flex items-center gap-2 text-lg font-bold text-slate-900"><ShieldCheck className="h-5 w-5" /> Administrator Data Backup</div>
                <p className="mt-1 text-xs text-slate-500">Firestore data export/import for the Budget Control collections.</p>
              </div>
              <button onClick={() => setOpen(false)} className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button>
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <button disabled={busy} onClick={() => runExport(false)} className="rounded-xl border border-slate-200 px-4 py-4 text-left hover:bg-slate-50 disabled:opacity-50"><Download className="mb-2 h-5 w-5" /><div className="font-semibold">Export JSON</div><div className="text-xs text-slate-500">Human-readable database backup.</div></button>
              <button disabled={busy} onClick={() => runExport(true)} className="rounded-xl border border-slate-200 px-4 py-4 text-left hover:bg-slate-50 disabled:opacity-50"><Download className="mb-2 h-5 w-5" /><div className="font-semibold">Export BIN</div><div className="text-xs text-slate-500">Base64-wrapped binary-compatible backup.</div></button>
              <button disabled={busy} onClick={() => inputRef.current?.click()} className="rounded-xl border border-indigo-200 bg-indigo-50 px-4 py-4 text-left hover:bg-indigo-100 disabled:opacity-50"><Upload className="mb-2 h-5 w-5 text-indigo-600" /><div className="font-semibold text-indigo-900">Import Backup</div><div className="text-xs text-indigo-700">Restore JSON or BIN into Firestore.</div></button>
            </div>
            <input ref={inputRef} type="file" accept=".json,.bin,application/json" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) void runImport(f); }} />
            {message && <div className="mt-4 rounded-lg bg-slate-50 p-3 text-sm text-slate-700">{message}</div>}
            {busy && <div className="mt-3 text-xs text-slate-500">Processing… please keep this window open.</div>}
          </div>
        </div>
      )}
    </>
  );
}
