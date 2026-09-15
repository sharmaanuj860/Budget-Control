import { collection, getDocs, writeBatch, doc, getDoc, db, auth } from './firebase';

export type BackupCollection = { name: string; documents: Record<string, unknown>[] };
export type BudgetControlBackup = { format: 'budget-control-backup'; version: 2; exportedAt: string; exportedBy: { uid: string; email: string | null } | null; collections: BackupCollection[] };
export type BudgetControlBinaryBackup = { format: 'budget-control-bin'; version: 1; encoding: 'base64-json'; payload: string };

export const BACKUP_COLLECTIONS = ['financialYears','ranges','schemes','sectors','activities','subActivities','soeHeads','approvedBudgets','treasuryReceipts','soeBudgets','allocations','expenditures','appSettings','bills','surrenders','payees','featureLocks','approvedBudgetFiles','distributedBudgetFiles','notifications','memos','auditLogs'] as const;
const ADMIN_EMAILS = new Set(['sharmaanuj860@gmail.com','admin@rajgarhforest.app']);

async function ensureAdmin() {
  const user = auth.currentUser;
  if (!user) throw new Error('You must be signed in as an administrator.');
  const snap = await getDoc(doc(db, 'users', user.uid));
  const role = snap.exists() ? snap.data()?.role : undefined;
  if (role !== 'admin' && !ADMIN_EMAILS.has(user.email ?? '')) throw new Error('Administrator access is required for backup and restore.');
  return user;
}

function serialise(value: unknown): unknown {
  if (value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(serialise);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value as Record<string, unknown>).map(([k, v]) => [k, serialise(v)]));
  return value;
}

export async function exportBudgetControlBackup(): Promise<BudgetControlBackup> {
  const user = await ensureAdmin();
  const groups: BackupCollection[] = [];
  for (const name of BACKUP_COLLECTIONS) {
    const snap = await getDocs(collection(db, name));
    groups.push({ name, documents: snap.docs.map(item => ({ id: item.id, ...(serialise(item.data()) as Record<string, unknown>) })) });
  }
  return { format: 'budget-control-backup', version: 2, exportedAt: new Date().toISOString(), exportedBy: { uid: user.uid, email: user.email }, collections: groups };
}

export function encodeBinaryBackup(backup: BudgetControlBackup): BudgetControlBinaryBackup {
  const bytes = new TextEncoder().encode(JSON.stringify(backup));
  let binary = ''; for (const byte of bytes) binary += String.fromCharCode(byte);
  return { format: 'budget-control-bin', version: 1, encoding: 'base64-json', payload: btoa(binary) };
}

export async function decodeBackupFile(file: File): Promise<unknown> {
  const text = (await file.text()).trim();
  if (!file.name.toLowerCase().endsWith('.bin')) return JSON.parse(text);
  const wrapper = JSON.parse(text) as BudgetControlBinaryBackup;
  if (wrapper.format !== 'budget-control-bin' || wrapper.version !== 1 || wrapper.encoding !== 'base64-json') throw new Error('Unsupported .bin backup format.');
  const binary = atob(wrapper.payload);
  const bytes = Uint8Array.from(binary, ch => ch.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

export async function restoreBudgetControlBackup(backup: unknown): Promise<{ written: number }> {
  await ensureAdmin();
  if (!backup || typeof backup !== 'object') throw new Error('Invalid backup file.');
  const candidate = backup as Partial<BudgetControlBackup>;
  if (candidate.format !== 'budget-control-backup' || ![1, 2].includes(candidate.version as number) || !Array.isArray(candidate.collections)) throw new Error('Unsupported or invalid Budget Control backup format.');
  const allowed = new Set(BACKUP_COLLECTIONS); let written = 0; let batch = writeBatch(db); let count = 0;
  const commit = async () => { if (!count) return; await batch.commit(); batch = writeBatch(db); count = 0; };
  for (const group of candidate.collections) {
    if (!group || !allowed.has(group.name as typeof BACKUP_COLLECTIONS[number]) || !Array.isArray(group.documents)) continue;
    for (const raw of group.documents) {
      if (!raw || typeof raw !== 'object') continue;
      const item = raw as Record<string, unknown>; const id = typeof item.id === 'string' ? item.id : null; if (!id) continue;
      const { id: _ignored, ...data } = item; batch.set(doc(db, group.name, id), data, { merge: false }); count++; written++; if (count >= 450) await commit();
    }
  }
  await commit(); return { written };
}
