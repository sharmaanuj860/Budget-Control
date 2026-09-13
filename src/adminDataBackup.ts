import {
  collection,
  getDocs,
  writeBatch,
  doc,
  getDoc,
  db,
  auth
} from './firebase';

export type BackupCollection = {
  name: string;
  documents: Record<string, unknown>[];
};

export type BudgetControlBackup = {
  format: 'budget-control-backup';
  version: 1;
  exportedAt: string;
  exportedBy: { uid: string; email: string | null } | null;
  collections: BackupCollection[];
};

// Keep this list explicit: importing/exporting only known application data
// avoids accidentally copying unrelated Firebase collections.
export const BACKUP_COLLECTIONS = [
  'financialYears',
  'ranges',
  'schemes',
  'sectors',
  'activities',
  'subActivities',
  'soeHeads',
  'approvedBudgets',
  'treasuryReceipts',
  'soeBudgets',
  'allocations',
  'expenditures',
  'appSettings',
  'bills',
  'surrenders',
  'payees',
  'featureLocks',
  'approvedBudgetFiles',
  'distributedBudgetFiles',
  'notifications',
  'memos',
  'auditLogs'
] as const;

const ensureAdmin = async () => {
  const user = auth.currentUser;
  if (!user) throw new Error('You must be signed in as an administrator.');

  const snap = await getDoc(doc(db, 'users', user.uid));
  const role = snap.exists() ? snap.data()?.role : undefined;
  if (role !== 'admin' && user.email !== 'sharmaanuj860@gmail.com' && user.email !== 'admin@rajgarhforest.app') {
    throw new Error('Administrator access is required for backup and restore.');
  }
  return user;
};

export async function exportBudgetControlBackup(): Promise<BudgetControlBackup> {
  const user = await ensureAdmin();
  const collections: BackupCollection[] = [];

  for (const name of BACKUP_COLLECTIONS) {
    const snap = await getDocs(collection(db, name));
    collections.push({
      name,
      documents: snap.docs.map((item) => ({ id: item.id, ...item.data() }))
    });
  }

  return {
    format: 'budget-control-backup',
    version: 1,
    exportedAt: new Date().toISOString(),
    exportedBy: { uid: user.uid, email: user.email },
    collections
  };
}

export async function restoreBudgetControlBackup(backup: unknown): Promise<{ written: number }> {
  await ensureAdmin();

  if (!backup || typeof backup !== 'object') throw new Error('Invalid backup file.');
  const candidate = backup as Partial<BudgetControlBackup>;
  if (candidate.format !== 'budget-control-backup' || candidate.version !== 1 || !Array.isArray(candidate.collections)) {
    throw new Error('Unsupported or invalid Budget Control backup format.');
  }

  const allowed = new Set(BACKUP_COLLECTIONS);
  let written = 0;
  let batch = writeBatch(db);
  let batchCount = 0;

  const commitBatch = async () => {
    if (batchCount === 0) return;
    await batch.commit();
    batch = writeBatch(db);
    batchCount = 0;
  };

  for (const group of candidate.collections) {
    if (!group || !allowed.has(group.name as typeof BACKUP_COLLECTIONS[number]) || !Array.isArray(group.documents)) continue;

    for (const raw of group.documents) {
      if (!raw || typeof raw !== 'object') continue;
      const item = raw as Record<string, unknown>;
      const id = typeof item.id === 'string' ? item.id : undefined;
      if (!id) continue;
      const { id: _ignored, ...data } = item;
      batch.set(doc(db, group.name, id), data, { merge: false });
      batchCount += 1;
      written += 1;
      if (batchCount >= 450) await commitBatch();
    }
  }

  await commitBatch();
  return { written };
}
