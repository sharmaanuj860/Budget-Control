import React, { useState, useMemo, useEffect, useRef, useCallback } from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid, LineChart, Line } from 'recharts';
import { IndianRupee, Wallet, TrendingDown, Landmark, Activity, FileText, Map, MapPin, Plus, Trash2, Download, LogOut, User, UserCheck, Shield, FileBarChart, Filter, Search, Menu, Table, Pencil, Edit2, Home, ChevronUp, ChevronDown, TreePine, Check, X, Unlock, RefreshCcw, RefreshCw, Save, Eye, EyeOff, ShieldCheck, Lock, TrendingUp, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, Printer, CornerUpLeft, Calendar, PieChart as PieChartIcon, Maximize2, Minimize2, Bell, MoveHorizontal, PlusCircle, Users, Send, History, Building2, DollarSign, AlertTriangle, CheckCircle, CheckCircle2, ArrowRight, Clock, ArrowUpRight, QrCode, Smartphone, Copy, ExternalLink, Share2, Scan, Undo2, Loader2, Inbox, Globe, Laptop, Wifi, WifiOff, Monitor, Tablet, Cpu, HardDrive, Layers, ShieldAlert, Radio, Sparkles } from 'lucide-react';
import { getDeviceHardwareProfile, fetchClientNetworkLocation, getCachedNetworkProfile, DeviceHardwareProfile, NetworkLocationProfile } from './deviceTracking';
import QRCode from 'qrcode';
import { initializeApp, getApps } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { 
  auth, db, storage, signInWithPopup, googleProvider, signOut, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, sendPasswordResetEmail, setPersistence, browserSessionPersistence, browserLocalPersistence,
  collection, doc, setDoc, getDoc, getDocs, onSnapshot, query, where, or, orderBy, addDoc, updateDoc, deleteDoc, getDocFromServer, firebaseConfig, runTransaction, writeBatch,
  ref, uploadBytes, uploadBytesResumable, getDownloadURL, deleteObject
} from './firebase';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { preloadDatabase } from './preloadData';
import { emergencyRestoreData } from './emergencyRestoreData';

// --- Types ---
type FinancialYear = { id: string; name: string };
type Range = { id: string; name: string };
type Scheme = { id: string; name: string };
type Sector = { id: string; schemeId: string; name: string };
type ActivityItem = { id: string; sectorId?: string; schemeId?: string; name: string };
type SubActivity = { id: string; activityId: string; name: string };
type SOE = { 
  id: string; 
  name: string; 
  isProvisional?: boolean;
  schemeId?: string; 
  sectorId?: string; 
  activityId?: string; 
  subActivityId?: string; 
  approvedBudget: number; 
  receivedInTry: number;
  fyId?: string;
  financialYear?: string;
  approvedBudgetAmount?: number;
  receivedInTryAmount?: number;
  tryAmount?: number;
  updatedAt?: number;
  createdAt?: number;
};

type Allocation = { 
  id: string; 
  rangeId: string; 
  schemeId: string; 
  sectorId?: string; 
  activityId?: string; 
  subActivityId?: string; 
  amount: number; 
  status: 'Pending SOE Funds' | 'Funded'; 
  fundedSOEs: { soeId: string; amount: number }[]; 
  fyId?: string; 
  financialYear?: string; 
  remarks?: string;
  updatedAt?: number;
  createdAt?: number;
};

type Expense = { 
  id: string; 
  allocationId: string; 
  soeId: string; 
  schemeId?: string;
  sectorId?: string;
  amount: number; 
  date: string; 
  description: string; 
  fyId?: string; 
  financialYear?: string; 
  status: 'pending' | 'approved' | 'rejected'; 
  isLocked: boolean; 
  approvalId?: number;
  createdBy?: string;
  createdByRole?: string;
  updatedBy?: string;
  updatedByRole?: string;
  updatedAt?: number;
  createdAt?: number;
  approvalReason?: string;
  payeeId?: string;
  payeeName?: string;
  rangeId?: string;
  deductionType?: 'None' | 'TDS' | 'TDS_GST' | 'Both';
  deductions?: { type: 'TDS' | 'TDS_GST'; amount: number }[];
  deductedAmount?: number;
  tdsAmount?: number;
  tdsGstAmount?: number;
  panNumber?: string;
  gstNumber?: string;
  netAmount?: number;
  syncedMemoId?: string;
  syncedMemoNo?: string;
};

type Notification = {
  id: string;
  name: string;
  url: string;
  fileData?: string;
  type: string;
  createdAt: number;
  uploadedBy: string;
  description?: string;
  targetRanges?: string[];
  category?: 'general' | 'budget_allocation' | 'budget_distribution' | 'system';
  amount?: number;
  schemeName?: string;
  sectorName?: string;
  rangeName?: string;
  soeName?: string;
  targetTab?: string;
  targetId?: string;
};

type Bill = {
  id: string;
  billNo: string;
  billDate: string;
  expenseIds: string[];
  fyId: string;
  financialYear: string;
  totalAmount: number;
  status: 'draft' | 'finalized';
  createdAt: number;
  updatedAt: number;
  remarks?: string;
};

type Payee = {
  id: string;
  name: string;
  address: string;
  accountNumber: string;
  ifscCode?: string;
  panNumber?: string;
  gstNumber?: string;
  treasuryCode?: string;
  rangeId?: string;
  mappedUserIds?: string[];
  mappedRangeIds?: string[];
  createdAt: number;
  updatedAt: number;
  createdBy: string;
};

type MemoSubVoucher = {
  id: string;
  voucherNo?: string;
  description?: string;
  amount: number;
};

type MemoPayeeEntry = {
  id?: string;
  payeeId?: string;
  name: string;
  address: string;
  accountNumber: string;
  ifscCode: string;
  treasuryCode?: string;
  panNumber: string;
  gstNumber?: string;
  totalAmount: number;
  deductITax?: boolean;
  iTaxPercent: number;
  iTaxAmount: number;
  deductGst?: boolean;
  gstPercent?: number;
  gstAmount?: number;
  netRtgsAmount: number;
  description?: string;
  subVouchers?: MemoSubVoucher[];
};

type MemoForFund = {
  id: string;
  memoNo: string;
  date: string;
  monthYear: string;
  schemeId?: string;
  schemeName?: string;
  sectorId?: string;
  sectorName?: string;
  soeId?: string;
  soeName?: string;
  rangeId?: string;
  rangeName?: string;
  toAuthority?: string;
  financialYear: string;
  fyId?: string;
  status: 'draft' | 'submitted' | 'correction';
  totalAmount: number;
  totalITax: number;
  totalGst?: number;
  totalNetRtgs: number;
  payeeEntries: MemoPayeeEntry[];
  createdBy: string;
  createdByRole?: string;
  createdByName?: string;
  createdAt: number;
  updatedAt: number;
  remarks?: string;
  correctionRemarks?: string;
  correctionRemarksBy?: string;
  correctionRemarksAt?: number;
  isApproved?: boolean;
  approvedBy?: string;
  approvedByRole?: string;
  approvedAt?: number;
  viewedByAdmin?: boolean;
  viewedAt?: number;
  viewedBy?: string;
  viewedByRole?: string;
  pulledBack?: boolean;
  pullBackRemarks?: string;
  pulledBackBy?: string;
  pulledBackAt?: number;
  copiedFromMemoNo?: string;
};

type BudgetFile = {
  id: string;
  name: string;
  url: string;
  schemeId: string;
  sectorId?: string;
  rangeId?: string;
  type: 'approved' | 'distributed';
  uploadedBy: string;
  uploadedAt: number;
  fyId: string;
};

type AppUser = { 
  id: string; 
  email: string; 
  role: 'admin' | 'deo' | 'approver' | 'DA' | 'Sarahan' | 'Narag' | 'Habban' | 'Division' | 'Rajgarh'; 
  password?: string;
  maxSessions?: number;
  activeSessions?: string[];
  isDisabled?: boolean;
  updatedAt?: number;
};

type FeatureLock = {
  id: string;
  feature: 'Allocation' | 'Expenditure' | 'Access' | 'Memo' | 'MemoSync';
  target: string; // role or rangeId
  isLocked: boolean;
  updatedBy: string;
  updatedAt: number;
};

type Surrender = {
  id: string;
  rangeId: string;
  schemeId: string;
  sectorId: string;
  activityId: string;
  subActivityId: string;
  soeId: string;
  amount: number;
  date: string;
  remarks: string;
  fyId: string;
  financialYear: string;
  createdAt: number;
  updatedAt: number;
};

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function convertNumberToWords(amount: number): string {
  if (isNaN(amount) || amount <= 0) return 'Zero';
  
  const single = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const double = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function numToWords(n: number): string {
    let str = '';
    if (n > 99) {
      str += single[Math.floor(n / 100)] + ' Hundred ';
      n %= 100;
    }
    if (n > 19) {
      str += tens[Math.floor(n / 10)] + ' ';
      n %= 10;
    } else if (n >= 10) {
      str += double[n - 10] + ' ';
      n = 0;
    }
    if (n > 0) {
      str += single[n] + ' ';
    }
    return str;
  }

  const rounded = Math.floor(Math.abs(amount));
  let str = '';
  
  const crore = Math.floor(rounded / 10000000);
  let rem = rounded % 10000000;
  
  const lakh = Math.floor(rem / 100000);
  rem %= 100000;
  
  const thousand = Math.floor(rem / 1000);
  rem %= 1000;
  
  const hundred = rem;

  if (crore > 0) str += numToWords(crore) + 'Crore ';
  if (lakh > 0) str += numToWords(lakh) + 'Lakh ';
  if (thousand > 0) str += numToWords(thousand) + 'Thousand ';
  if (hundred > 0) str += numToWords(hundred);

  return str.trim() || 'Zero';
}

function sanitizeFirestoreDoc<T>(obj: T): T {
  if (obj === undefined) return "" as any;
  if (obj === null) return null as any;
  if (Array.isArray(obj)) {
    return obj.map(item => sanitizeFirestoreDoc(item)) as any;
  }
  if (typeof obj === 'object' && !(obj instanceof Date)) {
    const cleaned: Record<string, any> = {};
    for (const [key, value] of Object.entries(obj)) {
      if (value !== undefined) {
        cleaned[key] = sanitizeFirestoreDoc(value);
      }
    }
    return cleaned as any;
  }
  return obj;
}

function getDeviceInfo(): string {
  if (typeof window === 'undefined') return 'Web Client';
  return getDeviceHardwareProfile().summary;
}

let globalAlertDispatcher: ((msg: string) => void) | null = null;

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errText = error instanceof Error ? error.message : String(error);
  const errInfo: FirestoreErrorInfo = {
    error: errText,
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email || undefined,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId || undefined,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  if (globalAlertDispatcher) {
    globalAlertDispatcher(`Database operation failed (${operationType} on ${path || 'database'}).\n\nPlease contact the administrator with this error:\n"${errText}"`);
  }
}

const TryUpdateInput = ({ soeId, initialValue, onUpdate }: { soeId: string, initialValue: number, onUpdate: (id: string, val: number) => void }) => {
  const [val, setVal] = useState(initialValue);
  
  useEffect(() => {
    setVal(initialValue);
  }, [initialValue]);

  return (
    <div className="flex items-center gap-1">
      <input 
        type="number" 
        value={val} 
        onChange={(e) => setVal(parseFloat(e.target.value) || 0)}
        className="w-24 p-1 text-xs border rounded focus:ring-1 focus:ring-indigo-500 outline-none"
      />
      <button 
        onClick={() => onUpdate(soeId, val)}
        className="p-1 bg-indigo-50 text-indigo-600 rounded hover:bg-indigo-100 transition-colors"
        title="Update Treasury"
      >
        <Check className="w-3 h-3" />
      </button>
    </div>
  );
};

const getApprovedBudget = (s: any) => {
  const val = s.approvedBudget || s.approvedBudgetAmount || s.approved_budget || s.budget || 0;
  return Number(val) || 0;
};

const getReceivedInTry = (s: any) => {
  const val = s.receivedInTry || s.receivedInTryAmount || s.tryAmount || s.received_in_try || 0;
  return Number(val) || 0;
};

const ALLOWED_SOES = ['20 OC', '21 Maint', '30MV', '33M&S', '36M&W', 'Provisional'];

// --- Payee Selector Component ---
const PayeeSelector = ({ 
  payees, 
  selectedPayees, 
  onSelect, 
  onRemove, 
  onAmountChange,
  ranges,
  availableBalance,
  selectedDeductions = [],
  gstNumber = ''
}: { 
  payees: Payee[], 
  selectedPayees: { payeeId: string, amount: string }[], 
  onSelect: (payeeId: string) => void, 
  onRemove: (payeeId: string) => void, 
  onAmountChange: (payeeId: string, amount: string) => void,
  ranges: Range[],
  availableBalance?: number,
  selectedDeductions?: string[],
  gstNumber?: string
}) => {
  const [search, setSearch] = useState('');
  const [showResults, setShowResults] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowResults(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredPayees = useMemo(() => {
    const lower = search.toLowerCase();
    const base = payees.filter(p => !selectedPayees.some(sp => sp.payeeId === p.id));
    const seen: Record<string, boolean> = {};
    const uniqueList: Payee[] = [];
    base.forEach(p => {
      if (p.id && !seen[p.id]) {
        seen[p.id] = true;
        uniqueList.push(p);
      }
    });
    if (!search) return uniqueList;
    return uniqueList.filter(p => 
      p.name?.toLowerCase().includes(lower) || 
      p.accountNumber?.toLowerCase().includes(lower) ||
      p.panNumber?.toLowerCase().includes(lower) ||
      p.gstNumber?.toLowerCase().includes(lower) ||
      p.treasuryCode?.toLowerCase().includes(lower)
    );
  }, [search, payees, selectedPayees]);

  const totalAmount = useMemo(() => 
    selectedPayees.reduce((sum, p) => sum + (parseFloat(p.amount) || 0), 0)
  , [selectedPayees]);

  const totalCalculatedDeductions = useMemo(() => {
    let totalDeductions = 0;
    selectedPayees.forEach(sp => {
      const p = payees.find(payee => payee.id === sp.payeeId);
      const pAmt = parseFloat(sp.amount) || 0;
      const pGst = p?.gstNumber || gstNumber;
      const pTds = (selectedDeductions.includes('TDS') && pAmt > 30000) ? Math.round(pAmt * 0.01) : 0;
      const pGstTds = (selectedDeductions.includes('TDS_GST') && pAmt > 250000) ? Math.round(pAmt * 0.02) : 0;
      totalDeductions += (pTds + pGstTds);
    });
    return totalDeductions;
  }, [selectedPayees, payees, selectedDeductions, gstNumber]);

  return (
    <div className="space-y-3 p-3 bg-gray-50 rounded-lg border border-gray-200" ref={dropdownRef}>
      <div className="flex justify-between items-center">
        <label className="block text-[10px] font-bold text-gray-500 uppercase">Payee Selection & Tax Breakdown</label>
        {availableBalance !== undefined && (
          <div className="flex items-center gap-2">
            <span className="text-[10px] font-bold text-gray-500 uppercase">Available:</span>
            <span className="text-xs font-bold text-blue-700">â‚¹{availableBalance.toLocaleString()}</span>
          </div>
        )}
      </div>
      
      <div className="space-y-2">
        <div className="flex flex-col gap-2">
          {/* Select Payee Dropdown */}
          <select 
            className="w-full p-2 text-sm border rounded bg-white shadow-sm focus:ring-2 focus:ring-emerald-500 outline-none font-medium text-gray-800"
            value=""
            onChange={(e) => {
              if (e.target.value) {
                onSelect(e.target.value);
                setSearch('');
                setShowResults(false);
              }
            }}
          >
            <option value="">-- Select Payee from List --</option>
            {filteredPayees.map((p, idx) => {
              const rName = ranges.find(r => r.id === p.rangeId)?.name || '';
              return (
                <option key={`opt-payee-${p.id}-${idx}`} value={p.id} title={`${p.name} (A/C: ${p.accountNumber})`}>
                  {p.name} (A/C: {p.accountNumber}){rName ? ` - [${rName}]` : ''}
                </option>
              );
            })}
          </select>

          {/* Search Input Box */}
          <div className="relative flex-1">
            <Search className="absolute left-2 top-2.5 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Or type to search payee..." 
              value={search}
              onChange={(e) => { setSearch(e.target.value); setShowResults(true); }}
              onFocus={() => setShowResults(true)}
              className="w-full pl-8 pr-8 p-2 text-sm border rounded bg-white shadow-sm focus:ring-2 focus:ring-emerald-500 outline-none transition-all"
            />
            <button 
              type="button"
              onClick={() => setShowResults(!showResults)}
              className="absolute right-2 top-2 text-gray-400 hover:text-gray-600"
            >
              {showResults ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
          </div>
        </div>

        {showResults && (
          <div className="relative w-full mt-1 bg-white border rounded-lg shadow-lg max-h-60 overflow-y-auto z-20">
            {filteredPayees.length > 0 ? (
              filteredPayees.map((p, idx) => (
                <button
                  key={`btn-payee-${p.id}-${idx}`}
                  type="button"
                  onClick={() => {
                    onSelect(p.id);
                    setSearch('');
                    setShowResults(false);
                  }}
                  className="w-full text-left px-3 py-2 text-sm hover:bg-emerald-50 border-b last:border-0 group transition-colors"
                >
                  <div className="flex justify-between items-start">
                    <div>
                      <div className="font-bold text-gray-800">{p.name}</div>
                      <div className="text-[10px] text-gray-500 font-mono">A/C: {p.accountNumber}</div>
                      <div className="flex flex-wrap gap-1.5 mt-0.5">
                        {p.treasuryCode && <span className="text-[9px] bg-emerald-50 border border-emerald-200 px-1 rounded text-emerald-800 font-bold font-mono">TRY: {p.treasuryCode}</span>}
                        {p.ifscCode && <span className="text-[9px] bg-blue-50 px-1 rounded text-blue-700">IFSC: {p.ifscCode}</span>}
                        {p.panNumber && <span className="text-[9px] bg-gray-100 px-1 rounded text-gray-600">PAN: {p.panNumber}</span>}
                        {p.gstNumber && <span className="text-[9px] bg-gray-100 px-1 rounded text-gray-600">GST: {p.gstNumber}</span>}
                      </div>
                    </div>
                    <div className="text-right">
                      <div className="text-[9px] font-bold text-emerald-600 truncate max-w-[80px]">{ranges.find(r => r.id === p.rangeId)?.name || 'N/A'}</div>
                    </div>
                  </div>
                </button>
              ))
            ) : (
              <div className="px-4 py-3 text-sm text-gray-500 italic text-center">No matching payees available</div>
            )}
          </div>
        )}
      </div>

      {selectedPayees.length > 0 && (
        <div className="space-y-2">
          {selectedPayees.map((sp, idx) => {
            const p = payees.find(payee => payee.id === sp.payeeId);
            const pAmt = parseFloat(sp.amount) || 0;
            const pGst = p?.gstNumber || gstNumber;
            const pTds = (selectedDeductions.includes('TDS') && pAmt > 30000) ? Math.round(pAmt * 0.01) : 0;
            const pGstTds = (selectedDeductions.includes('TDS_GST') && pAmt > 250000) ? Math.round(pAmt * 0.02) : 0;
            const pTax = pTds + pGstTds;
            const pNet = pAmt - pTax;

            return (
              <div key={`sel-payee-${sp.payeeId}-${idx}`} className="bg-white p-2.5 rounded-lg border border-gray-200 shadow-xs space-y-1.5">
                <div className="flex items-center gap-2">
                  <div className="flex-1 min-w-0">
                    <div className="text-xs font-bold text-gray-900 truncate">{p?.name}</div>
                    <div className="text-[10px] text-gray-500 font-mono truncate">A/C: {p?.accountNumber}{p?.ifscCode ? ` | IFSC: ${p.ifscCode}` : ''}</div>
                  </div>
                  <div className="w-36 shrink-0">
                    <div className="relative">
                      <span className="absolute left-2 top-2 text-xs font-bold text-gray-400">â‚¹</span>
                      <input 
                        type="text" 
                        inputMode="decimal"
                        placeholder="0.00" 
                        value={sp.amount}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val === '' || /^\d*\.?\d*$/.test(val)) {
                            onAmountChange(sp.payeeId, val);
                          }
                        }}
                        className="w-full pl-5 pr-2 py-1.5 text-sm border-2 border-emerald-100 rounded text-right font-bold text-emerald-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 outline-none transition-all"
                        required
                      />
                    </div>
                  </div>
                  <div className="shrink-0">
                    <button 
                      type="button"
                      onClick={() => onRemove(sp.payeeId)}
                      className="p-1.5 text-red-500 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                      title="Remove Payee"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Tax & Net Pay Info per Payee */}
                {pAmt > 0 && (
                  <div className="flex flex-wrap items-center justify-between gap-1.5 pt-1.5 border-t border-gray-100 text-[11px] bg-slate-50/90 px-2.5 py-1.5 rounded-md">
                    <div className="flex flex-wrap items-center gap-2 text-gray-700">
                      <span>Gross: <strong className="text-gray-900 font-bold">â‚¹{pAmt.toLocaleString('en-IN')}</strong></span>
                      {pTds > 0 && (
                        <span className="text-red-600 font-semibold bg-red-50 px-1.5 py-0.5 rounded border border-red-100">
                          TDS (1%): -â‚¹{pTds.toLocaleString('en-IN')}
                        </span>
                      )}
                      {pGstTds > 0 && (
                        <span className="text-purple-600 font-semibold bg-purple-50 px-1.5 py-0.5 rounded border border-purple-100">
                          TDS GST (2%): -â‚¹{pGstTds.toLocaleString('en-IN')}
                        </span>
                      )}
                      {pTax === 0 && selectedDeductions.length > 0 && (
                        <span className="text-gray-400 text-[10px]">
                          TDS: â‚¹0 {selectedDeductions.includes('TDS') && pAmt <= 30000 ? '(â‰¤â‚¹30k threshold)' : ''}
                        </span>
                      )}
                    </div>
                    <div className="font-extrabold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                      Net Pay (RTGS): â‚¹{pNet.toLocaleString('en-IN')}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
          <div className="space-y-1 pt-2 border-t border-gray-200">
            <div className="flex justify-between items-center">
              <span className="text-[10px] font-bold text-gray-500 uppercase">Total Gross Amount:</span>
              <span className="text-sm font-bold text-gray-900">â‚¹{totalAmount.toLocaleString()}</span>
            </div>
            {totalCalculatedDeductions > 0 && (
              <div className="flex justify-between items-center text-xs">
                <span className="text-[10px] font-bold text-red-600 uppercase">Total Tax Deductions:</span>
                <span className="font-bold text-red-600">-â‚¹{totalCalculatedDeductions.toLocaleString()}</span>
              </div>
            )}
            {selectedDeductions.length > 0 && (
              <div className="flex justify-between items-center text-xs">
                <span className="text-[10px] font-bold text-emerald-700 uppercase">Total Net Payable (RTGS):</span>
                <span className="text-sm font-black text-emerald-700">â‚¹{(totalAmount - totalCalculatedDeductions).toLocaleString()}</span>
              </div>
            )}
            {availableBalance !== undefined && (
              <div className="flex justify-between items-center">
                <span className="text-[10px] font-bold text-gray-500 uppercase">Remaining Balance:</span>
                <span className={`text-sm font-bold ${availableBalance - totalAmount < 0 ? 'text-red-600' : 'text-blue-600'}`}>
                  â‚¹{(availableBalance - totalAmount).toLocaleString()}
                </span>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState('Dashboard');
  const [allocationAmount, setAllocationAmount] = useState<string>('');
  const [expenseAmount, setExpenseAmount] = useState<string>('');
  const [expenseDate, setExpenseDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [expenseDescription, setExpenseDescription] = useState<string>('');
  const [trackerSearch, setTrackerSearch] = useState('');
  const [menuOpen, setMenuOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [deductionType, setDeductionType] = useState<'None' | 'TDS' | 'TDS_GST' | 'Both'>('None');
  const [selectedDeductions, setSelectedDeductions] = useState<('TDS' | 'TDS_GST')[]>([]);
  const [panNumber, setPanNumber] = useState('');
  const [gstNumber, setGstNumber] = useState('');
  const [payeePan, setPayeePan] = useState('');
  const [payeeGst, setPayeeGst] = useState('');
  const [uploadTasks, setUploadTasks] = useState<{[key: string]: any}>({});
  const [billSearchTerm, setBillSearchTerm] = useState('');
  const [payeeSearchTerm, setPayeeSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(25);
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [showAllBudget, setShowAllBudget] = useState(false);
  const [rangeSearch, setRangeSearch] = useState('');
  const [soeSearchTerm, setSoeSearchTerm] = useState('');
  const [soeAbstractSearch, setSoeAbstractSearch] = useState('');
  const [showAllRange, setShowAllRange] = useState(false);
  const [isFormExpanded, setIsFormExpanded] = useState(window.innerWidth > 1024);
  const [formWidth, setFormWidth] = useState<'normal' | 'wide' | 'extra'>(() => (localStorage.getItem('fbc_form_width') as 'normal' | 'wide' | 'extra') || 'wide');
  const [isSoeTrackerExpanded, setIsSoeTrackerExpanded] = useState(false);
  const [showMobileReportSearch, setShowMobileReportSearch] = useState(false);
  const [showReconSummary, setShowReconSummary] = useState(false);
  const [user, setUser] = useState<any>(null);
  const [userRole, setUserRole] = useState<'admin' | 'deo' | 'approver' | 'DA' | 'Sarahan' | 'Narag' | 'Habban' | 'Division' | 'Rajgarh' | null>(null);
  const [loading, setLoading] = useState(true);
  const [fundingAllocation, setFundingAllocation] = useState<Allocation | null>(null);
  const [isSoesLoaded, setIsSoesLoaded] = useState(false);

  const handleLogout = async () => {
    try {
      logAuditAction('User Logout', 'User logged out of session');
      if (user) {
        const userDoc = await getDoc(doc(db, 'users', user.uid));
        if (userDoc.exists()) {
          const activeSessions = userDoc.data().activeSessions || [];
          const updatedSessions = activeSessions.filter((id: string) => id !== sessionId);
          await updateDoc(doc(db, 'users', user.uid), { activeSessions: updatedSessions });
        }
      }
      await signOut(auth);
      setUser(null);
      setUserRole(null);
      setActiveTab('Dashboard');
    } catch (error) {
      console.error("Logout error:", error);
    }
  };

  const [alertModal, setAlertModal] = useState<{ isOpen: boolean; message: string }>({ isOpen: false, message: '' });
  const [confirmModal, setConfirmModal] = useState<{ isOpen: boolean; message: string; onConfirm: () => void }>({ isOpen: false, message: '', onConfirm: () => {} });

  const showAlert = (message: string) => setAlertModal({ isOpen: true, message });
  useEffect(() => {
    globalAlertDispatcher = showAlert;
    return () => { globalAlertDispatcher = null; };
  }, []);
  const showConfirm = (message: string, onConfirm: () => void) => setConfirmModal({ isOpen: true, message, onConfirm });

  // --- State ---
  const [fys, setFys] = useState<FinancialYear[]>([]);
  const [selectedFY, setSelectedFY] = useState<string>('2026-27');
  const [loginFY, setLoginFY] = useState<string>('2026-27');

  const fyOptions = useMemo(() => {
    const defaultFYs = ['2026-27', '2025-26', '2024-25'];
    const fetchedFYs = fys.map(f => f.name).filter(Boolean);
    return Array.from(new Set([...defaultFYs, ...fetchedFYs]));
  }, [fys]);
  const [isFyHiddenForUsers, setIsFyHiddenForUsers] = useState<boolean>(false);
  const [ranges, setRanges] = useState<Range[]>([]);
  const [schemes, setSchemes] = useState<Scheme[]>([]);
  const [sectors, setSectors] = useState<Sector[]>([]);
  const [activities, setActivities] = useState<ActivityItem[]>([]);
  const [subActivities, setSubActivities] = useState<SubActivity[]>([]);
  const [soes, setSoes] = useState<SOE[]>([]);
  const [surrenders, setSurrenders] = useState<Surrender[]>([]);
  const [surrenderFilters, setSurrenderFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
  const hasSeeded = React.useRef(false);

  const [allocations, setAllocations] = useState<Allocation[]>([]);
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [bills, setBills] = useState<Bill[]>([]);
  const [payees, setPayees] = useState<Payee[]>([]);
  const [users, setUsers] = useState<AppUser[]>([]);
  const [featureLocks, setFeatureLocks] = useState<FeatureLock[]>([]);
  const [sessionId] = useState(() => Math.random().toString(36).substring(2, 15));
  const [newUserEmail, setNewUserEmail] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserRole, setNewUserRole] = useState<'admin' | 'deo' | 'approver' | 'Sarahan' | 'Narag' | 'Habban' | 'Division' | 'Rajgarh'>('deo');
  const [visiblePasswords, setVisiblePasswords] = useState<{[key: string]: boolean}>({});
  const [editingPasswordId, setEditingPasswordId] = useState<string | null>(null);
  const [newPasswordInput, setNewPasswordInput] = useState('');
  const [selectedLockTarget, setSelectedLockTarget] = useState<string>('');

  // --- Filters ---
  const [expDateRange, setExpDateRange] = useState({ start: '', end: '' });
  const [expFilters, setExpFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '' });
  const [expenditureSubTab, setExpenditureSubTab] = useState<'list' | 'bills' | 'payees' | 'memo'>('list');
  const [showExpenditurePrintModal, setShowExpenditurePrintModal] = useState(false);
  const [showPayeePrintModal, setShowPayeePrintModal] = useState(false);
  const [treasuryCodeModalPayee, setTreasuryCodeModalPayee] = useState<Payee | null>(null);
  const [treasuryCodeInput, setTreasuryCodeInput] = useState<string>('');

  // Memo Sync state
  const [selectedSyncedMemo, setSelectedSyncedMemo] = useState<{ memoId: string; memoNo: string; entryId?: string } | null>(null);
  const [showMemoSyncModal, setShowMemoSyncModal] = useState<boolean>(false);
  const [memoSearchTerm, setMemoSearchTerm] = useState<string>('');

  // Audit Logs State & Network/Device Tracking
  const [auditLogs, setAuditLogs] = useState<any[]>([]);
  const [auditSearchTerm, setAuditSearchTerm] = useState<string>('');
  const [auditRoleFilter, setAuditRoleFilter] = useState<string>('all');
  const [auditActionFilter, setAuditActionFilter] = useState<string>('all');
  const [inspectAuditLog, setInspectAuditLog] = useState<any | null>(null);
  const [clientNetwork, setClientNetwork] = useState<NetworkLocationProfile | null>(null);
  const [clientIpAddress, setClientIpAddress] = useState<string>('Detecting...');
  const [isNetworkOnline, setIsNetworkOnline] = useState<boolean>(typeof navigator !== 'undefined' ? navigator.onLine : true);
  const [isRefreshingLogs, setIsRefreshingLogs] = useState<boolean>(false);

  useEffect(() => {
    const handleOnline = () => setIsNetworkOnline(true);
    const handleOffline = () => setIsNetworkOnline(false);
    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    let isMounted = true;
    fetchClientNetworkLocation().then((prof) => {
      if (isMounted && prof) {
        setClientNetwork(prof);
        setClientIpAddress(prof.ip || 'Direct Connection');
      }
    }).catch(() => {
      if (isMounted) setClientIpAddress('Direct Connection');
    });

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      isMounted = false;
    };
  }, []);

  const logAuditAction = async (action: string, details?: string, extraMetadata?: Record<string, any>) => {
    try {
      const userRangeNameForLog = userRole && ['Sarahan', 'Narag', 'Habban', 'Division', 'Rajgarh'].includes(userRole) ? userRole : (userRangeName || 'All Ranges');
      const hw = getDeviceHardwareProfile();
      const net = clientNetwork || getCachedNetworkProfile() || {
        ip: clientIpAddress !== 'Detecting...' ? clientIpAddress : (typeof window !== 'undefined' ? window.location.hostname : 'Direct IP'),
        city: '',
        region: 'Himachal Pradesh',
        country: 'India',
        countryCode: 'IN',
        postal: '',
        isp: 'Direct Network',
        org: '',
        locationString: 'Himachal Pradesh, India',
        timezone: 'Asia/Kolkata'
      };

      const logData = {
        action,
        details: details || '',
        userName: user?.displayName || user?.email?.split('@')[0] || 'User',
        userEmail: user?.email || '',
        userRole: userRole || 'User',
        userRange: userRangeNameForLog,
        // Public IP & Geolocation Location
        ipAddress: net.ip || clientIpAddress || 'Direct IP',
        ipLocation: net.locationString || 'Himachal Pradesh, India',
        ipCity: net.city || '',
        ipRegion: net.region || '',
        ipCountry: net.country || 'India',
        ipIsp: net.isp || net.org || '',
        // Hardware & Display Specifications
        screenResolution: hw.screenResolution,
        availableResolution: hw.availableResolution,
        viewportSize: hw.viewportSize,
        colorDepth: hw.colorDepth,
        pixelRatio: hw.pixelRatio,
        orientation: hw.orientation,
        hardwareConcurrency: hw.hardwareConcurrency,
        deviceMemory: hw.deviceMemory,
        touchSupport: hw.touchSupport,
        gpuRenderer: hw.gpuRenderer,
        hardwareSummary: hw.hardwareSummary,
        // Operating System & Browser Versions
        os: hw.os,
        browser: hw.browser,
        browserVersion: hw.browserVersion ? `${hw.browser} ${hw.browserVersion}` : hw.browser,
        deviceType: hw.deviceType,
        platform: hw.platform,
        timeZone: hw.timeZone,
        languages: hw.languages,
        deviceInfo: hw.summary,
        userAgent: hw.userAgent,
        timestamp: Date.now(),
        ...(extraMetadata || {})
      };
      await addDoc(collection(db, 'auditLogs'), sanitizeFirestoreDoc(logData));
    } catch (err) {
      console.warn("Could not log audit action:", err);
    }
  };

  useEffect(() => {
    if (!user) return;
    const isUserAdmin = userRole === 'admin' || user?.email?.toLowerCase() === 'admin@rajgarhforest.app' || user?.email?.toLowerCase() === 'sharmaanuj860@gmail.com';
    if (!isUserAdmin) {
      setAuditLogs([]);
      return;
    }
    const q = query(collection(db, 'auditLogs'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const logs = snapshot.docs.map(d => ({ id: d.id, ...d.data() }) as any);
      logs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
      setAuditLogs(logs.slice(0, 300));
    }, (error) => {
      console.warn("Audit logs listener error:", error);
    });
    return () => unsubscribe();
  }, [user, userRole]);

  const isMemoSyncEnabled = !featureLocks.some(l => l.feature === 'MemoSync' && (l.target === 'global' || l.target === 'all') && l.isLocked);
  const userRangeRoleName = userRole && ['Sarahan', 'Narag', 'Habban', 'Division', 'Rajgarh'].includes(userRole) ? userRole : '';
  const isMemoLockedForUser = userRole !== 'admin' && featureLocks.some(l => l.feature === 'Memo' && l.isLocked && (l.target === 'all' || l.target === 'global' || l.target === userRangeRoleName || l.target === userRole));

  // Memo for Fund states
  const [memos, setMemos] = useState<MemoForFund[]>([]);
  const [editingMemo, setEditingMemo] = useState<MemoForFund | null>(null);
  const [viewingMemo, setViewingMemo] = useState<MemoForFund | null>(null);

  const [viewingBillPdf, setViewingBillPdf] = useState<{ url: string; bill: any } | null>(null);
  // Bill QR Code & Mobile Verification State
  const [viewingBillQR, setViewingBillQR] = useState<{ bill: Bill; qrDataUrl: string } | null>(null);
  const [verifyingBillId, setVerifyingBillId] = useState<string | null>(null);
  const [verifyingBillData, setVerifyingBillData] = useState<Bill | null>(null);
  const [isVerifyingLoading, setIsVerifyingLoading] = useState<boolean>(false);
  const [verifySearchInput, setVerifySearchInput] = useState<string>('');
  const [copiedQRLink, setCopiedQRLink] = useState<boolean>(false);

  const [allocFilters, setAllocFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
  const [soeFilters, setSoeFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeName: '' });

  // Detect URL parameter for mobile QR code verification
  useEffect(() => {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const verifyParam = urlParams.get('verifyBill') || urlParams.get('billVerification') || urlParams.get('verify');
      if (verifyParam) {
        setVerifyingBillId(verifyParam.trim());
      }
    } catch (e) {
      console.error('Error parsing verifyBill URL param:', e);
    }
  }, []);

  // Fetch or resolve bill data for verification modal
  useEffect(() => {
    if (!verifyingBillId) {
      setVerifyingBillData(null);
      return;
    }
    const foundInState = bills.find(b => b.id === verifyingBillId || b.billNo.trim().toLowerCase() === verifyingBillId.trim().toLowerCase());
    if (foundInState) {
      setVerifyingBillData(foundInState);
      return;
    }

    setIsVerifyingLoading(true);
    getDoc(doc(db, 'bills', verifyingBillId)).then(snap => {
      if (snap.exists()) {
        setVerifyingBillData({ id: snap.id, ...snap.data() } as Bill);
      } else {
        getDocs(query(collection(db, 'bills'), where('billNo', '==', verifyingBillId))).then(qSnap => {
          if (!qSnap.empty) {
            const docData = qSnap.docs[0];
            setVerifyingBillData({ id: docData.id, ...docData.data() } as Bill);
          } else {
            setVerifyingBillData(null);
          }
        }).catch(() => setVerifyingBillData(null))
        .finally(() => setIsVerifyingLoading(false));
        return;
      }
    }).catch(() => setVerifyingBillData(null))
    .finally(() => setIsVerifyingLoading(false));
  }, [verifyingBillId, bills]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (showExpenditurePrintModal) setShowExpenditurePrintModal(false);
        if (showPayeePrintModal) setShowPayeePrintModal(false);
        if (viewingMemo) setViewingMemo(null);
        if (showMemoSyncModal) setShowMemoSyncModal(false);
        if (viewingBillPdf) setViewingBillPdf(null);
        if (viewingBillQR) setViewingBillQR(null);
        if (verifyingBillId) setVerifyingBillId(null);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [showExpenditurePrintModal, showPayeePrintModal, viewingMemo, showMemoSyncModal, viewingBillPdf, viewingBillQR, verifyingBillId]);

  // Form Header State
  const [memoNoInput, setMemoNoInput] = useState<string>('');
  const [memoDateInput, setMemoDateInput] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [memoMonthYearInput, setMemoMonthYearInput] = useState<string>('08/2026');
  const [memoSchemeIdInput, setMemoSchemeIdInput] = useState<string>('');
  const [memoSectorIdInput, setMemoSectorIdInput] = useState<string>('');
  const [memoSoeIdInput, setMemoSoeIdInput] = useState<string>('');
  const [memoFromInput, setMemoFromInput] = useState<string>('RFO Sarahan');
  const [memoToInput, setMemoToInput] = useState<string>('DCF Rajgarh');

  // Payee Entry Form State
  const [memoPayeeEntries, setMemoPayeeEntries] = useState<MemoPayeeEntry[]>([]);
  const [entryPayeeId, setEntryPayeeId] = useState<string>('');
  const [entryName, setEntryName] = useState<string>('');
  const [entryAddress, setEntryAddress] = useState<string>('');
  const [entryAccountNo, setEntryAccountNo] = useState<string>('');
  const [entryIfsc, setEntryIfsc] = useState<string>('');
  const [entryTreasuryCode, setEntryTreasuryCode] = useState<string>('');
  const [entryPan, setEntryPan] = useState<string>('');
  const [entryGst, setEntryGst] = useState<string>('');
  const [entryTotalAmount, setEntryTotalAmount] = useState<string>('');
  const [entryDeductITax, setEntryDeductITax] = useState<boolean>(true);
  const [entryITaxPercent, setEntryITaxPercent] = useState<string>('1');
  const [entryDeductGst, setEntryDeductGst] = useState<boolean>(false);
  const [entryGstPercent, setEntryGstPercent] = useState<string>('2');
  const [editingEntryIndex, setEditingEntryIndex] = useState<number | null>(null);

  // Payee Sub-Vouchers / Bills Breakdown State
  const [entrySubVouchers, setEntrySubVouchers] = useState<MemoSubVoucher[]>([]);
  const [showSubVoucherSection, setShowSubVoucherSection] = useState<boolean>(false);
  const [subVoucherNoInput, setSubVoucherNoInput] = useState<string>('');
  const [subVoucherDescInput, setSubVoucherDescInput] = useState<string>('');
  const [subVoucherAmountInput, setSubVoucherAmountInput] = useState<string>('');
  const [editingSubVoucherId, setEditingSubVoucherId] = useState<string | null>(null);
  const [isMemoPayeeListFullScreen, setIsMemoPayeeListFullScreen] = useState<boolean>(false);
  const [memoPayeeSearchTerm, setMemoPayeeSearchTerm] = useState<string>('');
  const [showMemoPayeeDropdown, setShowMemoPayeeDropdown] = useState<boolean>(false);
  const [duplicatePayeeModalData, setDuplicatePayeeModalData] = useState<{ existingPayee: Payee; enteredName: string; enteredAccountNo: string } | null>(null);
  const [memoCorrectionModalData, setMemoCorrectionModalData] = useState<{ memo: MemoForFund; remarks: string } | null>(null);
  const [isSendingCorrection, setIsSendingCorrection] = useState<boolean>(false);
  const [memoPullBackModalData, setMemoPullBackModalData] = useState<{ memo: MemoForFund; remarks: string } | null>(null);
  const [isPullingBack, setIsPullingBack] = useState<boolean>(false);
  const [copiedFromMemoInfo, setCopiedFromMemoInfo] = useState<{ memoNo: string; originalId: string } | null>(null);

  // --- Auto-Save Drafts Engine (Offline & Power Outage Protection) ---
  const [lastAutoSaveMemoTime, setLastAutoSaveMemoTime] = useState<string | null>(null);
  const [lastAutoSaveExpenseTime, setLastAutoSaveExpenseTime] = useState<string | null>(null);
  const memoDraftLoadedRef = useRef(false);
  const expenseDraftLoadedRef = useRef(false);

  // 1. Memo for Fund: Auto-Restore Draft on Initial Load
  useEffect(() => {
    if (memoDraftLoadedRef.current) return;
    try {
      const savedMemoDraft = localStorage.getItem('rajgarh_draft_memo_fund');
      if (savedMemoDraft) {
        const d = JSON.parse(savedMemoDraft);
        if (d && !editingMemo) {
          if (d.memoNoInput) setMemoNoInput(d.memoNoInput);
          if (d.memoDateInput) setMemoDateInput(d.memoDateInput);
          if (d.memoMonthYearInput) setMemoMonthYearInput(d.memoMonthYearInput);
          if (d.memoSchemeIdInput) setMemoSchemeIdInput(d.memoSchemeIdInput);
          if (d.memoSectorIdInput) setMemoSectorIdInput(d.memoSectorIdInput);
          if (d.memoSoeIdInput) setMemoSoeIdInput(d.memoSoeIdInput);
          if (d.memoFromInput) setMemoFromInput(d.memoFromInput);
          if (d.memoToInput) setMemoToInput(d.memoToInput);
          if (Array.isArray(d.memoPayeeEntries) && d.memoPayeeEntries.length > 0) {
            setMemoPayeeEntries(d.memoPayeeEntries);
          }
          if (d.entryName) setEntryName(d.entryName);
          if (d.entryAddress) setEntryAddress(d.entryAddress);
          if (d.entryAccountNo) setEntryAccountNo(d.entryAccountNo);
          if (d.entryIfsc) setEntryIfsc(d.entryIfsc);
          if (d.entryTreasuryCode) setEntryTreasuryCode(d.entryTreasuryCode);
          if (d.entryPan) setEntryPan(d.entryPan);
          if (d.entryGst) setEntryGst(d.entryGst);
          if (d.entryTotalAmount) setEntryTotalAmount(d.entryTotalAmount);
          if (d.entrySubVouchers && Array.isArray(d.entrySubVouchers)) {
            setEntrySubVouchers(d.entrySubVouchers);
            if (d.entrySubVouchers.length > 0) setShowSubVoucherSection(true);
          }
          if (d.timestamp) {
            setLastAutoSaveMemoTime(new Date(d.timestamp).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }));
          }
        }
      }
    } catch (e) {
      console.warn("Could not restore memo draft from localStorage:", e);
    }
    memoDraftLoadedRef.current = true;
  }, [editingMemo]);

  // 2. Memo for Fund: Continuous Auto-Save on Any Input Change
  useEffect(() => {
    if (editingMemo) return; // Do not overwrite draft when editing existing saved memo
    const hasData = memoPayeeEntries.length > 0 || memoNoInput.trim() || entryName.trim() || entryAccountNo.trim() || entryTotalAmount || entrySubVouchers.length > 0;
    if (!hasData) return;

    const timer = setTimeout(() => {
      try {
        const draftObj = {
          memoNoInput,
          memoDateInput,
          memoMonthYearInput,
          memoSchemeIdInput,
          memoSectorIdInput,
          memoSoeIdInput,
          memoFromInput,
          memoToInput,
          memoPayeeEntries,
          entryName,
          entryAddress,
          entryAccountNo,
          entryIfsc,
          entryTreasuryCode,
          entryPan,
          entryGst,
          entryTotalAmount,
          entryDeductITax,
          entryITaxPercent,
          entryDeductGst,
          entryGstPercent,
          entrySubVouchers,
          subVoucherNoInput,
          subVoucherDescInput,
          subVoucherAmountInput,
          timestamp: Date.now()
        };
        localStorage.setItem('rajgarh_draft_memo_fund', JSON.stringify(draftObj));
        setLastAutoSaveMemoTime(new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }));
      } catch (err) {
        console.warn("Auto-save memo draft failed:", err);
      }
    }, 600);

    return () => clearTimeout(timer);
  }, [memoNoInput, memoDateInput, memoMonthYearInput, memoSchemeIdInput, memoSectorIdInput, memoSoeIdInput, memoFromInput, memoToInput, memoPayeeEntries, entryName, entryAddress, entryAccountNo, entryIfsc, entryTreasuryCode, entryPan, entryGst, entryTotalAmount, entryDeductITax, entryITaxPercent, entryDeductGst, entryGstPercent, entrySubVouchers, subVoucherNoInput, subVoucherDescInput, subVoucherAmountInput, editingMemo]);

  const handleClearMemoDraft = () => {
    localStorage.removeItem('rajgarh_draft_memo_fund');
    setLastAutoSaveMemoTime(null);
    handleResetMemoForm();
  };

  const [selectedMemoPayeeId, setSelectedMemoPayeeId] = useState<string>('');
  const [selectedMemoSchemeId, setSelectedMemoSchemeId] = useState<string>('');
  const [selectedMemoSectorId, setSelectedMemoSectorId] = useState<string>('');
  const [memoMonthYear, setMemoMonthYear] = useState<string>(() => {
    const now = new Date();
    return now.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  });
  const [memoRefNo, setMemoRefNo] = useState<string>('RFO/RAJGARH/Memo/2026-27/01');
  const [memoDate, setMemoDate] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [manualMemoAmount, setManualMemoAmount] = useState<string>('');
  const [selectedExpensesForBill, setSelectedExpensesForBill] = useState<string[]>([]);
  const [billFilters, setBillFilters] = useState({ billNo: '', rangeId: '', soeId: '', amount: '' });
  const [billExpFilters, setBillExpFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
  const [isBillFormFullScreen, setIsBillFormFullScreen] = useState(false);
  const [isTableFullScreen, setIsTableFullScreen] = useState(false);

  const [notifSearchTerm, setNotifSearchTerm] = useState('');
  const [notifPage, setNotifPage] = useState(1);
  const [notifItemsPerPage, setNotifItemsPerPage] = useState<number | 'All'>(25);
  const [notifCategoryFilter, setNotifCategoryFilter] = useState<'All' | 'general' | 'budget_allocation' | 'budget_distribution'>('All');

  const [budgetSearchTerm, setBudgetSearchTerm] = useState('');
  const [budgetPage, setBudgetPage] = useState(1);
  const [budgetItemsPerPage, setBudgetItemsPerPage] = useState<number | 'All'>(25);

  const [reportFilters, setReportFilters] = useState({ scheme: '', sector: '', activity: '', subActivity: '', range: '', soe: '', deductionType: '' });
  const [ledgerFilters, setLedgerFilters] = useState({ scheme: '', sector: '', activity: '', subActivity: '', range: '', soe: '' });
  const [showLedgerFilters, setShowLedgerFilters] = useState(false);
  const [ledgerSearchTerm, setLedgerSearchTerm] = useState('');
  const [reportSubTab, setReportSubTab] = useState('summary');
  const [reportSearchTerm, setReportSearchTerm] = useState('');
  const [reportPage, setReportPage] = useState(1);
  const [reportItemsPerPage, setReportItemsPerPage] = useState(25);
  const [isApprovalModalOpen, setIsApprovalModalOpen] = useState(false);
  const [selectedExpenseForApproval, setSelectedExpenseForApproval] = useState<Expense | null>(null);
  const [approvalStatus, setApprovalStatus] = useState<'approved' | 'rejected'>('approved');
  const [approvalReason, setApprovalReason] = useState('');
  const [isExpFilterExpanded, setIsExpFilterExpanded] = useState(false);
  const [isAllocFilterExpanded, setIsAllocFilterExpanded] = useState(false);
  const [isSoeFilterExpanded, setIsSoeFilterExpanded] = useState(false);
  const [openDropdown, setOpenDropdown] = useState<string | null>(null);

  const menuItems = useMemo<({ name: string; icon: React.ReactNode; children?: { name: string; icon: React.ReactNode }[] })[]>(() => {
    const adminItems = [
      { name: 'Dashboard', icon: <Home className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Notifications', icon: <Bell className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Financial Years', icon: <Calendar className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Ranges', icon: <Map className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Manage Scheme', 
        icon: <Landmark className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Schemes', icon: <TreePine className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Sectors', icon: <Shield className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Activities', icon: <Activity className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Sub-Activities', icon: <TrendingUp className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { name: 'SOE Heads', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Manage Budget', 
        icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Approved Budget', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Distributed Budget', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Allocations', icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Expenditures', icon: <TrendingDown className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Surrender', icon: <CornerUpLeft className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { name: 'Reconciliation', icon: <RefreshCcw className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Reports', 
        icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Ledger', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Reports', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { name: 'Audit Log', icon: <History className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Users', icon: <User className="w-3 h-3 sm:w-4 sm:h-4" /> }
    ];

    const deoItems = [
      { name: 'Dashboard', icon: <Home className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Notifications', icon: <Bell className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Manage Budget', 
        icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Approved Budget', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Distributed Budget', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Allocations', icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Expenditures', icon: <TrendingDown className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Surrender', icon: <CornerUpLeft className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { name: 'Reconciliation', icon: <RefreshCcw className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Reports', 
        icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Ledger', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Reports', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      ];

    const daItems = [
      { name: 'Dashboard', icon: <Home className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Notifications', icon: <Bell className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Manage Budget', 
        icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Approved Budget', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Distributed Budget', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Allocations', icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Expenditures', icon: <TrendingDown className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { name: 'Reconciliation', icon: <RefreshCcw className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Reports', 
        icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Ledger', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Reports', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      ];

    const otherItems = [
      { name: 'Dashboard', icon: <Home className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { name: 'Notifications', icon: <Bell className="w-3 h-3 sm:w-4 sm:h-4" /> },
      { 
        name: 'Manage Budget', 
        icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Approved Budget', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Distributed Budget', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Allocations', icon: <Wallet className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Expenditures', icon: <TrendingDown className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Surrender', icon: <CornerUpLeft className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      },
      { 
        name: 'Reports', 
        icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" />,
        children: [
          { name: 'Ledger', icon: <FileText className="w-3 h-3 sm:w-4 sm:h-4" /> },
          { name: 'Reports', icon: <FileBarChart className="w-3 h-3 sm:w-4 sm:h-4" /> }
        ]
      }
    ];

    if (userRole === 'admin') return adminItems;
    if (userRole === 'deo') return deoItems;
    if (userRole === 'DA' || userRole === 'approver') return daItems;
    return otherItems;
  }, [userRole]);

    // Auto-collapse filters and reset all filters on tab change
    useEffect(() => {
      setIsExpFilterExpanded(false);
      setIsAllocFilterExpanded(false);
      setIsSoeFilterExpanded(false);
      setShowReportFilters(false);
      
      // Reset all filters when switching main tabs
      setReportFilters({ scheme: '', sector: '', activity: '', subActivity: '', range: '', soe: '', deductionType: '' });
      setLedgerFilters({ scheme: '', sector: '', activity: '', subActivity: '', range: '', soe: '' });
      setReportSearchTerm('');
      setLedgerSearchTerm('');
      setReportPage(1);
      setExpFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '' });
      setAllocFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
      setSoeFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeName: '' });
      setExpDateRange({ start: '', end: '' });
      setSearchTerm('');
      setBillSearchTerm('');
      setPayeeSearchTerm('');
      setSelectedExpensesForBill([]);
      setBillExpFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
      setDashboardSearch('');
      setRangeSearch('');
      setSoeSearchTerm('');
      setSoeAbstractSearch('');
      setTrackerSearch('');
      setDeductionType('None');
      setReconSearchTerm('');
      setReconSchemeId('');
      setBudgetFileSelection({ schemeId: '', sectorId: '', rangeId: '' });
      setShowLedgerFilters(false);
    }, [activeTab]);

    // Reset report filters on sub-tab change
    useEffect(() => {
      setReportFilters({ scheme: '', sector: '', activity: '', subActivity: '', range: '', soe: '', deductionType: '' });
      setReportSearchTerm('');
      setSoeAbstractSearch('');
      setLedgerSearchTerm('');
      setReportPage(1);
      setShowReportFilters(false);
    }, [reportSubTab]);
  const [showReportFilters, setShowReportFilters] = useState(false);
  const [surrenderFormSelection, setSurrenderFormSelection] = useState<any>({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', soeId: '', rangeId: '' });
  const [approvedBudgetFiles, setApprovedBudgetFiles] = useState<BudgetFile[]>([]);
  const [distributedBudgetFiles, setDistributedBudgetFiles] = useState<BudgetFile[]>([]);
  
  // Per-type upload status
  const [uploadStatus, setUploadStatus] = useState<{
    [key: string]: {
      isUploading: boolean;
      progress: number;
      fileName: string;
      transferred: number;
      total: number;
      error: string | null;
    }
  }>({
    approved: { isUploading: false, progress: 0, fileName: '', transferred: 0, total: 0, error: null },
    distributed: { isUploading: false, progress: 0, fileName: '', transferred: 0, total: 0, error: null }
  });

  const [budgetFileSelection, setBudgetFileSelection] = useState({ schemeId: '', sectorId: '', rangeId: '' });
  const [expenseFormSelection, setExpenseFormSelection] = useState<any>({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', soeId: '', rangeId: '' });
  const [showSoeAbstract, setShowSoeAbstract] = useState(true);
  const [showDetailedReport, setShowDetailedReport] = useState(true);
  const [allocationFormFilters, setAllocationFormFilters] = useState({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', soeId: '', fundingSoeName: '', rangeId: '' });
  const [reconSchemeId, setReconSchemeId] = useState('');
  const [reconSearchTerm, setReconSearchTerm] = useState('');
  const [reconData, setReconData] = useState<any>({});
  const [selectedPayeesForExpense, setSelectedPayeesForExpense] = useState<{ payeeId: string; amount: string }[]>([]);
  const [currentSoeBalance, setCurrentSoeBalance] = useState<number | undefined>(undefined);

  useEffect(() => {
    if (activeTab !== 'Expenditures') {
      setCurrentSoeBalance(undefined);
      setSelectedPayeesForExpense([]);
    }
  }, [activeTab]);

  // --- Editing State ---
  const [editingItem, setEditingItem] = useState<{ type: string; item: any } | null>(null);

  // 3. Expenditure Entry: Auto-Restore Draft on Initial Load
  useEffect(() => {
    if (expenseDraftLoadedRef.current) return;
    try {
      const savedExpDraft = localStorage.getItem('rajgarh_draft_expenditure');
      if (savedExpDraft && !editingItem) {
        const d = JSON.parse(savedExpDraft);
        if (d.expenseAmount) setExpenseAmount(d.expenseAmount);
        if (d.expenseDate) setExpenseDate(d.expenseDate);
        if (d.expenseDescription) setExpenseDescription(d.expenseDescription);
        if (Array.isArray(d.selectedPayeesForExpense)) setSelectedPayeesForExpense(d.selectedPayeesForExpense);
        if (Array.isArray(d.selectedDeductions)) setSelectedDeductions(d.selectedDeductions);
        if (d.panNumber) setPanNumber(d.panNumber);
        if (d.gstNumber) setGstNumber(d.gstNumber);
        if (d.timestamp) {
          setLastAutoSaveExpenseTime(new Date(d.timestamp).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }));
        }
      }
    } catch (e) {
      console.warn("Could not restore expenditure draft:", e);
    }
    expenseDraftLoadedRef.current = true;
  }, [editingItem]);

  // 4. Expenditure Entry: Continuous Auto-Save on Any Input Change
  useEffect(() => {
    if (editingItem) return;
    const hasData = expenseAmount.trim() || expenseDescription.trim() || selectedPayeesForExpense.length > 0 || panNumber.trim() || gstNumber.trim();
    if (!hasData) return;

    const timer = setTimeout(() => {
      try {
        const draftObj = {
          expenseAmount,
          expenseDate,
          expenseDescription,
          selectedPayeesForExpense,
          selectedDeductions,
          panNumber,
          gstNumber,
          timestamp: Date.now()
        };
        localStorage.setItem('rajgarh_draft_expenditure', JSON.stringify(draftObj));
        setLastAutoSaveExpenseTime(new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }));
      } catch (err) {
        console.warn("Auto-save expense draft failed:", err);
      }
    }, 600);

    return () => clearTimeout(timer);
  }, [expenseAmount, expenseDate, expenseDescription, selectedPayeesForExpense, selectedDeductions, panNumber, gstNumber, editingItem]);

  const handleClearExpenseDraft = () => {
    localStorage.removeItem('rajgarh_draft_expenditure');
    setLastAutoSaveExpenseTime(null);
    setExpenseAmount('');
    setExpenseDescription('');
    setSelectedPayeesForExpense([]);
    setSelectedDeductions([]);
    setPanNumber('');
    setGstNumber('');
  };

  useEffect(() => {
    if (!editingItem) {
      setExpenseAmount('');
      setSelectedPayeesForExpense([]);
      // Removed setCurrentSoeBalance(undefined) to prevent race conditions with CascadingDropdowns
    }
  }, [editingItem, expenseFormSelection.schemeId, expenseFormSelection.sectorId, expenseFormSelection.activityId, expenseFormSelection.subActivityId]);
  const [viewingSoeExp, setViewingSoeExp] = useState<{ soeId: string; soeName: string; hierarchy: string } | null>(null);
  const [showScrollTop, setShowScrollTop] = useState(false);

  useEffect(() => {
    const handleScroll = () => {
      setShowScrollTop(window.scrollY > 300);
    };
    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  // --- Auth & Role Check ---
  useEffect(() => {
    // Set persistence to local - will keep user logged in even if tab is closed
    // But we will manually check for the 10-minute grace period
    const initAuth = async () => {
      try {
        await setPersistence(auth, browserLocalPersistence);
      } catch (err) {
        console.error("Persistence error:", err);
      }
    };
    initAuth();

    // --- Grace Period Check ---
    const lastClosedTime = localStorage.getItem('lastClosedTime');
    if (lastClosedTime) {
      const timeDiff = Date.now() - parseInt(lastClosedTime);
      if (timeDiff > 10 * 60 * 1000) { // 10 minutes
        signOut(auth).catch(err => console.error("Sign out error:", err));
        localStorage.removeItem('lastClosedTime');
      }
    }

    const handleUnload = () => {
      localStorage.setItem('lastClosedTime', Date.now().toString());
    };

    window.addEventListener('beforeunload', handleUnload);

    async function testConnection() {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration. ");
        }
      }
    }
    testConnection();

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser) {
        setLoading(true);
        const email = currentUser.email?.toLowerCase();
        
        try {
          // Parallelize initial data fetching for faster load
          const [userDoc, locksSnap, rangesSnap] = await Promise.all([
            getDoc(doc(db, 'users', currentUser.uid)),
            getDocs(query(
              collection(db, 'featureLocks'),
              where('feature', '==', 'Access'),
              where('isLocked', '==', true)
            )),
            getDocs(collection(db, 'ranges'))
          ]);

          let userData = userDoc.exists() ? userDoc.data() as AppUser : null;

          // Hardcode roles for specific emails
          if (email === 'admin@rajgarhforest.app' || email === 'sharmaanuj860@gmail.com') {
            if (!userData || userData.role !== 'admin') {
              userData = { ...(userData || {}), id: currentUser.uid, email: currentUser.email!, role: 'admin', maxSessions: 999999, activeSessions: userData?.activeSessions || [] };
              await setDoc(doc(db, 'users', currentUser.uid), userData, { merge: true });
            }
          } else if (email === 'da123@rajgarhforest.app') {
            if (!userData) {
              userData = { id: currentUser.uid, email: currentUser.email!, role: 'deo', maxSessions: 999999, activeSessions: [] };
              await setDoc(doc(db, 'users', currentUser.uid), userData, { merge: true });
            }
          } else if (email === 'da789@rajgarhforest.app') {
            if (!userData) {
              userData = { id: currentUser.uid, email: currentUser.email!, role: 'approver', maxSessions: 999999, activeSessions: [] };
              await setDoc(doc(db, 'users', currentUser.uid), userData, { merge: true });
            }
          }

          if (userData) {
            // Check if user is disabled individually
            if (userData.isDisabled) {
              showAlert("Your account has been disabled by the administrator. Please contact support.");
              await signOut(auth);
              setLoading(false);
              return;
            }

            // Check if user's role or range is disabled via featureLocks
            if (userData.role !== 'admin') {
              const activeAccessLocks = locksSnap.docs.map(d => d.data() as FeatureLock);
              
              // Check individual user lock via featureLocks (legacy or fallback)
              if (activeAccessLocks.some(l => l.target === userData!.id)) {
                showAlert(`Access for your account has been disabled by the administrator.`);
                await signOut(auth);
                setLoading(false);
                return;
              }
              
              // Check role lock
              if (activeAccessLocks.some(l => l.target === userData!.role)) {
                showAlert(`Access for the ${userData.role.toUpperCase()} role has been disabled by the administrator.`);
                await signOut(auth);
                setLoading(false);
                return;
              }

              // Check range lock (if role is a range name)
              if (['Sarahan', 'Narag', 'Habban', 'Division', 'Rajgarh'].includes(userData.role)) {
                const userRange = rangesSnap.docs.find(d => d.data().name === userData!.role);
                if (userRange && activeAccessLocks.some(l => l.target === userRange.id)) {
                  showAlert(`Access for the ${userData.role} range has been disabled by the administrator.`);
                  await signOut(auth);
                  setLoading(false);
                  return;
                }
              }
            }

            // Session Validation
            const activeSessions = userData.activeSessions || [];
            if (!activeSessions.includes(sessionId)) {
              // Default to 999999 (unlimited) if not specified, but admins are always unlimited
              const maxSessions = userData.role === 'admin' ? 999999 : (userData.maxSessions || 999999);
              if (activeSessions.length >= maxSessions) {
                showAlert(`Maximum concurrent sessions (${maxSessions}) reached for this account. Please logout from other devices.`);
                await signOut(auth);
                setLoading(false);
                return;
              }
              // Add current session
              const updatedSessions = [...activeSessions, sessionId];
              try {
                await updateDoc(doc(db, 'users', currentUser.uid), { activeSessions: updatedSessions });
              } catch (err) {
                handleFirestoreError(err, OperationType.UPDATE, `users/${currentUser.uid}`);
              }
            }

            setUser(currentUser);
            setUserRole(userData.role);
            setActiveTab('Dashboard');
            // Clear last closed time as user is now active
            localStorage.removeItem('lastClosedTime');
          } else {
            // If first user ever, make admin
            try {
              const usersSnap = await getDocs(collection(db, 'users'));
              if (usersSnap.empty) {
                const newRole = 'admin';
                const newUserData = { email: currentUser.email, role: newRole, maxSessions: 999999, activeSessions: [sessionId] };
                await setDoc(doc(db, 'users', currentUser.uid), newUserData);
                setUser(currentUser);
                setUserRole(newRole);
              } else {
                setUserRole(null);
              }
            } catch (e) {
              console.warn("Could not check for first user (likely permission denied), assuming not first user.");
              setUserRole(null);
            }
          }
        } catch (error) {
          handleFirestoreError(error, OperationType.GET, `users/${currentUser.uid}`);
        }
      } else {
        setUser(null);
        setUserRole(null);
      }
      setLoading(false);
    });
    return () => {
      unsubscribe();
      window.removeEventListener('beforeunload', handleUnload);
    };
  }, []);

  // --- Session Expiry Logic (Activity Timer) ---
  useEffect(() => {
    if (!user || !userRole) return;

    let timeoutId: any;

    const resetTimer = () => {
      if (timeoutId) clearTimeout(timeoutId);
      // Admin: 20 mins, Others: 15 mins
      const timeoutMinutes = userRole === 'admin' ? 20 : 15;
      timeoutId = setTimeout(() => {
        showAlert(`Session expired due to inactivity (${timeoutMinutes} mins). Please login again.`);
        handleLogout();
      }, timeoutMinutes * 60 * 1000);
    };

    // Initial start
    resetTimer();

    // Listen for activity
    const activityEvents = ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'];
    activityEvents.forEach(event => {
      window.addEventListener(event, resetTimer);
    });

    return () => {
      if (timeoutId) clearTimeout(timeoutId);
      activityEvents.forEach(event => {
        window.removeEventListener(event, resetTimer);
      });
    };
  }, [user, userRole]);

  // --- Real-time Data Sync (Master Data) ---
  useEffect(() => {
    if (!user || !userRole) return;

    const unsubFys = onSnapshot(collection(db, 'financialYears'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as FinancialYear));
      setFys(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'financialYears'));

    const unsubRanges = onSnapshot(collection(db, 'ranges'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Range));
      setRanges(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'ranges'));

    const unsubSchemes = onSnapshot(collection(db, 'schemes'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Scheme));
      setSchemes(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'schemes'));

    const unsubSectors = onSnapshot(collection(db, 'sectors'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Sector));
      setSectors(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'sectors'));

    const unsubActivities = onSnapshot(collection(db, 'activities'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as ActivityItem));
      setActivities(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'activities'));

    const unsubSubActivities = onSnapshot(collection(db, 'subActivities'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as SubActivity));
      setSubActivities(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'subActivities'));

    const unsubUsers = userRole === 'admin' ? onSnapshot(collection(db, 'users'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as AppUser));
      setUsers(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'users')) : () => {};

    // Listen to current user's document for real-time disablement (for non-admins)
    const unsubCurrentUser = user && userRole !== 'admin' ? onSnapshot(doc(db, 'users', user.uid), (snap) => {
      const userData = snap.data() as AppUser;
      if (userData?.isDisabled) {
        showAlert("Your account has been disabled by the administrator. Please contact support.");
        handleLogout();
      }
    }, (error) => {
      // Ignore permission errors if user is already disabled and doc becomes unreadable
      if (!error.message.includes('insufficient permissions')) {
        handleFirestoreError(error, OperationType.GET, `users/${user.uid}`);
      }
    }) : () => {};

    const unsubLocks = onSnapshot(collection(db, 'featureLocks'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as FeatureLock));
      setFeatureLocks(data);
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'featureLocks'));

    const unsubPayees = onSnapshot(collection(db, 'payees'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Payee));
      setPayees(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'payees'));

    const unsubNotifications = onSnapshot(collection(db, 'notifications'), (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Notification));
      setNotifications(data.sort((a: any, b: any) => (b.createdAt || 0) - (a.createdAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'notifications'));

    return () => {
      unsubFys(); unsubRanges(); unsubSchemes(); unsubSectors(); unsubActivities();
      unsubSubActivities(); unsubUsers(); unsubCurrentUser(); unsubPayees(); unsubLocks();
      unsubNotifications();
    };
  }, [user, userRole]);

  // --- Real-time Data Sync (Transactional Data filtered by FY) ---
  useEffect(() => {
    if (!user || !userRole || !selectedFY) return;

    setIsSoesLoaded(false);

    const activeFy = fys.find(f => f.name === selectedFY || f.id === selectedFY);
    const fyQueryValues = activeFy ? Array.from(new Set([activeFy.id, activeFy.name])) : [selectedFY];
    if (fyQueryValues.length === 0 || (fyQueryValues.length === 1 && !fyQueryValues[0])) return;

    const soesQuery = query(
      collection(db, 'soeHeads'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubSoes = onSnapshot(soesQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as SOE));
      setSoes(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
      setIsSoesLoaded(true);
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'soeHeads'));

    const allocsQuery = query(
      collection(db, 'allocations'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubAllocations = onSnapshot(allocsQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Allocation));
      console.log('Fetched allocations for FY:', fyQueryValues, 'Count:', data.length);
      setAllocations(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'allocations'));

    const expensesQuery = query(
      collection(db, 'expenditures'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubExpenses = onSnapshot(expensesQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Expense));
      setExpenses(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'expenditures'));

    const billsQuery = query(
      collection(db, 'bills'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubBills = onSnapshot(billsQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Bill));
      setBills(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'bills'));

    const surrendersQuery = query(
      collection(db, 'surrenders'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubSurrenders = onSnapshot(surrendersQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as Surrender));
      setSurrenders(data.sort((a: any, b: any) => (b.updatedAt || 0) - (a.updatedAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'surrenders'));

    const approvedFilesQuery = query(
      collection(db, 'approvedBudgetFiles'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubApprovedFiles = onSnapshot(approvedFilesQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as BudgetFile));
      setApprovedBudgetFiles(data.sort((a, b) => b.uploadedAt - a.uploadedAt));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'approvedBudgetFiles'));

    const distributedFilesQuery = query(
      collection(db, 'distributedBudgetFiles'), 
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubDistributedFiles = onSnapshot(distributedFilesQuery, (snap) => {
      const data = snap.docs.map(d => ({ ...d.data(), id: d.id } as BudgetFile));
      setDistributedBudgetFiles(data.sort((a, b) => b.uploadedAt - a.uploadedAt));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'distributedBudgetFiles'));

    const memosQuery = query(
      collection(db, 'memos'),
      or(where('financialYear', 'in', fyQueryValues), where('fyId', 'in', fyQueryValues))
    );
    const unsubMemos = onSnapshot(memosQuery, (snap) => {
      const data = snap.docs.map(d => {
        const raw = { ...d.data(), id: d.id } as MemoForFund;
        const roundedPayeeEntries = (raw.payeeEntries || []).map(entry => {
          const tot = Math.round(Number(entry.totalAmount) || 0);
          const itax = Math.round(Number(entry.iTaxAmount) || 0);
          const gst = Math.round(Number(entry.gstAmount) || 0);
          const net = Math.round(Number(entry.netRtgsAmount) || (tot - itax - gst));
          return {
            ...entry,
            totalAmount: tot,
            iTaxAmount: itax,
            gstAmount: gst,
            netRtgsAmount: net,
          };
        });

        const totalGross = roundedPayeeEntries.length > 0
          ? roundedPayeeEntries.reduce((sum, e) => sum + e.totalAmount, 0)
          : Math.round(Number(raw.totalAmount) || 0);
        const totalITax = roundedPayeeEntries.length > 0
          ? roundedPayeeEntries.reduce((sum, e) => sum + e.iTaxAmount, 0)
          : Math.round(Number(raw.totalITax) || 0);
        const totalGst = roundedPayeeEntries.length > 0
          ? roundedPayeeEntries.reduce((sum, e) => sum + (e.gstAmount || 0), 0)
          : Math.round(Number(raw.totalGst) || 0);
        const totalNetRtgs = roundedPayeeEntries.length > 0
          ? roundedPayeeEntries.reduce((sum, e) => sum + e.netRtgsAmount, 0)
          : Math.round(Number(raw.totalNetRtgs) || (totalGross - totalITax - totalGst));

        return {
          ...raw,
          totalAmount: totalGross,
          totalITax: totalITax,
          totalGst: totalGst,
          totalNetRtgs: totalNetRtgs,
          payeeEntries: roundedPayeeEntries,
        };
      });
      setMemos(data.sort((a: any, b: any) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0)));
    }, (error) => handleFirestoreError(error, OperationType.LIST, 'memos'));

    return () => {
      unsubSoes(); unsubAllocations(); unsubExpenses(); unsubBills(); unsubSurrenders(); unsubApprovedFiles(); unsubDistributedFiles(); unsubMemos();
    };
  }, [user, userRole, selectedFY, fys]);

  // --- Silent Auto-Seeding for 2025-26 ---
  useEffect(() => {
    if (!user || !userRole || selectedFY !== '2025-26' || !isSoesLoaded) return;

    if (soes.length === 0 && !hasSeeded.current) {
      hasSeeded.current = true;
      const seedData = async () => {
        try {
          console.log("Silently seeding emergency restore data for FY 2025-26...");
          
          const fy = fys.find(f => f.name === '2025-26');
          const fyId = fy ? fy.id : '2025-26';

          let localSchemes = [...schemes];
          let localSectors = [...sectors];
          let localActivities = [...activities];
          let localSubActivities = [...subActivities];

          for (const item of emergencyRestoreData) {
            // Determine IDs based on hierarchy string
            let schemeId = null;
            let sectorId = null;
            let activityId = null;
            let subActivityId = null;

            const parts = item.hierarchy.split(' -> ');
            
            if (parts.length > 0) {
              let sch = localSchemes.find(s => s.name === parts[0]);
              if (!sch) {
                const docRef = await addDoc(collection(db, 'schemes'), { name: parts[0], createdAt: Date.now(), updatedAt: Date.now() });
                sch = { id: docRef.id, name: parts[0] };
                localSchemes.push(sch);
              }
              schemeId = sch.id;
            }
            if (parts.length > 1) {
              let sec = localSectors.find(s => s.name === parts[1] && s.schemeId === schemeId);
              if (!sec) {
                const docRef = await addDoc(collection(db, 'sectors'), { name: parts[1], schemeId, createdAt: Date.now(), updatedAt: Date.now() });
                sec = { id: docRef.id, name: parts[1], schemeId };
                localSectors.push(sec);
              }
              sectorId = sec.id;
            }
            if (parts.length > 2) {
              let act = localActivities.find(a => a.name === parts[2] && a.sectorId === sectorId);
              if (!act) {
                const docRef = await addDoc(collection(db, 'activities'), { name: parts[2], sectorId, createdAt: Date.now(), updatedAt: Date.now() });
                act = { id: docRef.id, name: parts[2], sectorId };
                localActivities.push(act);
              }
              activityId = act.id;
            }
            if (parts.length > 3) {
              let sa = localSubActivities.find(s => s.name === parts[3] && s.activityId === activityId);
              if (!sa) {
                const docRef = await addDoc(collection(db, 'subActivities'), { name: parts[3], activityId, createdAt: Date.now(), updatedAt: Date.now() });
                sa = { id: docRef.id, name: parts[3], activityId };
                localSubActivities.push(sa);
              }
              subActivityId = sa.id;
            }

            await addDoc(collection(db, 'soeHeads'), {
              name: item.soeName,
              schemeId,
              sectorId,
              activityId,
              subActivityId,
              approvedBudget: item.approvedBudget,
              approvedBudgetAmount: item.approvedBudget,
              receivedInTry: item.receivedInTry,
              receivedInTryAmount: item.receivedInTry,
              tryAmount: item.receivedInTry,
              financialYear: fyId,
              createdAt: Date.now(),
              updatedAt: Date.now()
            });
          }
          console.log("Silent seeding completed successfully.");
        } catch (error) {
          console.error("Failed to silently seed data:", error);
          handleFirestoreError(error, OperationType.CREATE, 'soeHeads');
        }
      };

      seedData();
    }
  }, [user, userRole, selectedFY, isSoesLoaded, soes.length, schemes, sectors, activities, subActivities, fys]);

  const [loginEmail, setLoginEmail] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [isRefreshingApp, setIsRefreshingApp] = useState(false);

  const handleForceAppRefresh = async () => {
    setIsRefreshingApp(true);
    try {
      // 1. Clear all browser and service worker caches
      if ('caches' in window) {
        const cacheKeys = await caches.keys();
        await Promise.all(cacheKeys.map(key => caches.delete(key)));
      }

      // 2. Unregister or update all active Service Workers
      if ('serviceWorker' in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        for (const reg of registrations) {
          await reg.update();
          if (reg.waiting) {
            reg.waiting.postMessage({ type: 'SKIP_WAITING' });
          }
        }
      }
    } catch (err) {
      console.error('Failed to clear cache during refresh:', err);
    } finally {
      // 3. Force hard reload bypassing cache with a timestamp query param
      const url = new URL(window.location.href);
      url.searchParams.set('_t', Date.now().toString());
      window.location.replace(url.toString());
    }
  };

  const handleLogin = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    setLoginError('');
    if (loginFY) {
      setSelectedFY(loginFY);
    }
    if (!navigator.onLine) {
      setLoginError('No internet connection. Please check your network.');
      return;
    }
    setLoading(true);
    try {
      const emailToUseTrimmed = loginEmail.trim();

      if (!emailToUseTrimmed) {
        setLoginError('Please enter your User ID or Email.');
        setLoading(false);
        return;
      }
      if (!loginPassword) {
        setLoginError('Please enter your password.');
        setLoading(false);
        return;
      }

      let emailToUse = emailToUseTrimmed;
      if (!emailToUse.includes('@')) {
        emailToUse = `${emailToUse}@rajgarhforest.app`;
      }
      // Ensure persistence is set before login
      await setPersistence(auth, browserLocalPersistence);
      await signInWithEmailAndPassword(auth, emailToUse, loginPassword);
      logAuditAction('User Login / Access', `User logged in from ${clientNetwork?.locationString || 'Himachal Pradesh, India'}`);
    } catch (error: any) {
      console.error('Auth error:', error);
      if (error.code === 'auth/network-request-failed') {
        setLoginError('Network request failed. This may be due to a poor connection or browser restrictions. Please refresh and try again.');
      } else if (error.code === 'auth/too-many-requests') {
        setLoginError('Too many failed login attempts. Please try again later or reset your password.');
      } else if (error.code === 'auth/operation-not-allowed') {
        setLoginError('Email/Password authentication is not enabled in your Firebase project. Please go to the Firebase Console -> Authentication -> Sign-in method, and enable "Email/Password".');
      } else if (error.code === 'auth/user-not-found' || error.code === 'auth/invalid-credential' || error.code === 'auth/wrong-password') {
        setLoginError('Invalid User ID or Password. Please try again.');
      } else if (error.code === 'auth/invalid-email') {
        setLoginError('The User ID or Email format is invalid.');
      } else {
        setLoginError(error.message || 'Authentication failed. Please check your credentials.');
      }
      setLoading(false);
    }
  };


  // --- Derived Data / Helpers ---
  const currentSchemes = schemes;
  const currentSectors = sectors;
  const currentActivities = activities;
  const currentSubActivities = subActivities;
  const userRangeId = useMemo(() => {
    console.log('Calculating userRangeId. userRole:', userRole);
    if (userRole && ['Sarahan', 'Narag', 'Habban', 'Division', 'Rajgarh'].includes(userRole)) {
      const r = ranges.find(r => r.name === userRole);
      console.log('Found range for role:', r);
      return r?.id;
    }
    return null;
  }, [userRole, ranges]);

  const filteredPayeesList = useMemo(() => {
    const seen: Record<string, boolean> = {};
    const uniquePayees: Payee[] = [];
    payees.forEach(p => {
      if (p.id && !seen[p.id]) {
        seen[p.id] = true;
        uniquePayees.push(p);
      }
    });
    if (!userRole) return uniquePayees;
    const roleLower = userRole.toLowerCase();
    if (roleLower === 'admin' || roleLower === 'deo' || roleLower === 'da' || roleLower === 'approver') {
      return uniquePayees;
    }
    if (userRangeId) {
      return uniquePayees.filter(p => 
        !p.rangeId || 
        p.rangeId === userRangeId || 
        p.createdBy === user?.uid ||
        (p.mappedUserIds && p.mappedUserIds.includes(user?.uid)) ||
        (p.mappedRangeIds && p.mappedRangeIds.includes(userRangeId))
      );
    }
    return uniquePayees.filter(p => 
      !p.rangeId || 
      p.createdBy === user?.uid || 
      (p.mappedUserIds && p.mappedUserIds.includes(user?.uid))
    );
  }, [payees, userRole, userRangeId, user?.uid]);

  const isAdmin = () => userRole === 'admin' || user?.email?.toLowerCase() === 'admin@rajgarhforest.app' || user?.email?.toLowerCase() === 'sharmaanuj860@gmail.com';
  const isDEO = () => userRole === 'deo';

  // --- Real-time Access Control Enforcement ---
  useEffect(() => {
    if (!user || userRole === 'admin' || featureLocks.length === 0) return;

    const accessLock = featureLocks.find(l => 
      l.feature === 'Access' && 
      l.isLocked && 
      (l.target === userRole || (userRangeId && l.target === userRangeId) || l.target === user.uid)
    );

    if (accessLock) {
      showAlert(`Access for your ${accessLock.target === userRole ? 'role' : (accessLock.target === user.uid ? 'account' : 'range')} has been disabled by the administrator.`);
      handleLogout();
    }
  }, [user, userRole, userRangeId, featureLocks]);

  const currentSoes = useMemo(() => {
    let filtered = soes.filter(s => ALLOWED_SOES.includes(s.name || 'Provisional'));
    
    // SOE Heads don't have rangeId directly, so we don't filter by userRangeId or soeFilters.rangeId
    
    if (soeFilters.schemeId) {
      filtered = filtered.filter(s => s.schemeId === soeFilters.schemeId);
    }
    if (soeFilters.sectorId) {
      filtered = filtered.filter(s => s.sectorId === soeFilters.sectorId);
    }
    if (soeFilters.activityId) {
      filtered = filtered.filter(s => s.activityId === soeFilters.activityId);
    }
    if (soeFilters.subActivityId) {
      filtered = filtered.filter(s => s.subActivityId === soeFilters.subActivityId);
    }
    if (soeFilters.soeName) {
      filtered = filtered.filter(s => s.name === soeFilters.soeName);
    }
    
    return filtered;
  }, [soes, soeFilters]);

  const baseAllocations = useMemo(() => {
    let filtered = allocations;
    if (userRangeId) {
      filtered = filtered.filter(a => a.rangeId === userRangeId);
    }
    return filtered;
  }, [allocations, userRangeId]);

  const baseExpenses = useMemo(() => {
    let filtered = expenses;
    if (userRangeId) {
      const userAllocIds = baseAllocations.map(a => a.id);
      filtered = filtered.filter(e => userAllocIds.includes(e.allocationId));
    }
    return filtered;
  }, [expenses, baseAllocations, userRangeId]);

  const currentAllocations = useMemo(() => {
    let filtered = baseAllocations;
    
    if (userRangeId) {
      filtered = filtered.filter(a => a.rangeId === userRangeId);
    }

    if (allocFilters.schemeId) {
      filtered = filtered.filter(a => a.schemeId === allocFilters.schemeId);
    }
    if (allocFilters.sectorId) {
      filtered = filtered.filter(a => a.sectorId === allocFilters.sectorId);
    }
    if (allocFilters.activityId) {
      filtered = filtered.filter(a => a.activityId === allocFilters.activityId);
    }
    if (allocFilters.subActivityId) {
      filtered = filtered.filter(a => a.subActivityId === allocFilters.subActivityId);
    }
    if (allocFilters.rangeId) {
      filtered = filtered.filter(a => a.rangeId === allocFilters.rangeId);
    }
    if (allocFilters.soeId) {
      filtered = filtered.filter(a => a.fundedSOEs && a.fundedSOEs.some(f => f.soeId === allocFilters.soeId));
    }

    // Filter out "Division" allocations with 0 amount to avoid cluttering as requested
    filtered = filtered.filter(a => {
      const r = ranges.find(range => range.id === a.rangeId);
      const isDivision = r?.name === 'Division' || r?.name === 'Rajgarh Forest Division';
      if (isDivision && a.amount === 0) return false;
      return true;
    });

    console.log('currentAllocations count:', filtered.length);
    return filtered;
  }, [baseAllocations, allocFilters, userRangeId, ranges]);

  const currentExpenses = useMemo(() => {
    let filtered = expenses;
    
    if (userRangeId) {
      const userAllocIds = currentAllocations.map(a => a.id);
      filtered = filtered.filter(e => userAllocIds.includes(e.allocationId));
    }

    if (expDateRange.start) filtered = filtered.filter(e => e.date >= expDateRange.start);
    if (expDateRange.end) filtered = filtered.filter(e => e.date <= expDateRange.end);
    
    if (expFilters.schemeId) {
      filtered = filtered.filter(e => {
        const alloc = allocations.find(a => a.id === e.allocationId);
        return alloc?.schemeId === expFilters.schemeId;
      });
    }
    
    if (expFilters.sectorId) {
      filtered = filtered.filter(e => {
        const alloc = allocations.find(a => a.id === e.allocationId);
        return alloc?.sectorId === expFilters.sectorId;
      });
    }
    
    if (expFilters.activityId) {
      filtered = filtered.filter(e => {
        const alloc = allocations.find(a => a.id === e.allocationId);
        return alloc?.activityId === expFilters.activityId;
      });
    }

    if (expFilters.subActivityId) {
      filtered = filtered.filter(e => {
        const alloc = allocations.find(a => a.id === e.allocationId);
        return alloc?.subActivityId === expFilters.subActivityId;
      });
    }

    if (expFilters.rangeId) {
      filtered = filtered.filter(e => {
        const alloc = allocations.find(a => a.id === e.allocationId);
        return alloc?.rangeId === expFilters.rangeId;
      });
    }

    return filtered;
  }, [expenses, currentAllocations, expDateRange, expFilters, allocations, userRangeId]);

  // Memo for Fund Computed Values
  const selectedPayeeObj = useMemo(() => {
    return payees.find(p => p.id === selectedMemoPayeeId);
  }, [payees, selectedMemoPayeeId]);

  const selectedSchemeObj = useMemo(() => {
    return currentSchemes.find(s => s.id === selectedMemoSchemeId);
  }, [currentSchemes, selectedMemoSchemeId]);

  const selectedSectorObj = useMemo(() => {
    return currentSectors.find(s => s.id === selectedMemoSectorId);
  }, [currentSectors, selectedMemoSectorId]);

  const isRangeRole = Boolean(userRole && ['Sarahan', 'Narag', 'Habban', 'Rajgarh', 'Division'].some(r => r.toLowerCase() === userRole.toLowerCase()));

  const userRangeName = useMemo(() => {
    const r = ranges.find(r => r.id === userRangeId);
    if (r) return r.name;
    if (userRole === 'Sarahan') return 'Sarahan Range';
    if (userRole === 'Narag') return 'Narag Range';
    if (userRole === 'Habban') return 'Habban Range';
    if (userRole === 'Rajgarh') return 'Rajgarh Range';
    if (userRole === 'Division') return 'Division Office';
    return 'Rajgarh Forest Division';
  }, [ranges, userRangeId, userRole]);

  const filteredMemos = useMemo(() => {
    return memos.filter(m => {
      if (isRangeRole) {
        const userRoleLower = userRole!.toLowerCase();
        const codeMap: Record<string, string> = {
          sarahan: 'SRH',
          narag: 'NRG',
          habban: 'HBN',
          rajgarh: 'RJG',
          division: 'DIV',
        };
        const code = codeMap[userRoleLower];

        const matchRangeId = userRangeId && m.rangeId === userRangeId;
        const matchRangeName = m.rangeName && (
          m.rangeName.toLowerCase().includes(userRoleLower) ||
          (userRangeName && m.rangeName.toLowerCase().includes(userRangeName.toLowerCase()))
        );
        const matchCreatedByRole = m.createdByRole && m.createdByRole.toLowerCase() === userRoleLower;
        const matchCreatedBy = m.createdBy && user?.uid && m.createdBy === user.uid;
        const matchMemoCode = code && m.memoNo && m.memoNo.toUpperCase().includes(`/${code}/`);

        if (!matchRangeId && !matchRangeName && !matchCreatedByRole && !matchCreatedBy && !matchMemoCode) {
          return false;
        }
      }
      return true;
    });
  }, [memos, isRangeRole, userRole, userRangeId, userRangeName, user?.uid]);

  const isAuthorizedUserOrAdmin = useMemo(() => {
    if (!user) return false;
    if (userRole === 'admin' || userRole === 'deo' || userRole === 'approver' || userRole === 'DA' || userRole === 'Division') return true;
    const email = user.email?.toLowerCase() || '';
    return email === 'sharmaanuj860@gmail.com' || email === 'admin@rajgarhforest.app' || email === 'da123@rajgarhforest.app' || email === 'da789@rajgarhforest.app';
  }, [user, userRole]);

  const isRangeUser = useMemo(() => {
    if (userRole && ['Sarahan', 'Narag', 'Habban', 'Rajgarh'].some(r => r.toLowerCase() === userRole.toLowerCase())) return true;
    if (userRole === 'admin' || userRole === 'deo' || userRole === 'approver' || userRole === 'DA') return false;
    const email = user?.email?.toLowerCase() || '';
    if (email === 'sharmaanuj860@gmail.com' || email === 'admin@rajgarhforest.app' || email === 'da123@rajgarhforest.app' || email === 'da789@rajgarhforest.app') return false;
    return true;
  }, [userRole, user?.email]);

  const filteredMemosForSync = useMemo(() => {
    return filteredMemos.filter(m => {
      // Strictly ONLY submitted memos can be synced / incurred as expenditure
      if (m.status !== 'submitted') {
        return false;
      }
      if (selectedFY && m.financialYear && m.financialYear !== selectedFY && m.fyId !== selectedFY) {
        return false;
      }
      if (memoSearchTerm.trim()) {
        const q = memoSearchTerm.toLowerCase();
        const matchNo = m.memoNo?.toLowerCase().includes(q);
        const matchMonth = m.monthYear?.toLowerCase().includes(q);
        const matchScheme = m.schemeName?.toLowerCase().includes(q);
        const matchPayee = m.payeeEntries?.some(p => p.name?.toLowerCase().includes(q) || p.accountNumber?.includes(q));
        return matchNo || matchMonth || matchScheme || matchPayee;
      }
      return true;
    });
  }, [filteredMemos, selectedFY, memoSearchTerm]);

  const memoFilteredExpenses = useMemo(() => {
    if (!selectedMemoPayeeId) return [];
    return currentExpenses.filter(e => {
      const alloc = allocations.find(a => a.id === e.allocationId);
      const eSchemeId = e.schemeId || alloc?.schemeId;
      const eSectorId = e.sectorId || alloc?.sectorId;
      const payeeMatch = e.payeeId === selectedMemoPayeeId || (selectedPayeeObj && e.payeeName === selectedPayeeObj.name);
      const schemeMatch = !selectedMemoSchemeId || eSchemeId === selectedMemoSchemeId;
      const sectorMatch = !selectedMemoSectorId || eSectorId === selectedMemoSectorId;
      return payeeMatch && schemeMatch && sectorMatch;
    });
  }, [currentExpenses, selectedMemoPayeeId, selectedMemoSchemeId, selectedMemoSectorId, selectedPayeeObj, allocations]);

  const memoTotalAmount = useMemo(() => {
    if (manualMemoAmount !== '' && !isNaN(Number(manualMemoAmount))) {
      return Number(manualMemoAmount);
    }
    return memoFilteredExpenses.reduce((sum, e) => sum + (Number(e.amount) || 0), 0);
  }, [memoFilteredExpenses, manualMemoAmount]);

  const comprehensiveReportData = useMemo(() => {
    return baseAllocations.map(a => {
      const allocExpenses = baseExpenses.filter(e => e.allocationId === a.id && e.status !== 'rejected');
      const totalExp = allocExpenses.reduce((sum, e) => sum + e.amount, 0);
      const range = ranges.find(r => r.id === a.rangeId);
      const scheme = schemes.find(s => s.id === a.schemeId);
      const sector = sectors.find(s => s.id === a.sectorId);
      const activity = activities.find(act => act.id === a.activityId);
      const subActivity = subActivities.find(sa => sa.id === a.subActivityId);

      const soeBreakdown = soes.reduce((acc: any, soe) => {
        const soeAlloc = a.fundedSOEs?.find(f => f.soeId === soe.id)?.amount || 0;
        const soeExp = allocExpenses
          .filter(e => e.soeId === soe.id)
          .reduce((sum, e) => sum + e.amount, 0);
        acc[soe.name] = { alloc: soeAlloc, exp: soeExp };
        return acc;
      }, {});

      return {
        id: a.id,
        range: range?.name === 'Rajgarh Forest Division' ? 'Division' : (range?.name || 'Unknown'),
        scheme: scheme?.name || 'Unknown',
        sector: sector?.name || 'Unknown',
        activity: activity?.name || 'Unknown',
        subActivity: subActivity?.name || 'Unknown',
        totalAlloc: a.amount,
        totalExp,
        balance: a.amount - totalExp,
        soeBreakdown
      };
    });
  }, [baseAllocations, baseExpenses, ranges, schemes, sectors, activities, subActivities, soes]);

  const allocationExpenditureData = useMemo(() => {
    return currentSoes.map(s => {
      const soeAllocations = baseAllocations.filter(a => 
        a.fundedSOEs?.some(f => f.soeId === s.id)
      );
      const soeExpenses = baseExpenses.filter(e => e.soeId === s.id && e.status !== 'rejected');
      
      const totalAllocated = soeAllocations.reduce((sum, a) => {
        const funded = a.fundedSOEs?.find(f => f.soeId === s.id)?.amount || 0;
        return sum + funded;
      }, 0);
      
      const totalExp = soeExpenses.reduce((sum, e) => sum + e.amount, 0);
      const approvedBudget = getApprovedBudget(s);

      return {
        soeName: s.name,
        approvedBudget,
        totalAllocated,
        totalExp,
        balance: totalAllocated - totalExp,
        treasuryBalance: approvedBudget - totalAllocated
      };
    });
  }, [currentSoes, baseAllocations, baseExpenses]);

  const combinedReportData = useMemo(() => {
    return [...comprehensiveReportData, ...allocationExpenditureData];
  }, [comprehensiveReportData, allocationExpenditureData]);

  const uniqueSchemes = useMemo(() => 
    Array.from(new Set(comprehensiveReportData.map((item: any) => item.scheme).filter(Boolean))).sort()
  , [comprehensiveReportData]);

  const uniqueSectors = useMemo(() => 
    Array.from(new Set(comprehensiveReportData.map((item: any) => item.sector).filter(Boolean))).sort()
  , [comprehensiveReportData]);

  const uniqueActivities = useMemo(() => 
    Array.from(new Set(comprehensiveReportData.map((item: any) => item.activity).filter(Boolean))).sort()
  , [comprehensiveReportData]);

  const uniqueSubActivities = useMemo(() => 
    Array.from(new Set(comprehensiveReportData.map((item: any) => item.subActivity).filter(Boolean))).sort()
  , [comprehensiveReportData]);

  const uniqueRangesList = useMemo(() => 
    Array.from(new Set(comprehensiveReportData.map((item: any) => item.range).filter(Boolean))).sort()
  , [comprehensiveReportData]);

  const uniqueSoes = useMemo(() => 
    Array.from(new Set(soes.map(s => s.name))).sort()
  , [soes]);

  const filteredLedgerData = useMemo(() => {
    const filtered = currentAllocations.filter(alloc => {
      const r = ranges.find(r => r.id === alloc.rangeId);
      const sa = subActivities.find(sa => sa.id === alloc.subActivityId);
      const act = activities.find(a => a.id === (alloc.subActivityId ? sa?.activityId : alloc.activityId));
      const sec = sectors.find(sec => sec.id === act?.sectorId);
      const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
      const soeNames = alloc.fundedSOEs?.map(f => soes.find(s => s.id === f.soeId)?.name).filter(Boolean) || [];
      
      let hierarchy = '';
      if (alloc.subActivityId) {
        hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' -> ');
      } else if (alloc.activityId) {
        hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' -> ');
      }

      const searchLower = ledgerSearchTerm.toLowerCase();
      const matchesSearch = !ledgerSearchTerm || (
        hierarchy.toLowerCase().includes(searchLower) ||
        soeNames.join(' ').toLowerCase().includes(searchLower) ||
        r?.name.toLowerCase().includes(searchLower) ||
        alloc.remarks?.toLowerCase().includes(searchLower) ||
        alloc.id.toLowerCase().includes(searchLower)
      );

      const matchesFilters = (
        (!ledgerFilters.scheme || sch?.name === ledgerFilters.scheme) &&
        (!ledgerFilters.sector || sec?.name === ledgerFilters.sector) &&
        (!ledgerFilters.activity || act?.name === ledgerFilters.activity) &&
        (!ledgerFilters.subActivity || sa?.name === ledgerFilters.subActivity) &&
        (!ledgerFilters.range || r?.name === ledgerFilters.range) &&
        (!ledgerFilters.soe || soeNames.includes(ledgerFilters.soe))
      );
      return matchesSearch && matchesFilters;
    });

    let totalCredit = 0;
    let totalDebit = 0;

    filtered.forEach(alloc => {
      totalCredit += alloc.amount;
      const allocExpenses = expenses.filter(e => e.allocationId === alloc.id && e.status !== 'rejected');
      totalDebit += allocExpenses.reduce((sum, e) => sum + e.amount, 0);
    });

    return {
      allocations: filtered,
      totals: {
        credit: totalCredit,
        debit: totalDebit,
        balance: totalCredit - totalDebit
      }
    };
  }, [currentAllocations, ledgerSearchTerm, ledgerFilters, ranges, subActivities, activities, sectors, schemes, soes, expenses]);

    const downloadLedgerPDF = () => {
      const doc = new jsPDF('landscape');
      doc.setFontSize(16);
      doc.text("Passbook Ledger Report", 14, 15);
      doc.setFontSize(10);
      doc.text(`Financial Year: ${fys.find(f => f.id === selectedFY)?.name || selectedFY}`, 14, 22);
      
      const headers = ["Date", "Range", "Hierarchy & SOE", "Description", "Approval ID", "Credit (Rs.)", "Debit (Rs.)", "Unspent Balance (Rs.)"];
      const body: any[] = [];
      
      filteredLedgerData.allocations.forEach(alloc => {
        const r = ranges.find(r => r.id === alloc.rangeId);
        const soeNames = alloc.fundedSOEs?.map(f => soes.find(s => s.id === f.soeId)?.name).filter(Boolean).join(', ') || 'Pending Funds';
        
        let hierarchy = '';
        if (alloc.subActivityId) {
          const sa = subActivities.find(sa => sa.id === alloc.subActivityId);
          const act = activities.find(a => a.id === sa?.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' -> ');
        } else if (alloc.activityId) {
          const act = activities.find(a => a.id === alloc.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' -> ');
        }

        const allocExpenses = expenses.filter(e => e.allocationId === alloc.id && e.status !== 'rejected').sort((a,b) => new Date(a.date).getTime() - new Date(b.date).getTime());
        
        let currentBalance = alloc.amount;
        
        // Initial Allocation Row
        body.push([
          "-",
          r?.name || 'N/A',
          `${hierarchy || 'N/A'}\n${soeNames}`,
          "Initial Allocation",
          "-",
          alloc.amount.toLocaleString('en-IN'),
          "-",
          currentBalance.toLocaleString('en-IN')
        ]);

        // Expense Rows
        allocExpenses.forEach(exp => {
          currentBalance -= exp.amount;
          body.push([
            exp.date ? exp.date.split('-').reverse().join('/') : '',
            r?.name || 'N/A',
            `${hierarchy || 'N/A'}\n${soeNames}`,
            exp.description,
            exp.approvalId ? `#${exp.approvalId}` : '-',
            "-",
            exp.amount.toLocaleString('en-IN'),
            currentBalance.toLocaleString('en-IN')
          ]);
        });
      });

      autoTable(doc, {
        head: [headers],
        body: body,
        startY: 30,
        styles: { fontSize: 7, cellPadding: 2, overflow: 'linebreak' },
        headStyles: { fillColor: [5, 150, 105] },
        columnStyles: {
          0: { cellWidth: 20 },
          1: { cellWidth: 25 },
          2: { cellWidth: 80 },
          3: { cellWidth: 60 },
          4: { cellWidth: 20 },
          5: { cellWidth: 25, halign: 'right' },
          6: { cellWidth: 25, halign: 'right' },
          7: { cellWidth: 25, halign: 'right' }
        }
      });

      doc.save(`ledger_report_${selectedFY}.pdf`);
    };

    const downloadLedgerExcel = async () => {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Passbook Ledger");
      
      const title = `Passbook Ledger Report - FY ${fys.find(f => f.id === selectedFY)?.name || selectedFY}`;
      const titleRow = sheet.addRow([title]);
      titleRow.font = { bold: true, size: 14 };
      sheet.mergeCells(1, 1, 1, 8);
      titleRow.alignment = { horizontal: 'center' };

      const headers = ["Date", "Range", "Hierarchy & SOE", "Description", "Approval ID", "Credit (Rs.)", "Debit (Rs.)", "Unspent Balance (Rs.)"];
      const headerRow = sheet.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE0E0E0' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      filteredLedgerData.allocations.forEach(alloc => {
        const r = ranges.find(r => r.id === alloc.rangeId);
        const soeNames = alloc.fundedSOEs?.map(f => soes.find(s => s.id === f.soeId)?.name).filter(Boolean).join(', ') || 'Pending Funds';
        
        let hierarchy = '';
        if (alloc.subActivityId) {
          const sa = subActivities.find(sa => sa.id === alloc.subActivityId);
          const act = activities.find(a => a.id === sa?.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' -> ');
        } else if (alloc.activityId) {
          const act = activities.find(a => a.id === alloc.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' -> ');
        }

        const allocExpenses = expenses.filter(e => e.allocationId === alloc.id && e.status !== 'rejected').sort((a,b) => new Date(a.date).getTime() - new Date(b.date).getTime());
        
        let currentBalance = alloc.amount;
        
        const row1 = sheet.addRow([
          "-",
          r?.name || 'N/A',
          `${hierarchy || 'N/A'}\n${soeNames}`,
          "Initial Allocation",
          "-",
          alloc.amount,
          "-",
          currentBalance
        ]);
        row1.eachCell(cell => {
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cell.alignment = { wrapText: true, vertical: 'middle' };
        });

        allocExpenses.forEach(exp => {
          currentBalance -= exp.amount;
          const row = sheet.addRow([
            exp.date ? exp.date.split('-').reverse().join('/') : '',
            r?.name || 'N/A',
            `${hierarchy || 'N/A'}\n${soeNames}`,
            exp.description,
            exp.approvalId ? `#${exp.approvalId}` : '-',
            "-",
            exp.amount,
            currentBalance
          ]);
          row.eachCell(cell => {
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { wrapText: true, vertical: 'middle' };
          });
        });
      });

      sheet.columns.forEach((column, i) => {
        column.width = [12, 15, 45, 40, 15, 15, 15, 15][i];
        if (i >= 5) {
          column.alignment = { horizontal: 'right', vertical: 'middle' };
          column.numFmt = '#,##0.00';
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `ledger_report_${selectedFY}.xlsx`);
    };

  const getSoeAllocated = (soeId: string) => {
    const allocated = baseAllocations.reduce((sum, a) => sum + (a.fundedSOEs?.find(f => f.soeId === soeId)?.amount || 0), 0);
    const surrendered = surrenders.filter(s => s.soeId === soeId && s.fyId === selectedFY).reduce((sum, s) => sum + s.amount, 0);
    return allocated - surrendered;
  };
  const getAllocSpent = (allocId: string) => currentExpenses.filter(e => e.allocationId === allocId).reduce((sum, e) => sum + e.amount, 0);

  const totalAllocated = baseAllocations.reduce((sum, a) => sum + a.amount, 0);
  const totalSpent = baseExpenses.reduce((sum, e) => sum + e.amount, 0);
  const totalBudget = userRangeId ? totalAllocated : currentSoes.reduce((sum, s) => sum + getApprovedBudget(s), 0);
  const totalReceivedInTry = userRangeId ? 0 : currentSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0);
  const remainingBalance = totalAllocated - totalSpent;
  const totalTryBalance = totalReceivedInTry - totalAllocated;

  const chartData = userRangeId ? [
    { name: 'Allocated (Unspent)', value: Math.max(0, totalAllocated - totalSpent), color: '#007bff' },
    { name: 'Spent', value: totalSpent, color: '#dc3545' }
  ] : [
    { name: 'Allocated (Unspent)', value: Math.max(0, totalAllocated - totalSpent), color: '#007bff' },
    { name: 'Spent', value: totalSpent, color: '#dc3545' },
    { name: 'Unallocated', value: Math.max(0, totalBudget - totalAllocated), color: '#28a745' }
  ];

  const soeAbstractData = useMemo(() => {
    return currentSoes.map(soe => {
      const sch = schemes.find(s => s.id === soe.schemeId);
      const sec = sectors.find(s => s.id === soe.sectorId);
      const act = activities.find(a => a.id === soe.activityId);
      const sa = subActivities.find(s => s.id === soe.subActivityId);
      const hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' > ');

      const allocated = baseAllocations.reduce((sum, a) => {
        const funded = a.fundedSOEs?.find(f => f.soeId === soe.id);
        return sum + (funded?.amount || 0);
      }, 0);

      const spent = baseExpenses.filter(e => e.soeId === soe.id).reduce((sum, e) => sum + e.amount, 0);
      const approvedBudget = getApprovedBudget(soe);
      const receivedInTry = getReceivedInTry(soe);

      return {
        id: soe.id,
        soeId: soe.id,
        soeName: soe.name,
        hierarchy,
        schemeName: sch?.name || '',
        sectorName: sec?.name || '',
        activityName: act?.name || '',
        subActivityName: sa?.name || '',
        approvedBudget,
        receivedInTry,
        allocated,
        toBeAllocated: receivedInTry - allocated,
        tryBalance: receivedInTry - allocated,
        spent,
        remainingToSpend: allocated - spent,
        schemeId: soe.schemeId,
        sectorId: soe.sectorId,
        activityId: soe.activityId,
        subActivityId: soe.subActivityId
      };
    })
    .filter(item => !userRangeId || item.allocated > 0)
    .sort((a, b) => a.hierarchy.localeCompare(b.hierarchy) || a.soeName.localeCompare(b.soeName));
  }, [currentSoes, baseAllocations, baseExpenses, schemes, sectors, activities, subActivities, userRangeId]);

  const surrenderBudgetStatus = useMemo(() => {
    const { rangeId, schemeId, sectorId, activityId, subActivityId, soeId } = surrenderFormSelection;
    if (!rangeId || !soeId) return null;

    const relevantAllocations = baseAllocations.filter(a => 
      a.rangeId === rangeId &&
      a.schemeId === schemeId &&
      (!sectorId || !a.sectorId || a.sectorId === sectorId) &&
      (!activityId || !a.activityId || a.activityId === activityId) &&
      (!subActivityId || !a.subActivityId || a.subActivityId === subActivityId)
    );

    let totalAllocated = 0;
    let totalSpent = 0;

    relevantAllocations.forEach(alloc => {
      const funded = alloc.fundedSOEs?.find(f => f.soeId === soeId);
      if (funded) {
        totalAllocated += funded.amount;
        const spent = baseExpenses.filter(e => 
          e.allocationId === alloc.id && 
          e.soeId === soeId &&
          e.status !== 'rejected'
        ).reduce((sum, e) => sum + e.amount, 0);
        totalSpent += spent;
      }
    });

    const totalSurrendered = surrenders.filter(s => 
      s.rangeId === rangeId &&
      s.schemeId === schemeId &&
      s.soeId === soeId &&
      (!sectorId || !s.sectorId || s.sectorId === sectorId) &&
      (!activityId || !s.activityId || s.activityId === activityId) &&
      (!subActivityId || !s.subActivityId || s.subActivityId === subActivityId)
    ).reduce((sum, s) => sum + s.amount, 0);

    return {
      allocated: totalAllocated,
      spent: totalSpent,
      surrendered: totalSurrendered,
      balance: totalAllocated - totalSpent - totalSurrendered
    };
  }, [surrenderFormSelection, baseAllocations, baseExpenses, surrenders]);

  const masterControlData = useMemo(() => {
    const map: Record<string, any> = {};
    
    baseAllocations.forEach(alloc => {
      const range = ranges.find(r => r.id === alloc.rangeId);
      const rangeName = range?.name === 'Rajgarh Forest Division' ? 'Division' : (range?.name || 'N/A');
      const sch = schemes.find(s => s.id === alloc.schemeId);
      const schName = sch?.name || 'N/A';
      const sec = sectors.find(s => s.id === alloc.sectorId);
      const secName = sec?.name || 'N/A';
      const act = activities.find(a => a.id === alloc.activityId);
      const actName = act?.name || 'N/A';
      const sa = subActivities.find(s => s.id === alloc.subActivityId);
      const saName = sa?.name || 'N/A';

      // Apply filters
      if (reportFilters.range && rangeName !== reportFilters.range) return;
      if (reportFilters.scheme && schName !== reportFilters.scheme) return;
      if (reportFilters.sector && secName !== reportFilters.sector) return;
      if (reportFilters.activity && actName !== reportFilters.activity) return;
      if (reportFilters.subActivity && saName !== reportFilters.subActivity) return;
      
      alloc.fundedSOEs?.forEach(funded => {
        const soe = soes.find(s => s.id === funded.soeId);
        const soeName = soe?.name || 'N/A';
        if (reportFilters.soe && soeName !== reportFilters.soe) return;

        const key = `${alloc.rangeId}-${alloc.schemeId}-${alloc.sectorId}-${alloc.activityId}-${alloc.subActivityId}-${funded.soeId}`;
        const spent = baseExpenses.filter(e => 
          e.allocationId === alloc.id && 
          e.soeId === funded.soeId &&
          e.status !== 'rejected'
        ).reduce((sum, e) => sum + e.amount, 0);
        
        if (map[key]) {
          map[key].allocated += funded.amount;
          map[key].expenditure += spent;
          map[key].balance = map[key].allocated - map[key].expenditure;
        } else {
          map[key] = {
            rangeName,
            schemeName: schName,
            sectorName: secName,
            activityName: actName,
            subActivityName: saName,
            soeName,
            allocated: funded.amount,
            expenditure: spent,
            balance: funded.amount - spent
          };
        }
      });
    });
    
    let result = Object.values(map);
    
    if (reportSearchTerm) {
      const lower = reportSearchTerm.toLowerCase();
      result = result.filter((item: any) => 
        item.rangeName.toLowerCase().includes(lower) ||
        item.schemeName.toLowerCase().includes(lower) ||
        item.soeName.toLowerCase().includes(lower) ||
        item.activityName.toLowerCase().includes(lower)
      );
    }

    return result.sort((a: any, b: any) => 
      a.rangeName.localeCompare(b.rangeName) || 
      a.schemeName.localeCompare(b.schemeName) || 
      a.soeName.localeCompare(b.soeName)
    );
  }, [ranges, baseAllocations, baseExpenses, schemes, sectors, activities, subActivities, soes, reportFilters, reportSearchTerm]);

  const soeAbstractForAllocations = useMemo(() => {
    return soeAbstractData.filter(item => {
      // Apply dynamic filtering based on form selection
      if (allocationFormFilters.soeId && item.id !== allocationFormFilters.soeId) return false;
      if (allocationFormFilters.subActivityId && item.subActivityId !== allocationFormFilters.subActivityId) return false;
      if (allocationFormFilters.activityId && item.activityId !== allocationFormFilters.activityId) return false;
      if (allocationFormFilters.sectorId && item.sectorId !== allocationFormFilters.sectorId) return false;
      if (allocationFormFilters.schemeId && item.schemeId !== allocationFormFilters.schemeId) return false;
      
      if (soeSearchTerm) {
        const lowerSearch = soeSearchTerm.toLowerCase();
        return item.soeName.toLowerCase().includes(lowerSearch) || 
               item.hierarchy.toLowerCase().includes(lowerSearch);
      }
      
      return true;
    });
  }, [soeAbstractData, allocationFormFilters, soeSearchTerm]);


  // --- Render Functions for Tabs ---
  const renderDashboard = () => {
    const rangeAllocationMap: Record<string, any> = {};
    baseAllocations.forEach(alloc => {
      const key = `${alloc.rangeId}-${alloc.schemeId}-${alloc.sectorId}-${alloc.activityId}`;
      const spent = baseExpenses.filter(e => e.allocationId === alloc.id).reduce((sum, e) => sum + e.amount, 0);
      
      if (rangeAllocationMap[key]) {
        const existing = rangeAllocationMap[key];
        existing.allocated += alloc.amount;
        existing.spent += spent;
        existing.balance = existing.allocated - existing.spent;
      } else {
        const r = ranges.find(r => r.id === alloc.rangeId);
        const sch = currentSchemes.find(s => s.id === alloc.schemeId);
        const sec = currentSectors.find(s => s.id === alloc.sectorId);
        const act = currentActivities.find(a => a.id === alloc.activityId);
        
        rangeAllocationMap[key] = {
          range: r?.name || 'N/A',
          scheme: sch?.name || 'N/A',
          sector: sec?.name || 'N/A',
          activity: act?.name || 'N/A',
          allocated: alloc.amount,
          spent: spent,
          balance: alloc.amount - spent
        };
      }
    });
    
    const rangeAllocationSummary = Object.values(rangeAllocationMap).sort((a, b) => {
      return a.range.localeCompare(b.range) || a.scheme.localeCompare(b.scheme) || a.sector.localeCompare(b.sector) || a.activity.localeCompare(b.activity);
    });

    // Group expenses by date for trend chart
    const expensesByDate = baseExpenses.reduce((acc, exp) => {
      acc[exp.date] = (acc[exp.date] || 0) + exp.amount;
      return acc;
    }, {} as Record<string, number>);
    
    const trendData = Object.keys(expensesByDate).sort().map(date => ({
      date,
      amount: expensesByDate[date]
    }));

    const schemeSummary = currentSchemes.map(sch => {
      const schAllocations = baseAllocations.filter(a => a.schemeId === sch.id);
      const totalAllocated = schAllocations.reduce((sum, a) => sum + a.amount, 0);
      const totalSpent = baseExpenses.filter(e => schAllocations.some(a => a.id === e.allocationId)).reduce((sum, e) => sum + e.amount, 0);
      
      const schemeSoes = currentSoes.filter(s => s.schemeId === sch.id);
      const totalSoeBudget = schemeSoes.reduce((sum, s) => sum + getApprovedBudget(s), 0);

      const displayBudget = userRangeId ? totalAllocated : totalSoeBudget;

      return {
        name: sch.name,
        budget: displayBudget,
        allocated: totalAllocated,
        spent: totalSpent,
        balance: totalAllocated - totalSpent
      };
    }).filter(s => !userRangeId || s.allocated > 0 || s.spent > 0);

    const sectorSummary = currentSectors.map(sec => {
      const secAllocations = baseAllocations.filter(a => a.sectorId === sec.id);
      const totalAllocated = secAllocations.reduce((sum, a) => sum + a.amount, 0);
      const totalSpent = baseExpenses.filter(e => secAllocations.some(a => a.id === e.allocationId)).reduce((sum, e) => sum + e.amount, 0);
      
      return {
        name: sec.name,
        allocated: totalAllocated,
        spent: totalSpent,
        balance: totalAllocated - totalSpent
      };
    }).filter(s => !userRangeId || s.allocated > 0 || s.spent > 0);

    const activitySummary = currentActivities.map(act => {
      const sec = currentSectors.find(s => s.id === act.sectorId);
      const sch = currentSchemes.find(s => s.id === (sec ? sec.schemeId : act.schemeId));

      const actAllocations = baseAllocations.filter(a => a.activityId === act.id);
      const totalAllocated = actAllocations.reduce((sum, a) => sum + a.amount, 0);
      const totalSpent = baseExpenses.filter(e => actAllocations.some(a => a.id === e.allocationId)).reduce((sum, e) => sum + e.amount, 0);
      
      return {
        scheme: sch?.name || 'N/A',
        sector: sec?.name || 'N/A',
        name: act.name,
        allocated: totalAllocated,
        spent: totalSpent,
        balance: totalAllocated - totalSpent
      };
    }).filter(a => !userRangeId || a.allocated > 0 || a.spent > 0).sort((a, b) => {
      const aHasEntry = a.allocated > 0 || a.spent > 0 ? 1 : 0;
      const bHasEntry = b.allocated > 0 || b.spent > 0 ? 1 : 0;
      if (aHasEntry !== bHasEntry) return bHasEntry - aHasEntry;
      return a.scheme.localeCompare(b.scheme) || a.sector.localeCompare(b.sector) || a.name.localeCompare(b.name);
    });

    const soeDashboardSummary = soeAbstractData.filter(item => {
      const lowerSearch = (dashboardSearch || searchTerm).toLowerCase();
      if (lowerSearch) {
        return item.soeName.toLowerCase().includes(lowerSearch) || 
               item.hierarchy.toLowerCase().includes(lowerSearch) ||
               item.schemeName.toLowerCase().includes(lowerSearch) ||
               item.sectorName.toLowerCase().includes(lowerSearch);
      }
      return true;
    });

    const dashboardStats = {
      totalBudget: soeDashboardSummary.reduce((sum, item) => sum + item.approvedBudget, 0),
      totalReceivedInTry: soeDashboardSummary.reduce((sum, item) => sum + item.receivedInTry, 0),
      totalAllocated: soeDashboardSummary.reduce((sum, item) => sum + item.allocated, 0),
      totalTryBalance: soeDashboardSummary.reduce((sum, item) => sum + item.toBeAllocated, 0),
      totalSpent: soeDashboardSummary.reduce((sum, item) => sum + item.spent, 0),
      remainingBalance: soeDashboardSummary.reduce((sum, item) => sum + item.remainingToSpend, 0)
    };

    // Filter notifications relevant to current user
    const relevantNotifs = notifications.filter(n => {
      if (userRole !== 'admin' && userRole !== 'deo' && userRole !== 'approver' && userRole !== 'DA') {
        if (n.targetRanges && n.targetRanges.length > 0 && !n.targetRanges.includes('All')) {
          const userRangeObj = ranges.find(r => r.id === userRangeId);
          const matches = n.targetRanges.includes(userRole || '') ||
                          (userRangeId && n.targetRanges.includes(userRangeId)) ||
                          (userRangeObj && n.targetRanges.includes(userRangeObj.name));
          if (!matches) return false;
        }
      }
      return true;
    }).sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));

    // Get latest allocations for quick notification preview on dashboard
    const recentAllocations = [...baseAllocations]
      .sort((a, b) => (b.createdAt || b.updatedAt || 0) - (a.createdAt || a.updatedAt || 0))
      .slice(0, 4)
      .map(alloc => {
        const r = ranges.find(rg => rg.id === alloc.rangeId);
        const sch = currentSchemes.find(s => s.id === alloc.schemeId);
        const sec = currentSectors.find(s => s.id === alloc.sectorId);
        const soeNames = alloc.fundedSOEs?.map(f => {
          const soeObj = soes.find(s => s.id === f.soeId);
          return soeObj ? `${soeObj.name} (â‚¹${(Number(f.amount) || 0).toLocaleString('en-IN')})` : '';
        }).filter(Boolean).join(', ');
        return {
          id: alloc.id,
          rangeName: r?.name || 'N/A',
          schemeName: sch?.name || 'N/A',
          sectorName: sec?.name || 'General',
          amount: alloc.amount,
          soeNames: soeNames || 'Direct Head',
          date: alloc.createdAt ? new Date(alloc.createdAt).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : 'Recently',
          status: alloc.status || 'Funded'
        };
      });

    // 10 Most Recent Actions (Expenditures added, Allocations updated/added) sorted by date descending
    const expenseActivities = baseExpenses.map(exp => {
      const alloc = allocations.find(a => a.id === exp.allocationId);
      const range = ranges.find(r => r.id === (alloc?.rangeId || (exp as any).rangeId));
      const sch = schemes.find(s => s.id === (alloc?.schemeId || exp.schemeId));
      const sec = sectors.find(s => s.id === (alloc?.sectorId || exp.sectorId));
      const soe = soes.find(s => s.id === exp.soeId);

      let ts = exp.createdAt || exp.updatedAt || 0;
      if (!ts && exp.date) {
        const parsed = new Date(exp.date).getTime();
        if (!isNaN(parsed)) ts = parsed;
      }

      const dateStr = ts > 0 
        ? new Date(ts).toLocaleString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
        : (exp.date ? exp.date.split('-').reverse().join('/') : 'Recently');

      return {
        id: `exp-${exp.id}`,
        rawId: exp.id,
        type: 'expenditure' as const,
        actionName: 'Expenditure Added',
        amount: exp.amount,
        dateStr,
        timestamp: ts,
        rangeName: range?.name || 'Division',
        schemeName: sch?.name || 'N/A',
        sectorName: sec?.name || '',
        soeName: soe?.name || 'SOE Head',
        details: exp.description || (exp.payeeName ? `Payee: ${exp.payeeName}` : '') || (exp.approvalId ? `Approval #${exp.approvalId}` : ''),
        status: exp.status || 'approved',
        statusLabel: exp.status === 'pending' ? 'Pending' : exp.status === 'rejected' ? 'Rejected' : 'Approved',
        statusColor: exp.status === 'pending' ? 'bg-amber-100 text-amber-800 border-amber-200' : exp.status === 'rejected' ? 'bg-red-100 text-red-800 border-red-200' : 'bg-emerald-100 text-emerald-800 border-emerald-200',
      };
    });

    const allocationActivities = baseAllocations.map(alloc => {
      const range = ranges.find(r => r.id === alloc.rangeId);
      const sch = schemes.find(s => s.id === alloc.schemeId);
      const sec = sectors.find(s => s.id === alloc.sectorId);
      const act = activities.find(a => a.id === alloc.activityId);

      let ts = alloc.updatedAt || alloc.createdAt || 0;
      const isUpdate = alloc.updatedAt && alloc.createdAt && alloc.updatedAt > alloc.createdAt;

      const dateStr = ts > 0 
        ? new Date(ts).toLocaleString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
        : 'Recently';

      const soeNames = alloc.fundedSOEs?.map(f => {
        const soeObj = soes.find(s => s.id === f.soeId);
        return soeObj ? soeObj.name : '';
      }).filter(Boolean).join(', ');

      return {
        id: `alloc-${alloc.id}`,
        rawId: alloc.id,
        type: 'allocation' as const,
        actionName: isUpdate ? 'Allocation Updated' : 'Allocation Added',
        amount: alloc.amount,
        dateStr,
        timestamp: ts,
        rangeName: range?.name || 'Division',
        schemeName: sch?.name || 'N/A',
        sectorName: sec?.name || '',
        soeName: soeNames || (act?.name || 'Budget Allocation'),
        details: alloc.remarks || (act?.name ? `Activity: ${act.name}` : ''),
        status: alloc.status || 'Funded',
        statusLabel: alloc.status === 'Pending SOE Funds' ? 'Pending Funds' : 'Funded',
        statusColor: alloc.status === 'Pending SOE Funds' ? 'bg-amber-100 text-amber-800 border-amber-200' : 'bg-blue-100 text-blue-800 border-blue-200',
      };
    });

    const recentActivitiesList = [...expenseActivities, ...allocationActivities]
      .sort((a, b) => b.timestamp - a.timestamp)
      .slice(0, 10);

    return (
      <div className="space-y-6">
        {/* Recent Updates & Budget Allocation Notifications Alert Bar */}
        {(relevantNotifs.length > 0 || recentAllocations.length > 0) && (
          <div className="bg-gradient-to-r from-emerald-50 via-teal-50 to-indigo-50/60 p-4 rounded-xl border border-emerald-200/80 shadow-sm">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 pb-2.5 border-b border-emerald-200/60">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-emerald-600 text-white rounded-lg shadow-xs animate-bounce">
                  <Bell className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-xs font-black text-emerald-950 uppercase tracking-wide flex items-center gap-2">
                    <span>Recent Updates & Budget Allocation Notifications</span>
                    <span className="bg-emerald-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">
                      {relevantNotifs.length + recentAllocations.length} New
                    </span>
                  </h3>
                  <p className="text-[11px] text-emerald-800 font-medium">
                    Live notifications for new budget allocations, scheme fund releases, and official updates.
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2 self-end sm:self-auto">
                <button
                  type="button"
                  onClick={() => {
                    setActiveTab('Notifications');
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                  }}
                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold shadow-xs transition-all flex items-center gap-1 cursor-pointer active:scale-95"
                >
                  <Eye className="w-3.5 h-3.5" />
                  <span>View All Notifications</span>
                </button>
              </div>
            </div>

            {/* Notification items carousel / grid */}
            <div className="mt-3 grid grid-cols-1 md:grid-cols-2 gap-2.5">
              {/* Latest Budget Allocation Cards */}
              {recentAllocations.slice(0, 2).map((al) => (
                <div
                  key={`dash-alloc-${al.id}`}
                  onClick={() => {
                    setActiveTab('Allocations');
                    setSearchTerm(al.schemeName !== 'N/A' ? al.schemeName : '');
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                  }}
                  className="bg-white/95 p-3 rounded-lg border border-emerald-200 shadow-2xs hover:border-emerald-500 hover:shadow-sm transition-all flex flex-col justify-between text-xs cursor-pointer group"
                  title="Click to view this budget allocation details"
                >
                  <div className="flex items-center justify-between gap-2 mb-1.5">
                    <span className="font-extrabold text-emerald-900 bg-emerald-100 px-2 py-0.5 rounded text-[10px] flex items-center gap-1">
                      <DollarSign className="w-3 h-3 text-emerald-700" />
                      Budget Allotted: â‚¹{al.amount.toLocaleString('en-IN')}
                    </span>
                    <span className="text-[10px] text-gray-500 font-medium">{al.date}</span>
                  </div>
                  <div className="space-y-0.5 text-gray-700">
                    <div className="font-bold text-gray-900 flex items-center justify-between gap-1.5">
                      <div className="flex items-center gap-1.5 truncate">
                        <MapPin className="w-3 h-3 text-indigo-600 shrink-0" />
                        <span>{al.rangeName}</span>
                        <span className="text-gray-400 font-normal">|</span>
                        <span className="text-emerald-800 font-semibold">{al.schemeName}</span>
                      </div>
                      <span className="text-[10px] text-emerald-600 font-bold flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity shrink-0">
                        View <ChevronRight className="w-3 h-3" />
                      </span>
                    </div>
                    <div className="text-[11px] text-gray-600 truncate">
                      <span className="font-semibold text-gray-700">Sector:</span> {al.sectorName} 
                      {al.soeNames && <span className="ml-1 text-emerald-950 font-medium">({al.soeNames})</span>}
                    </div>
                  </div>
                </div>
              ))}

              {/* Latest Official Notifications */}
              {relevantNotifs.slice(0, 2).map((nf) => (
                <div
                  key={`dash-notif-${nf.id}`}
                  className="bg-white/95 p-3 rounded-lg border border-indigo-200 shadow-2xs hover:border-indigo-500 hover:shadow-sm transition-all flex flex-col justify-between text-xs cursor-pointer group"
                  onClick={() => {
                    if (nf.url || nf.fileData) {
                      handleViewFile(nf);
                    } else if (nf.category === 'budget_allocation' || nf.type?.includes('allocation')) {
                      setActiveTab('Allocations');
                      setSearchTerm(nf.schemeName || '');
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    } else if (nf.category === 'budget_distribution') {
                      setActiveTab('Distributed Budget');
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    } else {
                      setActiveTab('Notifications');
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    }
                  }}
                  title={nf.url || nf.fileData ? "Click to open document directly" : "Click to view details"}
                >
                  <div className="flex items-center justify-between gap-2 mb-1.5">
                    <span className="font-extrabold text-indigo-900 bg-indigo-100 px-2 py-0.5 rounded text-[10px] flex items-center gap-1">
                      <FileText className="w-3 h-3 text-indigo-700" />
                      {nf.category === 'budget_allocation' ? 'Allocation Alert' : nf.category === 'budget_distribution' ? 'Distribution Allotment' : 'Notice / Circular'}
                    </span>
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] text-gray-500 font-medium">
                        {nf.createdAt ? new Date(nf.createdAt).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' }) : 'Recent'}
                      </span>
                      {(nf.url || nf.fileData) && (
                        <span className="p-1 bg-indigo-50 text-indigo-600 rounded hover:bg-indigo-100 transition-colors" title="Open PDF">
                          <Eye className="w-3 h-3" />
                        </span>
                      )}
                    </div>
                  </div>
                  <div className="space-y-0.5">
                    <div className="font-bold text-gray-900 truncate flex items-center justify-between gap-1.5" title={nf.name}>
                      <span className="truncate">{nf.name}</span>
                      <span className="text-[10px] text-indigo-600 font-bold flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity shrink-0">
                        {nf.url || nf.fileData ? 'Open PDF' : 'View'} <ChevronRight className="w-3 h-3" />
                      </span>
                    </div>
                    {nf.description && (
                      <div className="text-[11px] text-gray-600 truncate" title={nf.description}>
                        {nf.description}
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className={`grid grid-cols-2 sm:grid-cols-2 md:grid-cols-3 ${userRangeId ? 'lg:grid-cols-3' : 'xl:grid-cols-6 lg:grid-cols-3'} gap-2 sm:gap-3`}>
          {!userRangeId && <StatCard title="Total Approved Budget" amount={dashboardStats.totalBudget} icon={<Wallet />} color="text-blue-600" />}
          {!userRangeId && <StatCard title="Total Received (Try)" amount={dashboardStats.totalReceivedInTry} icon={<Landmark />} color="text-indigo-500" />}
          <StatCard title={userRangeId ? "My Allocated Budget" : "Total Allocated"} amount={dashboardStats.totalAllocated} icon={<Map />} color="text-indigo-600" />
          {!userRangeId && <StatCard title="To Be Allocated (Received)" amount={dashboardStats.totalTryBalance} icon={<IndianRupee />} color="text-orange-500" />}
          <StatCard 
            title={userRangeId ? "My Total Expenditure" : "Total Expenditure"} 
            amount={dashboardStats.totalSpent} 
            icon={<TrendingDown />} 
            color="text-red-600" 
            subtitle={
              userRangeId 
                ? (dashboardStats.totalAllocated > 0 ? `${((dashboardStats.totalSpent / dashboardStats.totalAllocated) * 100).toFixed(1)}% of My Allocation` : undefined)
                : (dashboardStats.totalBudget > 0 ? `${((dashboardStats.totalSpent / dashboardStats.totalBudget) * 100).toFixed(1)}% of Budget` : undefined)
            }
          />
          <StatCard title={userRangeId ? "My Remaining Balance" : "Remaining Balance (Unspent)"} amount={dashboardStats.remainingBalance} icon={<IndianRupee />} color="text-emerald-600" />
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 lg:col-span-1">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2">Budget Overview</h3>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={chartData} innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                    {chartData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
                  </Pie>
                  <Tooltip formatter={(value: number) => `â‚¹${value.toLocaleString()}`} />
                  <Legend verticalAlign="bottom" height={36}/>
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 lg:col-span-2">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
              <Activity className="h-5 w-5 text-gray-500" /> Scheme-wise Budget
            </h3>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={schemeSummary} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
                  <XAxis dataKey="name" axisLine={false} tickLine={false} />
                  <YAxis axisLine={false} tickLine={false} tickFormatter={(val) => `â‚¹${val/1000}k`} />
                  <Tooltip formatter={(value: number) => `â‚¹${value.toLocaleString()}`} cursor={{fill: '#f3f4f6'}} />
                  <Legend />
                  <Bar dataKey="allocated" name="Allocated" fill="#007bff" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="spent" name="Spent" fill="#dc3545" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="balance" name="Balance" fill="#28a745" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {trendData.length > 0 && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
              <TrendingUp className="h-5 w-5 text-red-600" /> Spending Trend (Expenditure Over Time)
            </h3>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={trendData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
                  <XAxis dataKey="date" axisLine={false} tickLine={false} tickFormatter={(val) => val && typeof val === 'string' ? val.split('-').reverse().slice(0, 2).join('/') : ''} />
                  <YAxis axisLine={false} tickLine={false} tickFormatter={(val) => `â‚¹${val/1000}k`} />
                  <Tooltip formatter={(value: number) => `â‚¹${value.toLocaleString()}`} />
                  <Legend />
                  <Line type="monotone" dataKey="amount" name="Expenditure" stroke="#dc3545" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4 border-b pb-2">
            <h3 className="text-lg font-semibold flex items-center gap-2">
              <Table className="h-5 w-5 text-gray-500" /> Budget Abstract (Scheme/Sector/Activity)
            </h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm min-w-[600px]">
              <thead>
                <tr className="bg-gray-50 text-gray-600 font-semibold">
                  <th className="p-3 border-b">Scheme</th>
                  <th className="p-3 border-b">Sector</th>
                  <th className="p-3 border-b">Activity</th>
                  <th className="p-3 border-b text-right">Allocated</th>
                  <th className="p-3 border-b text-right">Spent</th>
                  <th className="p-3 border-b text-right">Balance</th>
                </tr>
              </thead>
              <tbody>
                {activitySummary
                  .filter(act => {
                    const searchStr = (dashboardSearch || searchTerm).toLowerCase();
                    return (
                      act.scheme.toLowerCase().includes(searchStr) ||
                      act.sector.toLowerCase().includes(searchStr) ||
                      act.name.toLowerCase().includes(searchStr) ||
                      act.allocated.toString().includes(searchStr) ||
                      act.spent.toString().includes(searchStr) ||
                      act.balance.toString().includes(searchStr)
                    );
                  })
                  .slice(0, showAllBudget ? undefined : 5)
                  .map((act, idx) => (
                    <tr key={idx} className="border-b hover:bg-gray-50 transition-colors">
                      <td className="p-3 text-xs text-gray-500">{act.scheme}</td>
                      <td className="p-3 text-xs text-gray-500">{act.sector}</td>
                      <td className="p-3 font-medium text-gray-800">{act.name}</td>
                      <td className="p-3 text-right font-mono text-blue-600">â‚¹{act.allocated.toLocaleString()}</td>
                      <td className="p-3 text-right font-mono text-red-600">â‚¹{act.spent.toLocaleString()}</td>
                      <td className="p-3 text-right font-mono font-bold text-emerald-600">â‚¹{act.balance.toLocaleString()}</td>
                    </tr>
                  ))}
                {activitySummary.length === 0 && (
                  <tr>
                    <td colSpan={6} className="p-8 text-center text-gray-400 italic">No budget data found for this Financial Year.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          {activitySummary.length > 5 && (
            <div className="mt-4 text-center">
              <button 
                onClick={() => setShowAllBudget(!showAllBudget)}
                className="text-emerald-600 font-semibold hover:text-emerald-700 transition-colors flex items-center gap-1 mx-auto"
              >
                {showAllBudget ? 'Show Less' : 'Read More'}
              </button>
            </div>
          )}
        </div>

        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
            <Activity className="h-5 w-5 text-emerald-600" /> Live SOE Budget Tracker
          </h3>
          <div className="overflow-x-auto max-h-96">
            <table className="w-full text-left border-collapse text-sm min-w-[1000px]">
              <thead className="sticky top-0 bg-white shadow-sm">
                <tr className="bg-gray-50 text-gray-600 font-semibold">
                  <th className="p-3 border-b">Hierarchy (Scheme &gt; Sector &gt; Activity)</th>
                  <th className="p-3 border-b">SOE Head</th>
                  {!userRangeId && <th className="p-3 border-b text-right">Approved Budget</th>}
                  {!userRangeId && <th className="p-3 border-b text-right">Received in Try</th>}
                  <th className="p-3 border-b text-right">Allocated</th>
                  {!userRangeId && <th className="p-3 border-b text-right">To Be Allocated</th>}
                  {!userRangeId && <th className="p-3 border-b text-right">Try Balance</th>}
                  <th className="p-3 border-b text-right">Spent</th>
                  <th className="p-3 border-b text-right">Balance</th>
                </tr>
              </thead>
              <tbody>
                {soeDashboardSummary.map((item, idx) => (
                  <tr key={idx} className="border-b hover:bg-gray-50 transition-colors">
                    <td className="p-3 text-xs text-gray-500">{item.hierarchy || 'N/A'}</td>
                    <td className="p-3 font-medium text-gray-800">{item.soeName}</td>
                    {!userRangeId && <td className="p-3 text-right text-gray-700">â‚¹{item.approvedBudget.toLocaleString()}</td>}
                    {!userRangeId && <td className="p-3 text-right text-indigo-600">â‚¹{item.receivedInTry.toLocaleString()}</td>}
                    <td className="p-3 text-right text-blue-600">â‚¹{item.allocated.toLocaleString()}</td>
                    {!userRangeId && (
                      <td className={`p-3 text-right font-bold ${item.toBeAllocated > 0 ? 'text-emerald-600' : 'text-gray-400'}`}>
                        â‚¹{item.toBeAllocated.toLocaleString()}
                      </td>
                    )}
                    {!userRangeId && (
                      <td className={`p-3 text-right font-bold ${item.tryBalance >= 0 ? 'text-emerald-600' : 'text-red-600'}`}>
                        â‚¹{item.tryBalance.toLocaleString()}
                      </td>
                    )}
                    <td className="p-3 text-right text-red-600">â‚¹{item.spent.toLocaleString()}</td>
                    <td className={`p-3 text-right font-bold ${item.remainingToSpend > 0 ? 'text-emerald-600' : 'text-gray-400'}`}>
                      â‚¹{item.remainingToSpend.toLocaleString()}
                    </td>
                  </tr>
                ))}
                {soeDashboardSummary.length === 0 && (
                  <tr>
                    <td colSpan={userRangeId ? 5 : 9} className="p-4 text-center text-gray-500">No SOE data matching current selection.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
              <Activity className="h-5 w-5 text-gray-500" /> Sector-wise Budget
            </h3>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={sectorSummary} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
                  <XAxis dataKey="name" axisLine={false} tickLine={false} />
                  <YAxis axisLine={false} tickLine={false} tickFormatter={(val) => `â‚¹${val/1000}k`} />
                  <Tooltip formatter={(value: number) => `â‚¹${value.toLocaleString()}`} cursor={{fill: '#f3f4f6'}} />
                  <Legend />
                  <Bar dataKey="allocated" name="Allocated" fill="#007bff" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="spent" name="Spent" fill="#dc3545" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="balance" name="Balance" fill="#28a745" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
              <Table className="h-5 w-5 text-gray-500" /> Scheme-wise Budget
            </h3>
            <div className="overflow-x-auto h-64">
              <table className="w-full text-left border-collapse text-sm min-w-[600px]">
                <thead className="sticky top-0 bg-white shadow-sm">
                  <tr className="bg-gray-50 text-gray-600 font-semibold">
                    <th className="p-3 border-b">Scheme</th>
                    {!userRangeId && <th className="p-3 border-b text-right">SOE Budget</th>}
                    <th className="p-3 border-b text-right">Allocation</th>
                    <th className="p-3 border-b text-right">Expenditure</th>
                    <th className="p-3 border-b text-right">Balance</th>
                  </tr>
                </thead>
                <tbody>
                  {schemeSummary.map((sch, idx) => (
                    <tr key={idx} className="border-b hover:bg-gray-50">
                      <td className="p-3 font-medium text-gray-800">{sch.name}</td>
                      {!userRangeId && <td className="p-3 text-right">â‚¹{sch.budget.toLocaleString()}</td>}
                      <td className="p-3 text-right text-blue-600">â‚¹{sch.allocated.toLocaleString()}</td>
                      <td className="p-3 text-right text-red-600">â‚¹{sch.spent.toLocaleString()}</td>
                      <td className="p-3 text-right text-emerald-600 font-medium">â‚¹{sch.balance.toLocaleString()}</td>
                    </tr>
                  ))}
                  {schemeSummary.length === 0 && (
                    <tr>
                      <td colSpan={5} className="p-4 text-center text-gray-500">No scheme data available</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 lg:col-span-2">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4 border-b pb-2">
              <h3 className="text-lg font-semibold flex items-center gap-2">
                <Table className="h-5 w-5 text-gray-500" /> Range-wise Allocation Summary
              </h3>
              <div className="relative w-full md:w-64">
                <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                <input 
                  type="text" 
                  placeholder="Search unit..." 
                  value={rangeSearch}
                  onChange={(e) => setRangeSearch(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent"
                />
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-sm min-w-[800px]">
                <thead>
                  <tr className="bg-gray-50 text-gray-600 font-semibold">
                    <th className="p-3 border-b">Range</th>
                    <th className="p-3 border-b">Scheme</th>
                    <th className="p-3 border-b">Sector</th>
                    <th className="p-3 border-b">Activity</th>
                    <th className="p-3 border-b text-right">Allocated</th>
                    <th className="p-3 border-b text-right">Spent</th>
                    <th className="p-3 border-b text-right">Balance</th>
                  </tr>
                </thead>
                <tbody>
                  {rangeAllocationSummary
                    .filter(r => 
                      r.range.toLowerCase().includes(rangeSearch.toLowerCase()) ||
                      r.scheme.toLowerCase().includes(rangeSearch.toLowerCase()) ||
                      r.sector.toLowerCase().includes(rangeSearch.toLowerCase()) ||
                      r.activity.toLowerCase().includes(rangeSearch.toLowerCase())
                    )
                    .slice(0, showAllRange ? undefined : 5)
                    .map((r, idx) => (
                      <tr key={idx} className="border-b hover:bg-gray-50 transition-colors">
                        <td className="p-3 font-medium text-gray-800">{r.range}</td>
                        <td className="p-3 text-xs text-gray-500">{r.scheme}</td>
                        <td className="p-3 text-xs text-gray-500">{r.sector}</td>
                        <td className="p-3 text-xs text-gray-500">{r.activity}</td>
                        <td className="p-3 text-right font-mono text-blue-600">â‚¹{r.allocated.toLocaleString()}</td>
                        <td className="p-3 text-right font-mono text-red-600">â‚¹{r.spent.toLocaleString()}</td>
                        <td className="p-3 text-right font-mono font-bold text-emerald-600">â‚¹{r.balance.toLocaleString()}</td>
                      </tr>
                    ))}
                  {rangeAllocationSummary.length === 0 && (
                    <tr>
                      <td colSpan={7} className="p-4 text-center text-gray-500">No allocations found.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {rangeAllocationSummary.length > 5 && (
              <div className="mt-4 text-center">
                <button 
                  onClick={() => setShowAllRange(!showAllRange)}
                  className="text-emerald-600 hover:text-emerald-700 font-medium text-sm"
                >
                  {showAllRange ? 'Show Less' : `View All (${rangeAllocationSummary.length})`}
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Recent Activities Section */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4 border-b pb-3">
            <div className="flex items-center gap-2.5">
              <div className="p-2 bg-emerald-50 text-emerald-700 rounded-lg">
                <History className="h-5 w-5" />
              </div>
              <div>
                <h3 className="text-lg font-bold text-gray-900 flex items-center gap-2">
                  <span>Recent Activities</span>
                  <span className="bg-emerald-50 text-emerald-700 text-xs font-semibold px-2.5 py-0.5 rounded-full border border-emerald-200">
                    Latest 10 Actions
                  </span>
                </h3>
                <p className="text-xs text-gray-500">
                  Real-time timeline of recent expenditures recorded and budget allocations updated.
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 text-xs font-medium text-gray-500 self-start sm:self-auto">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-gray-50 rounded-md border border-gray-200 text-gray-600">
                <Clock className="w-3.5 h-3.5 text-gray-400" />
                <span>Sorted by Date (Descending)</span>
              </span>
            </div>
          </div>

          {/* Vertical list of 10 items */}
          <div className="divide-y divide-gray-100">
            {recentActivitiesList.length > 0 ? (
              recentActivitiesList.map((item, idx) => (
                <div
                  key={item.id || idx}
                  onClick={() => {
                    if (item.type === 'expenditure') {
                      setActiveTab('Expenditures');
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    } else {
                      setActiveTab('Allocations');
                      if (item.schemeName && item.schemeName !== 'N/A') {
                        setSearchTerm(item.schemeName);
                      }
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    }
                  }}
                  className="py-3 px-2 sm:px-3 hover:bg-gray-50/90 rounded-lg transition-colors cursor-pointer flex flex-col sm:flex-row sm:items-center justify-between gap-3 group"
                >
                  <div className="flex items-start gap-3 min-w-0">
                    {/* Action Icon Badge */}
                    <div className={`p-2 rounded-lg shrink-0 mt-0.5 sm:mt-0 ${
                      item.type === 'expenditure' ? 'bg-red-50 text-red-600' : 'bg-indigo-50 text-indigo-600'
                    }`}>
                      {item.type === 'expenditure' ? (
                        <TrendingDown className="w-4 h-4" />
                      ) : (
                        <MapPin className="w-4 h-4" />
                      )}
                    </div>

                    {/* Action Details */}
                    <div className="min-w-0 space-y-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="font-bold text-sm text-gray-900 group-hover:text-emerald-700 transition-colors">
                          {item.actionName}
                        </span>
                        <span className="font-extrabold text-sm text-gray-900 font-mono">
                          â‚¹{item.amount.toLocaleString('en-IN')}
                        </span>
                        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${item.statusColor}`}>
                          {item.statusLabel}
                        </span>
                      </div>

                      <div className="text-xs text-gray-600 flex flex-wrap items-center gap-x-2 gap-y-0.5">
                        <span className="font-semibold text-gray-800 flex items-center gap-1">
                          <Building2 className="w-3 h-3 text-gray-400" />
                          {item.rangeName}
                        </span>
                        <span className="text-gray-300">â€¢</span>
                        <span className="text-emerald-800 font-medium">{item.schemeName}</span>
                        {item.sectorName && (
                          <>
                            <span className="text-gray-300">â€¢</span>
                            <span className="text-gray-600">{item.sectorName}</span>
                          </>
                        )}
                        {item.soeName && (
                          <>
                            <span className="text-gray-300">â€¢</span>
                            <span className="text-gray-500 italic">({item.soeName})</span>
                          </>
                        )}
                        {item.details && (
                          <>
                            <span className="text-gray-300">â€¢</span>
                            <span className="text-gray-500 truncate max-w-xs">{item.details}</span>
                          </>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Timestamp and Arrow */}
                  <div className="flex items-center justify-between sm:justify-end gap-3 shrink-0 pl-11 sm:pl-0">
                    <div className="text-left sm:text-right">
                      <span className="text-[11px] font-mono text-gray-500 block">
                        {item.dateStr}
                      </span>
                      <span className="text-[10px] text-gray-400 capitalize">
                        {item.type}
                      </span>
                    </div>
                    <ChevronRight className="w-4 h-4 text-gray-400 group-hover:text-emerald-600 group-hover:translate-x-0.5 transition-all" />
                  </div>
                </div>
              ))
            ) : (
              <div className="py-8 text-center text-gray-400 text-sm italic">
                No recent activities recorded yet.
              </div>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 lg:col-span-2">
            <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
              <FileText className="h-5 w-5 text-gray-500" /> Latest Expenditures
            </h3>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[700px]">
                <thead>
                  <tr className="bg-gray-50 text-gray-600 text-sm">
                    <th className="p-3 border-b">Date</th>
                    <th className="p-3 border-b">Range</th>
                    <th className="p-3 border-b">SOE</th>
                    <th className="p-3 border-b text-right">Approval ID</th>
                    <th className="p-3 border-b text-right">Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {currentExpenses.slice().reverse().slice(0, 5).map((exp) => {
                    const alloc = allocations.find(a => a.id === exp.allocationId);
                    const range = ranges.find(r => r.id === alloc?.rangeId);
                    const soe = soes.find(s => s.id === exp.soeId);
                    return (
                      <tr key={exp.id} className="border-b last:border-0 hover:bg-gray-50">
                        <td className="p-3">{exp.date ? exp.date.split('-').reverse().join('/') : ''}</td>
                        <td className="p-3 font-medium">{range?.name}</td>
                        <td className="p-3 text-gray-600">{soe?.name}</td>
                        <td className="p-3 text-right font-mono text-xs">{exp.approvalId ? `#${exp.approvalId}` : '-'}</td>
                        <td className="p-3 text-right font-bold text-red-600">â‚¹{exp.amount.toLocaleString()}</td>
                      </tr>
                    );
                  })}
                  {expenses.length === 0 && <tr><td colSpan={4} className="p-4 text-center text-gray-500">No expenditures yet.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderMyRangeSummaryTable = () => {
    // Group currentAllocations by SOE Head
    const summaryMap: Record<string, { hierarchy: string, soeName: string, allocated: number, spent: number, remaining: number }> = {};
    
    currentAllocations.forEach(alloc => {
      alloc.fundedSOEs?.forEach(f => {
        const soe = soes.find(s => s.id === f.soeId);
        if (!soe) return;
        
        let hierarchy = '';
        if (alloc.subActivityId) {
          const sa = subActivities.find(sa => sa.id === alloc.subActivityId);
          const act = activities.find(a => a.id === sa?.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' > ');
        } else if (alloc.activityId) {
          const act = activities.find(a => a.id === alloc.activityId);
          const sec = sectors.find(sec => sec.id === act?.sectorId);
          const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
          hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' > ');
        }

        const spent = currentExpenses.filter(e => e.allocationId === alloc.id && e.soeId === f.soeId).reduce((sum, e) => sum + e.amount, 0);
        const key = `${alloc.id}-${f.soeId}`;

        if (summaryMap[key]) {
          summaryMap[key].allocated += f.amount;
          summaryMap[key].spent += spent;
          summaryMap[key].remaining = summaryMap[key].allocated - summaryMap[key].spent;
        } else {
          summaryMap[key] = {
            hierarchy,
            soeName: soe.name,
            allocated: f.amount,
            spent,
            remaining: f.amount - spent
          };
        }
      });
    });

    const summaryData = Object.values(summaryMap).sort((a, b) => a.hierarchy.localeCompare(b.hierarchy) || a.soeName.localeCompare(b.soeName));

    return (
      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-6">
        <div className="flex justify-between items-center mb-4 border-b pb-2">
          <h3 className="text-lg font-semibold flex items-center gap-2">
            <Activity className="h-5 w-5 text-emerald-600" /> My Range Summary
          </h3>
        </div>
        <div className="overflow-x-auto max-h-80">
          <table className="w-full text-left border-collapse text-sm min-w-[800px]">
            <thead className="sticky top-0 bg-white shadow-sm">
              <tr className="bg-gray-50 text-gray-600 font-semibold">
                <th className="p-3 border-b">Hierarchy</th>
                <th className="p-3 border-b">SOE Head</th>
                <th className="p-3 border-b text-right">Allocated to Me</th>
                <th className="p-3 border-b text-right">My Expenditure</th>
                <th className="p-3 border-b text-right">My Remaining Balance</th>
              </tr>
            </thead>
            <tbody>
              {summaryData.map((item, idx) => (
                <tr key={idx} className="border-b hover:bg-gray-50 transition-colors">
                  <td className="p-3 text-xs text-gray-500">{item.hierarchy || 'N/A'}</td>
                  <td className="p-3 font-medium text-gray-800">{item.soeName}</td>
                  <td className="p-3 text-right text-blue-600">â‚¹{item.allocated.toLocaleString()}</td>
                  <td className="p-3 text-right text-red-600">â‚¹{item.spent.toLocaleString()}</td>
                  <td className={`p-3 text-right font-bold ${item.remaining > 0 ? 'text-emerald-600' : 'text-gray-400'}`}>
                    â‚¹{item.remaining.toLocaleString()}
                  </td>
                </tr>
              ))}
              {summaryData.length === 0 && (
                <tr>
                  <td colSpan={5} className="p-4 text-center text-gray-500">No allocations found for your unit.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const renderSoeAbstractTable = () => (
    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-6">
      <div 
        className="flex justify-between items-center mb-4 border-b pb-2 cursor-pointer hover:bg-gray-50 -mx-6 px-6"
        onClick={() => setIsSoeTrackerExpanded(!isSoeTrackerExpanded)}
      >
        <h3 className="text-lg font-semibold flex items-center gap-2">
          <Activity className="h-5 w-5 text-emerald-600" /> Live SOE Budget Tracker
        </h3>
        <div className="flex items-center gap-4">
          {isSoeTrackerExpanded && (
            <div className="relative" onClick={(e) => e.stopPropagation()}>
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
              <input
                type="text"
                placeholder="Search tracker..."
                className="pl-9 pr-4 py-1.5 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 w-64"
                value={soeSearchTerm}
                onChange={(e) => setSoeSearchTerm(e.target.value)}
              />
            </div>
          )}
          <button type="button" className="text-gray-500 hover:text-gray-700">
            {isSoeTrackerExpanded ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
          </button>
        </div>
      </div>
      
      {isSoeTrackerExpanded && (
        <div className="overflow-x-auto max-h-80">
          <table className="w-full text-left border-collapse text-sm min-w-[1000px]">
            <thead className="sticky top-0 bg-white shadow-sm">
              <tr className="bg-gray-50 text-gray-600 font-semibold">
                <th className="p-3 border-b">Hierarchy</th>
                <th className="p-3 border-b">SOE Head</th>
                <th className="p-3 border-b text-right">Approved Budget</th>
                <th className="p-3 border-b text-right">Received in Try</th>
                <th className="p-3 border-b text-right">Allocated</th>
                <th className="p-3 border-b text-right">To Be Allocated</th>
                <th className="p-3 border-b text-right">Try Balance</th>
                <th className="p-3 border-b text-right">Spent</th>
                <th className="p-3 border-b text-right">Remaining to Spend</th>
              </tr>
            </thead>
            <tbody>
              {soeAbstractForAllocations.map((item) => (
                <tr key={item.id} className="border-b hover:bg-gray-50 transition-colors">
                  <td className="p-3 text-xs text-gray-500">{item.hierarchy || 'N/A'}</td>
                  <td className="p-3 font-medium text-gray-800">{item.soeName}</td>
                  <td className="p-3 text-right text-gray-700">â‚¹{item.approvedBudget.toLocaleString()}</td>
                  <td className="p-3 text-right text-indigo-600">â‚¹{item.receivedInTry.toLocaleString()}</td>
                  <td className="p-3 text-right text-blue-600">â‚¹{item.allocated.toLocaleString()}</td>
                  <td className={`p-3 text-right font-bold ${item.toBeAllocated > 0 ? 'text-emerald-600' : 'text-gray-400'}`}>
                    â‚¹{item.toBeAllocated.toLocaleString()}
                  </td>
                  <td className={`p-3 text-right font-bold ${item.tryBalance >= 0 ? 'text-emerald-600' : 'text-red-600'}`}>
                    â‚¹{item.tryBalance.toLocaleString()}
                  </td>
                  <td className="p-3 text-right text-red-600">â‚¹{item.spent.toLocaleString()}</td>
                  <td className={`p-3 text-right font-bold ${item.remainingToSpend > 0 ? 'text-emerald-600' : 'text-gray-400'}`}>
                    â‚¹{item.remainingToSpend.toLocaleString()}
                  </td>
                </tr>
              ))}
              {soeAbstractForAllocations.length === 0 && (
                <tr>
                  <td colSpan={7} className="p-4 text-center text-gray-500">No SOE data matching current selection.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  const getHierarchyText = (item: any) => {
    let hierarchy = '';
    if (item.subActivityId) {
      const sa = subActivities.find(sa => sa.id === item.subActivityId);
      const act = activities.find(a => a.id === sa?.activityId);
      const sec = sectors.find(sec => sec.id === act?.sectorId);
      const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
      hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' -> ');
    } else if (item.activityId) {
      const act = activities.find(a => a.id === item.activityId);
      const sec = sectors.find(sec => sec.id === act?.sectorId);
      const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
      hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' -> ');
    } else if (item.sectorId) {
      const sec = sectors.find(sec => sec.id === item.sectorId);
      const sch = schemes.find(sc => sc.id === sec?.schemeId);
      hierarchy = [sch?.name, sec?.name].filter(Boolean).join(' -> ');
    } else if (item.schemeId) {
      const sch = schemes.find(sc => sc.id === item.schemeId);
      hierarchy = sch?.name || '';
    }
    return hierarchy || 'Global (No Hierarchy)';
  };

  const formatHierarchyText = (item: any) => {
    const sch = schemes.find(sc => sc.id === item.schemeId);
    const sec = sectors.find(sec => sec.id === item.sectorId);
    const act = activities.find(a => a.id === item.activityId);
    const sa = subActivities.find(sa => sa.id === item.subActivityId);

    const parts = [];
    if (sch) parts.push(`Scheme: ${sch.name}`);
    if (sec) parts.push(`Sector: ${sec.name}`);
    if (act) parts.push(`Activity: ${act.name}`);
    if (sa) parts.push(`Sub-Activity: ${sa.name}`);

    return parts.length > 0 ? parts.join(' | ') : 'Global (No Hierarchy)';
  };

  const renderHierarchy = (item: any) => {
    return <span className="text-xs text-gray-500">{formatHierarchyText(item)}</span>;
  };



  const handleSaveReconciliation = async (allocationId: string) => {
    const distribution = reconData[allocationId] || {};
    const allocation = baseAllocations.find(a => a.id === allocationId);
    if (!allocation) return;

    const totalDistributed = Object.values(distribution).reduce<number>((sum, val) => sum + (parseFloat(val as string) || 0), 0);
    
    if (Math.abs(totalDistributed - allocation.amount) > 0.01) {
      showAlert("Total distributed amount must match the allocated amount.");
      return;
    }

    try {
      const fundedSOEs = Object.entries(distribution)
        .filter(([_, amount]) => (parseFloat(amount as string) || 0) > 0)
        .map(([soeId, amount]) => ({ soeId, amount: parseFloat(amount as string) }));

      await updateDoc(doc(db, 'allocations', allocationId), {
        fundedSOEs,
        status: 'Funded'
      });
      
      const newReconData = { ...reconData };
      delete newReconData[allocationId];
      setReconData(newReconData);
      
      showAlert("Reconciliation saved successfully!");
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'allocations');
    }
  };

  const renderBudgetTracker = () => {
    // Flatten all SOEs with their hierarchy for the table
    const flattenedData = currentSoes.map(soe => {
      const scheme = currentSchemes.find(s => s.id === soe.schemeId);
      const sector = currentSectors.find(s => s.id === soe.sectorId);
      
      const sanctioned = getApprovedBudget(soe);
      const approved = getReceivedInTry(soe);
      const allocated = getSoeAllocated(soe.id);
      const balance = approved - allocated;

      return {
        id: soe.id,
        schemeName: scheme?.name || 'Unknown Scheme',
        sectorName: sector?.name || 'Unknown Sector',
        soeName: soe.name,
        sanctioned,
        approved,
        allocated,
        balance
      };
    }).filter(item => item.sanctioned > 0 || item.approved > 0 || item.allocated > 0);

    // Apply search filter
    const filteredData = flattenedData.filter(item => {
      const searchStr = trackerSearch.toLowerCase() || searchTerm.toLowerCase();
      return (
        item.schemeName.toLowerCase().includes(searchStr) ||
        item.sectorName.toLowerCase().includes(searchStr) ||
        item.soeName.toLowerCase().includes(searchStr) ||
        item.sanctioned.toString().includes(searchStr) ||
        item.approved.toString().includes(searchStr) ||
        item.allocated.toString().includes(searchStr) ||
        item.balance.toString().includes(searchStr)
      );
    });

    return (
      <div className="bg-white rounded-xl shadow-sm border border-emerald-100 overflow-hidden mb-6">
        <div 
          className="bg-emerald-50 p-4 flex items-center justify-between cursor-pointer hover:bg-emerald-100 transition-colors"
          onClick={() => setIsSoeTrackerExpanded(!isSoeTrackerExpanded)}
        >
          <div className="flex items-center gap-2">
            <div className="bg-emerald-600 p-1.5 rounded-lg text-white">
              <TrendingDown className="w-4 h-4" />
            </div>
            <div>
              <h3 className="font-bold text-emerald-900 text-sm">Live Budget Tracker (Abstract Table)</h3>
              <p className="text-[10px] text-emerald-600 font-medium uppercase tracking-wider">Real-time SOE-wise Allocation Status</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <div className="relative" onClick={(e) => e.stopPropagation()}>
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-emerald-500" />
              <input
                type="text"
                placeholder="Search budget details..."
                value={trackerSearch}
                onChange={(e) => setTrackerSearch(e.target.value)}
                className="pl-8 pr-3 py-1.5 text-xs border border-emerald-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 w-48 md:w-64 bg-white"
              />
            </div>
            <span className="text-[10px] bg-emerald-200 text-emerald-800 px-2 py-0.5 rounded-full font-bold">
              {filteredData.length} Records
            </span>
            {isSoeTrackerExpanded ? <ChevronUp className="w-5 h-5 text-emerald-600" /> : <ChevronDown className="w-5 h-5 text-emerald-600" />}
          </div>
        </div>

        {isSoeTrackerExpanded && (
          <div className="p-0 overflow-x-auto animate-in fade-in slide-in-from-top-2 duration-300">
            <table className="w-full text-left border-collapse min-w-[1000px]">
              <thead>
                <tr className="bg-emerald-600 text-white text-[11px] uppercase tracking-wider">
                  <th className="px-4 py-3 font-bold border-r border-emerald-500">Scheme</th>
                  <th className="px-4 py-3 font-bold border-r border-emerald-500">Sector</th>
                  <th className="px-4 py-3 font-bold border-r border-emerald-500">SOE Head</th>
                  <th className="px-4 py-3 font-bold border-r border-emerald-500 text-right">Total Sanction (Approved)</th>
                  <th className="px-4 py-3 font-bold border-r border-emerald-500 text-right">Received Budget (Try)</th>
                  <th className="px-4 py-3 font-bold border-r border-emerald-500 text-right">Allocated to Ranges</th>
                  <th className="px-4 py-3 font-bold text-right">Balance to Allocate</th>
                </tr>
              </thead>
              <tbody className="text-[11px]">
                {filteredData.length > 0 ? (
                  filteredData.map((item, idx) => (
                    <tr key={item.id} className={`${idx % 2 === 0 ? 'bg-white' : 'bg-emerald-50/30'} hover:bg-emerald-100/50 transition-colors border-b border-emerald-50`}>
                      <td className="px-4 py-2.5 font-medium text-gray-900 border-r border-emerald-50 min-w-[150px]">{item.schemeName}</td>
                      <td className="px-4 py-2.5 text-gray-600 border-r border-emerald-50 min-w-[150px]">{item.sectorName}</td>
                      <td className="px-4 py-2.5 font-bold text-emerald-800 border-r border-emerald-50">{item.soeName}</td>
                      <td className="px-4 py-2.5 text-right font-semibold text-gray-500 border-r border-emerald-50">â‚¹{item.sanctioned.toLocaleString()}</td>
                      <td className="px-4 py-2.5 text-right font-semibold text-gray-900 border-r border-emerald-50">â‚¹{item.approved.toLocaleString()}</td>
                      <td className="px-4 py-2.5 text-right font-semibold text-blue-600 border-r border-emerald-50">â‚¹{item.allocated.toLocaleString()}</td>
                      <td className={`px-4 py-2.5 text-right font-bold ${item.balance > 0 ? 'text-orange-600' : 'text-emerald-600'}`}>
                        â‚¹{item.balance.toLocaleString()}
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={7} className="px-4 py-8 text-center text-gray-400 italic">
                      No budget records found matching your search.
                    </td>
                  </tr>
                )}
              </tbody>
              <tfoot className="bg-emerald-50 font-bold text-emerald-900 text-[11px]">
                <tr>
                  <td colSpan={3} className="px-4 py-3 text-right uppercase tracking-wider border-r border-emerald-100">Grand Total</td>
                  <td className="px-4 py-3 text-right border-r border-emerald-100">â‚¹{filteredData.reduce((sum, i) => sum + i.sanctioned, 0).toLocaleString()}</td>
                  <td className="px-4 py-3 text-right border-r border-emerald-100">â‚¹{filteredData.reduce((sum, i) => sum + i.approved, 0).toLocaleString()}</td>
                  <td className="px-4 py-3 text-right border-r border-emerald-100">â‚¹{filteredData.reduce((sum, i) => sum + i.allocated, 0).toLocaleString()}</td>
                  <td className="px-4 py-3 text-right">â‚¹{filteredData.reduce((sum, i) => sum + i.balance, 0).toLocaleString()}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>
    );
  };

  const handleSaveAllReconciliation = async () => {
    const provisionalSchemes = currentSchemes.filter(scheme => 
      soes.some(soe => soe.schemeId === scheme.id && soe.isProvisional)
    );

    const allocationsToSave = baseAllocations.filter(a => 
      (reconSchemeId === 'all' ? provisionalSchemes.some(ps => ps.id === a.schemeId) : a.schemeId === reconSchemeId) && a.status === 'Pending SOE Funds'
    );

    // Validate all variations are 0
    const invalid = allocationsToSave.some(alloc => {
      const distribution = reconData[alloc.id] || {};
      const totalDistributed = Object.values(distribution).reduce<number>((sum, val) => sum + (parseFloat(val as string) || 0), 0);
      return Math.abs(alloc.amount - totalDistributed) >= 0.01;
    });

    if (invalid) {
      showAlert("All rows must have zero variation before saving.");
      return;
    }

    setLoading(true);
    try {
      const promises = allocationsToSave.map(alloc => {
        const distribution = reconData[alloc.id] || {};
        const fundedSOEs = Object.entries(distribution)
          .filter(([_, amount]) => (parseFloat(amount as string) || 0) > 0)
          .map(([soeId, amount]) => ({ soeId, amount: parseFloat(amount as string) }));

        return updateDoc(doc(db, 'allocations', alloc.id), {
          fundedSOEs,
          status: 'Funded'
        });
      });

      await Promise.all(promises);
      setReconData({});
      showAlert("All reconciliations saved successfully!");
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'allocations');
    } finally {
      setLoading(false);
    }
  };

  const renderApprovalModal = () => {
    if (!isApprovalModalOpen || !selectedExpenseForApproval) return null;

    const handleConfirm = () => {
      showConfirm(`Are you sure you want to ${approvalStatus} this expenditure? This action will lock the entry.`, async () => {
        await handleUpdateExpenseStatus(selectedExpenseForApproval.id, approvalStatus, true, approvalReason);
        setIsApprovalModalOpen(false);
        setSelectedExpenseForApproval(null);
        setApprovalReason('');
      });
    };

    return (
      <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[100] p-4">
        <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200">
          <div className="bg-emerald-600 p-4 text-white flex justify-between items-center">
            <h3 className="font-bold">Expenditure Action</h3>
            <button onClick={() => setIsApprovalModalOpen(false)}><X className="w-5 h-5" /></button>
          </div>
          <div className="p-6 space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Action</label>
              <select 
                value={approvalStatus} 
                onChange={(e) => setApprovalStatus(e.target.value as 'approved' | 'rejected')}
                className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
              >
                <option value="approved">Accept / Approve</option>
                <option value="rejected">Reject</option>
              </select>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Reason / Remarks</label>
              <textarea 
                value={approvalReason}
                onChange={(e) => setApprovalReason(e.target.value)}
                placeholder="Enter reason for this action..."
                className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none min-h-[100px]"
              />
            </div>
            <div className="flex gap-3 pt-2">
              <button 
                onClick={handleConfirm}
                className={`flex-1 py-2 rounded-lg font-bold text-white transition-colors ${approvalStatus === 'approved' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-red-600 hover:bg-red-700'}`}
              >
                Confirm {approvalStatus === 'approved' ? 'Approval' : 'Rejection'}
              </button>
              <button 
                onClick={() => setIsApprovalModalOpen(false)}
                className="flex-1 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg font-bold transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderSoeExpModal = () => {
    if (!viewingSoeExp) return null;

    const relevantExpenses = expenses.filter(e => e.soeId === viewingSoeExp.soeId && e.status !== 'rejected');
    const total = relevantExpenses.reduce((sum, e) => sum + e.amount, 0);

    return (
      <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[200] p-4">
        <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl overflow-hidden animate-in zoom-in duration-200">
          <div className="bg-emerald-600 p-4 text-white flex justify-between items-center">
            <div>
              <h3 className="font-bold">Expenditure Details</h3>
              <div className="text-[10px] opacity-90 font-medium uppercase tracking-wider">{viewingSoeExp.hierarchy} | {viewingSoeExp.soeName}</div>
            </div>
            <button onClick={() => setViewingSoeExp(null)} className="hover:bg-emerald-700 p-1 rounded-full transition-colors">
              <X className="w-5 h-5" />
            </button>
          </div>
          <div className="p-6">
            <div className="overflow-y-auto max-h-[60vh]">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-600 text-[10px] uppercase font-bold">
                    <th className="p-2 border-b">Date</th>
                    <th className="p-2 border-b">Approval ID</th>
                    <th className="p-2 border-b">Description</th>
                    <th className="p-2 border-b text-right">Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {relevantExpenses.map(exp => (
                    <tr key={exp.id} className="border-b hover:bg-gray-50 text-xs">
                      <td className="p-2">{exp.date ? exp.date.split('-').reverse().join('/') : ''}</td>
                      <td className="p-2 font-mono text-gray-500">{exp.approvalId ? `#${exp.approvalId}` : '-'}</td>
                      <td className="p-2">{exp.description}</td>
                      <td className="p-2 text-right font-bold text-red-600">â‚¹{exp.amount.toLocaleString()}</td>
                    </tr>
                  ))}
                  {relevantExpenses.length === 0 && (
                    <tr>
                      <td colSpan={4} className="p-8 text-center text-gray-500 italic">No expenditures found for this SOE Head.</td>
                    </tr>
                  )}
                </tbody>
                {relevantExpenses.length > 0 && (
                  <tfoot>
                    <tr className="bg-gray-50 font-bold">
                      <td colSpan={3} className="p-2 text-right text-gray-700">TOTAL EXPENDITURE:</td>
                      <td className="p-2 text-right text-red-700">â‚¹{total.toLocaleString()}</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
            <div className="mt-6 flex justify-end">
              <button 
                onClick={() => setViewingSoeExp(null)}
                className="bg-gray-100 text-gray-600 px-4 py-2 rounded-lg hover:bg-gray-200 transition-colors font-medium text-sm"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderFundingModal = () => {
    if (!fundingAllocation) return null;

    const alreadyFundedTotal = fundingAllocation.fundedSOEs?.reduce((sum, f) => sum + f.amount, 0) || 0;
    const remainingToFund = fundingAllocation.amount - alreadyFundedTotal;

    // Filter SOEs that are relevant to this allocation's hierarchy
    const relevantSoes = currentSoes.filter(s => {
      if (s.schemeId !== fundingAllocation.schemeId) return false;
      if (fundingAllocation.sectorId && s.sectorId !== fundingAllocation.sectorId) return false;
      if (fundingAllocation.activityId && s.activityId !== fundingAllocation.activityId) return false;
      if (fundingAllocation.subActivityId && s.subActivityId !== fundingAllocation.subActivityId) return false;
      return true;
    });

    return (
      <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
        <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
          <div className="bg-emerald-600 p-4 text-white flex justify-between items-center">
            <h3 className="font-bold flex items-center gap-2">
              <Landmark className="w-5 h-5" />
              Assign SOE Funds
            </h3>
            <button onClick={() => setFundingAllocation(null)} className="hover:bg-white/20 rounded-full p-1">
              <X className="w-5 h-5" />
            </button>
          </div>
          
          <div className="p-6 space-y-4">
            <div className="bg-gray-50 p-4 rounded-xl border border-gray-100">
              <div className="text-xs text-gray-500 uppercase font-bold mb-2">Allocation Details</div>
              <div className="text-sm font-medium">{renderHierarchy(fundingAllocation)}</div>
              <div className="text-xs text-gray-400 mt-1">Range: {ranges.find(r => r.id === fundingAllocation.rangeId)?.name}</div>
              <div className="mt-3 flex justify-between items-end">
                <div>
                  <div className="text-[10px] text-gray-400 uppercase">Sanctioned</div>
                  <div className="font-bold text-gray-900">â‚¹{fundingAllocation.amount.toLocaleString()}</div>
                </div>
                <div className="text-right">
                  <div className="text-[10px] text-gray-400 uppercase">Remaining to Fund</div>
                  <div className="font-bold text-emerald-600">â‚¹{remainingToFund.toLocaleString()}</div>
                </div>
              </div>
            </div>

            <form onSubmit={handleFundAllocation} className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Select SOE Head (Budget Source)</label>
                <select name="soeId" required className="w-full p-2 border rounded text-sm">
                  <option value="">Select SOE</option>
                  {relevantSoes.map(s => {
                    const totalReceived = getReceivedInTry(s);
                    const totalFundedFromThisSoe = baseAllocations
                      .reduce((sum, a) => {
                        const funded = a.fundedSOEs?.find(f => f.soeId === s.id);
                        return sum + (funded?.amount || 0);
                      }, 0);
                    const available = totalReceived - totalFundedFromThisSoe;
                    return (
                      <option key={s.id} value={s.id} disabled={available <= 0}>
                        {s.name} (Available: â‚¹{available.toLocaleString()})
                      </option>
                    );
                  })}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Amount to Assign (â‚¹)</label>
                <input 
                  name="amount" 
                  type="number" 
                  step="0.01"
                  max={remainingToFund}
                  required 
                  placeholder="Enter amount"
                  className="w-full p-2 border rounded text-sm" 
                />
              </div>

              <div className="flex gap-3 pt-2">
                <button 
                  type="button" 
                  onClick={() => setFundingAllocation(null)}
                  className="flex-1 px-4 py-2 border rounded-lg text-sm font-medium hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  className="flex-1 px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 flex items-center justify-center gap-2"
                >
                  <Check className="w-4 h-4" />
                  Confirm Funding
                </button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  };

  const handleSurrender = async (e: any) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const rangeId = formData.get('rangeId') as string;
    const schemeId = formData.get('schemeId') as string;
    const sectorId = formData.get('sectorId') as string;
    const activityId = formData.get('activityId') as string;
    const subActivityId = formData.get('subActivityId') as string;
    const soeId = formData.get('soeId') as string;
    const amount = Number(formData.get('amount'));
    const remarks = formData.get('remarks') as string;
    const date = formData.get('date') as string;

    if (!rangeId || !soeId || amount <= 0) {
      showAlert("Please fill all required fields and enter a valid amount.");
      return;
    }

    const nSchemeId = schemeId || null;
    const nSectorId = sectorId || null;
    const nActivityId = activityId || null;
    const nSubActivityId = subActivityId || null;

    const rangeAllocs = allocations.filter(a => 
      a.rangeId === rangeId && 
      a.schemeId === nSchemeId && 
      (!nSectorId || !a.sectorId || a.sectorId === nSectorId) &&
      (!nActivityId || !a.activityId || a.activityId === nActivityId) &&
      (!nSubActivityId || !a.subActivityId || a.subActivityId === nSubActivityId) &&
      a.fundedSOEs.some(s => s.soeId === soeId)
    );

    if (rangeAllocs.length === 0) {
      showAlert("No allocation found for this selection in the source unit.");
      return;
    }

    // Find the best allocation to surrender from (one that has enough balance)
    let rangeAlloc = rangeAllocs.find(a => {
      const soeFund = a.fundedSOEs.find(s => s.soeId === soeId);
      if (!soeFund) return false;
      const spentForSoe = expenses
        .filter(e => e.allocationId === a.id && e.soeId === soeId && e.status !== 'rejected')
        .reduce((sum, e) => sum + e.amount, 0);
      return (soeFund.amount - amount) >= spentForSoe;
    });

    if (!rangeAlloc) {
      // If none has enough individually, pick the one with the most balance
      rangeAlloc = rangeAllocs.sort((a, b) => {
        const getBal = (alloc: any) => {
          const soeFund = alloc.fundedSOEs.find((s: any) => s.soeId === soeId);
          if (!soeFund) return 0;
          const spent = expenses
            .filter((e: any) => e.allocationId === alloc.id && e.soeId === soeId && e.status !== 'rejected')
            .reduce((sum: number, e: any) => sum + e.amount, 0);
          return soeFund.amount - spent;
        };
        return getBal(b) - getBal(a);
      })[0];
    }

    const soeFund = rangeAlloc.fundedSOEs.find(s => s.soeId === soeId);
    if (!soeFund) {
      showAlert("No funds found for this SOE in the selected allocation.");
      return;
    }

    // Validation: Ensure surrender doesn't leave allocation below expenditure for this SOE
    const spentForSoe = expenses
      .filter(e => e.allocationId === rangeAlloc.id && e.soeId === soeId && e.status !== 'rejected')
      .reduce((sum, e) => sum + e.amount, 0);

    if (soeFund.amount - amount < spentForSoe) {
      showAlert(`Cannot surrender â‚¹${amount.toLocaleString()}. Remaining SOE budget (â‚¹${(soeFund.amount - amount).toLocaleString()}) would be less than expenditure (â‚¹${spentForSoe.toLocaleString()}) for this SOE.`);
      return;
    }

    try {
      const activeFy = fys.find(f => f.name === selectedFY || f.id === selectedFY);
      const fyId = activeFy ? activeFy.id : selectedFY;

      // 1. Add Surrender record
      await addDoc(collection(db, 'surrenders'), {
        rangeId, 
        schemeId: nSchemeId, 
        sectorId: nSectorId, 
        activityId: nActivityId, 
        subActivityId: nSubActivityId, 
        soeId,
        amount, date, remarks, fyId, financialYear: selectedFY,
        createdAt: Date.now(), updatedAt: Date.now()
      });

      // 2. Decrease Source Unit Allocation (Optional: Log only, don't mutate original allocation to maintain audit trail)
      // We no longer mutate the allocation record directly to avoid double-counting issues.
      // The budget logic now correctly subtracts surrenders from the gross allocation.
      
      showAlert("Amount surrendered successfully. It is now available for reallocation from the Sector-wide budget.");
      setEditingItem(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'surrenders');
    }
  };

  const dataUrlToBlob = (dataUrl: string): Blob | null => {
    try {
      const parts = dataUrl.split(',');
      if (parts.length < 2) return null;
      const mimeMatch = parts[0].match(/:(.*?);/);
      const mime = mimeMatch ? mimeMatch[1] : 'application/pdf';
      const bstr = atob(parts[1]);
      let n = bstr.length;
      const u8arr = new Uint8Array(n);
      while (n--) {
        u8arr[n] = bstr.charCodeAt(n);
      }
      return new Blob([u8arr], { type: mime });
    } catch (e) {
      console.error("Error converting data URL to blob:", e);
      return null;
    }
  };

  const readFileAsDataUrl = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = (err) => reject(err);
      reader.readAsDataURL(file);
    });
  };

  const handleViewFile = async (fileObj: { url?: string; fileData?: string; name?: string; type?: string }) => {
    const fileSource = fileObj.fileData || fileObj.url;
    if (!fileSource) {
      showAlert("File data or link is missing.");
      return;
    }

    if (fileSource.startsWith('data:')) {
      const blob = dataUrlToBlob(fileSource);
      if (blob) {
        const blobUrl = URL.createObjectURL(blob);
        const newWin = window.open(blobUrl, '_blank');
        if (!newWin) {
          const a = document.createElement('a');
          a.href = blobUrl;
          a.target = '_blank';
          a.click();
        }
        return;
      }
    }

    if (fileSource.startsWith('http://') || fileSource.startsWith('https://')) {
      try {
        const response = await fetch(fileSource);
        if (!response.ok) {
          if (response.status === 402) {
            showAlert("Firebase Storage payment required (Error 402: Storage billing account delinquent). Please delete and re-upload this file so it embeds directly.");
            return;
          }
          throw new Error(`HTTP ${response.status}`);
        }
        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);
        window.open(blobUrl, '_blank');
      } catch (err) {
        console.warn("Direct fetch failed, falling back to window.open:", err);
        window.open(fileSource, '_blank');
      }
      return;
    }

    showAlert("Unable to open file.");
  };

  const handleDownloadFile = async (fileObj: { url?: string; fileData?: string; name?: string; type?: string }) => {
    const fileSource = fileObj.fileData || fileObj.url;
    if (!fileSource) {
      showAlert("File data or link is missing.");
      return;
    }

    const fileName = fileObj.name || 'document.pdf';

    if (fileSource.startsWith('data:')) {
      const blob = dataUrlToBlob(fileSource);
      if (blob) {
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
        return;
      }
    }

    if (fileSource.startsWith('http://') || fileSource.startsWith('https://')) {
      try {
        const response = await fetch(fileSource);
        if (!response.ok) {
          if (response.status === 402) {
            showAlert("Firebase Storage payment required (Error 402: Storage billing account delinquent). Please delete and re-upload this file so it embeds directly.");
            return;
          }
          throw new Error(`HTTP ${response.status}`);
        }
        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
      } catch (err) {
        console.warn("Fetch download failed, opening direct link:", err);
        const a = document.createElement('a');
        a.href = fileSource;
        a.download = fileName;
        a.target = '_blank';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
      }
      return;
    }

    showAlert("Unable to download file.");
  };

  const handleBudgetFileUpload = async (e: React.ChangeEvent<HTMLInputElement>, type: 'approved' | 'distributed') => {
    const files = e.target.files;
    if (!files || files.length === 0 || !selectedFY) return;
    
    const activeFy = fys.find(f => f.name === selectedFY || f.id === selectedFY);
    if (!activeFy) return;

    if (!budgetFileSelection.schemeId) {
      showAlert("Please select a scheme first.");
      return;
    }

    const fileList = Array.from(files);
    
    for (const file of fileList) {
      setUploadStatus(prev => ({
        ...prev,
        [type]: { isUploading: true, progress: 0, fileName: file.name, transferred: 0, total: file.size, error: null }
      }));

      try {
        let base64Data = '';
        if (file.size < 800 * 1024) {
          try {
            base64Data = await readFileAsDataUrl(file);
          } catch (readErr) {
            console.error("Error reading file as data URL:", readErr);
          }
        }

        let downloadURL = '';

        try {
          const storagePath = `budget_files/${type}/${Date.now()}_${file.name}`;
          const storageRef = ref(storage, storagePath);
          const uploadTask = uploadBytesResumable(storageRef, file);

          setUploadTasks(prev => ({ ...prev, [type]: uploadTask }));

          await new Promise<void>((resolve, reject) => {
            uploadTask.on('state_changed', 
              (snapshot) => {
                const progress = (snapshot.bytesTransferred / snapshot.totalBytes) * 100;
                setUploadStatus(prev => ({
                  ...prev,
                  [type]: { 
                    ...prev[type], 
                    progress, 
                    transferred: snapshot.bytesTransferred, 
                    total: snapshot.totalBytes 
                  }
                }));
              }, 
              (error) => {
                console.warn("Storage upload warning:", error);
                reject(error);
              }, 
              async () => {
                try {
                  const remoteUrl = await getDownloadURL(uploadTask.snapshot.ref);
                  downloadURL = remoteUrl;
                  resolve();
                } catch (e) {
                  console.warn("Could not fetch remote URL:", e);
                  reject(e);
                }
              }
            );
          });
        } catch (storageErr) {
          console.warn("Storage upload error:", storageErr);
        }

        let savedFileData = '';
        if (downloadURL && downloadURL.startsWith('http')) {
          savedFileData = '';
        } else if (base64Data && base64Data.length < 800000) {
          downloadURL = base64Data;
          savedFileData = base64Data;
        } else {
          throw new Error(`File upload failed. Storage upload did not complete and file size (${(file.size / (1024 * 1024)).toFixed(2)} MB) exceeds inline document limit.`);
        }

        const fileRecord: any = {
          name: file.name,
          url: downloadURL,
          fileData: savedFileData,
          schemeId: budgetFileSelection.schemeId,
          sectorId: budgetFileSelection.sectorId || '',
          type,
          uploadedBy: user?.uid,
          uploadedAt: Date.now(),
          fyId: activeFy.id,
          financialYear: activeFy.name
        };

        if (type === 'approved') {
          await addDoc(collection(db, 'approvedBudgetFiles'), fileRecord);
        } else {
          fileRecord.rangeId = budgetFileSelection.rangeId || '';
          await addDoc(collection(db, 'distributedBudgetFiles'), fileRecord);
        }

        // Live notification sync for newly uploaded budget distribution / approved budget document
        try {
          const schObj = schemes.find(s => s.id === budgetFileSelection.schemeId);
          const secObj = sectors.find(s => s.id === budgetFileSelection.sectorId);
          const rgObj = ranges.find(r => r.id === budgetFileSelection.rangeId);
          const rangeNameLabel = rgObj?.name || (type === 'distributed' ? 'All Ranges' : 'All');

          await addDoc(collection(db, 'notifications'), {
            name: `${type === 'approved' ? 'Approved Budget Document' : 'Budget Allotment Distribution'}: ${file.name}`,
            description: `Scheme: ${schObj?.name || 'All'} | Sector: ${secObj?.name || 'General'} | Range: ${rangeNameLabel} | FY: ${activeFy.name}`,
            url: downloadURL,
            fileData: savedFileData,
            type: 'application/pdf',
            category: type === 'distributed' ? 'budget_distribution' : 'budget_allocation',
            schemeName: schObj?.name || '',
            sectorName: secObj?.name || '',
            rangeName: rangeNameLabel,
            targetTab: type === 'distributed' ? 'Distributed Budget' : 'Approved Budget',
            targetRanges: budgetFileSelection.rangeId && rgObj ? [rgObj.name, 'All'] : ['All'],
            createdAt: Date.now(),
            uploadedBy: user?.uid || 'Admin'
          });
        } catch (notifErr) {
          console.warn("Could not create notification for budget file upload:", notifErr);
        }
        
        setUploadStatus(prev => ({
          ...prev,
          [type]: { ...prev[type], isUploading: false, progress: 100, error: null }
        }));
        
        setTimeout(() => {
          setUploadStatus(prev => ({
            ...prev,
            [type]: { ...prev[type], progress: 0, fileName: '', transferred: 0, total: 0 }
          }));
        }, 3000);
      } catch (err: any) {
        console.error("Error saving budget file:", err);
        const errorMsg = err?.message || "Failed to save file.";
        showAlert(`${file.name}: ${errorMsg}`);
        setUploadStatus(prev => ({
          ...prev,
          [type]: { ...prev[type], isUploading: false, error: errorMsg }
        }));
      } finally {
        setUploadTasks(prev => {
          const newTasks = { ...prev };
          delete newTasks[type];
          return newTasks;
        });
      }
    }
    
    e.target.value = '';
  };

  const handleBudgetFileDelete = async (file: BudgetFile) => {
    showConfirm("Are you sure you want to delete this file?", async () => {
      try {
        // Delete from storage
        try {
          const storageRef = ref(storage, file.url);
          await deleteObject(storageRef);
        } catch (storageError) {
          console.warn("Storage delete error (continuing with firestore delete):", storageError);
        }

        // Delete from firestore
        const collectionName = file.type === 'approved' ? 'approvedBudgetFiles' : 'distributedBudgetFiles';
        await deleteDoc(doc(db, collectionName, file.id));
        
        showAlert("File deleted successfully!");
      } catch (error) {
        console.error("Delete error:", error);
        showAlert("Failed to delete file.");
      }
    });
  };

  const renderBudgetFilesTab = (type: 'approved' | 'distributed') => {
    const files = type === 'approved' ? approvedBudgetFiles : distributedBudgetFiles;

    const filteredFiles = files.filter(f => {
      if (budgetFileSelection.schemeId && f.schemeId !== budgetFileSelection.schemeId) return false;
      if (budgetFileSelection.sectorId && f.sectorId !== budgetFileSelection.sectorId) return false;
      if (type === 'distributed' && budgetFileSelection.rangeId && f.rangeId !== budgetFileSelection.rangeId) return false;
      
      const matchesSearch = f.name.toLowerCase().includes(budgetSearchTerm.toLowerCase());
      return matchesSearch;
    });

    const paginatedFiles = budgetItemsPerPage === 'All' 
      ? filteredFiles 
      : filteredFiles.slice((budgetPage - 1) * budgetItemsPerPage, budgetPage * budgetItemsPerPage);

    return (
      <div className="space-y-6">
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100">
          <div className="flex flex-col md:flex-row gap-4 items-end">
            <div className="flex-1">
              <CascadingDropdowns
                onSelectionChange={setBudgetFileSelection}
                schemes={schemes}
                sectors={sectors}
                activities={activities}
                subActivities={subActivities}
                ranges={ranges}
                allocations={allocations}
                expenses={expenses}
                soes={soes}
                showRange={type === 'distributed'}
                showSector={true}
                showActivity={false}
                showSubActivity={false}
                showSoe={false}
                userRole={userRole}
                userRangeId={userRangeId}
                showConfirm={showConfirm}
                type="BudgetView"
              />
            </div>
            <div className="flex flex-col gap-2 w-full md:w-auto">
              <div className="flex items-center gap-1">
                <div className="relative flex-1">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search files..."
                    value={budgetSearchTerm}
                    onChange={(e) => { setBudgetSearchTerm(e.target.value); setBudgetPage(1); }}
                    className="pl-9 pr-4 py-2 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 w-full md:w-64"
                  />
                </div>
                <button 
                  className="p-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors shadow-sm"
                  title="Search"
                >
                  <Search className="w-4 h-4" />
                </button>
              </div>
            </div>
            {userRole === 'admin' && (
              <div className="flex flex-col gap-2">
                {(uploadStatus[type].isUploading || uploadStatus[type].progress === 100 || uploadStatus[type].error) && (
                  <div className={`flex flex-col gap-2 min-w-[250px] p-3 rounded-lg border ${uploadStatus[type].error ? 'bg-red-50 border-red-100' : uploadStatus[type].progress === 100 ? 'bg-green-50 border-green-100' : 'bg-blue-50 border-blue-100'}`}>
                    <div className="flex justify-between items-start gap-2">
                      <div className="flex flex-col overflow-hidden">
                        <span className="text-[10px] font-bold text-gray-700 truncate">{uploadStatus[type].fileName}</span>
                        <span className="text-[9px] text-gray-500">
                          {(uploadStatus[type].transferred / (1024 * 1024)).toFixed(2)} MB / {(uploadStatus[type].total / (1024 * 1024)).toFixed(2)} MB
                        </span>
                      </div>
                      <div className="flex items-center gap-1">
                        <span className={`text-[10px] font-bold ${uploadStatus[type].error ? 'text-red-600' : 'text-blue-600'}`}>
                          {uploadStatus[type].error ? 'Error' : `${Math.round(uploadStatus[type].progress)}%`}
                        </span>
                        {!uploadStatus[type].isUploading && (
                          <button
                            type="button"
                            onClick={() => setUploadStatus(prev => ({
                              ...prev,
                              [type]: { isUploading: false, progress: 0, fileName: '', transferred: 0, total: 0, error: null }
                            }))}
                            className="text-gray-400 hover:text-gray-700 p-0.5 rounded cursor-pointer ml-1"
                            title="Dismiss notification"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </div>
                    
                    <div className="w-full bg-gray-200 rounded-full h-1.5 overflow-hidden">
                      <div 
                        className={`h-full transition-all duration-300 ease-out ${uploadStatus[type].error ? 'bg-red-500' : uploadStatus[type].progress === 100 ? 'bg-green-500' : 'bg-blue-500'}`}
                        style={{ width: `${uploadStatus[type].progress}%` }}
                      />
                    </div>

                    {uploadStatus[type].error && (
                      <span className="text-[9px] text-red-600 font-medium leading-tight">
                        {uploadStatus[type].error}
                      </span>
                    )}
                    
                    {uploadStatus[type].progress === 100 && !uploadStatus[type].error && !uploadStatus[type].isUploading && (
                      <span className="text-[9px] text-green-600 font-bold flex items-center gap-1">
                        <Check className="w-3 h-3" /> Upload Complete
                      </span>
                    )}
                  </div>
                )}
                
                <div className="relative">
                  <input
                    type="file"
                    accept=".pdf"
                    multiple
                    onClick={(e) => { (e.target as HTMLInputElement).value = ''; }}
                    onChange={(e) => handleBudgetFileUpload(e, type)}
                    className="hidden"
                    id={`file-upload-${type}`}
                    disabled={uploadStatus[type].isUploading}
                  />
                  <div className="flex items-center gap-2">
                    <label
                      htmlFor={`file-upload-${type}`}
                      className={`flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors cursor-pointer shadow-sm ${uploadStatus[type].isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}
                    >
                      {uploadStatus[type].isUploading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />}
                      <span>{uploadStatus[type].isUploading ? 'Uploading...' : 'Upload PDF'}</span>
                    </label>
                    {uploadStatus[type].isUploading && (
                      <button
                        onClick={() => uploadTasks[type]?.cancel()}
                        className="p-2 bg-red-50 text-red-600 border border-red-100 rounded-lg hover:bg-red-100 transition-colors"
                        title="Cancel Upload"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="p-4 border-b border-gray-100 bg-gray-50 flex justify-between items-center">
            <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider flex items-center gap-2">
              <FileText className="w-4 h-4 text-emerald-600" />
              {type === 'approved' ? 'Approved Budget Files' : 'Distributed Budget Files'}
            </h3>
            <span className="text-xs font-medium text-gray-500 bg-gray-200 px-2 py-1 rounded-full">
              {filteredFiles.length} Files
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50">
                  <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider border-b">File Name</th>
                  <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider border-b">Uploaded At</th>
                  <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider border-b text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {paginatedFiles.length > 0 ? (
                  paginatedFiles.map(file => (
                    <tr key={file.id} className="hover:bg-gray-50 transition-colors group">
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-3">
                          <div className="p-2 bg-red-50 text-red-600 rounded-lg">
                            <FileText className="w-4 h-4" />
                          </div>
                          <span className="text-sm font-medium text-gray-700 truncate max-w-xs">{file.name}</span>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-xs text-gray-500">
                        {new Date(file.uploadedAt).toLocaleString()}
                      </td>
                      <td className="px-4 py-3 text-right">
                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => handleViewFile(file)}
                            className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer"
                            title="View PDF"
                          >
                            <Eye className="w-4 h-4" />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDownloadFile(file)}
                            className="p-2 text-emerald-600 hover:bg-emerald-50 rounded-lg transition-colors cursor-pointer"
                            title="Download PDF"
                          >
                            <Download className="w-4 h-4" />
                          </button>
                          {userRole === 'admin' && (
                            <button
                              type="button"
                              onClick={() => handleBudgetFileDelete(file)}
                              className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                              title="Delete"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={3} className="px-4 py-12 text-center text-gray-400 italic text-sm">
                      No files found for the selected criteria.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          {filteredFiles.length > 0 && (
            <div className="p-4 bg-gray-50 border-t border-gray-100">
              <Pagination
                totalItems={filteredFiles.length}
                itemsPerPage={budgetItemsPerPage}
                currentPage={budgetPage}
                onPageChange={setBudgetPage}
                onItemsPerPageChange={setBudgetItemsPerPage}
              />
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderSurrenderTab = () => {
    const filteredSurrenders = surrenders.filter(s => {
      // Filter by user's range if applicable
      if (userRangeId && s.rangeId !== userRangeId) return false;

      const r = ranges.find(range => range.id === s.rangeId);
      const sch = schemes.find(scheme => scheme.id === s.schemeId);
      const soe = soes.find(soe => soe.id === s.soeId);

      const matchesSearch = 
        (r?.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (sch?.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (soe?.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (s.remarks || '').toLowerCase().includes(searchTerm.toLowerCase());

      const matchesFilters = 
        (!surrenderFilters.rangeId || s.rangeId === surrenderFilters.rangeId) &&
        (!surrenderFilters.schemeId || s.schemeId === surrenderFilters.schemeId) &&
        (!surrenderFilters.sectorId || soe?.sectorId === surrenderFilters.sectorId) &&
        (!surrenderFilters.activityId || soe?.activityId === surrenderFilters.activityId) &&
        (!surrenderFilters.subActivityId || soe?.subActivityId === surrenderFilters.subActivityId) &&
        (!surrenderFilters.soeId || s.soeId === surrenderFilters.soeId);

      return matchesSearch && matchesFilters;
    });

    return renderSimpleManager(
      'Surrender',
      filteredSurrenders,
      [
        { key: 'date', label: 'Date', render: (val) => val ? val.split('-').reverse().join('/') : '' },
        { key: 'rangeId', label: 'Hierarchy / Unit', render: (_, item) => {
          const r = ranges.find(r => r.id === item.rangeId);
          const hText = getHierarchyText(item);
          return (
            <div className="max-w-[200px]">
              <div className="font-bold text-gray-900 truncate leading-tight">{r?.name === 'Rajgarh Forest Division' ? 'Division' : r?.name}</div>
              <div className="text-[9px] text-gray-500 truncate" title={hText}>{hText}</div>
            </div>
          );
        }, searchableText: (_, item) => getHierarchyText(item) },
        { key: 'soeId', label: 'SOE', render: (val) => <span className="font-medium text-gray-700">{soes.find(s => s.id === val)?.name || 'N/A'}</span> },
        { key: 'amount', label: 'Amount', render: (val) => <span className="font-bold text-red-600">â‚¹{val.toLocaleString()}</span> },
        { key: 'remarks', label: 'Remarks', render: (val) => <div className="text-[10px] italic text-gray-500 max-w-[150px] whitespace-normal break-words" title={val}>{val || '-'}</div> }
      ],
      handleSurrender,
      (id) => handleDelete('surrenders', id),
      (
        <div className="space-y-3">
          <CascadingDropdowns 
            schemes={currentSchemes} sectors={currentSectors} activities={currentActivities} subActivities={currentSubActivities} soes={currentSoes} soeBudgets={[]} allocations={baseAllocations} surrenders={surrenders} ranges={ranges} expenses={currentExpenses}
            editingItem={editingItem} type="Surrender" userRangeId={userRangeId} userRole={userRole} showConfirm={showConfirm}
            onSelectionChange={setSurrenderFormSelection}
          >
            {surrenderBudgetStatus && (
              <div className="bg-emerald-50 p-3 rounded-lg border border-emerald-100 mb-2">
                <div className="grid grid-cols-3 gap-2">
                  <div className="text-center">
                    <p className="text-[9px] font-bold text-emerald-800 uppercase">Allocated</p>
                    <p className="text-xs font-bold text-emerald-700">â‚¹{surrenderBudgetStatus.allocated.toLocaleString()}</p>
                  </div>
                  <div className="text-center border-x border-emerald-100">
                    <p className="text-[9px] font-bold text-red-800 uppercase">Spent</p>
                    <p className="text-xs font-bold text-red-700">â‚¹{surrenderBudgetStatus.spent.toLocaleString()}</p>
                  </div>
                  <div className="text-center">
                    <p className="text-[9px] font-bold text-blue-800 uppercase">Balance</p>
                    <p className="text-xs font-bold text-blue-700">â‚¹{surrenderBudgetStatus.balance.toLocaleString()}</p>
                  </div>
                </div>
              </div>
            )}
            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Date</label>
              <input type="date" name="date" required className="w-full p-1.5 border rounded text-sm" defaultValue={editingItem?.type === 'Surrender' ? editingItem.item.date : new Date().toISOString().split('T')[0]} />
            </div>
            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Amount to Surrender</label>
              <input type="number" name="amount" required defaultValue={editingItem?.type === 'Surrender' ? editingItem.item.amount : ''} placeholder="Amount" className="w-full p-1.5 border rounded text-sm" />
            </div>
            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Remarks</label>
              <textarea name="remarks" defaultValue={editingItem?.type === 'Surrender' ? editingItem.item.remarks : ''} placeholder="Remarks" className="w-full p-1.5 border rounded text-sm" rows={2}></textarea>
            </div>
          </CascadingDropdowns>
        </div>
      ),
      (item) => setEditingItem({ type: 'Surrender', item }),
      undefined,
      undefined,
      null,
      false,
      undefined,
      undefined,
      (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-3">
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Range</label>
            <select 
              value={surrenderFilters.rangeId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, rangeId: e.target.value })}
              className="w-full p-1.5 border rounded text-xs bg-white"
              disabled={!!userRangeId}
            >
              {userRangeId ? (
                <option value={userRangeId}>{ranges.find(r => r.id === userRangeId)?.name}</option>
              ) : (
                <>
                  <option value="">All Ranges</option>
                  {ranges.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                </>
              )}
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Scheme</label>
            <select 
              value={surrenderFilters.schemeId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, schemeId: e.target.value, sectorId: '', activityId: '', subActivityId: '', soeId: '' })}
              className="w-full p-1.5 border rounded text-xs bg-white"
            >
              <option value="">All Schemes</option>
              {schemes.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Sector</label>
            <select 
              value={surrenderFilters.sectorId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, sectorId: e.target.value, activityId: '', subActivityId: '', soeId: '' })}
              className="w-full p-1.5 border rounded text-xs bg-white"
            >
              <option value="">All Sectors</option>
              {sectors.filter(s => !surrenderFilters.schemeId || s.schemeId === surrenderFilters.schemeId).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Activity</label>
            <select 
              value={surrenderFilters.activityId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, activityId: e.target.value, subActivityId: '', soeId: '' })}
              className="w-full p-1.5 border rounded text-xs bg-white"
            >
              <option value="">All Activities</option>
              {activities.filter(a => {
                if (surrenderFilters.sectorId) return a.sectorId === surrenderFilters.sectorId;
                if (surrenderFilters.schemeId) return a.schemeId === surrenderFilters.schemeId;
                return true;
              }).map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Sub-Activity</label>
            <select 
              value={surrenderFilters.subActivityId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, subActivityId: e.target.value, soeId: '' })}
              className="w-full p-1.5 border rounded text-xs bg-white"
            >
              <option value="">All Sub-Activities</option>
              {subActivities.filter(sa => !surrenderFilters.activityId || sa.activityId === surrenderFilters.activityId).map(sa => <option key={sa.id} value={sa.id}>{sa.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">SOE</label>
            <select 
              value={surrenderFilters.soeId}
              onChange={(e) => setSurrenderFilters({ ...surrenderFilters, soeId: e.target.value })}
              className="w-full p-1.5 border rounded text-xs bg-white"
            >
              <option value="">All SOEs</option>
              {soes.filter(s => {
                if (surrenderFilters.subActivityId) return s.subActivityId === surrenderFilters.subActivityId;
                if (surrenderFilters.activityId) return s.activityId === surrenderFilters.activityId;
                if (surrenderFilters.sectorId) return s.sectorId === surrenderFilters.sectorId;
                if (surrenderFilters.schemeId) return s.schemeId === surrenderFilters.schemeId;
                return true;
              }).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <div className="xl:col-span-6 flex justify-end">
            <button 
              onClick={() => setSurrenderFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' })}
              className="text-xs text-red-600 hover:text-red-800 font-medium flex items-center gap-1"
            >
              <X className="w-3 h-3" />
              Reset Filters
            </button>
          </div>
        </div>
      )
    );
  };

  const renderReconciliation = () => {
    const provisionalSchemes = currentSchemes.filter(scheme => 
      soes.some(soe => soe.schemeId === scheme.id && soe.isProvisional)
    );

    const allocationsToReconcile = baseAllocations.filter(a => 
      (reconSchemeId === 'all' ? currentSchemes.some(ps => ps.id === a.schemeId) : a.schemeId === reconSchemeId) && a.status === 'Pending SOE Funds'
    );

    // Define the 4 SOE columns as requested
    const targetSoeNames = ['20 OC', '36 M&W', '21 Maint', '30MV']; // Mapping 36 M&S to 36 M&W if that's what's in data
    const displaySoeNames = ['20 OC', '36 M&S', '21 Maint', '30 MV'];

    // Hierarchical grouping
    const hierarchy: any[] = [];
    if (reconSchemeId) {
      const schemesToProcess = reconSchemeId === 'all' 
        ? currentSchemes 
        : [currentSchemes.find(s => s.id === reconSchemeId)].filter(Boolean);
      
      schemesToProcess.forEach(scheme => {
        const schemeSectors = currentSectors.filter(s => s.schemeId === scheme.id);
        const sectorRows: any[] = [];
        
        schemeSectors.forEach(sector => {
          const sectorActivities = currentActivities.filter(a => a.sectorId === sector.id);
          const activityRows: any[] = [];

          sectorActivities.forEach(activity => {
            const activitySubActivities = currentSubActivities.filter(sa => sa.activityId === activity.id);
            const subActivityRows: any[] = [];

            activitySubActivities.forEach(subActivity => {
              let subActivityAllocations = allocationsToReconcile.filter(a => a.subActivityId === subActivity.id);
              
              if (reconSearchTerm) {
                const lowerSearch = reconSearchTerm.toLowerCase();
                subActivityAllocations = subActivityAllocations.filter(a => {
                  const range = ranges.find(r => r.id === a.rangeId);
                  return (
                    (a.remarks || '').toLowerCase().includes(lowerSearch) ||
                    a.amount.toString().includes(lowerSearch) ||
                    (range?.name || '').toLowerCase().includes(lowerSearch) ||
                    subActivity.name.toLowerCase().includes(lowerSearch) ||
                    activity.name.toLowerCase().includes(lowerSearch) ||
                    sector.name.toLowerCase().includes(lowerSearch) ||
                    scheme.name.toLowerCase().includes(lowerSearch)
                  );
                });
              }

              if (subActivityAllocations.length > 0) {
                // Find SOE IDs for this sub-activity
                const subActivitySoes = currentSoes.filter(s => s.subActivityId === subActivity.id);
                const soeMap: Record<string, string> = {};
                targetSoeNames.forEach((name, idx) => {
                  const found = subActivitySoes.find(s => (s.name || '').includes(name.replace(/\s/g, '')));
                  if (found) soeMap[displaySoeNames[idx]] = found.id;
                });

                subActivityRows.push({
                  type: 'subActivity',
                  name: subActivity.name,
                  allocations: subActivityAllocations,
                  soeMap
                });
              }
            });

            if (subActivityRows.length > 0) {
              activityRows.push({
                type: 'activity',
                name: activity.name,
                subActivities: subActivityRows
              });
            }
          });

          if (activityRows.length > 0) {
            sectorRows.push({
              type: 'sector',
              name: sector.name,
              activities: activityRows
            });
          }
        });

        if (sectorRows.length > 0) {
          hierarchy.push({
            type: 'scheme',
            name: scheme.name,
            sectors: sectorRows
          });
        }
      });
    }

    const getRowTotals = (allocId: string) => {
      const distribution = reconData[allocId] || {};
      const total = Object.values(distribution).reduce<number>((sum, val) => sum + (parseFloat(val as string) || 0), 0);
      return total;
    };

    const isSchemeReady = allocationsToReconcile.length > 0 && allocationsToReconcile.every(a => {
      const total = getRowTotals(a.id);
      return Math.abs(a.amount - total) < 0.01;
    });

    const renderReconSummary = () => {
      if (!reconSchemeId) return null;
      
      let schemeSoes: any[] = [];
      let schemeAllocations: any[] = [];
      let title = "";

      if (reconSchemeId === 'all') {
        schemeSoes = currentSoes;
        schemeAllocations = baseAllocations;
        title = "All Schemes - SOE-wise Reconciliation Summary";
      } else {
        const scheme = currentSchemes.find(s => s.id === reconSchemeId);
        if (!scheme) return null;
        schemeSoes = currentSoes.filter(s => s.schemeId === reconSchemeId);
        schemeAllocations = baseAllocations.filter(a => a.schemeId === reconSchemeId);
        title = `${scheme.name} - SOE-wise Reconciliation Summary`;
      }

      if (reconSearchTerm) {
        const lowerSearch = reconSearchTerm.toLowerCase();
        schemeSoes = schemeSoes.filter(soe => {
          const schemeName = currentSchemes.find(s => s.id === soe.schemeId)?.name || '';
          return soe.name.toLowerCase().includes(lowerSearch) || schemeName.toLowerCase().includes(lowerSearch);
        });
      }

      return (
        <div className="animate-in fade-in duration-500">
          <div className="bg-emerald-900 text-white p-4 rounded-t-xl flex justify-between items-center">
            <h3 className="font-bold flex items-center gap-2">
              <Table className="w-5 h-5" />
              {title}
            </h3>
            <span className="text-xs bg-emerald-800 px-3 py-1 rounded-full border border-emerald-700">TRY Budget vs Allocated</span>
          </div>
          <div className="bg-white border border-gray-200 rounded-b-xl overflow-hidden shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[800px]">
                <thead>
                <tr className="bg-gray-50 text-gray-600 border-b">
                  <th className="p-4 text-left font-bold">SOE Head</th>
                  {reconSchemeId === 'all' && <th className="p-4 text-left font-bold">Scheme</th>}
                  <th className="p-4 text-right font-bold">TRY Budget (A)</th>
                  <th className="p-4 text-right font-bold">Total Reconciled (B)</th>
                  <th className="p-4 text-right font-bold">Balance (A - B)</th>
                  <th className="p-4 text-center font-bold">Utilization</th>
                </tr>
              </thead>
              <tbody>
                {schemeSoes.map(soe => {
                  const budget = getReceivedInTry(soe);
                  const reconciled = schemeAllocations.reduce((sum, alloc) => {
                    const funded = alloc.fundedSOEs?.find(f => f.soeId === soe.id);
                    return sum + (funded?.amount || 0);
                  }, 0);
                  const balance = budget - reconciled;
                  const percent = budget > 0 ? (reconciled / budget) * 100 : 0;
                  const schemeName = currentSchemes.find(s => s.id === soe.schemeId)?.name;

                  return (
                    <tr key={soe.id} className="border-b hover:bg-gray-50 transition-colors">
                      <td className="p-4 font-medium text-gray-900">{soe.name}</td>
                      {reconSchemeId === 'all' && <td className="p-4 text-gray-600 text-xs">{schemeName}</td>}
                      <td className="p-4 text-right font-mono">â‚¹{budget.toLocaleString()}</td>
                      <td className="p-4 text-right font-mono text-blue-600">â‚¹{reconciled.toLocaleString()}</td>
                      <td className={`p-4 text-right font-mono font-bold ${balance < 0 ? 'text-red-600' : 'text-emerald-600'}`}>
                        â‚¹{balance.toLocaleString()}
                      </td>
                      <td className="p-4">
                        <div className="w-full bg-gray-100 rounded-full h-2 max-w-[120px] mx-auto overflow-hidden">
                          <div 
                            className={`h-full rounded-full ${percent > 100 ? 'bg-red-500' : 'bg-emerald-500'}`} 
                            style={{ width: `${Math.min(percent, 100)}%` }}
                          />
                        </div>
                        <div className="text-[10px] text-center mt-1 font-bold text-gray-500">{percent.toFixed(1)}%</div>
                      </td>
                    </tr>
                  );
                })}
                {schemeSoes.length === 0 && (
                  <tr>
                    <td colSpan={reconSchemeId === 'all' ? 6 : 5} className="p-10 text-center text-gray-400 italic">No SOE Heads found for this scheme.</td>
                  </tr>
                )}
              </tbody>
              {schemeSoes.length > 0 && (
                <tfoot>
                  <tr className="bg-gray-50 font-bold border-t-2 border-gray-200">
                    <td className="p-4" colSpan={reconSchemeId === 'all' ? 2 : 1}>GRAND TOTAL</td>
                    <td className="p-4 text-right">â‚¹{schemeSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0).toLocaleString()}</td>
                    <td className="p-4 text-right text-blue-600">
                      â‚¹{schemeAllocations.reduce((sum, alloc) => {
                        return sum + (alloc.fundedSOEs?.reduce((s, f) => s + f.amount, 0) || 0);
                      }, 0).toLocaleString()}
                    </td>
                    <td className="p-4 text-right text-emerald-600">
                      â‚¹{(schemeSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0) - schemeAllocations.reduce((sum, alloc) => sum + (alloc.fundedSOEs?.reduce((s, f) => s + f.amount, 0) || 0), 0)).toLocaleString()}
                    </td>
                    <td></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      </div>
      );
    };

    return (
      <div className="space-y-6">
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
            <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
              <RefreshCcw className="w-5 h-5 text-emerald-600" />
              Budget Reconciliation (Provisional to SOE)
            </h2>
            <div className="flex flex-wrap items-center gap-3">
              <div className="flex bg-gray-100 p-1 rounded-lg border border-gray-200">
                <button 
                  onClick={() => setShowReconSummary(false)}
                  className={`px-4 py-1.5 rounded-md text-xs font-bold transition-all ${!showReconSummary ? 'bg-white text-emerald-700 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                >
                  Hierarchical Grid
                </button>
                <button 
                  onClick={() => setShowReconSummary(true)}
                  className={`px-4 py-1.5 rounded-md text-xs font-bold transition-all ${showReconSummary ? 'bg-white text-emerald-700 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                >
                  Summary View
                </button>
              </div>
              <div className="relative w-64">
                <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="Search anything..."
                  value={reconSearchTerm}
                  onChange={(e) => setReconSearchTerm(e.target.value)}
                  className="w-full pl-9 pr-3 py-2 border rounded-lg shadow-sm focus:ring-2 focus:ring-emerald-500 outline-none text-sm"
                />
              </div>
              <div className="w-64">
                <select 
                  value={reconSchemeId} 
                  onChange={(e) => { setReconSchemeId(e.target.value); setReconData({}); }}
                  className="w-full p-2 border rounded-lg shadow-sm focus:ring-2 focus:ring-emerald-500 outline-none text-sm"
                >
                  <option value="">-- Select Scheme --</option>
                  <option value="all">All Schemes</option>
                  {provisionalSchemes.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
            </div>
          </div>
          
          {!reconSchemeId ? (
            <div className="text-center py-20 bg-gray-50 rounded-xl border-2 border-dashed border-gray-200">
              <RefreshCcw className="w-12 h-12 text-gray-300 mx-auto mb-4 animate-spin-slow" />
              <p className="text-gray-500 font-medium">Please select a scheme to start the reconciliation process</p>
            </div>
          ) : showReconSummary ? renderReconSummary() : allocationsToReconcile.length === 0 ? (
            <div className="text-center py-20 bg-gray-50 rounded-xl border-2 border-dashed border-gray-200">
              <Check className="w-12 h-12 text-emerald-300 mx-auto mb-4" />
              <p className="text-gray-500 font-medium">No pending provisional allocations found for this scheme. Everything is reconciled!</p>
            </div>
          ) : (
            <div className="overflow-x-auto border rounded-xl shadow-sm">
              <table className="w-full border-collapse text-xs min-w-[1200px]">
                <thead>
                  <tr className="bg-emerald-800 text-white">
                    <th className="p-3 border border-emerald-700 text-left sticky left-0 bg-emerald-800 z-10" rowSpan={2}>Hierarchy (Sector/Activity/Sub-Activity)</th>
                    <th className="p-3 border border-emerald-700 text-center" rowSpan={2}>Approved Budget</th>
                    <th className="p-3 border border-emerald-700 text-left" rowSpan={2}>Range Name</th>
                    <th className="p-3 border border-emerald-700 text-right" rowSpan={2}>Amount Allocated</th>
                    <th className="p-3 border border-emerald-700 text-right" rowSpan={2}>Budget to be Allocated</th>
                    <th className="p-3 border border-emerald-700 text-center" colSpan={4}>SOE Distribution (Editable)</th>
                    <th className="p-3 border border-emerald-700 text-right" rowSpan={2}>Try SOE Total</th>
                    <th className="p-3 border border-emerald-700 text-right" rowSpan={2}>Variation</th>
                  </tr>
                  <tr className="bg-emerald-700 text-white">
                    {displaySoeNames.map(name => (
                      <th key={name} className="p-2 border border-emerald-600 text-center">{name}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {hierarchy.map((schemeLevel, idx) => (
                    <React.Fragment key={schemeLevel.name}>
                      {/* Scheme Row */}
                      {reconSchemeId === 'all' && (
                        <tr className="bg-gray-300 font-bold text-sm">
                          <td className="p-2 border border-gray-400 sticky left-0 bg-gray-300 z-10" colSpan={11}>SCHEME: {schemeLevel.name}</td>
                        </tr>
                      )}
                      
                      {schemeLevel.sectors.map((sector: any, sIdx: number) => (
                        <React.Fragment key={sector.name}>
                          {/* Sector Row */}
                          <tr className="bg-gray-200 font-bold">
                            <td className="p-2 border border-gray-300 sticky left-0 bg-gray-200 z-10" colSpan={11}>SECTOR: {sector.name}</td>
                          </tr>
                          {sector.activities.map((activity: any, aIdx: number) => (
                            <React.Fragment key={activity.name}>
                              {/* Activity Row */}
                              <tr className="bg-gray-100 font-semibold italic">
                                <td className="p-2 border border-gray-300 pl-6 sticky left-0 bg-gray-100 z-10" colSpan={11}>Activity: {activity.name}</td>
                              </tr>
                              {activity.subActivities.map((subActivity: any, saIdx: number) => {
                                const subActivityAllocations = subActivity.allocations;
                                const approvedBudget = currentSoes.filter(s => s.subActivityId === subActivityAllocations[0].subActivityId).reduce((sum, s) => sum + getReceivedInTry(s), 0);
                                
                                return (
                                  <React.Fragment key={subActivity.name}>
                                    {subActivityAllocations.map((alloc: any, alIdx: number) => {
                                      const distribution = reconData[alloc.id] || {};
                                      const tryTotal = getRowTotals(alloc.id);
                                      const variation = alloc.amount - tryTotal;
                                      const range = ranges.find(r => r.id === alloc.rangeId);
                                      
                                      return (
                                        <tr key={alloc.id} className="hover:bg-emerald-50 transition-colors">
                                          {alIdx === 0 && (
                                            <td className="p-2 border border-gray-300 pl-10 sticky left-0 bg-white z-10 font-medium" rowSpan={subActivityAllocations.length}>
                                              {subActivity.name}
                                            </td>
                                          )}
                                          {alIdx === 0 && (
                                            <td className="p-2 border border-gray-300 text-right font-bold text-emerald-700" rowSpan={subActivityAllocations.length}>
                                              â‚¹{approvedBudget.toLocaleString()}
                                            </td>
                                          )}
                                          <td className="p-2 border border-gray-300 italic text-gray-600">{range?.name}</td>
                                          <td className="p-2 border border-gray-300 text-right font-bold text-blue-600">â‚¹{alloc.amount.toLocaleString()}</td>
                                          <td className="p-2 border border-gray-300 text-right text-gray-400">â‚¹{(approvedBudget - alloc.amount).toLocaleString()}</td>
                                          
                                          {displaySoeNames.map(soeName => {
                                            const soeId = subActivity.soeMap[soeName];
                                            return (
                                              <td key={soeName} className="p-1 border border-gray-300">
                                                {soeId ? (
                                                  <input 
                                                    type="number"
                                                    value={distribution[soeId] || ''}
                                                    onChange={(e) => {
                                                      const val = e.target.value;
                                                      setReconData((prev: any) => ({
                                                        ...prev,
                                                        [alloc.id]: {
                                                          ...(prev[alloc.id] || {}),
                                                          [soeId]: val
                                                        }
                                                      }));
                                                    }}
                                                    placeholder="0"
                                                    className="w-full p-1 border-none focus:ring-1 focus:ring-emerald-500 text-right bg-transparent"
                                                  />
                                                ) : (
                                                  <div className="text-center text-gray-300">-</div>
                                                )}
                                              </td>
                                            );
                                          })}
                                          
                                          <td className="p-2 border border-gray-300 text-right font-bold text-emerald-600">â‚¹{tryTotal.toLocaleString()}</td>
                                          <td className={`p-2 border border-gray-300 text-right font-bold ${Math.abs(variation) < 0.01 ? 'text-green-600' : 'text-red-600'}`}>
                                            â‚¹{variation.toLocaleString()}
                                          </td>
                                        </tr>
                                      );
                                    })}
                                    {/* Sub-Activity Total Row */}
                                    <tr className="bg-gray-50 font-bold text-[10px]">
                                      <td className="p-2 border border-gray-300 text-right" colSpan={3}>Total for {subActivity.name}</td>
                                      <td className="p-2 border border-gray-300 text-right">â‚¹{subActivityAllocations.reduce((sum: number, a: any) => sum + a.amount, 0).toLocaleString()}</td>
                                      <td className="p-2 border border-gray-300" colSpan={5}></td>
                                      <td className="p-2 border border-gray-300 text-right">â‚¹{subActivityAllocations.reduce((sum: number, a: any) => sum + getRowTotals(a.id), 0).toLocaleString()}</td>
                                      <td className="p-2 border border-gray-300"></td>
                                    </tr>
                                  </React.Fragment>
                                );
                              })}
                              {/* Activity Total Row */}
                              <tr className="bg-blue-50 font-bold text-[11px]">
                                <td className="p-2 border border-gray-300 text-right" colSpan={3}>Total Activity: {activity.name}</td>
                                <td className="p-2 border border-gray-300 text-right">â‚¹{activity.subActivities.reduce((sum: number, sa: any) => sum + sa.allocations.reduce((s: number, a: any) => s + a.amount, 0), 0).toLocaleString()}</td>
                                <td className="p-2 border border-gray-300" colSpan={7}></td>
                              </tr>
                            </React.Fragment>
                          ))}
                          {/* Sector Total Row */}
                          <tr className="bg-emerald-50 font-bold text-xs">
                            <td className="p-2 border border-gray-300 text-right" colSpan={3}>Total Sector: {sector.name}</td>
                            <td className="p-2 border border-gray-300 text-right">â‚¹{sector.activities.reduce((sum: number, act: any) => sum + act.subActivities.reduce((s: number, sa: any) => s + sa.allocations.reduce((ss: number, a: any) => ss + a.amount, 0), 0), 0).toLocaleString()}</td>
                            <td className="p-2 border border-gray-300" colSpan={7}></td>
                          </tr>
                        </React.Fragment>
                      ))}
                      {/* Scheme Total Row (only if 'all' is selected) */}
                      {reconSchemeId === 'all' && (
                        <tr className="bg-emerald-100 font-bold text-sm">
                          <td className="p-2 border border-gray-300 text-right" colSpan={3}>Total Scheme: {schemeLevel.name}</td>
                          <td className="p-2 border border-gray-300 text-right">â‚¹{schemeLevel.sectors.reduce((sum: number, sec: any) => sum + sec.activities.reduce((s: number, act: any) => s + act.subActivities.reduce((ss: number, sa: any) => ss + sa.allocations.reduce((sss: number, a: any) => sss + a.amount, 0), 0), 0), 0).toLocaleString()}</td>
                          <td className="p-2 border border-gray-300" colSpan={7}></td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                  {/* Grand Total Row */}
                  <tr className="bg-emerald-900 text-white font-bold text-sm">
                    <td className="p-3 border border-emerald-800 text-right" colSpan={3}>GRAND TOTAL (SCHEME)</td>
                    <td className="p-3 border border-emerald-800 text-right">â‚¹{allocationsToReconcile.reduce((sum, a) => sum + a.amount, 0).toLocaleString()}</td>
                    <td className="p-3 border border-emerald-800" colSpan={5}></td>
                    <td className="p-3 border border-emerald-800 text-right">â‚¹{allocationsToReconcile.reduce((sum, a) => sum + getRowTotals(a.id), 0).toLocaleString()}</td>
                    <td className="p-3 border border-emerald-800 text-right">
                      â‚¹{allocationsToReconcile.reduce((sum, a) => sum + (a.amount - getRowTotals(a.id)), 0).toLocaleString()}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          )}

          {reconSchemeId && allocationsToReconcile.length > 0 && (
            <div className="mt-8 flex justify-end items-center gap-6 p-6 bg-gray-50 rounded-xl border border-gray-200">
              <div className="text-right">
                <p className="text-xs text-gray-500 uppercase font-bold">Reconciliation Status</p>
                <p className={`text-sm font-bold ${isSchemeReady ? 'text-green-600' : 'text-orange-600'}`}>
                  {isSchemeReady ? 'âœ“ All variations are zero. Ready to save.' : 'âš  Please fix variations to zero to enable saving.'}
                </p>
              </div>
              <button
                disabled={!isSchemeReady || loading}
                onClick={handleSaveAllReconciliation}
                className={`px-8 py-3 rounded-lg font-bold shadow-lg transition-all flex items-center gap-2 ${isSchemeReady ? 'bg-emerald-600 text-white hover:bg-emerald-700 hover:scale-105 active:scale-95' : 'bg-gray-200 text-gray-400 cursor-not-allowed'}`}
              >
                {loading ? <RefreshCcw className="w-5 h-5 animate-spin" /> : <Save className="w-5 h-5" />}
                Save All Reconciliation Data
              </button>
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderSOEHeads = () => {
    const filteredItems = currentSoes.filter(s => {
      // Search filter
      const search = searchTerm.toLowerCase();
      const nameMatch = (s.name || '').toLowerCase().includes(search);
      const schemeMatch = s.schemeId && schemes.find(sch => sch.id === s.schemeId)?.name.toLowerCase().includes(search);
      const sectorMatch = s.sectorId && sectors.find(sec => sec.id === s.sectorId)?.name.toLowerCase().includes(search);
      const activityMatch = s.activityId && activities.find(act => act.id === s.activityId)?.name.toLowerCase().includes(search);
      const subActivityMatch = s.subActivityId && subActivities.find(sub => sub.id === s.subActivityId)?.name.toLowerCase().includes(search);
      const matchesSearch = !searchTerm || nameMatch || schemeMatch || sectorMatch || activityMatch || subActivityMatch;

      // Hierarchy filters
      const matchesScheme = !soeFilters.schemeId || s.schemeId === soeFilters.schemeId;
      const matchesSector = !soeFilters.sectorId || s.sectorId === soeFilters.sectorId;
      const matchesActivity = !soeFilters.activityId || s.activityId === soeFilters.activityId;
      const matchesSubActivity = !soeFilters.subActivityId || s.subActivityId === soeFilters.subActivityId;
      const matchesSoeName = !soeFilters.soeName || s.name === soeFilters.soeName;
      const matchesRange = !soeFilters.rangeId || allocations.some(a => a.rangeId === soeFilters.rangeId && a.fundedSOEs?.some(f => f.soeId === s.id));

      return matchesSearch && matchesScheme && matchesSector && matchesActivity && matchesSubActivity && matchesSoeName && matchesRange;
    });

    const soeFormSpan = formWidth === 'extra' ? 'lg:col-span-6' : formWidth === 'wide' ? 'lg:col-span-5' : 'lg:col-span-4';
    const soeTableSpan = formWidth === 'extra' ? 'lg:col-span-6' : formWidth === 'wide' ? 'lg:col-span-7' : 'lg:col-span-8';

    return (
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <div className={soeFormSpan}>
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 sticky top-6">
            <div className="flex items-center justify-between mb-4 pb-2 border-b">
              <div className="flex items-center gap-2">
                <h3 className="text-lg font-semibold">Add SOE Head</h3>
                <button
                  type="button"
                  onClick={() => {
                    const nextWidth = formWidth === 'normal' ? 'wide' : formWidth === 'wide' ? 'extra' : 'normal';
                    setFormWidth(nextWidth);
                    localStorage.setItem('fbc_form_width', nextWidth);
                  }}
                  className="hidden lg:flex items-center gap-1.5 px-2 py-0.5 hover:bg-gray-100 rounded text-gray-700 text-[11px] font-bold border border-gray-200 bg-gray-50 transition-colors ml-1"
                  title="Form Size: Click to cycle between Standard (33%), Wide (42%), and Extra Wide (50%)"
                >
                  <MoveHorizontal className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="text-[10px] text-gray-700 whitespace-nowrap">
                    {formWidth === 'wide' ? 'Width: 42%' : formWidth === 'extra' ? 'Width: 50%' : 'Width: 33%'}
                  </span>
                </button>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => setIsFormExpanded(!isFormExpanded)} className="lg:hidden">
                  {isFormExpanded ? <ChevronUp /> : <ChevronDown />}
                </button>
              </div>
            </div>
            
            {isFormExpanded && (
              <form onSubmit={handleAddSoeName} className="space-y-4">
                <CascadingDropdowns 
                  schemes={currentSchemes} sectors={currentSectors} activities={currentActivities} subActivities={currentSubActivities} soes={currentSoes} soeBudgets={[]} allocations={baseAllocations} surrenders={surrenders} ranges={ranges} expenses={currentExpenses}
                  editingItem={editingItem} type="SOE Name" userRangeId={userRangeId} userRole={userRole} showConfirm={showConfirm}
                >
                  <select 
                    name="name" 
                    defaultValue={editingItem?.type === 'SOE Name' ? editingItem.item.name : ''} 
                    className="w-full p-2 border rounded" 
                  >
                    <option value="">Select SOE Head (Optional - defaults to Provisional)</option>
                    {ALLOWED_SOES.filter(n => n !== 'Provisional').map(n => <option key={n} value={n}>{n}</option>)}
                  </select>
                  <div className="grid grid-cols-2 gap-2">
                    <input 
                      name="approvedBudget" 
                      type="number" 
                      defaultValue={editingItem?.type === 'SOE Name' ? (getApprovedBudget(editingItem.item) || '') : ''} 
                      placeholder="Approved Budget (â‚¹) (Optional)" 
                      className="w-full p-2 border rounded text-sm" 
                    />
                    <input 
                      name="receivedInTry" 
                      type="number" 
                      defaultValue={editingItem?.type === 'SOE Name' ? (getReceivedInTry(editingItem.item) || '') : ''} 
                      placeholder="Received in TRY (â‚¹) (Optional)" 
                      className="w-full p-2 border rounded text-sm" 
                    />
                  </div>
                </CascadingDropdowns>
                <button type="submit" className="w-full bg-emerald-600 text-white py-2 rounded-lg hover:bg-emerald-700 transition-colors flex items-center justify-center gap-2">
                  <Plus className="w-4 h-4" />
                  <span>{editingItem?.type === 'SOE Name' ? 'Update SOE Head' : 'Add SOE Head'}</span>
                </button>
                {editingItem && (
                  <button type="button" onClick={() => setEditingItem(null)} className="w-full bg-gray-100 text-gray-600 py-2 rounded-lg hover:bg-gray-200 transition-colors">
                    Cancel Edit
                  </button>
                )}
              </form>
            )}
          </div>
        </div>

        <div className={`${soeTableSpan} space-y-4`}>
          <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 space-y-4">
            <div className="flex items-center gap-4">
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                <input 
                  type="text" 
                  placeholder="Search SOE Heads..." 
                  className="w-full pl-10 pr-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>
              <button 
                onClick={() => setIsSoeFilterExpanded(!isSoeFilterExpanded)}
                className={`flex items-center gap-1 px-3 py-2 border rounded-lg text-sm transition-colors ${isSoeFilterExpanded ? 'bg-emerald-50 border-emerald-500 text-emerald-700' : 'bg-white hover:bg-gray-50'}`}
              >
                <Filter className="w-4 h-4" />
                <span>Filters</span>
                {isSoeFilterExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
            </div>

            {isSoeFilterExpanded && (
              <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-6 gap-3 p-3 bg-gray-50 rounded-lg border border-gray-200 animate-in fade-in slide-in-from-top-2">
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Scheme</label>
                  <select 
                    value={soeFilters.schemeId}
                    onChange={(e) => setSoeFilters({ ...soeFilters, schemeId: e.target.value, sectorId: '', activityId: '', subActivityId: '' })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All Schemes</option>
                    {schemes.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Sector</label>
                  <select 
                    value={soeFilters.sectorId}
                    onChange={(e) => setSoeFilters({ ...soeFilters, sectorId: e.target.value, activityId: '', subActivityId: '' })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All Sectors</option>
                    {sectors.filter(s => !soeFilters.schemeId || s.schemeId === soeFilters.schemeId).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Activity</label>
                  <select 
                    value={soeFilters.activityId}
                    onChange={(e) => setSoeFilters({ ...soeFilters, activityId: e.target.value, subActivityId: '' })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All Activities</option>
                    {activities.filter(a => {
                      if (soeFilters.sectorId) return a.sectorId === soeFilters.sectorId;
                      if (soeFilters.schemeId) return a.schemeId === soeFilters.schemeId || sectors.find(s => s.id === a.sectorId)?.schemeId === soeFilters.schemeId;
                      return true;
                    }).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Sub-Activity</label>
                  <select 
                    value={soeFilters.subActivityId}
                    onChange={(e) => setSoeFilters({ ...soeFilters, subActivityId: e.target.value })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All Sub-Activities</option>
                    {subActivities.filter(sa => {
                      if (soeFilters.activityId) return sa.activityId === soeFilters.activityId;
                      if (soeFilters.sectorId) {
                        const act = activities.find(a => a.id === sa.activityId);
                        return act?.sectorId === soeFilters.sectorId;
                      }
                      if (soeFilters.schemeId) {
                        const act = activities.find(a => a.id === sa.activityId);
                        return act?.schemeId === soeFilters.schemeId || sectors.find(s => s.id === act?.sectorId)?.schemeId === soeFilters.schemeId;
                      }
                      return true;
                    }).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Range</label>
                  <select 
                    value={soeFilters.rangeId}
                    onChange={(e) => setSoeFilters({ ...soeFilters, rangeId: e.target.value })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All Ranges</option>
                    {ranges.map(s => <option key={s.id} value={s.id}>{s.name === 'Rajgarh Forest Division' ? 'Division' : s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">SOE Name</label>
                  <select 
                    value={soeFilters.soeName}
                    onChange={(e) => setSoeFilters({ ...soeFilters, soeName: e.target.value })}
                    className="w-full p-1.5 border border-gray-300 rounded text-xs bg-white"
                  >
                    <option value="">All SOEs</option>
                    {ALLOWED_SOES.filter(n => {
                      if (!soeFilters.schemeId && !soeFilters.sectorId && !soeFilters.activityId && !soeFilters.subActivityId) return true;
                      return currentSoes.some(s => {
                        const matchesScheme = !soeFilters.schemeId || s.schemeId === soeFilters.schemeId;
                        const matchesSector = !soeFilters.sectorId || s.sectorId === soeFilters.sectorId;
                        const matchesActivity = !soeFilters.activityId || s.activityId === soeFilters.activityId;
                        const matchesSubActivity = !soeFilters.subActivityId || s.subActivityId === soeFilters.subActivityId;
                        return s.name === n && matchesScheme && matchesSector && matchesActivity && matchesSubActivity;
                      });
                    }).map(n => <option key={n} value={n}>{n}</option>)}
                  </select>
                </div>
                <div className="md:col-span-3 lg:col-span-6 flex justify-end">
                  <button 
                    onClick={() => {
                      setSoeFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeName: '' });
                      setSearchTerm('');
                    }}
                    className="text-[10px] text-red-600 hover:text-red-800 font-bold uppercase flex items-center gap-1"
                  >
                    Reset Filters
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-x-auto">
            <table className="w-full text-sm text-left">
              <thead className="bg-gray-50 text-gray-600 font-medium border-b">
                <tr>
                  <th className="px-4 py-4">SrNo</th>
                  <th className="px-4 py-4">Hierarchy</th>
                  <th className="px-4 py-4">SOE Name</th>
                  <th className="px-4 py-4 text-right">Approved Budget</th>
                  <th className="px-4 py-4 text-right">Received in TRY</th>
                  <th className="px-4 py-4 text-right">Allocated</th>
                  <th className="px-4 py-4 text-right">To Be Allocated</th>
                  <th className="px-4 py-4 text-right">Actions</th>
                </tr>
              </thead>
            <tbody className="divide-y divide-gray-100">
                {filteredItems
                  .sort((a, b) => {
                    // For SOE Heads, we might want to prioritize those with received budget
                    const hasBudgetA = (a.receivedInTry || 0) > 0;
                    const hasBudgetB = (b.receivedInTry || 0) > 0;
                    if (hasBudgetA && !hasBudgetB) return -1;
                    if (!hasBudgetA && hasBudgetB) return 1;

                    return (b.updatedAt || 0) - (a.updatedAt || 0);
                  })
                  .slice((currentPage - 1) * itemsPerPage, itemsPerPage === -1 ? filteredItems.length : currentPage * itemsPerPage)
                  .map((s, index) => {
                  const allocated = allocations.reduce((sum, a) => {
                    const funded = a.fundedSOEs?.find(f => f.soeId === s.id);
                    return sum + (funded?.amount || 0);
                  }, 0);
                  const toBeAllocated = getReceivedInTry(s) - allocated;

                  return (
                    <tr key={s.id} className="hover:bg-gray-50 transition-colors">
                      <td className="px-4 py-4 text-gray-500 font-medium">{(currentPage - 1) * itemsPerPage + index + 1}</td>
                      <td className="px-4 py-4">{renderHierarchy(s)}</td>
                      <td className="px-4 py-4 font-medium">{s.name || '-'}</td>
                      <td className="px-4 py-4 text-right">â‚¹{getApprovedBudget(s).toLocaleString()}</td>
                      <td className="px-4 py-4 text-right">
                        <TryUpdateInput 
                          soeId={s.id} 
                          initialValue={getReceivedInTry(s)} 
                          onUpdate={handleUpdateSoeTry} 
                        />
                      </td>
                      <td className="px-4 py-4 text-right font-medium text-emerald-600">â‚¹{allocated.toLocaleString()}</td>
                      <td className="px-4 py-4 text-right font-medium text-orange-600">â‚¹{toBeAllocated.toLocaleString()}</td>
                      <td className="px-4 py-4 text-right space-x-2">
                        <button onClick={() => {
                          setEditingItem({ type: 'SOE Name', item: s });
                          setIsFormExpanded(true);
                          window.scrollTo({ top: 0, behavior: 'smooth' });
                        }} className="text-blue-600 hover:text-blue-800"><Edit2 className="w-4 h-4" /></button>
                        <button onClick={() => handleDelete('soeHeads', s.id)} className="text-red-600 hover:text-red-800"><Trash2 className="w-4 h-4" /></button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <Pagination 
            totalEntries={filteredItems.length} 
            currentPage={currentPage} 
            itemsPerPage={itemsPerPage} 
            onPageChange={setCurrentPage} 
            onItemsPerPageChange={(val) => { setItemsPerPage(val === 'All' ? -1 : val); setCurrentPage(1); }}
          />
        </div>
      </div>
    );
  };

  useEffect(() => {
    setCurrentPage(1);
  }, [activeTab, searchTerm, expFilters, allocFilters, billExpFilters, selectedFY, expDateRange, expenditureSubTab]);

  const renderSimpleManager = (
    title: string, 
    items: any[], 
    columns: {key: string, label: string, render?: (val: any, item: any) => React.ReactNode, searchableText?: (val: any, item: any) => string}[], 
    onAdd: (e: React.FormEvent) => void, 
    onDelete: (id: string) => void,
    formContent: React.ReactNode,
    onEdit?: (item: any) => void,
    canEditDelete?: (item: any) => boolean,
    extraContent?: React.ReactNode,
    customActions?: (item: any) => React.ReactNode,
    isSubmitDisabled: boolean = false,
    isFilterExpanded?: boolean,
    setIsFilterExpanded?: (val: boolean) => void,
    filterContent?: React.ReactNode,
    onResetFilters?: () => void,
    isFullScreen?: boolean,
    setIsFullScreen?: (val: boolean) => void,
    getRowClassName?: (item: any) => string,
    customSearchTerm?: string,
    customSetSearchTerm?: (val: string) => void,
    canEdit?: (item: any) => boolean,
    canDelete?: (item: any) => boolean
  ) => {
    let filteredItems = items;
    const currentSearchTerm = customSearchTerm !== undefined ? customSearchTerm : searchTerm;
    const currentSetSearchTerm = customSetSearchTerm !== undefined ? customSetSearchTerm : setSearchTerm;

    if (currentSearchTerm) {
      const lowerSearch = currentSearchTerm.toLowerCase();
      filteredItems = items.filter(item => {
        return columns.some(c => {
          if (c.searchableText) {
            return c.searchableText(item[c.key], item).toLowerCase().includes(lowerSearch);
          }
          const val = c.render ? c.render(item[c.key], item) : item[c.key];
          if (typeof val === 'string' || typeof val === 'number') {
            return String(val).toLowerCase().includes(lowerSearch);
          }
          if (typeof item[c.key] === 'string' || typeof item[c.key] === 'number') {
            return String(item[c.key]).toLowerCase().includes(lowerSearch);
          }
          return false;
        });
      });
    }

    const isFormVisible = (
      userRole === 'admin' || 
      userRole === 'deo' || 
      ((title === 'Expenditure' || title === 'Bill' || title === 'Payee') && userRole !== 'approver' && userRole !== 'DA') || 
      (editingItem?.type === title && (isAdmin() || isDEO()))
    );

    const formColSpanClass = isFullScreen
      ? 'lg:col-span-12 z-50 fixed inset-0 m-4 overflow-y-auto'
      : formWidth === 'extra'
        ? 'lg:col-span-6 lg:sticky lg:top-6 max-h-[calc(100vh-120px)] overflow-y-auto'
        : formWidth === 'wide'
          ? 'lg:col-span-5 lg:sticky lg:top-6 max-h-[calc(100vh-120px)] overflow-y-auto'
          : 'lg:col-span-4 lg:sticky lg:top-6 max-h-[calc(100vh-120px)] overflow-y-auto';

    const tableColSpanClass = isTableFullScreen
      ? 'fixed inset-0 z-[60] bg-white p-6 overflow-y-auto'
      : isFullScreen
        ? 'hidden'
        : !isFormVisible
          ? 'lg:col-span-12'
          : formWidth === 'extra'
            ? 'lg:col-span-6'
            : formWidth === 'wide'
              ? 'lg:col-span-7'
              : 'lg:col-span-8';

    return (
    <div className={`grid grid-cols-1 ${isFullScreen ? '' : 'lg:grid-cols-12'} gap-6 items-start relative`}>
      {isFormVisible && !isTableFullScreen && (
        <div className={`bg-white p-4 rounded-2xl shadow-sm border border-gray-100 ${formColSpanClass} custom-scrollbar transition-all duration-300`}>
          <div 
            className="flex justify-between items-center mb-2 border-b pb-1.5 cursor-pointer hover:bg-gray-50 -mx-3 px-3 pt-0.5" 
            onClick={() => setIsFormExpanded(!isFormExpanded)}
          >
            <div className="flex items-center gap-2 flex-wrap">
              <h3 className="text-sm font-semibold">
                {editingItem?.type === title ? `Edit ${title}` : `Add ${title}`}
              </h3>
              {editingItem?.type === title && (
                <button 
                  onClick={(e) => { e.stopPropagation(); setEditingItem(null); }}
                  className="text-[9px] bg-blue-50 text-blue-600 px-1.5 py-0.5 rounded hover:bg-blue-100 font-bold uppercase"
                >
                  New
                </button>
              )}
              {/* Form Width Controls */}
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  const nextWidth = formWidth === 'normal' ? 'wide' : formWidth === 'wide' ? 'extra' : 'normal';
                  setFormWidth(nextWidth);
                  localStorage.setItem('fbc_form_width', nextWidth);
                }}
                className="hidden lg:flex items-center gap-1.5 px-2 py-0.5 hover:bg-gray-100 rounded text-gray-700 text-[11px] font-bold border border-gray-200 bg-gray-50 transition-colors"
                title="Form Size: Click to switch between Standard (33%), Wide (42%), and Extra Wide (50%)"
              >
                <MoveHorizontal className="w-3.5 h-3.5 text-emerald-600" />
                <span className="text-[10px] text-gray-700 whitespace-nowrap">
                  {formWidth === 'wide' ? 'Width: 42%' : formWidth === 'extra' ? 'Width: 50%' : 'Width: 33%'}
                </span>
              </button>
              {setIsFullScreen && (
                <button
                  type="button"
                  onClick={(e) => { e.stopPropagation(); setIsFullScreen(!isFullScreen); }}
                  className="p-1 hover:bg-gray-100 rounded text-gray-500"
                  title={isFullScreen ? "Exit Full Screen" : "Full Screen Modal"}
                >
                  {isFullScreen ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                </button>
              )}
            </div>
            <button type="button" className="text-gray-500 hover:text-gray-700">
              {isFormExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
          </div>
          <div className={`${isFormExpanded ? 'block' : 'hidden'} pb-2`}>
            <form key={editingItem?.item?.id || 'new'} onSubmit={onAdd} className="space-y-2">
              {formContent}
              <div className="flex gap-2 pt-2 border-t mt-2 sticky bottom-0 bg-white pb-1">
                <button 
                  type="submit" 
                  disabled={isSubmitDisabled}
                  className={`flex-1 py-1.5 rounded text-sm font-medium flex items-center justify-center gap-2 transition-colors ${isSubmitDisabled ? 'bg-gray-300 cursor-not-allowed text-gray-500' : 'bg-emerald-600 hover:bg-emerald-700 text-white'}`}
                >
                  {editingItem?.type === title ? <Activity className="w-3.5 h-3.5" /> : <Plus className="w-3.5 h-3.5" />}
                  {editingItem?.type === title ? 'Update' : 'Add'}
                </button>
                {editingItem?.type === title && (
                  <button 
                    type="button" 
                    onClick={() => setEditingItem(null)}
                    className="px-3 py-1.5 border border-gray-300 rounded text-sm text-gray-600 hover:bg-gray-50"
                  >
                    Cancel
                  </button>
                )}
                {isFullScreen && setIsFullScreen && (
                   <button 
                    type="button" 
                    onClick={() => setIsFullScreen(false)}
                    className="px-3 py-1.5 border border-gray-300 rounded text-sm text-gray-600 hover:bg-gray-50"
                  >
                    Close
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}
      <div className={`space-y-6 ${tableColSpanClass}`}>
        {extraContent}
        <div className={`bg-white ${isTableFullScreen ? '' : 'p-4 rounded-2xl shadow-sm border border-gray-100'}`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4 border-b pb-3">
              <div className="flex items-center gap-3">
                <h3 className="text-lg font-bold text-gray-800">Existing {title}s</h3>
                <button
                  type="button"
                  onClick={() => setIsTableFullScreen(!isTableFullScreen)}
                  className="p-2 hover:bg-gray-100 rounded-xl text-gray-500 transition-colors border border-gray-100"
                  title={isTableFullScreen ? "Exit Full Screen" : "Expand to Full Screen"}
                >
                  {isTableFullScreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
                </button>
              </div>
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-1">
                <div className="relative">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                  <input
                    type="text"
                    placeholder={`Search ${title}s...`}
                    value={currentSearchTerm}
                    onChange={(e) => currentSetSearchTerm(e.target.value)}
                    className="pl-9 pr-4 py-1.5 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 w-full sm:w-64"
                  />
                </div>
                <button 
                  onClick={() => {
                    // Trigger search logic if needed, but it's already reactive
                  }}
                  className="p-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors shadow-sm"
                  title="Search"
                >
                  <Search className="w-4 h-4" />
                </button>
              </div>
              {filterContent && setIsFilterExpanded && (

                <button 
                  onClick={() => setIsFilterExpanded(!isFilterExpanded)}
                  className={`flex items-center gap-1 px-2 py-1.5 border rounded-lg text-xs transition-colors ${isFilterExpanded ? 'bg-emerald-50 border-emerald-500 text-emerald-700' : 'bg-white hover:bg-gray-50'}`}
                >
                  <Filter className="w-3.5 h-3.5" />
                  <span className="hidden sm:inline">Filters</span>
                  {isFilterExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                </button>
              )}
              {title === 'Expenditure' && (
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={downloadExpenditureListPDF}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer"
                    title="Export PDF Document"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span className="hidden sm:inline">Export PDF</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowExpenditurePrintModal(true)}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer"
                    title="Print Expenditure List Preview"
                  >
                    <Printer className="w-3.5 h-3.5" />
                    <span className="hidden sm:inline">Print List</span>
                  </button>
                </div>
              )}
            </div>
          </div>

          {isFilterExpanded && filterContent && (
            <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-2 p-2 bg-gray-50 rounded-lg border border-gray-200 mb-3 animate-in fade-in slide-in-from-top-2">
              {filterContent}
              {onResetFilters && (
                <div className="md:col-span-3 lg:col-span-5 flex justify-end">
                  <button 
                    onClick={onResetFilters}
                    className="text-[9px] text-red-600 hover:text-red-800 font-bold uppercase flex items-center gap-1"
                  >
                    Reset Filters
                  </button>
                </div>
              )}
            </div>
          )}
          <div className="overflow-x-auto rounded-xl border border-gray-100">
            <table className="w-full text-left border-collapse">
              <thead className="bg-gray-50/50 text-gray-500 font-bold uppercase tracking-wider text-[10px]">
                <tr>
                  <th className="p-3 border-b border-gray-100 whitespace-nowrap w-12 text-center">#</th>
                  {columns.map(c => <th key={c.key} className="p-3 border-b border-gray-100 whitespace-nowrap">{c.label}</th>)}
                  {(customActions || userRole === 'admin' || userRole === 'deo' || title === 'Expenditure' || (canEditDelete && items.some(canEditDelete)) || (canEdit && items.some(canEdit)) || (canDelete && items.some(canDelete))) && <th className="p-3 border-b border-gray-100 text-right whitespace-nowrap">Actions</th>}
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {filteredItems
                  .sort((a, b) => {
                    const statusA = a.status === 'Funded' || a.status === 'approved';
                    const statusB = b.status === 'Funded' || b.status === 'approved';
                    if (statusA !== statusB) return statusA ? -1 : 1;
                    return (b.updatedAt || 0) - (a.updatedAt || 0);
                  })
                  .slice((currentPage - 1) * itemsPerPage, itemsPerPage === -1 ? filteredItems.length : currentPage * itemsPerPage)
                  .map((item, index) => (
                  <tr key={item.id} className={`hover:bg-emerald-50/30 transition-colors group ${getRowClassName ? getRowClassName(item) : ''}`}>
                    <td className="p-3 text-gray-400 font-medium text-center text-[11px]">{(currentPage - 1) * itemsPerPage + index + 1}</td>
                    {columns.map(c => <td key={c.key} className="p-3 text-gray-600 text-[11px] leading-relaxed">{c.render ? c.render(item[c.key], item) : item[c.key]}</td>)}
                    {(customActions || title === 'Expenditure' || (canEditDelete && canEditDelete(item)) || (canEdit && canEdit(item)) || (canDelete && canDelete(item)) || (userRole === 'admin' || userRole === 'deo')) && (
                      <td className="p-3 text-right">
                        <div className="flex justify-end gap-1 transition-opacity">
                          {customActions && customActions(item)}
                          {(canEditDelete ? canEditDelete(item) : (canEdit ? canEdit(item) : (userRole === 'admin' || userRole === 'deo' || (title === 'Expenditure' && userRole !== 'approver')))) && (
                            <button 
                              onClick={() => {
                                if (title === 'Allocation' && isFeatureLocked('Allocation')) {
                                  showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
                                  return;
                                }
                                if (title === 'Expenditure' && isFeatureLocked('Expenditure')) {
                                  showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
                                  return;
                                }
                                onEdit?.(item);
                                setIsFormExpanded(true);
                                window.scrollTo({ top: 0, behavior: 'smooth' });
                              }} 
                              className={`rounded-lg p-1.5 transition-colors ${((title === 'Allocation' && isFeatureLocked('Allocation')) || (title === 'Expenditure' && isFeatureLocked('Expenditure'))) ? 'text-gray-300 cursor-not-allowed' : 'text-blue-500 hover:bg-blue-100'}`}
                              title={((title === 'Allocation' && isFeatureLocked('Allocation')) || (title === 'Expenditure' && isFeatureLocked('Expenditure'))) ? "Locked by Admin" : "Edit"}
                            >
                              <Pencil className="w-3.5 h-3.5"/>
                            </button>
                          )}
                          {(canEditDelete ? canEditDelete(item) : (canDelete ? canDelete(item) : (userRole === 'admin' || userRole === 'deo' || (title === 'Expenditure' && userRole !== 'approver')))) && (
                            <button 
                              onClick={() => {
                                if (title === 'Allocation' && isFeatureLocked('Allocation')) {
                                  showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
                                  return;
                                }
                                if (title === 'Expenditure' && isFeatureLocked('Expenditure')) {
                                  showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
                                  return;
                                }
                                onDelete(item.id);
                              }} 
                              className={`rounded-lg p-1.5 transition-colors ${((title === 'Allocation' && isFeatureLocked('Allocation')) || (title === 'Expenditure' && isFeatureLocked('Expenditure'))) ? 'text-gray-300 cursor-not-allowed' : 'text-red-500 hover:bg-red-100'}`}
                              title={((title === 'Allocation' && isFeatureLocked('Allocation')) || (title === 'Expenditure' && isFeatureLocked('Expenditure'))) ? "Locked by Admin" : "Delete"}
                            >
                              <Trash2 className="w-3.5 h-3.5"/>
                            </button>
                          )}
                        </div>
                      </td>
                    )}
                  </tr>
                ))}
                {filteredItems.length === 0 && <tr><td colSpan={columns.length + 2} className="p-8 text-center text-gray-400 text-sm italic">No records found matching your criteria.</td></tr>}
              </tbody>
            </table>
          </div>
        <Pagination 
          totalEntries={filteredItems.length} 
          currentPage={currentPage} 
          itemsPerPage={itemsPerPage} 
          onPageChange={setCurrentPage} 
          onItemsPerPageChange={(val) => { setItemsPerPage(val === 'All' ? -1 : val); setCurrentPage(1); }}
        />
      </div>
    </div>
    </div>
    );
  };

  useEffect(() => {
    if (editingItem?.type === 'Allocation') {
      setAllocationAmount(editingItem.item.amount.toString());
    } else if (editingItem?.type === 'Expenditure') {
      setExpenseAmount(editingItem.item.amount.toString());
      setSelectedPayeesForExpense([]);
    } else if (!editingItem) {
      setAllocationAmount('');
      setExpenseAmount('');
      setSelectedPayeesForExpense([]);
    }
  }, [editingItem]);

  // --- PWA Install Logic ---
  const [installPrompt, setInstallPrompt] = useState<any>(null);
  const [isInstallable, setIsInstallable] = useState(false);

  useEffect(() => {
    const handler = (e: any) => {
      e.preventDefault();
      setInstallPrompt(e);
      setIsInstallable(true);
    };
    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  const handleInstallClick = async () => {
    if (!installPrompt) return;
    installPrompt.prompt();
    const { outcome } = await installPrompt.userChoice;
    if (outcome === 'accepted') {
      setIsInstallable(false);
      setInstallPrompt(null);
    }
  };

  const isFeatureLocked = (feature: 'Allocation' | 'Expenditure') => {
    if (userRole === 'admin') return false;
    const roleLock = featureLocks.find(l => l.feature === feature && l.target === userRole);
    if (roleLock?.isLocked) return true;
    if (userRangeId) {
      const rangeLock = featureLocks.find(l => l.feature === feature && l.target === userRangeId);
      if (rangeLock?.isLocked) return true;
    }
    return false;
  };

  // --- Handlers ---
  const handleAddFy = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    try {
      if (editingItem?.type === 'Financial Year') {
        await updateDoc(doc(db, 'financialYears', editingItem.item.id), { name, updatedAt: Date.now() });
        logAuditAction('FY Updated', `Financial Year: ${name}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'financialYears'), { name, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('FY Created', `Financial Year: ${name}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem ? OperationType.UPDATE : OperationType.CREATE, 'financialYears');
    }
  };

  const handleAddRange = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    try {
      if (editingItem?.type === 'Range') {
        await updateDoc(doc(db, 'ranges', editingItem.item.id), { name, updatedAt: Date.now() });
        logAuditAction('Range Updated', `Range Name: ${name}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'ranges'), { name, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('Range Created', `Range Name: ${name}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Range' ? OperationType.UPDATE : OperationType.CREATE, 'ranges');
    }
  };

  const handleAddScheme = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    try {
      if (editingItem?.type === 'Scheme') {
        await updateDoc(doc(db, 'schemes', editingItem.item.id), { name, updatedAt: Date.now() });
        logAuditAction('Scheme Updated', `Scheme Name: ${name}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'schemes'), { name, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('Scheme Created', `Scheme Name: ${name}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Scheme' ? OperationType.UPDATE : OperationType.CREATE, 'schemes');
    }
  };

  const handleAddSector = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    const schemeId = e.target.schemeId.value;
    try {
      if (editingItem?.type === 'Sector') {
        await updateDoc(doc(db, 'sectors', editingItem.item.id), { name, schemeId, updatedAt: Date.now() });
        const schemeName = schemes.find(s => s.id === schemeId)?.name || schemeId;
        logAuditAction('Sector Updated', `Sector: ${name}, Scheme: ${schemeName}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'sectors'), { name, schemeId, createdAt: Date.now(), updatedAt: Date.now() });
        const schemeName = schemes.find(s => s.id === schemeId)?.name || schemeId;
        logAuditAction('Sector Created', `Sector: ${name}, Scheme: ${schemeName}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Sector' ? OperationType.UPDATE : OperationType.CREATE, 'sectors');
    }
  };

  const handleAddActivity = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    const sectorId = e.target.sectorId?.value || null;
    const schemeId = e.target.schemeId?.value || null;

    if (!sectorId && !schemeId) {
      showAlert("Please select either a Sector or a Scheme");
      return;
    }

    try {
      if (editingItem?.type === 'Activity') {
        await updateDoc(doc(db, 'activities', editingItem.item.id), { name, sectorId, schemeId, updatedAt: Date.now() });
        logAuditAction('Activity Updated', `Activity: ${name}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'activities'), { sectorId, schemeId, name, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('Activity Created', `Activity: ${name}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Activity' ? OperationType.UPDATE : OperationType.CREATE, 'activities');
    }
  };

  const handleAddSubActivity = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    const activityId = e.target.activityId.value;
    
    if (!activityId) {
      showAlert("Activity is mandatory for Sub-Activity.");
      return;
    }

    try {
      if (editingItem?.type === 'Sub-Activity') {
        await updateDoc(doc(db, 'subActivities', editingItem.item.id), { name, activityId, updatedAt: Date.now() });
        logAuditAction('Sub-Activity Updated', `Sub-Activity: ${name}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'subActivities'), { activityId, name, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('Sub-Activity Created', `Sub-Activity: ${name}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Sub-Activity' ? OperationType.UPDATE : OperationType.CREATE, 'subActivities');
    }
  };

  const handleAddSoeName = async (e: any) => {
    e.preventDefault();
    const rawName = e.target.name.value;
    const name = rawName || 'Provisional';
    const schemeId = e.target.schemeId?.value || null;
    const sectorId = e.target.sectorId?.value || null;
    const activityId = e.target.activityId?.value || null;
    const subActivityId = e.target.subActivityId?.value || null;
    const approvedBudget = parseFloat(e.target.approvedBudget?.value) || 0;
    const receivedInTry = parseFloat(e.target.receivedInTry?.value) || 0;

    if (!schemeId) {
      showAlert("Scheme is mandatory.");
      return;
    }

    try {
      const data = { 
        name, 
        isProvisional: !rawName,
        schemeId, 
        sectorId, 
        activityId, 
        subActivityId, 
        approvedBudget, 
        approvedBudgetAmount: approvedBudget,
        receivedInTry,
        receivedInTryAmount: receivedInTry,
        tryAmount: receivedInTry,
        financialYear: selectedFY,
        updatedAt: Date.now()
      };
      if (editingItem?.type === 'SOE Name') {
        await updateDoc(doc(db, 'soeHeads', editingItem.item.id), { ...data, updatedAt: Date.now() });
        logAuditAction('SOE Head Updated', `SOE: ${name}, Scheme: ${schemes.find(s => s.id === schemeId)?.name || 'N/A'}`);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'soeHeads'), { ...data, createdAt: Date.now(), updatedAt: Date.now() });
        logAuditAction('SOE Head Created', `SOE: ${name}, Scheme: ${schemes.find(s => s.id === schemeId)?.name || 'N/A'}`);
      }
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'SOE Name' ? OperationType.UPDATE : OperationType.CREATE, 'soeHeads');
    }
  };

  const handleUpdateSoeTry = async (soeId: string, amount: number) => {
    try {
      await updateDoc(doc(db, 'soeHeads', soeId), { 
        receivedInTry: amount,
        receivedInTryAmount: amount,
        tryAmount: amount,
        updatedAt: Date.now()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'soeHeads');
    }
  };



  useEffect(() => {
    setIsSoeTrackerExpanded(false);
  }, [activeTab]);

  const allocationBudgetStatus = useMemo(() => {
    if (activeTab !== 'Allocations' || !allocationFormFilters.schemeId) return { isInvalid: false, remaining: 0, availableBudget: 0, currentAllocated: 0 };
    const amount = parseFloat(allocationAmount);
    const isEditing = editingItem?.type === 'Allocation';

    const { schemeId, sectorId, activityId, subActivityId, fundingSoeName } = allocationFormFilters;

    // Check against expenditure if editing
    if (isEditing) {
      const spent = expenses
        .filter(e => e.allocationId === editingItem.item.id && e.status !== 'rejected')
        .reduce((sum, e) => sum + e.amount, 0);
      if (!isNaN(amount) && amount < spent) return { isInvalid: true, remaining: 0, availableBudget: 0, currentAllocated: 0, error: `Amount cannot be less than expenditure (â‚¹${spent.toLocaleString()})` };
    }

    // 1. Identify relevant SOEs (the "Pool")
    // Use path-inclusive logic: if we are at a deep level, we can fund from SOEs at this level or any parent level.
    const poolSoes = currentSoes.filter((s: any) => {
      if (s.schemeId !== schemeId) return false;
      if (fundingSoeName && s.name !== fundingSoeName) return false;
      if (subActivityId && s.subActivityId && s.subActivityId !== subActivityId) return false;
      if (activityId && s.activityId && s.activityId !== activityId) return false;
      if (sectorId && s.sectorId && s.sectorId !== sectorId) return false;
      return true;
    });

    // 2. Calculate Available Budget in this Pool
    const availableBudget = poolSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0);

    // 3. Calculate Already Allocated from this Pool
    const totalAllocated = currentAllocations.reduce((sum, a) => {
      const currentAllocId = isEditing ? editingItem.item.id : null;
      if (a.id === currentAllocId) return sum;

      // Count funded amounts from our pool
      const fundedFromOurPool = a.fundedSOEs?.filter((f: any) => poolSoes.some(s => s.id === f.soeId)) || [];
      const fundedAmount = fundedFromOurPool.reduce((s: number, f: any) => s + f.amount, 0);
      
      // Count unfunded portion if it belongs to this hierarchy
      // This ensures that even if an allocation isn't funded yet, it's reserved from the pool.
      const matchesHierarchy = 
        a.schemeId === schemeId &&
        (!sectorId || a.sectorId === sectorId) &&
        (!activityId || a.activityId === activityId) &&
        (!subActivityId || a.subActivityId === subActivityId);
      
      if (matchesHierarchy) {
        const totalFunded = a.fundedSOEs?.reduce((s, f) => s + f.amount, 0) || 0;
        const unfunded = Math.max(0, a.amount - totalFunded);
        return sum + fundedAmount + unfunded;
      }

      return sum + fundedAmount;
    }, 0);

    // 4. Subtract surrenders to get the NET allocated amount
    const totalSurrendered = surrenders.reduce((sum, s) => {
      if (poolSoes.some(ps => ps.id === s.soeId)) {
        return sum + s.amount;
      }
      return sum;
    }, 0);

    const currentNetAllocated = totalAllocated - totalSurrendered;
    const remaining = availableBudget - currentNetAllocated;
    const isInvalid = !isNaN(amount) && amount > 0 && amount > remaining;

    return { isInvalid, remaining, availableBudget, currentAllocated: currentNetAllocated };
  }, [activeTab, allocationAmount, allocationFormFilters, currentSoes, currentAllocations, expenses, surrenders, editingItem]);

  const isAllocationInvalid = allocationBudgetStatus.isInvalid || !allocationFormFilters.rangeId;

  const handleAddAllocation = async (e: any) => {
    e.preventDefault();
    if (isFeatureLocked('Allocation')) {
      showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
      return;
    }
    const rangeId = e.target.rangeId.value;
    const amount = parseFloat(e.target.amount.value);
    const remarks = e.target.remarks.value || '';
    const schemeId = e.target.schemeId.value || null;
    const sectorId = e.target.sectorId.value || null;
    const activityId = e.target.activityId.value || null;
    const subActivityId = e.target.subActivityId.value || null;
    const targetFyId = selectedFY;
    
    const fundingSoeName = e.target.fundingSoeName?.value || null;
    
    if (isNaN(amount) || amount <= 0) {
      showAlert("Please enter a valid positive amount.");
      return;
    }

    // Validation: Check against expenditure if editing
    if (editingItem?.type === 'Allocation') {
      const spent = expenses
        .filter(e => e.allocationId === editingItem.item.id && e.status !== 'rejected')
        .reduce((sum, e) => sum + e.amount, 0);
      if (amount < spent) {
        showAlert(`Cannot reduce allocation below expenditure. Already spent: â‚¹${spent.toLocaleString()}`);
        return;
      }
    }
    
    // Validation: Check against Available Budget
    // Use path-inclusive logic to find relevant SOEs (allows pulling from higher levels)
    const branchSoes = soes.filter((s: any) => {
      if (s.schemeId !== schemeId) return false;
      if (subActivityId && s.subActivityId && s.subActivityId !== subActivityId) return false;
      if (activityId && s.activityId && s.activityId !== activityId) return false;
      if (sectorId && s.sectorId && s.sectorId !== sectorId) return false;
      return true;
    });
    
    // If a specific SOE name is selected, validate against that SOE's balance in the branch
    if (fundingSoeName) {
      const matchedSoes = branchSoes.filter(s => s.name === fundingSoeName);
      const received = matchedSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0);
      const allocated = allocations.reduce((sum, a) => {
        const currentAllocId = editingItem?.type === 'Allocation' ? editingItem.item.id : null;
        if (a.id === currentAllocId) return sum;
        const fundedFromThese = a.fundedSOEs?.filter((f: any) => matchedSoes.some(s => s.id === f.soeId)) || [];
        return sum + fundedFromThese.reduce((s: number, f: any) => s + f.amount, 0);
      }, 0);
      const surrendered = surrenders.reduce((sum, s) => {
        if (matchedSoes.some(ms => ms.id === s.soeId)) return sum + s.amount;
        return sum;
      }, 0);
      const remaining = received - (allocated - surrendered);

      if (amount > remaining) {
        showAlert(`Cannot allocate. Amount â‚¹${amount.toLocaleString()} exceeds the remaining balance of SOE ${fundingSoeName} (â‚¹${remaining.toLocaleString()}).`);
        return;
      }
    } else {
      // General branch-wide validation
      const totalReceived = branchSoes.reduce((sum, s) => sum + getReceivedInTry(s), 0);
      const totalAllocated = allocations.reduce((sum, a) => {
        const currentAllocId = editingItem?.type === 'Allocation' ? editingItem.item.id : null;
        if (a.id === currentAllocId) return sum;
        const fundedFromBranch = a.fundedSOEs?.filter((f: any) => branchSoes.some(s => s.id === f.soeId)) || [];
        return sum + fundedFromBranch.reduce((s: number, f: any) => s + f.amount, 0);
      }, 0);
      const totalSurrendered = surrenders.reduce((sum, s) => {
        if (branchSoes.some(bs => bs.id === s.soeId)) return sum + s.amount;
        return sum;
      }, 0);

      const remaining = totalReceived - (totalAllocated - totalSurrendered);

      if (amount > remaining) {
        showAlert(`Cannot allocate. Amount â‚¹${amount.toLocaleString()} exceeds the remaining Available Budget of â‚¹${remaining.toLocaleString()}.`);
        return;
      }
    }

    try {
      let fundedSOEs: any[] = [];
      let status = 'Pending SOE Funds';
      
      const rangeName = ranges.find(r => r.id === rangeId)?.name || rangeId;
      const schemeName = schemes.find(s => s.id === schemeId)?.name || 'General';
      const sectorName = sectors.find(s => s.id === sectorId)?.name || 'General';
      const activityName = activities.find(a => a.id === activityId)?.name || 'General';

      if (fundingSoeName) {
        const matchedSoes = branchSoes.filter(s => s.name === fundingSoeName);
        let remainingToFund = amount;
        
        for (const soe of matchedSoes) {
          if (remainingToFund <= 0) break;
          const received = getReceivedInTry(soe);
          const allocated = allocations.reduce((sum, a) => {
            const fundedFromThis = a.fundedSOEs?.find((f: any) => f.soeId === soe.id);
            const currentAllocId = editingItem?.type === 'Allocation' ? editingItem.item.id : null;
            if (a.id === currentAllocId) return sum;
            return sum + (fundedFromThis?.amount || 0);
          }, 0);
          const surrendered = (surrenders || []).filter(s => s.soeId === soe.id).reduce((sum, s) => sum + s.amount, 0);
          const available = received - (allocated - surrendered);
          
          if (available > 0) {
            const fundAmount = Math.min(available, remainingToFund);
            fundedSOEs.push({ soeId: soe.id, amount: fundAmount });
            remainingToFund -= fundAmount;
          }
        }
        
        if (remainingToFund <= 0) {
          status = 'Funded';
        }
      }

      if (editingItem?.type === 'Allocation') {
        await updateDoc(doc(db, 'allocations', editingItem.item.id), { 
          rangeId, amount, remarks, schemeId, sectorId, activityId, subActivityId, financialYear: targetFyId,
          status, fundedSOEs,
          updatedAt: Date.now()
        });
        logAuditAction('Allocation Updated', `Range: ${rangeName}, Amount: â‚¹${amount.toLocaleString()}, SOE: ${fundingSoeName || 'Multiple'}`);
        
        try {
          await addDoc(collection(db, 'notifications'), {
            name: `Budget Allocation Updated: â‚¹${amount.toLocaleString('en-IN')} to ${rangeName}`,
            description: `Scheme: ${schemeName} | Sector: ${sectorName || 'General'} | Activity: ${activityName} | SOE: ${fundingSoeName || 'Multiple'} | FY: ${targetFyId}`,
            url: '',
            type: 'system/allocation',
            category: 'budget_allocation',
            amount: amount,
            rangeName: rangeName,
            schemeName: schemeName,
            sectorName: sectorName || '',
            soeName: fundingSoeName || '',
            targetRanges: [rangeName, 'All'],
            createdAt: Date.now(),
            uploadedBy: user?.uid || 'Admin'
          });
        } catch (notifErr) {
          console.warn('Could not post notification for allocation update:', notifErr);
        }

        setEditingItem(null);
      } else {
        // Always create a new entry for every allocation as requested
        await addDoc(collection(db, 'allocations'), { 
          rangeId, amount, remarks, schemeId, sectorId, activityId, subActivityId, financialYear: targetFyId,
          status,
          fundedSOEs,
          createdAt: Date.now(),
          updatedAt: Date.now()
        });
        logAuditAction('Allocation Created', `Range: ${rangeName}, Amount: â‚¹${amount.toLocaleString()}, SOE: ${fundingSoeName || 'Multiple'}`);
        
        try {
          await addDoc(collection(db, 'notifications'), {
            name: `New Budget Allocation: â‚¹${amount.toLocaleString('en-IN')} to ${rangeName}`,
            description: `Scheme: ${schemeName} | Sector: ${sectorName || 'General'} | Activity: ${activityName} | SOE: ${fundingSoeName || 'Multiple'} | FY: ${targetFyId}`,
            url: '',
            type: 'system/allocation',
            category: 'budget_allocation',
            amount: amount,
            rangeName: rangeName,
            schemeName: schemeName,
            sectorName: sectorName || '',
            soeName: fundingSoeName || '',
            targetRanges: [rangeName, 'All'],
            createdAt: Date.now(),
            uploadedBy: user?.uid || 'Admin'
          });
        } catch (notifErr) {
          console.warn('Could not post notification for allocation create:', notifErr);
        }
      }
      e.target.reset();
      setAllocationAmount('');
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Allocation' ? OperationType.UPDATE : OperationType.CREATE, 'allocations');
    }
  };

  const handleFundAllocation = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!fundingAllocation) return;

    const form = e.target as HTMLFormElement;
    const formData = new FormData(form);
    const soeId = formData.get('soeId') as string;
    const amount = parseFloat(formData.get('amount') as string);

    if (!soeId || isNaN(amount) || amount <= 0) {
      showAlert("Please select an SOE and enter a valid amount.");
      return;
    }

    // Check treasury availability for this SOE
    const soe = currentSoes.find(s => s.id === soeId);
    const totalReceived = soe ? getReceivedInTry(soe) : 0;

    const totalFundedFromThisSoe = baseAllocations
      .reduce((sum, a) => {
        const funded = a.fundedSOEs?.find(f => f.soeId === soeId);
        return sum + (funded?.amount || 0);
      }, 0);

    const availableInTry = totalReceived - totalFundedFromThisSoe;

    if (amount > availableInTry) {
      showAlert(`Insufficient funds in Treasury for this SOE. Available: â‚¹${availableInTry.toLocaleString()}`);
      return;
    }

    // Check if this funding exceeds the allocation's remaining amount
    const alreadyFundedTotal = fundingAllocation.fundedSOEs?.reduce((sum, f) => sum + f.amount, 0) || 0;
    const remainingToFund = fundingAllocation.amount - alreadyFundedTotal;

    if (amount > remainingToFund) {
      showAlert(`Funding amount â‚¹${amount.toLocaleString()} exceeds the remaining allocation requirement of â‚¹${remainingToFund.toLocaleString()}.`);
      return;
    }

    try {
      const updatedFundedSOEs = [...(fundingAllocation.fundedSOEs || [])];
      const existingIdx = updatedFundedSOEs.findIndex(f => f.soeId === soeId);
      if (existingIdx >= 0) {
        updatedFundedSOEs[existingIdx].amount += amount;
      } else {
        updatedFundedSOEs.push({ soeId, amount });
      }

      const totalFundedNow = updatedFundedSOEs.reduce((sum, f) => sum + f.amount, 0);
      const newStatus = totalFundedNow >= fundingAllocation.amount ? 'Funded' : 'Pending SOE Funds';

      await updateDoc(doc(db, 'allocations', fundingAllocation.id), {
        fundedSOEs: updatedFundedSOEs,
        status: newStatus,
        updatedAt: Date.now()
      });

      setFundingAllocation(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'allocations');
    }
  };

  const handleUpdateExpenseStatus = async (expenseId: string, status: 'approved' | 'rejected' | 'pending', isLocked: boolean, reason?: string) => {
    try {
      if (status === 'approved') {
        await runTransaction(db, async (transaction) => {
          const counterDocRef = doc(db, 'appSettings', 'counters');
          const counterDoc = await transaction.get(counterDocRef);
          
          let nextId = 100;
          if (counterDoc.exists()) {
            nextId = (counterDoc.data().lastApprovalId || 99) + 1;
          }
          
          transaction.set(counterDocRef, { lastApprovalId: nextId }, { merge: true });
          transaction.update(doc(db, 'expenditures', expenseId), { 
            status, 
            isLocked, 
            approvalId: nextId,
            approvalReason: reason || '',
            updatedAt: Date.now()
          });
          logAuditAction('Expenditure Approved', `Status: ${status}, Approval ID: ${nextId}`);
        });
      } else {
        await updateDoc(doc(db, 'expenditures', expenseId), { 
          status, 
          isLocked,
          approvalReason: reason || '',
          updatedAt: Date.now(),
          ...(status === 'pending' ? { approvalId: null } : {})
        });
        logAuditAction('Expenditure Status Updated', `Status: ${status}, Reason: ${reason || 'N/A'}`);
      }
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'expenditures');
      showAlert(`Error updating status: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  };

  const handleResetUnbilledExpenses = async () => {
    const unbilledApproved = expenses.filter(e => 
      e.status === 'approved' && 
      !bills.some(b => b.expenseIds.includes(e.id))
    );

    if (unbilledApproved.length === 0) {
      showAlert("No unbilled approved expenditures found.");
      return;
    }

    showConfirm(`Are you sure you want to reset ${unbilledApproved.length} unbilled approved expenditures to pending?`, async () => {
      try {
        const batch = writeBatch(db);
        unbilledApproved.forEach(exp => {
          batch.update(doc(db, 'expenditures', exp.id), {
            status: 'pending',
            isLocked: false,
            approvalId: null,
            updatedAt: Date.now()
          });
        });
        await batch.commit();
        showAlert(`Successfully reset ${unbilledApproved.length} expenditures to pending.`);
      } catch (error) {
        handleFirestoreError(error, OperationType.UPDATE, 'expenditures');
      }
    });
  };

  const isExpenseInvalid = useMemo(() => {
    // Basic validation for required selections
    if (!expenseFormSelection.schemeId || !expenseFormSelection.rangeId || !expenseFormSelection.soeId || !expenseFormSelection.allocationId) {
      return true;
    }
    
    if (currentSoeBalance === undefined) {
      return true;
    }

    if (!expenseDate) {
      return true;
    }
    
    let amount = parseFloat(expenseAmount || '0');
    if (selectedPayeesForExpense.length > 0) {
      amount = selectedPayeesForExpense.reduce((sum, p) => sum + parseFloat(p.amount || '0'), 0);
    }
    
    if (amount <= 0) return true;
    return amount > currentSoeBalance;
  }, [currentSoeBalance, expenseAmount, selectedPayeesForExpense, expenseFormSelection, expenseDate, expenseDescription]);

  useEffect(() => {
    if (editingItem?.type === 'Expenditure') {
      setExpenseAmount(String(editingItem.item.amount));
      setExpenseDate(editingItem.item.date);
      setExpenseDescription(editingItem.item.description || '');
      setSelectedPayeesForExpense(editingItem.item.payeeId ? [{ payeeId: editingItem.item.payeeId, amount: String(editingItem.item.amount) }] : []);
      const deds: ('TDS' | 'TDS_GST')[] = [];
      if (editingItem.item.tdsAmount || editingItem.item.deductionType === 'TDS' || editingItem.item.deductionType === 'Both') deds.push('TDS');
      if (editingItem.item.tdsGstAmount || editingItem.item.deductionType === 'TDS_GST' || editingItem.item.deductionType === 'Both') deds.push('TDS_GST');
      setSelectedDeductions(deds);
      setPanNumber(editingItem.item.panNumber || '');
      setGstNumber(editingItem.item.gstNumber || '');
    } else {
      setExpenseAmount('');
      setExpenseDate(new Date().toISOString().split('T')[0]);
      setExpenseDescription('');
      setSelectedPayeesForExpense([]);
      setSelectedDeductions([]);
      setPanNumber('');
      setGstNumber('');
    }
  }, [editingItem]);

  const handleAddExpense = async (e: any) => {
    e.preventDefault();
    if (isFeatureLocked('Expenditure')) {
      showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
      return;
    }

    // Safely extract field values even if multiple elements share the same name in DOM
    const getFieldValue = (fieldName: string) => {
      const field = e.target[fieldName];
      if (!field) return '';
      if (field.value !== undefined && typeof field.value === 'string') return field.value;
      if (field instanceof HTMLCollection || field instanceof NodeList) {
        for (let i = 0; i < field.length; i++) {
          const val = (field[i] as any).value;
          if (val) return val;
        }
      }
      return '';
    };

    const allocationId = getFieldValue('allocationId') || expenseFormSelection?.allocationId;
    const soeId = getFieldValue('soeId') || expenseFormSelection?.soeId;
    const amount = selectedPayeesForExpense.length > 0 
      ? selectedPayeesForExpense.reduce((sum, p) => sum + (parseFloat(p.amount) || 0), 0)
      : parseFloat(expenseAmount || '0');
    const date = expenseDate;
    const description = expenseDescription;
    const targetFyId = selectedFY;

    const today = new Date().toISOString().split('T')[0];
    if (date > today) {
      showAlert("Cannot add expenditure for a future date.");
      return;
    }

    if (!allocationId) {
      showAlert("Please select a valid SOE Head / Allocation before adding expenditure.");
      return;
    }

    const alloc = allocations.find(a => a.id === allocationId);
    if (!alloc) {
      showAlert("Selected allocation could not be found. Please re-select the SOE Head.");
      return;
    }

    const selectedSoe = soes.find(s => s.id === soeId);
    const selectedName = selectedSoe?.name || 'Unnamed SOE';

    // Aggregate allocation for this hierarchy and SOE Name
    const totalAllocatedForSoe = allocations.filter(a => 
      a.rangeId === alloc.rangeId &&
      a.schemeId === alloc.schemeId &&
      (a.sectorId || null) === (alloc.sectorId || null) &&
      (a.activityId || null) === (alloc.activityId || null) &&
      (a.subActivityId || null) === (alloc.subActivityId || null)
    ).reduce((sum, a) => {
      const funded = a.fundedSOEs?.find((f: any) => {
        const s = soes.find((soe: any) => soe.id === f.soeId);
        return (s?.name || 'Unnamed SOE') === selectedName;
      });
      return sum + (funded?.amount || 0);
    }, 0);

    // Aggregate expenditure for this hierarchy and SOE Name
    const totalSpentForSoe = expenses.filter(e => {
      const eAlloc = allocations.find(a => a.id === e.allocationId);
      const eSoeName = soes.find(s => s.id === e.soeId)?.name;
      return (
        eAlloc &&
        eAlloc.rangeId === alloc.rangeId &&
        eAlloc.schemeId === alloc.schemeId &&
        (eAlloc.sectorId || null) === (alloc.sectorId || null) &&
        (eAlloc.activityId || null) === (alloc.activityId || null) &&
        (eAlloc.subActivityId || null) === (alloc.subActivityId || null) &&
        eSoeName === selectedName &&
        e.status !== 'rejected' &&
        (editingItem?.type === 'Expenditure' ? e.id !== editingItem.item.id : true)
      );
    }).reduce((sum, e) => sum + e.amount, 0);

    const currentBalance = totalAllocatedForSoe - totalSpentForSoe;

    // If we have selected payees, we create multiple expenditures
    if (selectedPayeesForExpense.length > 0 && editingItem?.type !== 'Expenditure') {
      const totalAmount = selectedPayeesForExpense.reduce((sum, p) => sum + (parseFloat(p.amount) || 0), 0);
      
      if (totalAmount > currentBalance) {
        showAlert(`Insufficient funds in SOE ${selectedName}. Remaining: â‚¹${currentBalance.toLocaleString()}`);
        return;
      }

      try {
        const batch = writeBatch(db);
        for (const p of selectedPayeesForExpense) {
          const pAmt = Math.round(parseFloat(p.amount) || 0);
          const pPayeeObj = payees.find(payee => payee.id === p.payeeId);
          const pGst = pPayeeObj?.gstNumber || gstNumber;
          const pPan = pPayeeObj?.panNumber || panNumber;
          let pTdsAmt = 0;
          let pTdsGstAmt = 0;
          
          if (selectedDeductions.includes('TDS') && pAmt > 30000) {
            pTdsAmt = Math.round(pAmt * 0.01);
          }
          if (selectedDeductions.includes('TDS_GST') && pAmt > 250000) {
            pTdsGstAmt = Math.round(pAmt * 0.02);
          }
          
          const pTotalDeducted = pTdsAmt + pTdsGstAmt;
          const pNetAmt = pAmt - pTotalDeducted;

          const docRef = doc(collection(db, 'expenditures'));
          batch.set(docRef, {
            allocationId, soeId, amount: pAmt, date, description, financialYear: targetFyId,
            rangeId: alloc.rangeId,
            createdBy: user.uid,
            createdByRole: userRole,
            status: 'pending',
            isLocked: false,
            createdAt: Date.now(),
            updatedAt: Date.now(),
            payeeId: p.payeeId,
            deductionType: selectedDeductions.length > 0 ? (selectedDeductions.length > 1 ? 'Both' : selectedDeductions[0]) : 'None',
            deductions: selectedDeductions.map(type => ({ type, amount: type === 'TDS' ? pTdsAmt : pTdsGstAmt })),
            deductedAmount: pTotalDeducted || null,
            tdsAmount: pTdsAmt || null,
            tdsGstAmount: pTdsGstAmt || null,
            netAmount: pNetAmt,
            panNumber: pPan || null,
            gstNumber: pGst || null,
            syncedMemoId: selectedSyncedMemo?.memoId || null,
            syncedMemoNo: selectedSyncedMemo?.memoNo || null
          });
        }
        await batch.commit();
        logAuditAction('Expenditure Batch Created', `${selectedPayeesForExpense.length} payees added, Total: â‚¹${totalAmount.toLocaleString()}`);
        setSelectedPayeesForExpense([]);
        setCurrentSoeBalance(undefined);
        setExpenseAmount('');
        setExpenseDescription('');
        setSelectedDeductions([]);
        setPanNumber('');
        setGstNumber('');
        setSelectedSyncedMemo(null);
        showAlert(`${selectedPayeesForExpense.length} expenditures added successfully.`);
      } catch (error) {
        handleFirestoreError(error, OperationType.CREATE, 'expenditures');
      }
      return;
    }

    if (!amount || amount <= 0) {
      showAlert("Please provide a valid amount.");
      return;
    }

    if (amount > currentBalance) {
      showAlert(`Insufficient funds in SOE ${selectedName}. Remaining: â‚¹${currentBalance.toLocaleString()}`);
      return;
    }

    // Calculate deductions
    const amt = Math.round(amount);
    let tdsAmt = 0;
    let tdsGstAmt = 0;
    const hasGst = Boolean(gstNumber && gstNumber.trim());
    
    if (selectedDeductions.includes('TDS') && amt > 30000) {
      tdsAmt = Math.round(amt * 0.01);
    }
    if (selectedDeductions.includes('TDS_GST') && amt > 250000) {
      tdsGstAmt = Math.round(amt * 0.02);
    }
    
    const finalDeductedAmount = tdsAmt + tdsGstAmt;
    const netAmount = amt - finalDeductedAmount;

    try {
      if (editingItem?.type === 'Expenditure') {
        const payeeId = e.target.payeeId?.value;
        const payeeName = e.target.payeeName?.value;
        await updateDoc(doc(db, 'expenditures', editingItem.item.id), { 
          allocationId, soeId, amount, date, description, financialYear: targetFyId, rangeId: alloc.rangeId,
          payeeId: payeeId || null,
          payeeName: payeeName || null,
          deductionType: selectedDeductions.length > 0 ? (selectedDeductions.length > 1 ? 'Both' : selectedDeductions[0]) : 'None',
          deductions: selectedDeductions.map(type => ({ type, amount: type === 'TDS' ? tdsAmt : tdsGstAmt })),
          deductedAmount: finalDeductedAmount || null,
          tdsAmount: tdsAmt || null,
          tdsGstAmount: tdsGstAmt || null,
          netAmount: netAmount,
          panNumber: panNumber || null,
          gstNumber: gstNumber || null,
          updatedBy: user.uid,
          updatedByRole: userRole,
          updatedAt: Date.now()
        });
        logAuditAction('Expenditure Updated', `Amount: â‚¹${amount.toLocaleString()}, SOE: ${selectedName}`);
        setEditingItem(null);
        setExpenseAmount('');
        setExpenseDescription('');
        setSelectedDeductions([]);
        setPanNumber('');
        setGstNumber('');
        setCurrentSoeBalance(undefined);
      } else {
        const payeeName = e.target.payeeName?.value;
        const payeeId = e.target.payeeId?.value;
        await addDoc(collection(db, 'expenditures'), { 
          allocationId, soeId, amount, date, description, financialYear: targetFyId,
          rangeId: alloc.rangeId,
          createdBy: user.uid,
          createdByRole: userRole,
          status: 'pending',
          isLocked: false,
          createdAt: Date.now(),
          updatedAt: Date.now(),
          payeeName: payeeName || null,
          payeeId: payeeId || null,
          deductionType: selectedDeductions.length > 0 ? (selectedDeductions.length > 1 ? 'Both' : selectedDeductions[0]) : 'None',
          deductions: selectedDeductions.map(type => ({ type, amount: type === 'TDS' ? tdsAmt : tdsGstAmt })),
          deductedAmount: finalDeductedAmount || null,
          tdsAmount: tdsAmt || null,
          tdsGstAmount: tdsGstAmt || null,
          netAmount: netAmount,
          panNumber: panNumber || null,
          gstNumber: gstNumber || null,
          syncedMemoId: selectedSyncedMemo?.memoId || null,
          syncedMemoNo: selectedSyncedMemo?.memoNo || null
        });
        logAuditAction('Expenditure Created', `Amount: â‚¹${amount.toLocaleString()}, SOE: ${selectedName}`);
        setCurrentSoeBalance(undefined);
        setExpenseAmount('');
        setExpenseDescription('');
        setSelectedDeductions([]);
        setPanNumber('');
        setGstNumber('');
        setSelectedSyncedMemo(null);
      }
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Expenditure' ? OperationType.UPDATE : OperationType.CREATE, 'expenditures');
    }
  };

  const handleAddPayee = async (e: any) => {
    e.preventDefault();
    const name = e.target.name.value;
    const address = e.target.address.value;
    const accountNumber = e.target.accountNumber.value;
    const ifscCode = e.target.ifscCode?.value || '';
    const panNumber = e.target.panNumber?.value || '';
    const gstNumber = e.target.gstNumber?.value || '';
    const treasuryCode = (isAdmin() || isDEO()) ? (e.target.treasuryCode?.value || '') : (editingItem?.type === 'Payee' ? (editingItem.item.treasuryCode || '') : '');
    const rangeId = e.target.rangeId?.value || null;

    // Check duplicate account number
    const cleanAcc = (accountNumber || '').trim().replace(/\s+/g, '').toLowerCase();
    const duplicate = payees.find(p => {
      if (editingItem?.type === 'Payee' && p.id === editingItem.item.id) return false;
      const pAcc = (p.accountNumber || '').trim().replace(/\s+/g, '').toLowerCase();
      if (!pAcc || !cleanAcc) return false;
      return pAcc === cleanAcc;
    });

    if (duplicate) {
      setDuplicatePayeeModalData({
        existingPayee: duplicate,
        enteredName: (name || '').trim(),
        enteredAccountNo: (accountNumber || '').trim()
      });
      return;
    }

    try {
      if (editingItem?.type === 'Payee') {
        await updateDoc(doc(db, 'payees', editingItem.item.id), {
          name, address, accountNumber, ifscCode, panNumber, gstNumber, treasuryCode: treasuryCode ? treasuryCode.trim() : null, rangeId: rangeId || null,
          updatedAt: Date.now()
        });
        logAuditAction('Payee Updated', `Payee: ${name}, Treasury Code: ${treasuryCode || 'N/A'}`);
        setEditingItem(null);
        e.target.reset();
      } else {
        await addDoc(collection(db, 'payees'), {
          name, address, accountNumber, ifscCode, panNumber, gstNumber, treasuryCode: treasuryCode ? treasuryCode.trim() : null, rangeId: rangeId || null,
          createdBy: user.uid,
          createdAt: Date.now(),
          updatedAt: Date.now()
        });
        logAuditAction('Payee Created', `Payee: ${name}, Treasury Code: ${treasuryCode || 'N/A'}`);
        e.target.reset();
      }
      showAlert(`Payee ${editingItem?.type === 'Payee' ? 'updated' : 'added'} successfully.`);
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Payee' ? OperationType.UPDATE : OperationType.CREATE, 'payees');
    }
  };

  const handleMapPayeeToUser = async (existingPayee: Payee) => {
    if (!user) return;
    try {
      const currentMappedUserIds = existingPayee.mappedUserIds || [];
      const currentMappedRangeIds = existingPayee.mappedRangeIds || [];

      const newMappedUserIds = currentMappedUserIds.includes(user.uid)
        ? currentMappedUserIds
        : [...currentMappedUserIds, user.uid];

      const newMappedRangeIds = (userRangeId && !currentMappedRangeIds.includes(userRangeId))
        ? [...currentMappedRangeIds, userRangeId]
        : currentMappedRangeIds;

      await updateDoc(doc(db, 'payees', existingPayee.id), {
        mappedUserIds: newMappedUserIds,
        mappedRangeIds: newMappedRangeIds,
        updatedAt: Date.now()
      });

      logAuditAction(
        'Payee Mapped',
        `Mapped Payee ${existingPayee.name} (${existingPayee.accountNumber}) to User ${user.email || user.uid} / Range ${userRangeName || userRole || 'N/A'}`
      );

      setDuplicatePayeeModalData(null);
      showAlert(`Payee "${existingPayee.name}" has been mapped to your account (${userRangeName || userRole || 'your range'}) successfully! It is now available in your Payee Directory and Memo for Fund.`);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `payees/${existingPayee.id}`);
    }
  };

  const downloadPayeesPDF = () => {
    try {
      const doc = new jsPDF('landscape');
      doc.setFontSize(16);
      doc.text('Forest Budget Control System - Payee Details', 14, 15);
      doc.setFontSize(10);
      doc.text(`Role / Range: ${userRangeName || userRole || 'All'} | Date: ${new Date().toLocaleDateString('en-IN')}`, 14, 22);

      const tableData = filteredPayeesList.map((p, index) => [
        index + 1,
        p.name || '',
        p.address || '',
        p.treasuryCode || 'N/A',
        `${p.accountNumber || ''}\n${p.ifscCode || ''}`,
        p.panNumber || 'N/A',
        p.gstNumber || 'N/A',
        ranges.find(r => r.id === p.rangeId)?.name || 'N/A'
      ]);

      autoTable(doc, {
        startY: 28,
        head: [['S.No', 'Payee Name', 'Address', 'Try Code', 'Bank Account Details', 'PAN No', 'GST No', 'Range']],
        body: tableData,
        theme: 'grid',
        headStyles: { fillColor: [0, 77, 64] },
        styles: { fontSize: 8 }
      });

      doc.save(`Payees_List_${userRole || 'All'}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error("PDF export error:", err);
      showAlert("Failed to export PDF.");
    }
  };

  const markMemoAsViewedByAdmin = useCallback(async (memo: MemoForFund) => {
    if (!memo || !memo.id || memo.viewedByAdmin) return;
    if (!isAuthorizedUserOrAdmin) return;

    try {
      const roleLabel = userRole === 'admin' ? 'Admin' : (userRole === 'deo' ? 'DEO' : (userRole === 'approver' ? 'Approver' : userRole || 'Headquarter'));
      const authorLabel = user?.displayName ? `${roleLabel} (${user.displayName})` : (user?.email ? `${roleLabel} (${user.email})` : roleLabel);
      await updateDoc(doc(db, 'memos', memo.id), {
        viewedByAdmin: true,
        viewedAt: Date.now(),
        viewedBy: authorLabel,
        viewedByRole: userRole || 'admin',
        updatedAt: Date.now()
      });
      logAuditAction('Memo Viewed by Admin', `Memo No: ${memo.memoNo} viewed by ${authorLabel}`);
    } catch (err) {
      console.warn("Failed to mark memo as viewed by admin:", err);
    }
  }, [isAuthorizedUserOrAdmin, userRole, user?.displayName, user?.email]);

  const handleViewMemo = useCallback((memo: MemoForFund) => {
    setViewingMemo(memo);
    markMemoAsViewedByAdmin(memo);
  }, [markMemoAsViewedByAdmin]);

  const downloadMemoPDF = (memo: MemoForFund) => {
    try {
      markMemoAsViewedByAdmin(memo);
      const doc = new jsPDF('landscape', 'mm', 'a4');
      const rangeTitle = memo.rangeName ? memo.rangeName.replace(/^RFO\s*/i, '').replace(/\s*Range$/i, '').replace(/\s*Office$/i, '') : (userRangeName || 'Sarahan');

      doc.setFontSize(11);
      doc.setFont("helvetica", "bold");
      doc.text("H.P. FOREST DEPARTMENT", 148.5, 12, { align: "center" });
      doc.setFontSize(13);
      doc.text(`OFFICE OF THE RANGE FOREST OFFICER, ${rangeTitle.toUpperCase()}`, 148.5, 18, { align: "center" });
      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.text("Forest Division Rajgarh, District Sirmaur (H.P.)", 148.5, 23, { align: "center" });

      doc.setLineWidth(0.4);
      doc.line(12, 26, 285, 26);

      if (memo.isApproved) {
        doc.setFontSize(8.5);
        doc.setTextColor(5, 122, 85);
        doc.setFont("helvetica", "bold");
        doc.text(`[SANCTIONED & APPROVED FOR FUND RELEASE - ${memo.approvedBy || 'Office of DCF Rajgarh'}]`, 148.5, 30, { align: "center" });
        doc.setTextColor(0, 0, 0);
      }

      doc.setFontSize(9.5);
      doc.setFont("helvetica", "bold");
      doc.text(`No. ${memo.memoNo || ''}`, 12, 31);
      const dateFormatted = memo.date ? memo.date.split('-').reverse().join('.') : '';
      doc.text(`Dated: ${dateFormatted}`, 285, 31, { align: "right" });

      doc.setLineWidth(0.2);
      doc.line(12, 33, 285, 33);

      // Single-line From & To Header
      doc.setFontSize(9);
      doc.setFont("helvetica", "bold");
      doc.text("From:", 12, 38);
      doc.setFont("helvetica", "normal");
      doc.text(`Range Forest Officer, ${memo.rangeName || rangeTitle}.`, 24, 38);

      doc.setFont("helvetica", "bold");
      doc.text("To:", 155, 38);
      doc.setFont("helvetica", "normal");
      doc.text("The Divisional Forest Officer, Rajgarh Forest Division (H.P.).", 162, 38);

      doc.setFont("helvetica", "bold");
      doc.text("Subject:", 12, 44);
      const schemeText = memo.schemeName || 'All Schemes';
      const sectorText = memo.sectorName ? ` (${memo.sectorName})` : '';
      const soeText = memo.soeName ? ` [SOE: ${memo.soeName}]` : '';
      const subjStr = `Memo for Fund for the month of ${memo.monthYear} under scheme ${schemeText}${sectorText}${soeText}.`;
      doc.text(subjStr, 28, 44, { maxWidth: 255 });

      doc.text("Sir,", 12, 50);
      const totalGrossAmt = Math.round(Number(memo.totalAmount) || 0);
      const totalITaxAmt = Math.round(Number(memo.totalITax) || 0);
      const totalGstAmt = Math.round(Number(memo.totalGst) || 0);
      const totalNetRtgsAmt = Math.round(Number(memo.totalNetRtgs) || totalGrossAmt);

      const bodyText = `It is submitted that this Range wishes to make payment to the payee(s) for the execution of departmental forestry works / liabilities for the month of ${memo.monthYear} as per the details tabulated below. You are kindly requested to sanction and release the total expenditure amount of Rs. ${totalGrossAmt.toLocaleString('en-IN')} (Total Net RTGS Amount: Rs. ${totalNetRtgsAmt.toLocaleString('en-IN')}) and arrange payment through RTGS / Treasury e-Transfer mode to the respective payees at the earliest.`;
      const splitBody = doc.splitTextToSize(bodyText, 273);
      doc.text(splitBody, 12, 55);

      const nextY = 55 + (splitBody.length * 4.2) + 2;

      // Table columns requested:
      // 1. Name & Address
      // 2. Try Code
      // 3. Bank Account Details
      // 4. Total Amount
      // 5. Sub-voucher amount with description
      // 6. Combined Deduction (I/Tax + GST)
      // 7. Net Amount
      // 8. PAN & GST Number
      const tableHead = [['Sr.', 'Name & Address', 'Try Code', 'Bank Account Details', 'Total (Rs)', 'Sub Voucher Details', 'Deductions\n(I.Tax + GST)', 'Net RTGS (Rs)', 'PAN & GSTIN']];
      const tableData = (memo.payeeEntries || []).map((e, idx) => {
        const subVouchersText = e.subVouchers && e.subVouchers.length > 0
          ? e.subVouchers.map((sv, svIdx) => `${svIdx + 1}. Rs. ${Math.round(Number(sv.amount) || 0).toLocaleString('en-IN')}${sv.voucherNo ? ' [' + sv.voucherNo + ']' : ''}${sv.description ? ' (' + sv.description + ')' : ''}`).join('\n')
          : `Single Bill: Rs. ${Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN')}`;

        const iTax = Math.round(Number(e.iTaxAmount) || 0);
        const gst = Math.round(Number(e.gstAmount) || 0);
        const totDed = iTax + gst;
        const dedText = totDed > 0
          ? `Rs. ${totDed.toLocaleString('en-IN')}${iTax > 0 ? `\n(IT: Rs. ${iTax.toLocaleString('en-IN')})` : ''}${gst > 0 ? `\n(GST: Rs. ${gst.toLocaleString('en-IN')})` : ''}`
          : 'Nil (Rs. 0)';

        // Auto lookup for missing Try Code
        const lookupPayee = payees.find(p => 
          (e.payeeId && p.id === e.payeeId) || 
          (p.accountNumber === e.accountNumber && p.name === e.name)
        );
        const displayTreasuryCode = e.treasuryCode || lookupPayee?.treasuryCode || '-';

        return [
          idx + 1,
          `${e.name || ''}${e.address ? '\n' + e.address : ''}`,
          displayTreasuryCode,
          `${e.accountNumber || ''}\n${e.ifscCode || ''}`,
          Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN'),
          subVouchersText,
          dedText,
          Math.round(Number(e.netRtgsAmount) || 0).toLocaleString('en-IN'),
          `PAN: ${e.panNumber || 'N/A'}${e.gstNumber ? '\nGST: ' + e.gstNumber : ''}`
        ];
      });

      const footRow = [
        '',
        'TOTAL: -',
        '',
        '',
        totalGrossAmt.toLocaleString('en-IN'),
        '',
        (totalITaxAmt + totalGstAmt).toLocaleString('en-IN'),
        totalNetRtgsAmt.toLocaleString('en-IN'),
        ''
      ];

      autoTable(doc, {
        startY: nextY,
        head: tableHead,
        body: tableData,
        foot: [footRow],
        theme: 'grid',
        showHead: 'everyPage',
        pageBreak: 'auto',
        styles: { fontSize: 7, cellPadding: 1.2, overflow: 'linebreak' },
        headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' },
        footStyles: { fillColor: [245, 245, 245], textColor: [0, 0, 0], fontStyle: 'bold' },
        columnStyles: {
          0: { halign: 'center', cellWidth: 8 },
          1: { cellWidth: 44 },
          2: { halign: 'center', cellWidth: 24 },
          3: { cellWidth: 34 },
          4: { halign: 'right', cellWidth: 26, fontStyle: 'bold' },
          5: { cellWidth: 46 },
          6: { halign: 'right', cellWidth: 28 },
          7: { halign: 'right', cellWidth: 30, fontStyle: 'bold' },
          8: { cellWidth: 30 }
        }
      });

      let finalY = (doc as any).lastAutoTable?.finalY || nextY + 40;
      if (finalY > 175) {
        doc.addPage();
        finalY = 20;
      } else {
        finalY += 5;
      }

      const words = convertNumberToWords(totalNetRtgsAmt);
      doc.setFontSize(8.5);
      doc.setFont("helvetica", "bold");
      doc.text(`Total Net Amount Payable (in words): Rupees ${words} Only`, 12, finalY);

      finalY += 15;
      if (finalY > 185) {
        doc.addPage();
        finalY = 25;
      }
      doc.text("Range Forest Officer", 245, finalY, { align: "center" });
      doc.setFont("helvetica", "normal");
      doc.text(memo.rangeName || rangeTitle, 245, finalY + 4, { align: "center" });
      doc.text("Rajgarh Forest Division", 245, finalY + 8, { align: "center" });

      doc.save(`Memo_For_Fund_${(memo.memoNo || 'Memo').replace(/[/\\?%*:|"<>]/g, '_')}_${memo.monthYear || ''}.pdf`);
    } catch (err) {
      console.error("Error generating Memo PDF:", err);
      showAlert("Failed to download Memo PDF.");
    }
  };

  const handlePrintMemo = (targetMemo?: MemoForFund) => {
    const memoToPrint = targetMemo || viewingMemo;
    if (!memoToPrint) return;
    markMemoAsViewedByAdmin(memoToPrint);
    try {
      const rangeTitle = memoToPrint.rangeName ? memoToPrint.rangeName.replace(/^RFO\s*/i, '').replace(/\s*Range$/i, '').replace(/\s*Office$/i, '') : (userRangeName || 'Sarahan');
      const dateFormatted = memoToPrint.date ? memoToPrint.date.split('-').reverse().join('.') : '';
      const totalGrossAmt = Math.round(Number(memoToPrint.totalAmount) || 0);
      const totalITaxAmt = Math.round(Number(memoToPrint.totalITax) || 0);
      const totalGstAmt = Math.round(Number(memoToPrint.totalGst) || 0);
      const totalNetRtgsAmt = Math.round(Number(memoToPrint.totalNetRtgs) || totalGrossAmt);
      const words = convertNumberToWords(totalNetRtgsAmt);

      const tableRowsHtml = (memoToPrint.payeeEntries || []).map((e, idx) => {
        const subVouchersHtml = e.subVouchers && e.subVouchers.length > 0
          ? e.subVouchers.map((sv, svIdx) => `<div style="font-size:10px; color:#047857; line-height:1.3;">${svIdx + 1}. ${sv.voucherNo ? '<strong>' + sv.voucherNo + '</strong>: ' : ''}â‚¹${Math.round(Number(sv.amount) || 0).toLocaleString('en-IN')}${sv.description ? ' <em>(' + sv.description + ')</em>' : ''}</div>`).join('')
          : `<div style="font-size:10px; color:#666;">Single Bill: â‚¹${Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN')}</div>`;

        const iTax = Math.round(Number(e.iTaxAmount) || 0);
        const gst = Math.round(Number(e.gstAmount) || 0);
        const totDed = iTax + gst;
        const dedHtml = totDed > 0
          ? `<strong>â‚¹${totDed.toLocaleString('en-IN')}</strong>${iTax > 0 ? `<div style="font-size:9.5px; color:#b91c1c;">(IT: â‚¹${iTax.toLocaleString('en-IN')})</div>` : ''}${gst > 0 ? `<div style="font-size:9.5px; color:#6b21a8;">(GST: â‚¹${gst.toLocaleString('en-IN')})</div>` : ''}`
          : `<span style="color:#666; font-size:10px;">Nil (â‚¹0)</span>`;

        // Automatically look up Try Code from directory if missing in existing memo
        const lookupPayee = payees.find(p => 
          (e.payeeId && p.id === e.payeeId) || 
          (p.accountNumber === e.accountNumber && p.name === e.name)
        );
        const displayTreasuryCode = e.treasuryCode || lookupPayee?.treasuryCode || '-';

        return `
          <tr>
            <td style="text-align:center; padding:5px 3px; border:1px solid #444; font-size:11px;">${idx + 1}</td>
            <td style="padding:5px 6px; border:1px solid #444; font-size:11px;">
              <strong>${e.name || ''}</strong>
              ${e.address ? `<div style="font-size:10px; color:#555; margin-top:2px;">${e.address}</div>` : ''}
            </td>
            <td style="padding:5px; border:1px solid #444; font-family:monospace; font-size:10.5px; font-weight:700; color:#065f46; text-align:center;">${displayTreasuryCode}</td>
            <td style="padding:5px; border:1px solid #444; font-family:monospace; font-size:10.5px;">
              <div style="font-weight:bold;">${e.accountNumber || ''}</div>
              <div style="font-size:9px; color:#666;">${e.ifscCode || ''}</div>
            </td>
            <td style="text-align:right; padding:5px; border:1px solid #444; font-weight:bold; font-size:11px;">â‚¹${Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN')}</td>
            <td style="padding:5px; border:1px solid #444; font-size:10.5px;">${subVouchersHtml}</td>
            <td style="text-align:right; padding:5px; border:1px solid #444; font-size:10.5px;">${dedHtml}</td>
            <td style="text-align:right; padding:5px; border:1px solid #444; font-weight:bold; color:#064e3b; font-size:11px;">â‚¹${Math.round(Number(e.netRtgsAmount) || 0).toLocaleString('en-IN')}</td>
            <td style="padding:5px; border:1px solid #444; font-size:10px;">PAN: ${e.panNumber || 'N/A'}${e.gstNumber ? `<br/>GST: ${e.gstNumber}` : ''}</td>
          </tr>
        `;
      }).join('');

      const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8" />
          <title>Memo No. ${memoToPrint.memoNo}</title>
          <style>
            @page {
              size: A4 landscape;
              margin: 10mm 12mm 12mm 12mm;
            }
            body {
              font-family: Arial, "Helvetica Neue", Helvetica, sans-serif;
              color: #111;
              margin: 0;
              padding: 0;
              font-size: 11.5px;
              line-height: 1.4;
              -webkit-print-color-adjust: exact;
              print-color-adjust: exact;
            }
            .header {
              text-align: center;
              border-bottom: 2px solid #222;
              padding-bottom: 6px;
              margin-bottom: 6px;
            }
            .header h2 {
              margin: 0;
              font-size: 11px;
              letter-spacing: 2px;
              text-transform: uppercase;
              color: #333;
            }
            .header h1 {
              margin: 3px 0 2px 0;
              font-size: 14px;
              font-weight: bold;
              text-transform: uppercase;
              color: #000;
            }
            .header p {
              margin: 0;
              font-size: 10.5px;
              color: #444;
            }
            .ref-date {
              display: flex;
              justify-content: space-between;
              border-bottom: 1px solid #666;
              padding-bottom: 4px;
              margin-bottom: 6px;
              font-size: 11px;
            }
            .from-to-row {
              display: flex;
              justify-content: space-between;
              border-bottom: 1px solid #eee;
              padding-bottom: 4px;
              margin-bottom: 6px;
              font-size: 11px;
            }
            .subject-box {
              background: #f8fafc;
              border: 1px solid #cbd5e1;
              padding: 6px 10px;
              margin-bottom: 8px;
              border-radius: 4px;
              font-weight: bold;
              font-size: 11px;
            }
            .letter-text {
              margin-bottom: 6px;
              text-align: justify;
              font-size: 11px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 6px;
              margin-bottom: 8px;
              page-break-inside: auto;
            }
            thead {
              display: table-header-group;
            }
            tfoot {
              display: table-footer-group;
            }
            tr {
              page-break-inside: avoid;
            }
            th {
              background-color: #f1f5f9;
              font-weight: bold;
              text-align: center;
              border: 1px solid #444;
              padding: 5px 3px;
              font-size: 10px;
            }
            tfoot td {
              font-weight: bold;
              background-color: #f8fafc;
              border: 1px solid #444;
              padding: 5px;
              font-size: 11px;
            }
            .words-box {
              background: #f8fafc;
              border: 1px solid #cbd5e1;
              padding: 6px 10px;
              margin-top: 6px;
              border-radius: 4px;
              font-weight: bold;
              font-size: 11px;
            }
            .signature {
              margin-top: 25px;
              float: right;
              width: 240px;
              text-align: center;
              page-break-inside: avoid;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <h2>H.P. FOREST DEPARTMENT</h2>
            <h1>OFFICE OF THE RANGE FOREST OFFICER, ${rangeTitle.toUpperCase()}</h1>
            <p>Forest Division Rajgarh, District Sirmaur (H.P.)</p>
          </div>

          <div class="ref-date">
            <div>No. <strong>${memoToPrint.memoNo}</strong></div>
            <div>Dated: <strong>${dateFormatted}</strong></div>
          </div>

          <div class="from-to-row">
            <div><strong>From:</strong> Range Forest Officer, ${memoToPrint.rangeName || rangeTitle}.</div>
            <div><strong>To:</strong> The Divisional Forest Officer, Rajgarh Forest Division (H.P.).</div>
          </div>

          <div class="subject-box">
            Subject: - <u>Memo for Fund for the month of ${memoToPrint.monthYear} under scheme ${memoToPrint.schemeName || 'All Schemes'}${memoToPrint.sectorName ? ` (${memoToPrint.sectorName})` : ''}${memoToPrint.soeName ? ` [SOE: ${memoToPrint.soeName}]` : ''}.</u>
          </div>

          <div class="letter-text">
            <p style="margin: 0 0 3px 0;"><strong>Sir,</strong></p>
            <p style="margin: 0 0 3px 0;">
              It is submitted that this Range wishes to make payment to the payee(s) for the execution of departmental forestry works / liabilities for the month of <strong>${memoToPrint.monthYear}</strong> as per the details tabulated below.
            </p>
            <p style="margin: 0;">
              You are kindly requested to sanction and release the total expenditure amount of <strong>â‚¹${totalGrossAmt.toLocaleString('en-IN')}</strong> (Total Net RTGS Amount: <strong>â‚¹${totalNetRtgsAmt.toLocaleString('en-IN')}</strong>) and arrange payment through RTGS / Treasury e-Transfer mode to the respective payees at the earliest.
            </p>
          </div>

          <table>
            <thead>
              <tr>
                <th style="width: 24px;">Sr.</th>
                <th>Name & Address</th>
                <th style="width: 70px;">Try Code</th>
                <th style="width: 120px;">Bank Account Details</th>
                <th style="width: 85px;">Total Amount</th>
                <th>Sub Voucher Details</th>
                <th style="width: 90px;">Deductions<br/>(I.Tax / GST)</th>
                <th style="width: 95px;">Net RTGS</th>
                <th style="width: 95px;">PAN & GSTIN</th>
              </tr>
            </thead>
            <tbody>
              ${tableRowsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="4" style="text-align: right; font-weight: bold;">TOTAL: -</td>
                <td style="text-align: right; font-weight: bold;">â‚¹${totalGrossAmt.toLocaleString('en-IN')}</td>
                <td style="text-align: center;">-</td>
                <td style="text-align: right; color:#b91c1c; font-weight: bold;">â‚¹${(totalITaxAmt + totalGstAmt).toLocaleString('en-IN')}</td>
                <td style="text-align: right; color:#064e3b; font-weight: bold;">â‚¹${totalNetRtgsAmt.toLocaleString('en-IN')}</td>
                <td></td>
              </tr>
            </tfoot>
          </table>

          <div class="words-box">
            Total Net Amount Payable (in words): <em>Rupees ${words} Only</em>
          </div>

          <div class="signature">
            <br/><br/>
            <p style="margin:2px 0; font-weight:bold;">Range Forest Officer</p>
            <p style="margin:2px 0;">${memoToPrint.rangeName || rangeTitle}</p>
            <p style="margin:2px 0; font-size:10px; color:#555;">Rajgarh Forest Division</p>
          </div>
        </body>
        </html>
      `;

      // Create a hidden print iframe
      const printFrame = document.createElement('iframe');
      printFrame.style.position = 'fixed';
      printFrame.style.right = '0';
      printFrame.style.bottom = '0';
      printFrame.style.width = '0';
      printFrame.style.height = '0';
      printFrame.style.border = '0';
      document.body.appendChild(printFrame);

      const frameDoc = printFrame.contentWindow?.document || printFrame.contentDocument;
      if (frameDoc) {
        frameDoc.open();
        frameDoc.write(htmlContent);
        frameDoc.close();

        setTimeout(() => {
          try {
            printFrame.contentWindow?.focus();
            printFrame.contentWindow?.print();
          } catch (err) {
            console.warn("Iframe print fallback to window.print", err);
            window.print();
          } finally {
            setTimeout(() => {
              if (document.body.contains(printFrame)) {
                document.body.removeChild(printFrame);
              }
            }, 1500);
          }
        }, 300);
      } else {
        window.print();
      }
    } catch (e) {
      console.error("Print error:", e);
      window.print();
    }
  };

  const downloadMemoWord = (targetMemo?: MemoForFund) => {
    const memoToExport = targetMemo || viewingMemo;
    if (!memoToExport) return;
    if (userRole !== 'admin' && userRole !== 'deo' && !isAdmin() && !isDEO()) {
      showAlert("Only DEO and Admin users are permitted to download editable Memo DOC files.");
      return;
    }

    try {
      const rangeTitle = memoToExport.rangeName
        ? memoToExport.rangeName.replace(/^RFO\s*/i, '').replace(/\s*Range$/i, '').replace(/\s*Office$/i, '')
        : (userRangeName || 'Sarahan');
      const dateFormatted = memoToExport.date ? memoToExport.date.split('-').reverse().join('.') : '';
      const totalGrossAmt = Math.round(Number(memoToExport.totalAmount) || 0);
      const totalITaxAmt = Math.round(Number(memoToExport.totalITax) || 0);
      const totalGstAmt = Math.round(Number(memoToExport.totalGst) || 0);
      const totalNetRtgsAmt = Math.round(Number(memoToExport.totalNetRtgs) || totalGrossAmt);
      const words = convertNumberToWords(totalNetRtgsAmt);

      const schemeText = memoToExport.schemeName || 'All Schemes';
      const sectorText = memoToExport.sectorName ? ` (${memoToExport.sectorName})` : '';
      const soeText = memoToExport.soeName ? ` [SOE: ${memoToExport.soeName}]` : '';

      const tableRowsHtml = (memoToExport.payeeEntries || []).map((e, idx) => {
        const subVouchersHtml = e.subVouchers && e.subVouchers.length > 0
          ? e.subVouchers.map((sv, svIdx) => `<div>${svIdx + 1}. ${sv.voucherNo ? '<b>' + sv.voucherNo + '</b>: ' : ''}Rs. ${Math.round(Number(sv.amount) || 0).toLocaleString('en-IN')}${sv.description ? ' <i>(' + sv.description + ')</i>' : ''}</div>`).join('')
          : `<div>Single Bill: Rs. ${Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN')}</div>`;

        const iTax = Math.round(Number(e.iTaxAmount) || 0);
        const gst = Math.round(Number(e.gstAmount) || 0);
        const totDed = iTax + gst;
        const dedHtml = totDed > 0
          ? `<b>Rs. ${totDed.toLocaleString('en-IN')}</b>${iTax > 0 ? `<br/><small>IT: Rs. ${iTax.toLocaleString('en-IN')}</small>` : ''}${gst > 0 ? `<br/><small>GST: Rs. ${gst.toLocaleString('en-IN')}</small>` : ''}`
          : 'Nil (Rs. 0)';

        const lookupPayee = payees.find(p => 
          (e.payeeId && p.id === e.payeeId) || 
          (p.accountNumber === e.accountNumber && p.name === e.name)
        );
        const displayTreasuryCode = e.treasuryCode || lookupPayee?.treasuryCode || '-';

        return `
          <tr>
            <td style="text-align: center; vertical-align: top;">${idx + 1}</td>
            <td style="vertical-align: top;"><b>${e.name || ''}</b>${e.address ? '<br/><small>' + e.address + '</small>' : ''}</td>
            <td style="text-align: center; font-family: monospace; font-weight: bold; vertical-align: top;">${displayTreasuryCode}</td>
            <td style="font-family: monospace; vertical-align: top;"><b>${e.accountNumber || '-'}</b><br/><small>${e.ifscCode || '-'}</small></td>
            <td style="text-align: right; font-weight: bold; vertical-align: top;">Rs. ${Math.round(Number(e.totalAmount) || 0).toLocaleString('en-IN')}</td>
            <td style="vertical-align: top; font-size: 9pt;">${subVouchersHtml}</td>
            <td style="text-align: right; vertical-align: top;">${dedHtml}</td>
            <td style="text-align: right; font-weight: bold; color: #004d40; vertical-align: top;">Rs. ${Math.round(Number(e.netRtgsAmount) || 0).toLocaleString('en-IN')}</td>
            <td style="font-family: monospace; font-size: 8.5pt; vertical-align: top;">PAN: ${e.panNumber || 'N/A'}${e.gstNumber ? '<br/>GST: ' + e.gstNumber : ''}</td>
          </tr>
        `;
      }).join('');

      const html = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset="utf-8">
          <title>Memo for Fund - ${memoToExport.memoNo || 'Memo'}</title>
          <style>
            @page {
              size: A4 landscape;
              margin: 1.5cm;
            }
            body { 
              font-family: 'Calibri', 'Arial', sans-serif; 
              font-size: 10.5pt; 
              color: #111827; 
              line-height: 1.4;
            }
            h1, h2, h3, p { margin: 0 0 6px 0; }
            .header-box { text-align: center; border-bottom: 2px solid #111; padding-bottom: 8px; margin-bottom: 12px; }
            .header-dept { font-size: 12pt; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; color: #374151; }
            .header-office { font-size: 14pt; font-weight: 900; text-transform: uppercase; }
            .header-sub { font-size: 10pt; font-weight: bold; }
            .from-to-bar { width: 100%; margin-bottom: 10px; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }
            .subject-box { background-color: #f3f4f6; border: 1px solid #d1d5db; padding: 8px 10px; margin-bottom: 12px; font-weight: bold; }
            .body-text { text-align: justify; margin-bottom: 12px; font-size: 10pt; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 14px; }
            th, td { border: 1px solid #000; padding: 5px 6px; font-size: 9.5pt; }
            th { background-color: #f3f4f6; color: #000; font-weight: bold; text-align: center; }
            tfoot tr td { background-color: #f9fafb; font-weight: bold; }
            .words-box { margin-top: 10px; font-weight: bold; font-size: 10pt; }
          </style>
        </head>
        <body>
          <div class="header-box">
            <div class="header-dept">H.P. FOREST DEPARTMENT</div>
            <div class="header-office">OFFICE OF THE RANGE FOREST OFFICER, ${rangeTitle.toUpperCase()}</div>
            <div class="header-sub">Forest Division Rajgarh, District Sirmaur (H.P.)</div>
          </div>

          <table style="border:none; width:100%; margin-bottom: 6px;">
            <tr style="border:none;">
              <td style="border:none; text-align:left; font-weight:bold; font-size: 10pt; padding:0;">No. <span style="font-family: monospace;">${memoToExport.memoNo}</span></td>
              <td style="border:none; text-align:right; font-weight:bold; font-size: 10pt; padding:0;">Dated: ${dateFormatted}</td>
            </tr>
          </table>

          <table style="border:none; width:100%; margin-bottom: 10px; border-bottom: 1px solid #ccc;">
            <tr style="border:none;">
              <td style="border:none; text-align:left; padding: 4px 0;"><b>From:</b> Range Forest Officer, ${memoToExport.rangeName || rangeTitle}.</td>
              <td style="border:none; text-align:right; padding: 4px 0;"><b>To:</b> The Divisional Forest Officer, Rajgarh Forest Division (H.P.).</td>
            </tr>
          </table>

          <div class="subject-box">
            <b>Subject: - </b><u>Memo for Fund for the month of ${memoToExport.monthYear} under scheme ${schemeText}${sectorText}${soeText}.</u>
          </div>

          <div class="body-text">
            <p><b>Sir,</b></p>
            <p>It is submitted that this Range wishes to make payment to the payee(s) for the execution of departmental forestry works / liabilities for the month of <b>${memoToExport.monthYear}</b> as per the details tabulated below.</p>
            <p>You are kindly requested to sanction and release the total expenditure amount of <b>Rs. ${totalGrossAmt.toLocaleString('en-IN')}</b> (Total Net RTGS Amount: <b>Rs. ${totalNetRtgsAmt.toLocaleString('en-IN')}</b>) and arrange payment through RTGS / Treasury e-Transfer mode to the respective payees at the earliest.</p>
          </div>

          <table>
            <thead>
              <tr>
                <th style="width: 4%;">#</th>
                <th style="width: 18%;">Name & Address</th>
                <th style="width: 9%;">Try Code</th>
                <th style="width: 14%;">Bank Account Details</th>
                <th style="width: 9%;">Total Amt (Rs.)</th>
                <th style="width: 18%;">Sub Voucher Details</th>
                <th style="width: 10%;">Deductions (IT+GST)</th>
                <th style="width: 10%;">Net RTGS (Rs.)</th>
                <th style="width: 8%;">PAN & GSTIN</th>
              </tr>
            </thead>
            <tbody>
              ${tableRowsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="4" style="text-align: right; text-transform: uppercase;"><b>TOTAL: -</b></td>
                <td style="text-align: right;"><b>Rs. ${totalGrossAmt.toLocaleString('en-IN')}</b></td>
                <td></td>
                <td style="text-align: right;"><b>Rs. ${(totalITaxAmt + totalGstAmt).toLocaleString('en-IN')}</b></td>
                <td style="text-align: right; color: #004d40;"><b>Rs. ${totalNetRtgsAmt.toLocaleString('en-IN')}</b></td>
                <td></td>
              </tr>
            </tfoot>
          </table>

          <div class="words-box">
            Total Net Amount Payable (in words): <b>Rupees ${words} Only</b>
          </div>

          <div style="width: 100%; margin-top: 35px;">
            <table style="border:none; width: 100%;">
              <tr style="border:none;">
                <td style="border:none; width: 50%;"></td>
                <td style="border:none; width: 50%; text-align: center;">
                  <b>Range Forest Officer</b><br/>
                  ${memoToExport.rangeName || rangeTitle}<br/>
                  Rajgarh Forest Division
                </td>
              </tr>
            </table>
          </div>
        </body>
        </html>
      `;

      const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const cleanFileName = `Memo_For_Fund_${(memoToExport.memoNo || 'Memo').replace(/[/\\?%*:|"<>]/g, '_')}_${memoToExport.monthYear || ''}.doc`;
      a.download = cleanFileName;
      a.click();
      URL.revokeObjectURL(url);
      showAlert(`Downloaded Memo as editable Word Document (.doc) successfully.`);
    } catch (err) {
      console.error("Word export error:", err);
      showAlert("Failed to export Memo Word document.");
    }
  };

  const generateExpenditurePDFDoc = () => {
    const doc = new jsPDF('landscape');
    const fyName = fys.find(f => f.id === selectedFY)?.name || selectedFY;
    const rangeNameDisplay = ranges.find(r => r.id === userRangeId)?.name || userRangeName || 'Rajgarh Forest Division';

    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text("Department of Forests, Himachal Pradesh", 14, 12);
    
    doc.setFontSize(11);
    doc.text(`Rajgarh Forest Division â€” Expenditure Report List (FY ${fyName})`, 14, 18);
    
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.text(`Range/Unit: ${rangeNameDisplay} | Generated On: ${new Date().toLocaleDateString('en-GB')} | Total Entries: ${currentExpenses.length}`, 14, 24);

    const tableData = currentExpenses.map((exp, idx) => {
      const p = payees.find(p => p.id === exp.payeeId);
      const payeeName = p?.name || exp.payeeName || 'N/A';
      const payeeAcc = p?.accountNumber ? `A/C: ${p.accountNumber}` : '';
      const payeeInfo = payeeAcc ? `${payeeName}\n${payeeAcc}` : payeeName;
      
      const al = allocations.find(a => a.id === exp.allocationId);
      const r = ranges.find(r => r.id === al?.rangeId);
      const s = soes.find(s => s.id === exp.soeId);
      const locationInfo = `${r?.name || 'N/A'} / ${s?.name || 'N/A'}`;

      return [
        (idx + 1).toString(),
        exp.date ? exp.date.split('-').reverse().join('/') : '',
        payeeInfo,
        locationInfo,
        exp.description || '-',
        `Rs. ${(Number(exp.amount) || 0).toLocaleString('en-IN')}`,
        exp.status || 'pending'
      ];
    });

    const totalAmt = currentExpenses.reduce((sum, e) => sum + (Number(e.amount) || 0), 0);

    autoTable(doc, {
      startY: 28,
      head: [['#', 'Date', 'Payee Name & Account', 'Unit / Range / SOE', 'Description / Particulars', 'Amount (Rs.)', 'Status']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: [5, 150, 105], textColor: 255, fontStyle: 'bold' },
      styles: { fontSize: 8, cellPadding: 2 },
      columnStyles: {
        0: { cellWidth: 10, halign: 'center' },
        1: { cellWidth: 22 },
        2: { cellWidth: 50 },
        3: { cellWidth: 40 },
        4: { cellWidth: 80 },
        5: { cellWidth: 35, halign: 'right', fontStyle: 'bold' },
        6: { cellWidth: 25, halign: 'center' }
      },
      foot: [[
        { content: 'Total Expenditure Amount:', colSpan: 5, styles: { halign: 'right', fontStyle: 'bold' } },
        { content: `Rs. ${totalAmt.toLocaleString('en-IN')}`, styles: { halign: 'right', fontStyle: 'bold', textColor: [5, 150, 105] } },
        ''
      ]]
    });

    const finalY = (doc as any).lastAutoTable?.finalY || 150;
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text("Prepared By: Dealing Assistant / Data Entry Operator", 14, finalY + 15);
    doc.text("Verified & Authorised By: Range Forest Officer / DFO", 200, finalY + 15);

    return { doc, fyName };
  };

  const downloadExpenditureListPDF = () => {
    try {
      const { doc, fyName } = generateExpenditurePDFDoc();
      doc.save(`Expenditure_List_FY_${fyName}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error("PDF Export Error:", err);
      showAlert("Failed to export Expenditure PDF.");
    }
  };

  const printHtmlViaIframe = (htmlContent: string) => {
    try {
      const printFrame = document.createElement('iframe');
      printFrame.style.position = 'fixed';
      printFrame.style.right = '0';
      printFrame.style.bottom = '0';
      printFrame.style.width = '0';
      printFrame.style.height = '0';
      printFrame.style.border = '0';
      document.body.appendChild(printFrame);

      const frameDoc = printFrame.contentWindow?.document || printFrame.contentDocument;
      if (frameDoc) {
        frameDoc.open();
        frameDoc.write(htmlContent);
        frameDoc.close();

        setTimeout(() => {
          try {
            printFrame.contentWindow?.focus();
            printFrame.contentWindow?.print();
          } catch (err) {
            console.warn("Iframe print fallback to window.print", err);
            window.print();
          } finally {
            setTimeout(() => {
              if (document.body.contains(printFrame)) {
                document.body.removeChild(printFrame);
              }
            }, 1500);
          }
        }, 300);
      } else {
        window.print();
      }
    } catch (e) {
      console.error("Print error:", e);
      window.print();
    }
  };

  const handlePrintExpenditureReport = () => {
    try {
      const fyName = fys.find(f => f.id === selectedFY)?.name || selectedFY;
      const rangeNameDisplay = ranges.find(r => r.id === userRangeId)?.name || userRangeName || 'Rajgarh Forest Division';
      const totalGrossAmt = currentExpenses.reduce((sum, e) => sum + (Number(e.amount) || 0), 0);
      const totalTdsAmt = currentExpenses.reduce((sum, e) => sum + Math.round(e.tdsAmount || 0), 0);
      const totalGstTdsAmt = currentExpenses.reduce((sum, e) => sum + Math.round(e.tdsGstAmount || 0), 0);
      const totalDeductedAmt = totalTdsAmt + totalGstTdsAmt;
      const totalNetPayableAmt = currentExpenses.reduce((sum, e) => {
        const gross = Number(e.amount) || 0;
        const ded = Math.round(e.deductedAmount || ((e.tdsAmount || 0) + (e.tdsGstAmount || 0)));
        return sum + Math.round(e.netAmount ?? (gross - ded));
      }, 0);

      const rowsHtml = currentExpenses.map((exp, idx) => {
        const p = payees.find(p => p.id === exp.payeeId);
        const payeeName = p?.name || exp.payeeName || 'N/A';
        const payeeAcc = p?.accountNumber ? `A/C: ${p.accountNumber}` : '';
        const payeeTryCode = p?.treasuryCode ? `<br/><span style="font-size:9px; color:#065f46; font-weight:bold; font-family:monospace;">TRY: ${p.treasuryCode}</span>` : '';
        const al = allocations.find(a => a.id === exp.allocationId);
        const r = ranges.find(r => r.id === al?.rangeId);
        const s = soes.find(s => s.id === exp.soeId);

        const gross = Number(exp.amount) || 0;
        const tds = Math.round(exp.tdsAmount || 0);
        const gstTds = Math.round(exp.tdsGstAmount || 0);
        const totalDed = Math.round(exp.deductedAmount || (tds + gstTds));
        const netPay = Math.round(exp.netAmount ?? (gross - totalDed));

        let deductionDisplay = '-';
        if (totalDed > 0) {
          const parts = [];
          if (tds > 0) parts.push(`TDS: â‚¹${tds.toLocaleString('en-IN')}`);
          if (gstTds > 0) parts.push(`GST: â‚¹${gstTds.toLocaleString('en-IN')}`);
          deductionDisplay = parts.join('<br/>') || `â‚¹${totalDed.toLocaleString('en-IN')}`;
        }

        return `
          <tr>
            <td style="text-align: center;">${idx + 1}</td>
            <td style="white-space: nowrap;">${exp.date ? exp.date.split('-').reverse().join('/') : '-'}</td>
            <td>
              <strong>${payeeName}</strong>
              ${payeeAcc ? `<br/><span style="font-size:9px; color:#555; font-family:monospace;">${payeeAcc}</span>` : ''}
              ${payeeTryCode}
            </td>
            <td><strong>${r?.name || 'N/A'}</strong><br/><span style="font-size:9px; color:#555;">SOE: ${s?.name || 'N/A'}</span></td>
            <td style="font-style: italic;">${exp.description || '-'}</td>
            <td style="text-align: right; font-weight: bold;">â‚¹${gross.toLocaleString('en-IN')}</td>
            <td style="text-align: right; color: #b91c1c; font-size: 9.5px;">${deductionDisplay}</td>
            <td style="text-align: right; font-weight: bold; color: #064e3b;">â‚¹${netPay.toLocaleString('en-IN')}</td>
            <td style="text-align: center; text-transform: uppercase; font-size: 9px; font-weight: bold;">${exp.status || 'pending'}</td>
          </tr>
        `;
      }).join('');

      const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
          <title>Expenditure Report List - Rajgarh Forest Division</title>
          <meta charset="utf-8" />
          <style>
            @page {
              size: A4 landscape;
              margin: 10mm 12mm 12mm 12mm;
            }
            @media print {
              body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
              thead { display: table-header-group; }
              tfoot { display: table-footer-group; }
              tr { page-break-inside: avoid; }
            }
            body {
              font-family: Arial, sans-serif;
              font-size: 10px;
              color: #111;
              line-height: 1.35;
              padding: 0;
              margin: 0;
            }
            .header-block {
              text-align: center;
              margin-bottom: 12px;
              border-bottom: 2px solid #047857;
              padding-bottom: 8px;
            }
            .header-block h1 {
              font-size: 15px;
              margin: 0 0 3px 0;
              text-transform: uppercase;
              letter-spacing: 0.5px;
              color: #064e3b;
            }
            .header-block h2 {
              font-size: 12px;
              margin: 0 0 3px 0;
              color: #1f2937;
            }
            .meta-bar {
              display: flex;
              justify-content: space-between;
              font-size: 9.5px;
              color: #374151;
              margin-top: 4px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 8px;
            }
            th, td {
              border: 1px solid #9ca3af;
              padding: 5px 6px;
              vertical-align: middle;
            }
            th {
              background-color: #f3f4f6;
              color: #111827;
              font-weight: bold;
              text-align: left;
              font-size: 9.5px;
            }
            tfoot tr td {
              background-color: #f9fafb;
              font-weight: bold;
              border-top: 2px solid #374151;
            }
            .signature-block {
              margin-top: 35px;
              display: flex;
              justify-content: space-between;
              font-size: 10px;
              page-break-inside: avoid;
            }
          </style>
        </head>
        <body>
          <div class="header-block">
            <h1>Department of Forests, Himachal Pradesh</h1>
            <h2>Rajgarh Forest Division â€” Expenditure Report List</h2>
            <div class="meta-bar">
              <span><strong>Financial Year:</strong> ${fyName}</span>
              <span><strong>Range / Unit:</strong> ${rangeNameDisplay}</span>
              <span><strong>Total Entries:</strong> ${currentExpenses.length}</span>
              <span><strong>Date Generated:</strong> ${new Date().toLocaleDateString('en-GB')}</span>
            </div>
          </div>

          <table>
            <thead>
              <tr>
                <th style="width: 24px; text-align: center;">#</th>
                <th style="width: 70px;">Date</th>
                <th>Payee Name & Account</th>
                <th style="width: 130px;">Unit / SOE</th>
                <th>Description / Particulars</th>
                <th style="width: 80px; text-align: right;">Gross (â‚¹)</th>
                <th style="width: 85px; text-align: right;">Deductions</th>
                <th style="width: 85px; text-align: right;">Net RTGS (â‚¹)</th>
                <th style="width: 60px; text-align: center;">Status</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="5" style="text-align: right; text-transform: uppercase;">Total Expenditure Summary:</td>
                <td style="text-align: right; color: #111827;">â‚¹${totalGrossAmt.toLocaleString('en-IN')}</td>
                <td style="text-align: right; color: #b91c1c;">â‚¹${totalDeductedAmt.toLocaleString('en-IN')}</td>
                <td style="text-align: right; color: #064e3b; font-size: 11px;">â‚¹${totalNetPayableAmt.toLocaleString('en-IN')}</td>
                <td></td>
              </tr>
            </tfoot>
          </table>

          <div class="signature-block">
            <div>
              <p style="margin: 0; font-weight: bold;">Prepared By:</p>
              <p style="margin: 30px 0 0 0; color: #4b5563;">Dealing Assistant / Data Entry Operator</p>
            </div>
            <div style="text-align: right;">
              <p style="margin: 0; font-weight: bold;">Verified & Authorised By:</p>
              <p style="margin: 30px 0 0 0; color: #4b5563;">Range Forest Officer / Divisional Forest Officer</p>
            </div>
          </div>
        </body>
        </html>
      `;

      printHtmlViaIframe(htmlContent);
    } catch (e) {
      console.error("Print error:", e);
      window.print();
    }
  };

  const handlePrintPayees = () => {
    try {
      const rangeNameDisplay = userRangeName || (userRole === 'admin' ? 'Division Headquarter' : userRole) || 'All Ranges';
      const rowsHtml = filteredPayeesList.map((p, index) => {
        const rName = ranges.find(r => r.id === p.rangeId)?.name || 'Not Specified';
        return `
          <tr>
            <td style="text-align: center;">${index + 1}</td>
            <td><strong>${p.name || ''}</strong></td>
            <td>${p.address || ''}</td>
            <td style="font-family: monospace; font-weight: bold;">${p.accountNumber || ''}</td>
            <td style="font-family: monospace; color: #065f46; font-weight: bold;">${p.treasuryCode || '<span style="color:#9ca3af; font-weight:normal;">-</span>'}</td>
            <td style="font-family: monospace;">${p.ifscCode || '<span style="color:#9ca3af;">-</span>'}</td>
            <td style="font-family: monospace;">${p.panNumber || '<span style="color:#9ca3af;">-</span>'}</td>
            <td style="font-family: monospace;">${p.gstNumber || '<span style="color:#9ca3af;">-</span>'}</td>
            <td>${rName}</td>
          </tr>
        `;
      }).join('');

      const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
          <title>Payee Directory - Rajgarh Forest Division</title>
          <meta charset="utf-8" />
          <style>
            @page {
              size: A4 landscape;
              margin: 10mm 12mm 12mm 12mm;
            }
            @media print {
              body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
              thead { display: table-header-group; }
              tfoot { display: table-footer-group; }
              tr { page-break-inside: avoid; }
            }
            body {
              font-family: Arial, sans-serif;
              font-size: 10px;
              color: #111;
              line-height: 1.35;
              padding: 0;
              margin: 0;
            }
            .header-block {
              text-align: center;
              margin-bottom: 12px;
              border-bottom: 2px solid #047857;
              padding-bottom: 8px;
            }
            .header-block h1 {
              font-size: 15px;
              margin: 0 0 3px 0;
              text-transform: uppercase;
              letter-spacing: 0.5px;
              color: #064e3b;
            }
            .header-block h2 {
              font-size: 12px;
              margin: 0 0 3px 0;
              color: #1f2937;
            }
            .meta-bar {
              display: flex;
              justify-content: space-between;
              font-size: 9.5px;
              color: #374151;
              margin-top: 4px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 8px;
            }
            th, td {
              border: 1px solid #9ca3af;
              padding: 5px 6px;
              vertical-align: middle;
            }
            th {
              background-color: #f3f4f6;
              color: #111827;
              font-weight: bold;
              text-align: left;
              font-size: 9.5px;
            }
            tfoot tr td {
              background-color: #f9fafb;
              font-weight: bold;
              border-top: 2px solid #374151;
            }
            .signature-block {
              margin-top: 35px;
              display: flex;
              justify-content: space-between;
              font-size: 10px;
              page-break-inside: avoid;
            }
          </style>
        </head>
        <body>
          <div class="header-block">
            <h1>Department of Forests, Himachal Pradesh</h1>
            <h2>Rajgarh Forest Division â€” Payee Directory Details</h2>
            <div class="meta-bar">
              <span><strong>Scope:</strong> ${rangeNameDisplay}</span>
              <span><strong>Total Registered Payees:</strong> ${filteredPayeesList.length}</span>
              <span><strong>Date:</strong> ${new Date().toLocaleDateString('en-GB')}</span>
            </div>
          </div>

          <table>
            <thead>
              <tr>
                <th style="width: 24px; text-align: center;">#</th>
                <th style="width: 130px;">Payee Name</th>
                <th>Address</th>
                <th style="width: 110px;">Account Number</th>
                <th style="width: 100px;">Treasury Code</th>
                <th style="width: 85px;">IFSC Code</th>
                <th style="width: 85px;">PAN Number</th>
                <th style="width: 95px;">GST Number</th>
                <th style="width: 90px;">Range</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="9" style="text-align: right; font-size: 9px; color: #4b5563;">
                  Total Records: ${filteredPayeesList.length} Payee(s)
                </td>
              </tr>
            </tfoot>
          </table>

          <div class="signature-block">
            <div>
              <p style="margin: 0; font-weight: bold;">Verified By:</p>
              <p style="margin: 30px 0 0 0; color: #4b5563;">Data Entry Operator / Dealing Assistant</p>
            </div>
            <div style="text-align: right;">
              <p style="margin: 0; font-weight: bold;">Divisional Forest Officer / Range Officer</p>
              <p style="margin: 30px 0 0 0; color: #4b5563;">Rajgarh Forest Division, H.P.</p>
            </div>
          </div>
        </body>
        </html>
      `;

      printHtmlViaIframe(htmlContent);
    } catch (e) {
      console.error("Print error:", e);
      window.print();
    }
  };

  const handleSaveTreasuryCode = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!treasuryCodeModalPayee) return;
    if (!isAdmin() && !isDEO()) {
      showAlert("Only Admin or DEO can add or update the Treasury Code.");
      return;
    }
    try {
      const code = treasuryCodeInput.trim();
      await updateDoc(doc(db, 'payees', treasuryCodeModalPayee.id), {
        treasuryCode: code || null,
        updatedAt: Date.now()
      });
      logAuditAction('Payee Treasury Code Updated', `Payee: ${treasuryCodeModalPayee.name}, Treasury Code: ${code || 'Cleared'}`);
      showAlert(`Treasury Code for "${treasuryCodeModalPayee.name}" saved successfully.`);
      setTreasuryCodeModalPayee(null);
      setTreasuryCodeInput('');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `payees/${treasuryCodeModalPayee.id}`);
    }
  };

  const downloadPayeesWord = () => {
    try {
      const payeeData = filteredPayeesList;
      let html = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><title>Payee List</title>
        <style>
          body { font-family: Arial, sans-serif; }
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid #000; padding: 6px 8px; text-align: left; font-size: 10pt; }
          th { background-color: #004d40; color: #ffffff; }
          h2 { color: #004d40; }
        </style>
        </head>
        <body>
          <h2>Forest Budget Control System - Payee Details</h2>
          <p><b>Role / Range:</b> ${userRangeName || userRole || 'All'} | <b>Date:</b> ${new Date().toLocaleDateString('en-IN')}</p>
          <table>
            <thead>
              <tr>
                <th>S.No</th>
                <th>Payee Name</th>
                <th>Address</th>
                <th>Try Code</th>
                <th>Bank Account Details</th>
                <th>PAN Number</th>
                <th>GST Number</th>
                <th>Range</th>
              </tr>
            </thead>
            <tbody>
              ${payeeData.map((p, index) => `
                <tr>
                  <td>${index + 1}</td>
                  <td>${p.name || ''}</td>
                  <td>${p.address || ''}</td>
                  <td>${p.treasuryCode || 'N/A'}</td>
                  <td style="font-family: monospace;"><b>${p.accountNumber || ''}</b><br/><small>${p.ifscCode || ''}</small></td>
                  <td>${p.panNumber || 'N/A'}</td>
                  <td>${p.gstNumber || 'N/A'}</td>
                  <td>${ranges.find(r => r.id === p.rangeId)?.name || 'N/A'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </body>
        </html>
      `;
      const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Payees_List_${userRole || 'All'}_${new Date().toISOString().split('T')[0]}.doc`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Word export error:", err);
      showAlert("Failed to export Word document.");
    }
  };

  // --- Derived default From authority and auto-generated Memo Number ---
  const defaultFromAuthority = useMemo(() => {
    if (userRole && ['Sarahan', 'Narag', 'Habban', 'Rajgarh', 'Division'].some(r => r.toLowerCase() === userRole.toLowerCase())) {
      const match = ['Sarahan', 'Narag', 'Habban', 'Rajgarh', 'Division'].find(r => r.toLowerCase() === userRole.toLowerCase());
      return `RFO ${match || userRole}`;
    }
    if (userRangeName) {
      return `RFO ${userRangeName.replace(/^RFO\s*/i, '').replace(/\s*Range$/i, '').replace(/\s*Office$/i, '')}`;
    }
    return 'RFO Sarahan';
  }, [userRole, userRangeName]);

  // Helper to extract range key and 3-letter range code from string
  const getRangeCodeInfo = (fromAuthorityStr: string) => {
    const fromStr = (fromAuthorityStr || 'Sarahan').toLowerCase();
    if (fromStr.includes('narag') || fromStr.includes('nrg')) {
      return { rangeKey: 'Narag', code: 'NRG' };
    } else if (fromStr.includes('habban') || fromStr.includes('hbn')) {
      return { rangeKey: 'Habban', code: 'HBN' };
    } else if (fromStr.includes('rajgarh') || fromStr.includes('rjg')) {
      return { rangeKey: 'Rajgarh', code: 'RJG' };
    } else if (fromStr.includes('division') || fromStr.includes('div')) {
      return { rangeKey: 'Division', code: 'DIV' };
    } else {
      return { rangeKey: 'Sarahan', code: 'SRH' };
    }
  };

  // Helper to compute fresh next sequential memo number starting from 100
  const getNextSequentialMemoNo = useCallback((fromAuthorityStr: string, fy: string, allMemos: MemoForFund[]) => {
    const { rangeKey, code } = getRangeCodeInfo(fromAuthorityStr);
    const currentFY = fy || '2026-27';

    const fyRangeMemos = allMemos.filter(m => {
      const memoFY = m.financialYear || m.fyId;
      if (memoFY && memoFY !== currentFY) return false;

      if (m.createdByRole && m.createdByRole.toLowerCase() === rangeKey.toLowerCase()) return true;
      if (m.rangeName && m.rangeName.toLowerCase().includes(rangeKey.toLowerCase())) return true;
      if (m.memoNo && m.memoNo.toUpperCase().includes(`/${code}/`)) return true;
      return false;
    });

    const existingNums = fyRangeMemos
      .map(m => {
        if (!m.memoNo) return null;
        const match = m.memoNo.match(/Memo\/(\d+)/i) || m.memoNo.match(/(\d+)$/);
        return match ? parseInt(match[1], 10) : null;
      })
      .filter((n): n is number => n !== null && !isNaN(n));

    // Sequence counting begins at 100 for each Financial Year
    const maxNum = existingNums.length > 0 ? Math.max(...existingNums) : 99;
    const nextNum = Math.max(100, maxNum + 1);

    return `RFO/${code}/Memo/${nextNum}`;
  }, []);

  const autoMemoNo = useMemo(() => {
    if (editingMemo) return editingMemo.memoNo;
    const fromStr = memoFromInput || defaultFromAuthority || userRole || userRangeName || 'Sarahan';
    return getNextSequentialMemoNo(fromStr, selectedFY || '2026-27', memos);
  }, [memos, editingMemo, userRole, userRangeName, memoFromInput, defaultFromAuthority, selectedFY, getNextSequentialMemoNo]);

  useEffect(() => {
    if (!editingMemo) {
      setMemoNoInput(autoMemoNo);
    }
  }, [autoMemoNo, editingMemo]);

  useEffect(() => {
    if (!editingMemo) {
      setMemoFromInput(defaultFromAuthority);
      setMemoToInput('DCF Rajgarh');
    }
  }, [defaultFromAuthority, editingMemo]);

  // Automatic self-healing effect: Detect and re-sequence duplicate or collided memo numbers in existing Firestore database
  const isRepairingMemosRef = useRef(false);
  const resequenceDuplicateMemos = useCallback(async (manualAlert = false) => {
    if (isRepairingMemosRef.current || memos.length === 0) return;

    // Group memos by (FY + RangeCode)
    const groups: Record<string, MemoForFund[]> = {};
    memos.forEach(m => {
      const fy = m.financialYear || m.fyId || '2026-27';
      const { code } = getRangeCodeInfo(m.rangeName || m.createdByRole || m.memoNo || '');
      const key = `${fy}__${code}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(m);
    });

    let repairedCount = 0;

    for (const [key, memoList] of Object.entries(groups)) {
      const [fy, code] = key.split('__');

      // Sort chronologically (earliest created or dated first)
      const sorted = [...memoList].sort((a, b) => {
        const timeA = a.createdAt || (a.date ? new Date(a.date).getTime() : 0);
        const timeB = b.createdAt || (b.date ? new Date(b.date).getTime() : 0);
        if (timeA !== timeB) return timeA - timeB;
        return a.id.localeCompare(b.id);
      });

      // Check if duplicate memo numbers exist
      const seenNos = new Set<string>();
      let hasDuplicates = false;
      for (const m of sorted) {
        if (seenNos.has(m.memoNo)) {
          hasDuplicates = true;
          break;
        }
        seenNos.add(m.memoNo);
      }

      // If duplicate found or manual repair triggered, assign strict sequential numbers (100, 101, 102...)
      if (hasDuplicates || manualAlert) {
        isRepairingMemosRef.current = true;
        for (let i = 0; i < sorted.length; i++) {
          const m = sorted[i];
          const correctMemoNo = `RFO/${code}/Memo/${100 + i}`;
          if (m.memoNo !== correctMemoNo) {
            try {
              await updateDoc(doc(db, 'memos', m.id), {
                memoNo: correctMemoNo,
                updatedAt: Date.now()
              });
              repairedCount++;
              logAuditAction('Memo Re-sequencing', `Repaired Memo ${m.id} to ${correctMemoNo} (FY: ${fy})`);
            } catch (err) {
              console.error('Error auto-repairing memo number:', err);
            }
          }
        }
        isRepairingMemosRef.current = false;
      }
    }

    if (manualAlert) {
      if (repairedCount > 0) {
        showAlert(`Successfully re-sequenced and repaired ${repairedCount} memo reference number(s) (100, 101, 102...).`);
      } else {
        showAlert('All memo reference numbers are already strictly sequential with no duplicates.');
      }
    }
  }, [memos]);

  // Run auto-deduplication check when memos update
  useEffect(() => {
    resequenceDuplicateMemos(false);
  }, [resequenceDuplicateMemos]);

  // --- SOE Available Options & Live Budget Monitor for Memo for Fund ---
  const memoAvailableSoes = useMemo(() => {
    let targetRangeId = userRangeId;
    if (!targetRangeId) {
      const fromStr = (memoFromInput || userRole || userRangeName || '').toLowerCase();
      if (fromStr.includes('narag') || fromStr.includes('nrg')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('narag'))?.id;
      } else if (fromStr.includes('habban') || fromStr.includes('hbn')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('habban'))?.id;
      } else if (fromStr.includes('division') || fromStr.includes('div')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('division'))?.id;
      } else if (fromStr.includes('rajgarh') || fromStr.includes('rjg')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('rajgarh') && !r.name.toLowerCase().includes('division'))?.id 
          || ranges.find(r => r.name.toLowerCase().includes('rajgarh'))?.id;
      } else if (fromStr.includes('sarahan') || fromStr.includes('srh')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('sarahan'))?.id;
      }
    }

    // Filter allocations for target range, FY, scheme, sector
    const relevantAllocs = allocations.filter(a => {
      if (targetRangeId && a.rangeId !== targetRangeId) return false;
      const aFY = a.financialYear || a.fyId;
      if (selectedFY && aFY && aFY !== selectedFY) return false;
      if (memoSchemeIdInput && a.schemeId !== memoSchemeIdInput) return false;
      if (memoSectorIdInput && a.sectorId && a.sectorId !== memoSectorIdInput) return false;
      return true;
    });

    const soeAllocMap: Record<string, { id: string; name: string; allocatedAmount: number; spentAmount: number; availableAmount: number }> = {};

    // 1. Collect from allocations' funded SOEs
    relevantAllocs.forEach(a => {
      if (Array.isArray(a.fundedSOEs) && a.fundedSOEs.length > 0) {
        a.fundedSOEs.forEach((f: any) => {
          const s = soes.find(soe => soe.id === f.soeId);
          const name = s?.name || 'Provisional';
          const id = f.soeId || s?.id || name;
          if (!soeAllocMap[id]) {
            soeAllocMap[id] = { id, name, allocatedAmount: 0, spentAmount: 0, availableAmount: 0 };
          }
          soeAllocMap[id].allocatedAmount += Number(f.amount) || 0;
        });
      } else if ((a as any).soeId) {
        const legacySoeId = (a as any).soeId;
        const s = soes.find(soe => soe.id === legacySoeId);
        const name = s?.name || 'Provisional';
        const id = legacySoeId;
        if (!soeAllocMap[id]) {
          soeAllocMap[id] = { id, name, allocatedAmount: 0, spentAmount: 0, availableAmount: 0 };
        }
        soeAllocMap[id].allocatedAmount += Number(a.amount) || 0;
      }
    });

    // 2. Compute spent and available for each SOE with allocated amount
    const relevantAllocIds = relevantAllocs.map(a => a.id);
    Object.keys(soeAllocMap).forEach(soeId => {
      const targetSoe = soes.find(s => s.id === soeId);
      const targetSoeName = targetSoe?.name || soeAllocMap[soeId].name;
      let spent = 0;
      expenses.forEach(e => {
        if (e.status === 'rejected') return;
        const isMatchAlloc = targetRangeId ? (relevantAllocIds.includes(e.allocationId) || allocations.find(a => a.id === e.allocationId)?.rangeId === targetRangeId) : true;
        const eSoe = soes.find(s => s.id === e.soeId);
        const isMatchSoe = e.soeId === soeId || eSoe?.name === targetSoeName;
        if (isMatchAlloc && isMatchSoe) {
          spent += Number(e.amount) || 0;
        }
      });
      soeAllocMap[soeId].spentAmount = spent;
      soeAllocMap[soeId].availableAmount = Math.max(0, soeAllocMap[soeId].allocatedAmount - spent);
    });

    // ONLY return SOEs in which budget is allotted (allocatedAmount > 0)
    return Object.values(soeAllocMap)
      .filter(s => s.allocatedAmount > 0)
      .sort((a, b) => b.allocatedAmount - a.allocatedAmount || a.name.localeCompare(b.name));
  }, [allocations, expenses, soes, userRangeId, ranges, memoFromInput, userRole, userRangeName, selectedFY, memoSchemeIdInput, memoSectorIdInput]);

  const memoBudgetInfo = useMemo(() => {
    let targetRangeId = userRangeId;
    if (!targetRangeId) {
      const fromStr = (memoFromInput || userRole || userRangeName || '').toLowerCase();
      if (fromStr.includes('narag') || fromStr.includes('nrg')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('narag'))?.id;
      } else if (fromStr.includes('habban') || fromStr.includes('hbn')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('habban'))?.id;
      } else if (fromStr.includes('division') || fromStr.includes('div')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('division'))?.id;
      } else if (fromStr.includes('rajgarh') || fromStr.includes('rjg')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('rajgarh') && !r.name.toLowerCase().includes('division'))?.id 
          || ranges.find(r => r.name.toLowerCase().includes('rajgarh'))?.id;
      } else if (fromStr.includes('sarahan') || fromStr.includes('srh')) {
        targetRangeId = ranges.find(r => r.name.toLowerCase().includes('sarahan'))?.id;
      }
    }

    const relevantAllocs = allocations.filter(a => {
      if (targetRangeId && a.rangeId !== targetRangeId) return false;
      const aFY = a.financialYear || a.fyId;
      if (selectedFY && aFY && aFY !== selectedFY) return false;
      if (memoSchemeIdInput && a.schemeId !== memoSchemeIdInput) return false;
      if (memoSectorIdInput && a.sectorId && a.sectorId !== memoSectorIdInput) return false;
      return true;
    });

    let allocatedBudget = 0;
    let totalSpent = 0;

    if (memoSoeIdInput) {
      const targetSoe = soes.find(s => s.id === memoSoeIdInput);
      const targetSoeName = targetSoe?.name;

      relevantAllocs.forEach(a => {
        const funded = a.fundedSOEs?.find(f => {
          if (f.soeId === memoSoeIdInput) return true;
          const s = soes.find(soe => soe.id === f.soeId);
          return s?.name === targetSoeName;
        });
        if (funded) {
          allocatedBudget += Number(funded.amount) || 0;
        }
      });

      const relevantAllocIds = relevantAllocs.map(a => a.id);
      expenses.forEach(e => {
        if (e.status === 'rejected') return;
        const isMatchAlloc = targetRangeId ? (relevantAllocIds.includes(e.allocationId) || allocations.find(a => a.id === e.allocationId)?.rangeId === targetRangeId) : true;
        const eSoe = soes.find(s => s.id === e.soeId);
        const isMatchSoe = e.soeId === memoSoeIdInput || eSoe?.name === targetSoeName;
        if (isMatchAlloc && isMatchSoe) {
          totalSpent += Number(e.amount) || 0;
        }
      });
    } else if (memoSchemeIdInput || memoSectorIdInput) {
      relevantAllocs.forEach(a => {
        allocatedBudget += Number(a.amount) || 0;
      });
      const relevantAllocIds = relevantAllocs.map(a => a.id);
      expenses.forEach(e => {
        if (e.status === 'rejected') return;
        if (relevantAllocIds.includes(e.allocationId)) {
          totalSpent += Number(e.amount) || 0;
        }
      });
    }

    const availableBudget = Math.max(0, allocatedBudget - totalSpent);
    
    // Sum of payees currently added in memo table
    const currentMemoTotal = memoPayeeEntries.reduce((sum, e) => sum + (Math.round(Number(e.totalAmount)) || 0), 0);
    
    // If editing a specific row, exclude its previous amount to compute remaining available allowance for this row
    const editingOldAmount = (editingEntryIndex !== null && memoPayeeEntries[editingEntryIndex]) ? (Math.round(Number(memoPayeeEntries[editingEntryIndex].totalAmount)) || 0) : 0;
    const memoTotalWithoutEditingRow = Math.max(0, currentMemoTotal - editingOldAmount);
    
    // Available balance specifically for this next/edited payee entry
    const availableForCurrentPayee = Math.max(0, availableBudget - memoTotalWithoutEditingRow);

    // Current amount typed in the input field
    const typedAmount = Math.round(parseFloat(entryTotalAmount) || 0);

    // Live remaining balance of SOE after subtracting typed amount
    const liveRemainingAfterTyping = availableForCurrentPayee - typedAmount;

    // Projected total of memo if current payee amount is added
    const projectedMemoTotal = memoTotalWithoutEditingRow + typedAmount;

    // Status flags
    const isExceeded = memoSoeIdInput ? (allocatedBudget > 0 && currentMemoTotal > availableBudget) : false;
    const isTypingExceeded = memoSoeIdInput ? (allocatedBudget > 0 && typedAmount > availableForCurrentPayee) : false;
    const remainingAfterMemo = availableBudget - currentMemoTotal;

    return {
      allocatedBudget,
      totalSpent,
      availableBudget,
      currentMemoTotal,
      memoTotalWithoutEditingRow,
      availableForCurrentPayee,
      typedAmount,
      liveRemainingAfterTyping,
      projectedMemoTotal,
      isExceeded,
      isTypingExceeded,
      remainingAfterMemo
    };
  }, [allocations, expenses, userRangeId, ranges, memoFromInput, userRole, userRangeName, selectedFY, memoSchemeIdInput, memoSectorIdInput, memoSoeIdInput, soes, memoPayeeEntries, editingEntryIndex, entryTotalAmount]);

  // --- Memo for Fund Handlers ---
  const searchedMemoPayees = useMemo(() => {
    if (!memoPayeeSearchTerm.trim()) return filteredPayeesList;
    const term = memoPayeeSearchTerm.toLowerCase().trim();
    return filteredPayeesList.filter(p => {
      const name = (p.name || '').toLowerCase();
      const acc = (p.accountNumber || '').toLowerCase();
      const ifsc = (p.ifscCode || '').toLowerCase();
      const pan = (p.panNumber || '').toLowerCase();
      const gst = (p.gstNumber || '').toLowerCase();
      const tryCode = (p.treasuryCode || '').toLowerCase();
      const rangeName = (ranges.find(r => r.id === p.rangeId)?.name || '').toLowerCase();
      const address = (p.address || '').toLowerCase();
      return (
        name.includes(term) ||
        acc.includes(term) ||
        ifsc.includes(term) ||
        pan.includes(term) ||
        gst.includes(term) ||
        tryCode.includes(term) ||
        rangeName.includes(term) ||
        address.includes(term)
      );
    });
  }, [filteredPayeesList, memoPayeeSearchTerm, ranges]);

  const handleSelectPayeeForMemoEntry = (payeeId: string) => {
    setEntryPayeeId(payeeId);
    setShowMemoPayeeDropdown(false);
    if (!payeeId) {
      setEntryName('');
      setEntryAddress('');
      setEntryAccountNo('');
      setEntryIfsc('');
      setEntryTreasuryCode('');
      setEntryPan('');
      setEntryGst('');
      return;
    }
    const p = payees.find(item => item.id === payeeId);
    if (p) {
      setEntryName(p.name || '');
      setEntryAddress(p.address || '');
      setEntryAccountNo(p.accountNumber || '');
      setEntryIfsc(p.ifscCode || '');
      setEntryTreasuryCode(p.treasuryCode || '');
      setEntryPan(p.panNumber || '');
      setEntryGst(p.gstNumber || '');
      setMemoPayeeSearchTerm('');
    }
  };

  const handleTotalAmountInputChange = (val: string) => {
    setEntryTotalAmount(val);
    const tot = parseFloat(val);
    if (!isNaN(tot) && tot > 0) {
      if (tot > 30000) {
        setEntryDeductITax(true);
      } else {
        setEntryDeductITax(false);
      }
      if (tot > 250000) {
        setEntryDeductGst(true);
      } else {
        setEntryDeductGst(false);
      }
    } else {
      setEntryDeductITax(false);
      setEntryDeductGst(false);
    }
  };

  const handleStartEditSubVoucher = (sv: MemoSubVoucher) => {
    setEditingSubVoucherId(sv.id);
    setSubVoucherNoInput(sv.voucherNo || '');
    setSubVoucherDescInput(sv.description || '');
    setSubVoucherAmountInput(String(sv.amount || ''));
  };

  const handleCancelEditSubVoucher = () => {
    setEditingSubVoucherId(null);
    setSubVoucherNoInput('');
    setSubVoucherDescInput('');
    setSubVoucherAmountInput('');
  };

  const handleAddSubVoucher = () => {
    const amt = parseFloat(subVoucherAmountInput);
    if (isNaN(amt) || amt <= 0) {
      showAlert('Please enter a valid Sub-voucher Amount.');
      return;
    }

    const otherSubSum = entrySubVouchers
      .filter(v => v.id !== editingSubVoucherId)
      .reduce((s, v) => s + v.amount, 0);
    const newTotalSum = otherSubSum + Math.round(amt);

    // Enforce SOE Budget Restriction if SOE is selected
    if (memoSoeIdInput && memoBudgetInfo.allocatedBudget > 0) {
      const availableForThis = memoBudgetInfo.availableForCurrentPayee;
      if (newTotalSum > availableForThis) {
        const soeName = soes.find(s => s.id === memoSoeIdInput)?.name || 'Selected SOE';
        showAlert(
          `Cannot ${editingSubVoucherId ? 'update' : 'add'} Sub-voucher: Total amount (â‚¹${newTotalSum.toLocaleString('en-IN')}) will exceed the available SOE budget balance of â‚¹${Math.max(0, availableForThis).toLocaleString('en-IN')} for ${soeName}.`
        );
        return;
      }
    }

    let updatedSubVouchers: MemoSubVoucher[];
    if (editingSubVoucherId) {
      updatedSubVouchers = entrySubVouchers.map(sv => {
        if (sv.id === editingSubVoucherId) {
          return {
            ...sv,
            voucherNo: subVoucherNoInput.trim() || '',
            description: subVoucherDescInput.trim() || '',
            amount: Math.round(amt)
          };
        }
        return sv;
      });
      setEditingSubVoucherId(null);
    } else {
      const newSubVoucher: MemoSubVoucher = {
        id: Math.random().toString(36).substring(2, 9),
        voucherNo: subVoucherNoInput.trim() || '',
        description: subVoucherDescInput.trim() || '',
        amount: Math.round(amt)
      };
      updatedSubVouchers = [...entrySubVouchers, newSubVoucher];
    }

    setEntrySubVouchers(updatedSubVouchers);

    // Automatically recalculate sum and update entryTotalAmount
    const totalSum = updatedSubVouchers.reduce((s, v) => s + v.amount, 0);
    handleTotalAmountInputChange(String(totalSum));

    // Clear sub-voucher inputs
    setSubVoucherNoInput('');
    setSubVoucherDescInput('');
    setSubVoucherAmountInput('');
  };

  const handleRemoveSubVoucher = (id: string) => {
    if (editingSubVoucherId === id) {
      handleCancelEditSubVoucher();
    }
    const updated = entrySubVouchers.filter(v => v.id !== id);
    setEntrySubVouchers(updated);
    const totalSum = updated.reduce((s, v) => s + v.amount, 0);
    handleTotalAmountInputChange(String(totalSum));
  };

  const handleAddOrUpdatePayeeEntry = (e: React.FormEvent) => {
    e.preventDefault();
    if (!entryName.trim()) {
      showAlert('Please enter Payee Name.');
      return;
    }
    if (!entryAccountNo.trim()) {
      showAlert('Please enter Payee Bank Account Number.');
      return;
    }

    // Check duplicate account number in current memo list
    const cleanAcc = entryAccountNo.trim().replace(/\s+/g, '').toLowerCase();
    const duplicateInMemo = memoPayeeEntries.find((p, idx) => {
      if (editingEntryIndex !== null && idx === editingEntryIndex) return false;
      const pAcc = (p.accountNumber || '').trim().replace(/\s+/g, '').toLowerCase();
      return Boolean(pAcc && pAcc === cleanAcc);
    });

    if (duplicateInMemo) {
      showAlert(`Payee with Account Number "${entryAccountNo.trim()}" (${duplicateInMemo.name}) is already added in this Memo.`);
      return;
    }

    // Must have at least one sub-voucher added
    if (entrySubVouchers.length === 0) {
      showAlert('Please add at least one Sub-Voucher / Bill in the breakdown form below to calculate the Expenditure Amount.');
      return;
    }

    const tot = parseFloat(entryTotalAmount);
    if (isNaN(tot) || tot <= 0) {
      showAlert('Please enter a valid Total Amount through the Sub-Vouchers breakdown below.');
      return;
    }

    // STRICT BUDGET RESTRICTION: Do not allow exceeding remaining available budget under selected SOE
    if (memoSoeIdInput && memoBudgetInfo.allocatedBudget > 0) {
      const availableForThis = memoBudgetInfo.availableForCurrentPayee;
      if (Math.round(tot) > availableForThis) {
        const soeName = soes.find(s => s.id === memoSoeIdInput)?.name || 'Selected SOE';
        showAlert(
          `Cannot add Payee: Amount (â‚¹${Math.round(tot).toLocaleString('en-IN')}) exceeds the remaining available SOE budget of â‚¹${Math.max(0, availableForThis).toLocaleString('en-IN')} for ${soeName}.\n\nTotal Allotted: â‚¹${memoBudgetInfo.allocatedBudget.toLocaleString('en-IN')} | Already Spent: â‚¹${memoBudgetInfo.totalSpent.toLocaleString('en-IN')} | Added in Memo: â‚¹${memoBudgetInfo.memoTotalWithoutEditingRow.toLocaleString('en-IN')} | Remaining: â‚¹${Math.max(0, availableForThis).toLocaleString('en-IN')}`
        );
        return;
      }
    }

    const isITaxApplicable = entryDeductITax && tot > 30000;
    const isGstApplicable = entryDeductGst && tot > 250000;

    const iTaxP = isITaxApplicable ? (parseFloat(entryITaxPercent) || 0) : 0;
    const gstP = isGstApplicable ? (parseFloat(entryGstPercent) || 0) : 0;

    const iTaxAmt = isITaxApplicable ? Math.round((tot * iTaxP) / 100) : 0;
    const gstAmt = isGstApplicable ? Math.round((tot * gstP) / 100) : 0;
    const netRtgs = Math.round(tot) - iTaxAmt - gstAmt;

    const sanitizedSubVouchers = entrySubVouchers.map(sv => ({
      id: sv.id || Math.random().toString(36).substring(2, 9),
      voucherNo: (sv.voucherNo || '').trim(),
      description: (sv.description || '').trim(),
      amount: Math.round(Number(sv.amount) || 0)
    }));

    const newEntry: MemoPayeeEntry = {
      payeeId: (entryPayeeId || '').trim(),
      name: entryName.trim(),
      address: (entryAddress || '').trim(),
      accountNumber: entryAccountNo.trim(),
      ifscCode: (entryIfsc || '').trim(),
      treasuryCode: (entryTreasuryCode || '').trim(),
      panNumber: (entryPan || '').trim(),
      gstNumber: (entryGst || '').trim(),
      totalAmount: Math.round(tot),
      deductITax: isITaxApplicable,
      iTaxPercent: parseFloat(entryITaxPercent) || 0,
      iTaxAmount: iTaxAmt,
      deductGst: isGstApplicable,
      gstPercent: parseFloat(entryGstPercent) || 0,
      gstAmount: gstAmt,
      netRtgsAmount: netRtgs,
      subVouchers: sanitizedSubVouchers
    };

    if (editingEntryIndex !== null) {
      const updated = [...memoPayeeEntries];
      updated[editingEntryIndex] = newEntry;
      setMemoPayeeEntries(updated);
      setEditingEntryIndex(null);
    } else {
      setMemoPayeeEntries(prev => [...prev, newEntry]);
    }

    // Reset entry input fields
    setEntryPayeeId('');
    setEntryName('');
    setEntryAddress('');
    setEntryAccountNo('');
    setEntryIfsc('');
    setEntryTreasuryCode('');
    setEntryPan('');
    setEntryGst('');
    setEntryTotalAmount('');
    setEntryDeductITax(true);
    setEntryITaxPercent('1');
    setEntryDeductGst(false);
    setEntryGstPercent('2');
    setEntrySubVouchers([]);
    setShowSubVoucherSection(false);
    setEditingSubVoucherId(null);
    setSubVoucherNoInput('');
    setSubVoucherDescInput('');
    setSubVoucherAmountInput('');
  };

  const handleEditPayeeEntryInForm = (index: number) => {
    const entry = memoPayeeEntries[index];
    if (!entry) return;
    setEditingEntryIndex(index);
    setEntryPayeeId(entry.payeeId || '');
    setEntryName(entry.name || '');
    setEntryAddress(entry.address || '');
    setEntryAccountNo(entry.accountNumber || '');
    setEntryIfsc(entry.ifscCode || '');
    setEntryTreasuryCode(entry.treasuryCode || '');
    setEntryPan(entry.panNumber || '');
    setEntryGst(entry.gstNumber || '');
    setEntryTotalAmount(entry.totalAmount ? String(entry.totalAmount) : '');
    setEntryDeductITax(entry.deductITax !== undefined ? entry.deductITax : (entry.iTaxAmount > 0));
    setEntryITaxPercent(entry.iTaxPercent !== undefined ? String(entry.iTaxPercent) : '1');
    setEntryDeductGst(entry.deductGst !== undefined ? entry.deductGst : ((entry.gstAmount || 0) > 0));
    setEntryGstPercent(entry.gstPercent !== undefined ? String(entry.gstPercent) : '2');
    
    if (entry.subVouchers && entry.subVouchers.length > 0) {
      setEntrySubVouchers(entry.subVouchers.map(sv => ({
        id: sv.id || Math.random().toString(36).substring(2, 9),
        voucherNo: sv.voucherNo || '',
        description: sv.description || '',
        amount: Math.round(Number(sv.amount) || 0)
      })));
    } else if (entry.totalAmount && entry.totalAmount > 0) {
      setEntrySubVouchers([{
        id: Math.random().toString(36).substring(2, 9),
        voucherNo: 'V-01',
        description: 'Expenditure Voucher',
        amount: entry.totalAmount
      }]);
    } else {
      setEntrySubVouchers([]);
    }
    setShowSubVoucherSection(true);
    setEditingSubVoucherId(null);
    setSubVoucherNoInput('');
    setSubVoucherDescInput('');
    setSubVoucherAmountInput('');
  };

  const handleRemovePayeeEntryFromForm = (index: number) => {
    setMemoPayeeEntries(prev => prev.filter((_, i) => i !== index));
    if (editingEntryIndex === index) {
      setEditingEntryIndex(null);
      setEntryPayeeId('');
      setEntryName('');
      setEntryAddress('');
      setEntryAccountNo('');
      setEntryIfsc('');
      setEntryPan('');
      setEntryGst('');
      setEntryTotalAmount('');
      setEntryDeductITax(true);
      setEntryITaxPercent('1');
      setEntryDeductGst(false);
      setEntryGstPercent('2');
      setEntrySubVouchers([]);
      setShowSubVoucherSection(false);
      setEditingSubVoucherId(null);
      setSubVoucherNoInput('');
      setSubVoucherDescInput('');
      setSubVoucherAmountInput('');
    }
  };

  const handleSaveMemo = async (status: 'draft' | 'submitted') => {
    // 1. Validate Memo Date
    if (!memoDateInput || !memoDateInput.trim()) {
      showAlert('Please select or enter the Memo Date.');
      return;
    }

    // 2. Validate Memo Reference Number
    if (!memoNoInput || !memoNoInput.trim()) {
      showAlert('Please enter Memo Reference Number (e.g., 101, 102).');
      return;
    }

    // 3. Validate Month / Period
    if (!memoMonthYearInput || !memoMonthYearInput.trim()) {
      showAlert('Please enter Month / Period (e.g., February 2026 or 02/2026).');
      return;
    }

    // 4. Validate From Authority
    const effectiveFromAuthority = (memoFromInput || defaultFromAuthority || userRangeName || 'Sarahan').trim();
    if (!effectiveFromAuthority) {
      showAlert('Please select or specify From Authority / Range.');
      return;
    }

    // 5. Validate To Authority
    if (!memoToInput || !memoToInput.trim()) {
      showAlert('Please enter To Authority (e.g., DCF Rajgarh).');
      return;
    }

    // 6. Validate Scheme
    if (!memoSchemeIdInput || !memoSchemeIdInput.trim()) {
      showAlert('Please select a Scheme for this Memo.');
      return;
    }

    // 7. Validate Payees presence
    if (memoPayeeEntries.length === 0) {
      if (entryName.trim() || entryAccountNo.trim() || entryTotalAmount) {
        showAlert(`You have entered details for payee "${entryName.trim() || 'Payee'}" in Step 2, but have not clicked "+ Add Payee to Memo" yet.\n\nPlease click the "+ Add Payee to Memo" button to add this payee to the memo before saving.`);
        return;
      }
      showAlert('Please add at least one Payee to the Memo before saving.\n\nFill out the payee details in Step 2 above and click "+ Add Payee to Memo".');
      return;
    }

    // 8. Validate each Payee in the list
    for (let i = 0; i < memoPayeeEntries.length; i++) {
      const p = memoPayeeEntries[i];
      if (!p.name || !p.name.trim()) {
        showAlert(`Payee #${i + 1} is missing a Name. Please edit and specify the payee name.`);
        return;
      }
      if (!p.accountNumber || !p.accountNumber.trim()) {
        showAlert(`Payee #${i + 1} (${p.name}) is missing a Bank Account Number. Please edit and specify the account number.`);
        return;
      }
      const pTot = Number(p.totalAmount) || 0;
      if (pTot <= 0) {
        showAlert(`Payee #${i + 1} (${p.name}) has an invalid Total Amount (â‚¹${pTot}). Amount must be greater than zero.`);
        return;
      }
    }

    const schemeObj = currentSchemes.find(s => s.id === memoSchemeIdInput);
    const sectorObj = currentSectors.find(s => s.id === memoSectorIdInput);
    const soeObj = soes.find(s => s.id === memoSoeIdInput);

    const totalGross = memoPayeeEntries.reduce((sum, e) => sum + (Math.round(Number(e.totalAmount)) || 0), 0);
    const totalITax = memoPayeeEntries.reduce((sum, e) => sum + (Math.round(Number(e.iTaxAmount)) || 0), 0);
    const totalGst = memoPayeeEntries.reduce((sum, e) => sum + (Math.round(Number(e.gstAmount)) || 0), 0);
    const totalNetRtgs = memoPayeeEntries.reduce((sum, e) => sum + (Math.round(Number(e.netRtgsAmount)) || 0), 0);

    // Budget check: If SOE is selected and budget is allotted, prevent saving if memo total exceeds available budget
    if (memoSoeIdInput && memoBudgetInfo.allocatedBudget > 0 && memoBudgetInfo.isExceeded) {
      const soeName = soeObj?.name || 'the selected SOE';
      showAlert(`Cannot save Memo: Total memo amount (â‚¹${totalGross.toLocaleString('en-IN')}) exceeds the available allocated budget balance (â‚¹${memoBudgetInfo.availableBudget.toLocaleString('en-IN')}) for ${soeName}. Please adjust payee amounts or choose appropriate head.`);
      return;
    }

    // Determine final guaranteed unique reference number
    let finalMemoNo = editingMemo ? editingMemo.memoNo : (memoNoInput.trim() || autoMemoNo);
    if (!editingMemo) {
      const freshNextNo = getNextSequentialMemoNo(effectiveFromAuthority, selectedFY || '2026-27', memos);
      const isColliding = memos.some(m => {
        const mFY = m.financialYear || m.fyId;
        return (mFY === (selectedFY || '2026-27')) && m.memoNo === finalMemoNo;
      });
      if (isColliding) {
        finalMemoNo = freshNextNo;
      }
    }

    // Deeply sanitize payee entries to guarantee no undefined values exist
    const sanitizedPayees = memoPayeeEntries.map(p => ({
      payeeId: (p.payeeId || '').trim(),
      name: (p.name || '').trim(),
      address: (p.address || '').trim(),
      accountNumber: (p.accountNumber || '').trim(),
      ifscCode: (p.ifscCode || '').trim(),
      treasuryCode: (p.treasuryCode || '').trim(),
      panNumber: (p.panNumber || '').trim(),
      gstNumber: (p.gstNumber || '').trim(),
      totalAmount: Math.round(Number(p.totalAmount) || 0),
      deductITax: Boolean(p.deductITax),
      iTaxPercent: Number(p.iTaxPercent) || 0,
      iTaxAmount: Math.round(Number(p.iTaxAmount) || 0),
      deductGst: Boolean(p.deductGst),
      gstPercent: Number(p.gstPercent) || 0,
      gstAmount: Math.round(Number(p.gstAmount) || 0),
      netRtgsAmount: Math.round(Number(p.netRtgsAmount) || 0),
      subVouchers: (p.subVouchers || []).map(sv => ({
        id: sv.id || Math.random().toString(36).substring(2, 9),
        voucherNo: (sv.voucherNo || '').trim(),
        description: (sv.description || '').trim(),
        amount: Math.round(Number(sv.amount) || 0)
      }))
    }));

    const memoData = {
      memoNo: finalMemoNo,
      date: memoDateInput,
      monthYear: memoMonthYearInput.trim(),
      schemeId: memoSchemeIdInput || '',
      schemeName: schemeObj ? schemeObj.name : 'All Schemes',
      sectorId: memoSectorIdInput || '',
      sectorName: sectorObj ? sectorObj.name : '',
      soeId: memoSoeIdInput || '',
      soeName: soeObj ? soeObj.name : '',
      rangeId: userRangeId || '',
      rangeName: effectiveFromAuthority,
      toAuthority: memoToInput.trim() || 'DCF Rajgarh',
      financialYear: selectedFY || '2026-27',
      status,
      totalAmount: totalGross,
      totalITax: totalITax,
      totalGst: totalGst,
      totalNetRtgs: totalNetRtgs,
      payeeEntries: sanitizedPayees,
      createdBy: editingMemo?.createdBy || user?.uid || '',
      createdByRole: editingMemo?.createdByRole || userRole || '',
      createdByName: editingMemo?.createdByName || user?.email || '',
      updatedAt: Date.now(),
      ...(copiedFromMemoInfo ? { copiedFromMemoNo: copiedFromMemoInfo.memoNo } : {}),
      ...(editingMemo?.pulledBack ? {
        pulledBack: editingMemo.pulledBack,
        pullBackRemarks: editingMemo.pullBackRemarks,
        pulledBackBy: editingMemo.pulledBackBy,
        pulledBackAt: editingMemo.pulledBackAt
      } : {})
    };

    try {
      const cleanData = sanitizeFirestoreDoc(memoData);
      if (editingMemo) {
        await updateDoc(doc(db, 'memos', editingMemo.id), {
          ...cleanData,
          memoNo: editingMemo.memoNo // Strictly maintain original reference number when editing / submitting to admin
        });
        localStorage.removeItem('rajgarh_draft_memo_fund');
        setLastAutoSaveMemoTime(null);
        logAuditAction('Memo Updated', `Memo No: ${editingMemo.memoNo}, Status: ${status}, Payees: ${sanitizedPayees.length}, Amount: â‚¹${totalGross}`);
        showAlert(`Memo for Fund ${status === 'submitted' ? 'submitted and locked' : 'updated as draft'} successfully.`);
      } else {
        await addDoc(collection(db, 'memos'), {
          ...cleanData,
          createdAt: Date.now()
        });
        localStorage.removeItem('rajgarh_draft_memo_fund');
        setLastAutoSaveMemoTime(null);
        logAuditAction('Memo Created', `Memo No: ${finalMemoNo}, Status: ${status}, Payees: ${sanitizedPayees.length}, Amount: â‚¹${totalGross}`);
        showAlert(`Memo for Fund ${status === 'submitted' ? 'submitted and locked' : 'saved as draft'} successfully.`);
      }
      handleResetMemoForm();
    } catch (error: any) {
      console.error('Firestore Error saving memo:', error);
      const errMsg = error?.message || error?.code || String(error);
      showAlert(`Unable to save Memo for Fund.\n\nPlease contact the administrator with this error:\n"${errMsg}"`);
    }
  };

    const handleDuplicateMemo = (memo: MemoForFund) => {
    // Generate the next sequential memo number in line for this Range/Authority in this FY
    const fromAuthority = memo.rangeName || defaultFromAuthority || userRangeName || 'Sarahan';
    const nextNo = getNextSequentialMemoNo(fromAuthority, selectedFY || '2026-27', memos);

    setEditingMemo(null);
    setMemoNoInput(nextNo);
    setMemoDateInput(new Date().toISOString().split('T')[0]);
    setMemoMonthYearInput(memo.monthYear || '');
    setMemoSchemeIdInput(memo.schemeId || '');
    setMemoSectorIdInput(memo.sectorId || '');
    setMemoSoeIdInput(memo.soeId || '');
    setMemoFromInput(memo.rangeName || defaultFromAuthority);
    setMemoToInput(memo.toAuthority || 'DCF Rajgarh');

    // Deep clone payees so the user can freely edit or delete rows without affecting original memo
    const clonedPayees: MemoPayeeEntry[] = (memo.payeeEntries || []).map(p => ({
      ...p,
      subVouchers: p.subVouchers ? p.subVouchers.map(sv => ({ ...sv })) : []
    }));

    setMemoPayeeEntries(clonedPayees);
    setEditingEntryIndex(null);
    setEntryPayeeId('');
    setEntryName('');
    setEntryAddress('');
    setEntryAccountNo('');
    setEntryIfsc('');
    setEntryTreasuryCode('');
    setEntryPan('');
    setEntryGst('');
    setEntryTotalAmount('');
    setEntrySubVouchers([]);
    setShowSubVoucherSection(false);

    setCopiedFromMemoInfo({ memoNo: memo.memoNo, originalId: memo.id });

    showAlert(`Created duplicate draft from Memo #${memo.memoNo} with new sequential Memo #${nextNo}. You can now edit payee amounts, delete unwanted payees, or add new payees, then save as Draft or submit.`);
  };

  const handleOpenPullBackModal = (memo: MemoForFund) => {
    if (memo.status !== 'submitted') {
      showAlert('Only submitted memos can be pulled back.');
      return;
    }
    if (memo.viewedByAdmin) {
      showAlert(`This memo has already been reviewed by Headquarter (${memo.viewedBy || 'Admin'}${memo.viewedAt ? ` on ${new Date(memo.viewedAt).toLocaleDateString('en-GB')}` : ''}). It cannot be pulled back by the range user. Please contact Headquarter/Admin to return it for correction.`);
      return;
    }
    setMemoPullBackModalData({ memo, remarks: '' });
  };

  const handleConfirmPullBack = async () => {
    if (!memoPullBackModalData) return;
    const { memo, remarks } = memoPullBackModalData;
    if (!remarks.trim()) {
      showAlert('Please enter the reason / remarks for pulling back the memo.');
      return;
    }

    try {
      setIsPullingBack(true);
      const roleLabel = userRole || 'Range User';
      const authorLabel = user?.displayName ? `${roleLabel} (${user.displayName})` : (user?.email ? `${roleLabel} (${user.email})` : roleLabel);

      await updateDoc(doc(db, 'memos', memo.id), {
        status: 'draft',
        pulledBack: true,
        pullBackRemarks: remarks.trim(),
        pulledBackBy: authorLabel,
        pulledBackAt: Date.now(),
        updatedAt: Date.now()
      });

      logAuditAction(
        'Memo Pulled Back',
        `Memo No: ${memo.memoNo} pulled back to Draft by ${authorLabel} with remarks: "${remarks.trim()}"`
      );

      showAlert(`Memo No. ${memo.memoNo} has been successfully pulled back to Draft. You can now edit and re-submit it.`);
      setMemoPullBackModalData(null);

      // Automatically load into form for editing
      handleEditMemo({
        ...memo,
        status: 'draft',
        pulledBack: true,
        pullBackRemarks: remarks.trim(),
        pulledBackBy: authorLabel,
        pulledBackAt: Date.now()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `memos/${memo.id}`);
    } finally {
      setIsPullingBack(false);
    }
  };

  const handleEditMemo = (memo: MemoForFund) => {
    setEditingMemo(memo);
    setCopiedFromMemoInfo(null);
    setMemoNoInput(memo.memoNo || '');
    setMemoDateInput(memo.date || new Date().toISOString().split('T')[0]);
    setMemoMonthYearInput(memo.monthYear || '');
    setMemoSchemeIdInput(memo.schemeId || '');
    setMemoSectorIdInput(memo.sectorId || '');
    setMemoSoeIdInput(memo.soeId || '');
    setMemoFromInput(memo.rangeName || defaultFromAuthority);
    setMemoToInput(memo.toAuthority || 'DCF Rajgarh');
    setMemoPayeeEntries(memo.payeeEntries || []);
    setEditingEntryIndex(null);
  };

  const handleDeleteMemo = (memo: MemoForFund) => {
    showConfirm(`Are you sure you want to delete Memo No. ${memo.memoNo}?`, async () => {
      try {
        await deleteDoc(doc(db, 'memos', memo.id));
        logAuditAction('Memo Deleted', `Memo No: ${memo.memoNo}`);
        showAlert('Memo deleted successfully.');
        if (editingMemo?.id === memo.id) {
          handleResetMemoForm();
        }
        if (viewingMemo?.id === memo.id) {
          setViewingMemo(null);
        }
      } catch (error) {
        handleFirestoreError(error, OperationType.DELETE, `memos/${memo.id}`);
      }
    });
  };

  const handleOpenSendBackModal = (memo: MemoForFund) => {
    if (userRole !== 'admin' && userRole !== 'deo') {
      showAlert('Only Admin or Data Entry Operator can send back memo for correction.');
      return;
    }
    setMemoCorrectionModalData({ memo, remarks: memo.correctionRemarks || memo.remarks || '' });
  };

  const handleConfirmSendBackForCorrection = async () => {
    if (!memoCorrectionModalData) return;
    const { memo, remarks } = memoCorrectionModalData;
    if (!remarks.trim()) {
      showAlert('Please enter correction remarks / reason so the user knows what errors need to be fixed in the memo.');
      return;
    }

    try {
      setIsSendingCorrection(true);
      const roleLabel = userRole === 'admin' ? 'Admin' : (userRole === 'deo' ? 'DEO' : (userRole || 'User'));
      const authorLabel = user?.displayName ? `${roleLabel} (${user.displayName})` : roleLabel;

      await updateDoc(doc(db, 'memos', memo.id), {
        status: 'correction',
        correctionRemarks: remarks.trim(),
        remarks: remarks.trim(),
        correctionRemarksBy: authorLabel,
        correctionRemarksAt: Date.now(),
        updatedAt: Date.now()
      });

      logAuditAction(
        'Memo Returned for Correction',
        `Memo No: ${memo.memoNo} returned by ${authorLabel} with remarks: "${remarks.trim()}"`
      );

      showAlert(`Memo No. ${memo.memoNo} has been returned to ${memo.rangeName || 'the user'} for correction.`);
      setMemoCorrectionModalData(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `memos/${memo.id}`);
    } finally {
      setIsSendingCorrection(false);
    }
  };

  const handleApproveMemo = async (memo: MemoForFund) => {
    if (!isAdmin() && !isDEO() && userRole !== 'approver' && userRole !== 'DA') {
      showAlert('Only Admin or Data Entry Operator can approve memos for fund.');
      return;
    }

    try {
      const roleLabel = userRole === 'admin' ? 'Admin' : (userRole === 'deo' ? 'DEO' : (userRole === 'approver' ? 'Approver' : userRole || 'Headquarter'));
      const authorLabel = user?.displayName ? `${roleLabel} (${user.displayName})` : (user?.email ? `${roleLabel} (${user.email})` : roleLabel);

      await updateDoc(doc(db, 'memos', memo.id), {
        isApproved: true,
        approvedBy: authorLabel,
        approvedByRole: userRole || 'admin',
        approvedAt: Date.now(),
        viewedByAdmin: true,
        viewedAt: memo.viewedAt || Date.now(),
        viewedBy: memo.viewedBy || authorLabel,
        viewedByRole: memo.viewedByRole || userRole || 'admin',
        updatedAt: Date.now()
      });

      logAuditAction(
        'Memo Approved for Fund',
        `Memo No: ${memo.memoNo} approved for funding by ${authorLabel}`
      );

      showAlert(`Memo No. ${memo.memoNo} has been successfully Approved for Fund! The Range user can now see the "Approved by HQ" status.`);
      if (viewingMemo?.id === memo.id) {
        setViewingMemo({
          ...memo,
          isApproved: true,
          approvedBy: authorLabel,
          approvedByRole: userRole || 'admin',
          approvedAt: Date.now(),
          viewedByAdmin: true,
          viewedAt: memo.viewedAt || Date.now(),
          viewedBy: memo.viewedBy || authorLabel
        });
      }
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `memos/${memo.id}`);
    }
  };

  const handleRevokeMemoApproval = async (memo: MemoForFund) => {
    if (!isAdmin() && !isDEO() && userRole !== 'approver' && userRole !== 'DA') {
      showAlert('Only Admin or Data Entry Operator can revoke memo approval.');
      return;
    }

    showConfirm(`Are you sure you want to revoke approval for Memo No. ${memo.memoNo}?`, async () => {
      try {
        await updateDoc(doc(db, 'memos', memo.id), {
          isApproved: false,
          approvedBy: '',
          approvedByRole: '',
          approvedAt: 0,
          updatedAt: Date.now()
        });

        logAuditAction(
          'Memo Approval Revoked',
          `Approval for Memo No: ${memo.memoNo} was revoked by ${user?.email || userRole}`
        );

        showAlert(`Approval for Memo No. ${memo.memoNo} has been revoked.`);
        if (viewingMemo?.id === memo.id) {
          setViewingMemo({
            ...memo,
            isApproved: false,
            approvedBy: undefined,
            approvedByRole: undefined,
            approvedAt: undefined
          });
        }
      } catch (error) {
        handleFirestoreError(error, OperationType.UPDATE, `memos/${memo.id}`);
      }
    });
  };

  const handleMarkMemoAsRead = async (memo: MemoForFund) => {
    if (!isAuthorizedUserOrAdmin) return;
    try {
      const roleLabel = userRole === 'admin' ? 'Admin' : (userRole === 'deo' ? 'DEO' : (userRole === 'approver' ? 'Approver' : userRole || 'Headquarter'));
      const authorLabel = user?.displayName ? `${roleLabel} (${user.displayName})` : (user?.email ? `${roleLabel} (${user.email})` : roleLabel);
      await updateDoc(doc(db, 'memos', memo.id), {
        viewedByAdmin: true,
        viewedAt: Date.now(),
        viewedBy: authorLabel,
        viewedByRole: userRole || 'admin',
        updatedAt: Date.now()
      });
      logAuditAction('Memo Marked Reviewed', `Memo No: ${memo.memoNo} marked as reviewed by ${authorLabel}`);
      showAlert(`Memo No. ${memo.memoNo} marked as Read / Reviewed.`);
      if (viewingMemo?.id === memo.id) {
        setViewingMemo({
          ...memo,
          viewedByAdmin: true,
          viewedAt: Date.now(),
          viewedBy: authorLabel
        });
      }
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `memos/${memo.id}`);
    }
  };

  const handleResetMemoForm = () => {
    setEditingMemo(null);
    setCopiedFromMemoInfo(null);
    setMemoNoInput(autoMemoNo);
    setMemoDateInput(new Date().toISOString().split('T')[0]);
    setMemoMonthYearInput('08/2026');
    setMemoSchemeIdInput('');
    setMemoSectorIdInput('');
    setMemoSoeIdInput('');
    setMemoFromInput(defaultFromAuthority);
    setMemoToInput('DCF Rajgarh');
    setMemoPayeeEntries([]);
    setEditingEntryIndex(null);
    setEntryPayeeId('');
    setEntryName('');
    setEntryAddress('');
    setEntryAccountNo('');
    setEntryIfsc('');
    setEntryTreasuryCode('');
    setEntryPan('');
    setEntryGst('');
    setEntryTotalAmount('');
    setEntryDeductITax(true);
    setEntryITaxPercent('1');
    setEntryDeductGst(false);
    setEntryGstPercent('2');
    setEntrySubVouchers([]);
    setShowSubVoucherSection(false);
    setEditingSubVoucherId(null);
    setIsMemoPayeeListFullScreen(false);
    setSubVoucherNoInput('');
    setSubVoucherDescInput('');
    setSubVoucherAmountInput('');
  };

  const handleNotificationUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const fileList = Array.from(files);
    for (const file of fileList) {
      const type = 'notification';
      setUploadStatus(prev => ({
        ...prev,
        [type]: { isUploading: true, progress: 0, fileName: file.name, transferred: 0, total: file.size, error: null }
      }));

      try {
        let base64Data = '';
        if (file.size < 800 * 1024) {
          try {
            base64Data = await readFileAsDataUrl(file);
          } catch (readErr) {
            console.error("Error reading file:", readErr);
          }
        }

        let downloadURL = '';

        try {
          const storagePath = `notifications/${Date.now()}_${file.name}`;
          const storageRef = ref(storage, storagePath);
          const uploadTask = uploadBytesResumable(storageRef, file);

          setUploadTasks(prev => ({ ...prev, [type]: uploadTask }));

          await new Promise<void>((resolve, reject) => {
            uploadTask.on('state_changed',
              (snapshot) => {
                const progress = (snapshot.bytesTransferred / snapshot.totalBytes) * 100;
                setUploadStatus(prev => ({
                  ...prev,
                  [type]: { ...prev[type], progress, transferred: snapshot.bytesTransferred, total: snapshot.totalBytes }
                }));
              },
              (error) => {
                console.warn("Storage upload warning:", error);
                reject(error);
              },
              async () => {
                try {
                  const remoteUrl = await getDownloadURL(uploadTask.snapshot.ref);
                  downloadURL = remoteUrl;
                  resolve();
                } catch (err) {
                  console.warn("Could not get remote download URL:", err);
                  reject(err);
                }
              }
            );
          });
        } catch (storageErr) {
          console.warn("Storage upload error:", storageErr);
        }

        let savedFileData = '';
        if (downloadURL && downloadURL.startsWith('http')) {
          savedFileData = '';
        } else if (base64Data && base64Data.length < 800000) {
          downloadURL = base64Data;
          savedFileData = base64Data;
        } else {
          throw new Error(`Notification upload failed. Storage upload did not complete and file size (${(file.size / (1024 * 1024)).toFixed(2)} MB) exceeds inline document limit.`);
        }

        await addDoc(collection(db, 'notifications'), {
          name: file.name,
          url: downloadURL,
          fileData: savedFileData,
          type: file.type || 'application/pdf',
          category: 'general',
          description: 'Official Notification / Circular Document',
          targetRanges: ['All'],
          createdAt: Date.now(),
          uploadedBy: user?.uid
        });

        setUploadStatus(prev => ({ ...prev, [type]: { ...prev[type], isUploading: false, progress: 100, error: null } }));
      } catch (error) {
        console.error("Notification upload error:", error);
        showAlert("Failed to upload notification file.");
      } finally {
        setUploadTasks(prev => {
          const newTasks = { ...prev };
          delete newTasks[type];
          return newTasks;
        });
      }
    }
    e.target.value = '';
  };

  const handleDeleteNotification = async (notif: Notification) => {
    if (!isAdmin()) return;
    showConfirm(`Are you sure you want to delete "${notif.name}"?`, async () => {
      try {
        await deleteDoc(doc(db, 'notifications', notif.id));
        // Optionally delete from storage too
        try {
          const fileRef = ref(storage, notif.url);
          await deleteObject(fileRef);
        } catch (e) {
          console.warn("Could not delete file from storage:", e);
        }
        showAlert("Notification deleted successfully.");
      } catch (error) {
        handleFirestoreError(error, OperationType.DELETE, 'notifications');
      }
    });
  };

  const renderNotificationsTab = () => {
    const filteredNotifications = notifications.filter(n => {
      // Visibility Check for non-admins
      if (userRole !== 'admin' && userRole !== 'deo' && userRole !== 'approver' && userRole !== 'DA') {
        if (n.targetRanges && n.targetRanges.length > 0 && !n.targetRanges.includes('All')) {
          const userRangeObj = ranges.find(r => r.id === userRangeId);
          const matches = n.targetRanges.includes(userRole || '') ||
                          (userRangeId && n.targetRanges.includes(userRangeId)) ||
                          (userRangeObj && n.targetRanges.includes(userRangeObj.name));
          if (!matches) return false;
        }
      }

      // Category filter check
      if (notifCategoryFilter !== 'All') {
        if (notifCategoryFilter === 'general' && n.category && n.category !== 'general') return false;
        if (notifCategoryFilter === 'budget_allocation' && n.category !== 'budget_allocation') return false;
        if (notifCategoryFilter === 'budget_distribution' && n.category !== 'budget_distribution') return false;
      }

      const query = notifSearchTerm.toLowerCase();
      return n.name.toLowerCase().includes(query) ||
             (n.description || '').toLowerCase().includes(query) ||
             (n.schemeName || '').toLowerCase().includes(query) ||
             (n.sectorName || '').toLowerCase().includes(query) ||
             (n.rangeName || '').toLowerCase().includes(query);
    });

    const paginatedNotifications = notifItemsPerPage === 'All' 
      ? filteredNotifications 
      : filteredNotifications.slice((notifPage - 1) * notifItemsPerPage, notifPage * notifItemsPerPage);

    const countAll = notifications.length;
    const countCirculars = notifications.filter(n => !n.category || n.category === 'general').length;
    const countAllocations = notifications.filter(n => n.category === 'budget_allocation').length;
    const countDistributions = notifications.filter(n => n.category === 'budget_distribution').length;

    return (
      <div className="space-y-6">
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
            <div>
              <div className="flex items-center gap-2">
                <div className="p-2 bg-emerald-600 text-white rounded-lg shadow-xs">
                  <Bell className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-gray-900">Notifications & Circulars Hub</h2>
                  <p className="text-xs text-gray-500">Live notifications for budget distributions, allocations, and official departmental circulars.</p>
                </div>
              </div>
            </div>
            <div className="flex flex-col sm:flex-row items-center gap-3 w-full md:w-auto">
              <div className="flex items-center gap-1 w-full sm:w-auto">
                <div className="relative w-full sm:w-64">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search notifications, schemes..."
                    value={notifSearchTerm}
                    onChange={(e) => { setNotifSearchTerm(e.target.value); setNotifPage(1); }}
                    className="pl-9 pr-4 py-2 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 w-full"
                  />
                  {notifSearchTerm && (
                    <button
                      onClick={() => setNotifSearchTerm('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>
              </div>
              {isAdmin() && (
                <div className="flex items-center gap-4 w-full sm:w-auto">
                  {uploadStatus['notification']?.isUploading && (
                    <div className="flex items-center gap-3 bg-emerald-50 px-4 py-2 rounded-lg border border-emerald-100">
                      <div className="flex flex-col">
                        <span className="text-[10px] font-bold text-emerald-700 uppercase">Uploading...</span>
                        <span className="text-xs font-medium text-emerald-600 truncate max-w-[150px]">{uploadStatus['notification'].fileName}</span>
                      </div>
                      <div className="w-24 h-1.5 bg-emerald-200 rounded-full overflow-hidden">
                        <div className="h-full bg-emerald-600 transition-all duration-300" style={{ width: `${uploadStatus['notification'].progress}%` }}></div>
                      </div>
                      <button 
                        onClick={() => uploadTasks['notification']?.cancel()}
                        className="p-1 hover:bg-emerald-100 rounded-full text-emerald-600 cursor-pointer"
                        title="Cancel Upload"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>
                  )}
                  <label className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-medium transition-all cursor-pointer shadow-sm w-full sm:w-auto active:scale-95 text-sm">
                    <Plus className="w-4 h-4" />
                    Upload Notice / Circular
                    <input type="file" multiple accept="image/*,application/pdf" className="hidden" onChange={handleNotificationUpload} />
                  </label>
                </div>
              )}
            </div>
          </div>

          {/* Category Filter Chips */}
          <div className="flex flex-wrap items-center gap-2 mb-4 pb-3 border-b border-gray-100">
            <button
              onClick={() => { setNotifCategoryFilter('All'); setNotifPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                notifCategoryFilter === 'All'
                  ? 'bg-gray-900 text-white shadow-xs'
                  : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
              }`}
            >
              <Bell className="w-3.5 h-3.5" />
              <span>All Notifications</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${notifCategoryFilter === 'All' ? 'bg-white/20 text-white' : 'bg-gray-200 text-gray-700'}`}>
                {countAll}
              </span>
            </button>

            <button
              onClick={() => { setNotifCategoryFilter('general'); setNotifPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                notifCategoryFilter === 'general'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
              }`}
            >
              <FileText className="w-3.5 h-3.5" />
              <span>Notices & Circulars</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${notifCategoryFilter === 'general' ? 'bg-white/20 text-white' : 'bg-blue-100 text-blue-800'}`}>
                {countCirculars}
              </span>
            </button>

            <button
              onClick={() => { setNotifCategoryFilter('budget_allocation'); setNotifPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                notifCategoryFilter === 'budget_allocation'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
              }`}
            >
              <DollarSign className="w-3.5 h-3.5" />
              <span>Budget Allocations</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${notifCategoryFilter === 'budget_allocation' ? 'bg-white/20 text-white' : 'bg-emerald-100 text-emerald-800'}`}>
                {countAllocations}
              </span>
            </button>

            <button
              onClick={() => { setNotifCategoryFilter('budget_distribution'); setNotifPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                notifCategoryFilter === 'budget_distribution'
                  ? 'bg-indigo-600 text-white shadow-xs'
                  : 'bg-indigo-50 text-indigo-700 hover:bg-indigo-100'
              }`}
            >
              <FileBarChart className="w-3.5 h-3.5" />
              <span>Budget Distribution Allotments</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] ${notifCategoryFilter === 'budget_distribution' ? 'bg-white/20 text-white' : 'bg-indigo-100 text-indigo-800'}`}>
                {countDistributions}
              </span>
            </button>
          </div>

          <div className="bg-white rounded-xl border border-gray-100 overflow-hidden shadow-2xs">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50/75 border-b border-gray-200">
                    <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider">Document / Update Details</th>
                    <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider">Category</th>
                    <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider">Date & Time</th>
                    <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider">Visible To</th>
                    <th className="px-4 py-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider text-right">Quick Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {paginatedNotifications.length > 0 ? (
                    paginatedNotifications.map((notif) => {
                      const isPdfOrFile = Boolean(notif.url || notif.fileData);
                      const isAllocation = notif.category === 'budget_allocation' || notif.type?.includes('allocation');
                      const isDistribution = notif.category === 'budget_distribution';

                      return (
                        <tr 
                          key={notif.id} 
                          className="hover:bg-slate-50/80 transition-colors group cursor-pointer"
                          onClick={() => {
                            if (isPdfOrFile) {
                              handleViewFile(notif);
                            } else if (isAllocation) {
                              setActiveTab('Allocations');
                              setSearchTerm(notif.schemeName || '');
                              window.scrollTo({ top: 0, behavior: 'smooth' });
                            } else if (isDistribution) {
                              setActiveTab('Distributed Budget');
                              window.scrollTo({ top: 0, behavior: 'smooth' });
                            }
                          }}
                        >
                          <td className="px-4 py-3">
                            <div className="flex items-start gap-3">
                              <div className={`p-2 rounded-lg shrink-0 mt-0.5 ${
                                isAllocation 
                                  ? 'bg-emerald-100 text-emerald-700' 
                                  : isDistribution 
                                  ? 'bg-indigo-100 text-indigo-700' 
                                  : 'bg-blue-100 text-blue-700'
                              }`}>
                                {isAllocation ? (
                                  <DollarSign className="w-4 h-4" />
                                ) : isDistribution ? (
                                  <FileBarChart className="w-4 h-4" />
                                ) : (
                                  <FileText className="w-4 h-4" />
                                )}
                              </div>
                              <div className="min-w-0">
                                <div className="text-sm font-bold text-gray-900 group-hover:text-emerald-700 transition-colors flex items-center gap-2">
                                  <span>{notif.name}</span>
                                  {isPdfOrFile && (
                                    <span className="text-[9px] bg-red-50 text-red-700 border border-red-200 px-1 py-0.2 rounded font-bold uppercase">
                                      PDF
                                    </span>
                                  )}
                                </div>
                                {notif.description && (
                                  <p className="text-xs text-gray-500 mt-0.5 line-clamp-1">{notif.description}</p>
                                )}
                                {(notif.schemeName || notif.amount) && (
                                  <div className="flex flex-wrap items-center gap-2 mt-1 text-[11px] text-gray-600">
                                    {notif.amount && (
                                      <span className="font-bold text-emerald-800 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200">
                                        Amount: â‚¹{Number(notif.amount).toLocaleString('en-IN')}
                                      </span>
                                    )}
                                    {notif.schemeName && (
                                      <span className="font-medium bg-gray-100 px-1.5 py-0.5 rounded text-gray-700">
                                        Scheme: {notif.schemeName}
                                      </span>
                                    )}
                                    {notif.rangeName && (
                                      <span className="font-medium bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded border border-indigo-100">
                                        Range: {notif.rangeName}
                                      </span>
                                    )}
                                  </div>
                                )}
                              </div>
                            </div>
                          </td>

                          <td className="px-4 py-3">
                            {isAllocation ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                                <DollarSign className="w-3 h-3" /> Allocation
                              </span>
                            ) : isDistribution ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-bold bg-indigo-100 text-indigo-800 border border-indigo-200">
                                <FileBarChart className="w-3 h-3" /> Distribution
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-bold bg-blue-100 text-blue-800 border border-blue-200">
                                <FileText className="w-3 h-3" /> Notice / Circular
                              </span>
                            )}
                          </td>

                          <td className="px-4 py-3 text-xs text-gray-500 whitespace-nowrap">
                            {notif.createdAt ? new Date(notif.createdAt).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : 'N/A'}
                          </td>

                          <td className="px-4 py-3" onClick={(e) => e.stopPropagation()}>
                            {isAdmin() ? (
                              <div className="flex flex-wrap gap-1 max-w-xs">
                                {['All', 'Narag', 'Sarahan', 'Rajgarh', 'Habban', 'Division'].map(rangeName => {
                                  const currentRanges = notif.targetRanges || ['All'];
                                  const isChecked = currentRanges.includes(rangeName);
                                  return (
                                    <button
                                      key={rangeName}
                                      type="button"
                                      onClick={async () => {
                                        let updated: string[];
                                        if (rangeName === 'All') {
                                          updated = isChecked ? [] : ['All'];
                                        } else {
                                          let filtered = currentRanges.filter(r => r !== 'All');
                                          if (isChecked) {
                                            updated = filtered.filter(r => r !== rangeName);
                                          } else {
                                            updated = [...filtered, rangeName];
                                          }
                                          if (updated.length === 0) updated = ['All'];
                                        }
                                        try {
                                          await updateDoc(doc(db, 'notifications', notif.id), { targetRanges: updated });
                                        } catch (e) {
                                          console.error("Error updating notification target ranges:", e);
                                        }
                                      }}
                                      className={`px-1.5 py-0.5 rounded text-[10px] font-bold transition-all border cursor-pointer ${
                                        isChecked 
                                          ? 'bg-emerald-100 text-emerald-800 border-emerald-300' 
                                          : 'bg-gray-100 text-gray-400 border-gray-200 hover:bg-gray-200'
                                      }`}
                                      title={`Click to toggle visibility for ${rangeName}`}
                                    >
                                      {rangeName}
                                    </button>
                                  );
                                })}
                              </div>
                            ) : (
                              <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-100">
                                {(notif.targetRanges && notif.targetRanges.length > 0) ? notif.targetRanges.join(', ') : 'All Ranges'}
                              </span>
                            )}
                          </td>

                          <td className="px-4 py-3 text-right" onClick={(e) => e.stopPropagation()}>
                            <div className="flex justify-end items-center gap-1.5">
                              {isPdfOrFile && (
                                <>
                                  <button 
                                    type="button"
                                    onClick={() => handleViewFile(notif)}
                                    className="px-2.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg text-xs font-semibold flex items-center gap-1 transition-colors cursor-pointer border border-emerald-200"
                                    title="View PDF Document"
                                  >
                                    <Eye className="w-3.5 h-3.5" />
                                    <span>View PDF</span>
                                  </button>
                                  <button 
                                    type="button"
                                    onClick={() => handleDownloadFile(notif)}
                                    className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer border border-transparent hover:border-blue-200"
                                    title="Download File"
                                  >
                                    <Download className="w-4 h-4" />
                                  </button>
                                </>
                              )}

                              {isAllocation && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setActiveTab('Allocations');
                                    setSearchTerm(notif.schemeName || '');
                                    window.scrollTo({ top: 0, behavior: 'smooth' });
                                  }}
                                  className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1 transition-colors cursor-pointer shadow-xs"
                                  title="Go to Budget Allocation Tab"
                                >
                                  <span>View in Allocation</span>
                                  <ArrowRight className="w-3.5 h-3.5" />
                                </button>
                              )}

                              {isDistribution && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setActiveTab('Distributed Budget');
                                    window.scrollTo({ top: 0, behavior: 'smooth' });
                                  }}
                                  className="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1 transition-colors cursor-pointer shadow-xs"
                                  title="Go to Distributed Budget Tab"
                                >
                                  <span>Go to Distribution</span>
                                  <ArrowRight className="w-3.5 h-3.5" />
                                </button>
                              )}

                              {isAdmin() && (
                                <button 
                                  onClick={() => handleDeleteNotification(notif)}
                                  className="p-1.5 text-red-500 hover:text-red-700 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                                  title="Delete Notification"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={5} className="px-4 py-12 text-center text-gray-400 italic text-sm">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <Bell className="w-8 h-8 text-gray-300" />
                          <p>No notifications found matching your filter criteria.</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {filteredNotifications.length > 0 && (
              <div className="p-4 bg-gray-50 border-t border-gray-100">
                <Pagination
                  totalItems={filteredNotifications.length}
                  itemsPerPage={notifItemsPerPage}
                  currentPage={notifPage}
                  onPageChange={setNotifPage}
                  onItemsPerPageChange={setNotifItemsPerPage}
                />
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };
  
  const handleDelete = async (collectionName: string, id: string) => {
    if (collectionName === 'allocations' && isFeatureLocked('Allocation')) {
      showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
      return;
    }
    if (collectionName === 'expenditures' && isFeatureLocked('Expenditure')) {
      showAlert("This feature is currently locked by Admin. Please contact Admin for permission.");
      return;
    }
    showConfirm(`Are you sure you want to delete this ${collectionName.slice(0, -1)}?`, async () => {
      try {
        await deleteDoc(doc(db, collectionName, id));
        logAuditAction('Item Deleted', `Collection: ${collectionName}, ID: ${id}`);
        showAlert("Deleted successfully.");
        if (editingItem?.item?.id === id) {
          setEditingItem(null);
        }
      } catch (error) {
        handleFirestoreError(error, OperationType.DELETE, collectionName);
      }
    });
  };

  const handleCreateBill = async (e: any) => {
    e.preventDefault();
    const billNo = e.target.billNo.value;
    const billDate = e.target.billDate.value;
    const remarks = e.target.remarks.value;
    const isFinalizing = e.nativeEvent.submitter?.name === 'finalize';

    if (!billNo || !billDate || selectedExpensesForBill.length === 0) {
      showAlert('Please provide Bill No, Date and select at least one expenditure.');
      return;
    }

    const selectedExpObjects = expenses.filter(ex => selectedExpensesForBill.includes(ex.id));
    
    // Validate same SOE
    const soeIds = new Set(selectedExpObjects.map(ex => ex.soeId));
    if (soeIds.size > 1) {
      showAlert('A bill can only contain expenditures from a single SOE head.');
      return;
    }

    const totalAmount = selectedExpObjects.reduce((sum, ex) => sum + ex.amount, 0);
    const activeFy = fys.find(f => f.name === selectedFY || f.id === selectedFY);
    const currentFyId = activeFy?.id || selectedFY;

    // Check for duplicate bill number
    const isDuplicate = bills.some(b => 
      b.billNo === billNo && 
      b.fyId === currentFyId && 
      (editingItem?.type === 'Bill' ? b.id !== editingItem.item.id : true)
    );

    if (isDuplicate) {
      showAlert(`A bill with number "${billNo}" already exists for the selected financial year.`);
      return;
    }

    try {
      const billData = {
        billNo,
        billDate,
        expenseIds: selectedExpensesForBill,
        fyId: currentFyId,
        financialYear: activeFy?.name || selectedFY,
        totalAmount,
        status: isFinalizing ? 'finalized' : (editingItem?.type === 'Bill' ? editingItem.item.status : 'draft'),
        remarks: remarks || '',
        updatedAt: Date.now(),
        createdBy: editingItem?.type === 'Bill' ? editingItem.item.createdBy : user.uid
      };

      if (editingItem?.type === 'Bill') {
        await updateDoc(doc(db, 'bills', editingItem.item.id), billData);
        setEditingItem(null);
      } else {
        await addDoc(collection(db, 'bills'), { ...billData, createdAt: Date.now() });
      }
      setSelectedExpensesForBill([]);
      setBillExpFilters({ schemeId: '', sectorId: '', activityId: '', subActivityId: '', rangeId: '', soeId: '' });
      e.target.reset();
    } catch (error) {
      handleFirestoreError(error, editingItem?.type === 'Bill' ? OperationType.UPDATE : OperationType.CREATE, 'bills');
    }
  };

  const handleRemoveExpenseFromBill = async (billId: string, expenseId: string) => {
    const bill = bills.find(b => b.id === billId);
    if (!bill) return;

    const newExpenseIds = bill.expenseIds.filter(id => id !== expenseId);
    if (newExpenseIds.length === 0) {
      showAlert('A bill must have at least one expenditure. Delete the bill instead.');
      return;
    }

    const newTotalAmount = expenses
      .filter(e => newExpenseIds.includes(e.id))
      .reduce((sum, e) => sum + e.amount, 0);

    try {
      await updateDoc(doc(db, 'bills', billId), {
        expenseIds: newExpenseIds,
        totalAmount: newTotalAmount,
        updatedAt: Date.now()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'bills');
    }
  };

  const generateBillPdf = async (bill: Bill) => {
    const doc = new jsPDF();
    const activeFy = fys.find(f => f.id === bill.fyId || f.name === bill.financialYear);
    
    // Generate QR Code data URL for the bill
    const verificationUrl = `${window.location.origin}${window.location.pathname}?verifyBill=${encodeURIComponent(bill.id)}`;
    let qrDataUrl = '';
    try {
      qrDataUrl = await QRCode.toDataURL(verificationUrl, {
        margin: 1,
        width: 140,
        color: {
          dark: '#064e3b',
          light: '#ffffff'
        }
      });
    } catch (e) {
      console.error('Error creating QR for bill PDF:', e);
    }

    // Header
    doc.setFontSize(16);
    doc.text('TREASURY BILL', 105, 14, { align: 'center' });
    doc.setFontSize(9.5);
    doc.text(`Financial Year: ${activeFy?.name || bill.financialYear}`, 105, 20, { align: 'center' });
    doc.setFontSize(9);
    doc.text('H.P. Forest Department â€¢ Rajgarh Forest Division', 105, 25, { align: 'center' });

    // Place QR Code at top right if generated
    if (qrDataUrl) {
      doc.addImage(qrDataUrl, 'PNG', 165, 8, 28, 28);
      doc.setFontSize(6.5);
      doc.setTextColor(5, 150, 105);
      doc.text('Scan for Live Status', 179, 39, { align: 'center' });
      doc.setTextColor(0, 0, 0);
    }
    
    doc.setFontSize(10.5);
    doc.text(`Bill No: ${bill.billNo}`, 15, 33);
    doc.text(`Bill Date: ${bill.billDate ? bill.billDate.split('-').reverse().join('/') : 'N/A'}`, 15, 39);
    doc.text(`Status: ${bill.status.toUpperCase()}`, 15, 45);
    if (bill.remarks) {
      doc.text(`Remarks: ${bill.remarks}`, 15, 51);
    }

    const billExpenses = expenses.filter(e => bill.expenseIds.includes(e.id));
    
    const tableData = billExpenses.map((exp, index) => {
      const s = soes.find(s => s.id === exp.soeId);
      const al = allocations.find(a => a.id === exp.allocationId);
      const r = ranges.find(r => r.id === al?.rangeId);
      
      let hierarchy = '';
      if (al?.subActivityId) {
        const sa = subActivities.find(sa => sa.id === al.subActivityId);
        const act = activities.find(a => a.id === sa?.activityId);
        const sec = sectors.find(sec => sec.id === act?.sectorId);
        const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
        hierarchy = [sch?.name, sec?.name, act?.name, sa?.name].filter(Boolean).join(' -> ');
      } else if (al?.activityId) {
        const act = activities.find(a => a.id === al.activityId);
        const sec = sectors.find(sec => sec.id === act?.sectorId);
        const sch = schemes.find(sc => sc.id === (sec ? sec.schemeId : act?.schemeId));
        hierarchy = [sch?.name, sec?.name, act?.name].filter(Boolean).join(' -> ');
      }

      return [
        index + 1,
        exp.date ? exp.date.split('-').reverse().join('/') : 'N/A',
        r?.name || 'N/A',
        s?.name || 'N/A',
        hierarchy,
        exp.description,
        `Rs. ${exp.amount.toLocaleString()}`
      ];
    });

    autoTable(doc, {
      startY: bill.remarks ? 56 : 50,
      head: [['SrNo', 'Date', 'Range', 'SOE', 'Hierarchy', 'Description', 'Amount']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: [16, 185, 129] },
      styles: { fontSize: 8 },
      columnStyles: {
        6: { halign: 'right', fontStyle: 'bold' }
      }
    });

    const finalY = (doc as any).lastAutoTable.finalY || 70;
    doc.setFontSize(11);
    doc.text(`Total Amount: Rs. ${bill.totalAmount.toLocaleString()}`, 195, finalY + 8, { align: 'right' });

    // Official Security verification footer
    doc.setFontSize(7);
    doc.setTextColor(120, 120, 120);
    doc.text(`Official Digital Record â€¢ Rajgarh Forest Division â€¢ Verified with Live QR System â€¢ Printed on ${new Date().toLocaleDateString('en-IN')}`, 105, finalY + 18, { align: 'center' });
    doc.setTextColor(0, 0, 0);

    return doc;
  };

  const handleDownloadBill = async (bill: Bill) => {
    try {
      const doc = await generateBillPdf(bill);
      doc.save(`bill_${bill.billNo}.pdf`);
    } catch (err) {
      console.error('Error downloading bill PDF:', err);
      showAlert('Failed to generate PDF for download.');
    }
  };

  const handleViewBill = async (bill: Bill) => {
    try {
      const doc = await generateBillPdf(bill);
      const blob = doc.output('blob');
      const url = URL.createObjectURL(blob);
      setViewingBillPdf({ url, bill });
    } catch (err) {
      console.error('Error viewing bill PDF:', err);
      showAlert('Failed to open PDF preview.');
    }
  };

  // Open Bill QR Code Dialog for Field Officers
  const handleOpenBillQR = async (bill: Bill) => {
    try {
      const verificationUrl = `${window.location.origin}${window.location.pathname}?verifyBill=${encodeURIComponent(bill.id)}`;
      const qrDataUrl = await QRCode.toDataURL(verificationUrl, {
        width: 360,
        margin: 1,
        color: {
          dark: '#064e3b',
          light: '#ffffff'
        },
        errorCorrectionLevel: 'M'
      });
      setViewingBillQR({ bill, qrDataUrl });
      setCopiedQRLink(false);
    } catch (err) {
      console.error('Error generating bill QR code:', err);
      showAlert('Failed to generate QR code for this bill.');
    }
  };

  // Download Bill QR Image
  const handleDownloadBillQR = (bill: Bill, qrDataUrl: string) => {
    const link = document.createElement('a');
    link.href = qrDataUrl;
    link.download = `Bill_QR_${bill.billNo.replace(/[^a-zA-Z0-9_-]/g, '_')}.png`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Copy Bill Mobile Verification URL
  const handleCopyBillVerificationLink = (bill: Bill) => {
    const url = `${window.location.origin}${window.location.pathname}?verifyBill=${encodeURIComponent(bill.id)}`;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(url).then(() => {
        setCopiedQRLink(true);
        setTimeout(() => setCopiedQRLink(false), 2500);
      }).catch(() => {
        showAlert(`Verification Link:\n${url}`);
      });
    } else {
      showAlert(`Verification Link:\n${url}`);
    }
  };

  // Print Bill QR Verification Pass / Slip
  const handlePrintBillQRPad = (bill: Bill, qrDataUrl: string) => {
    const printWin = window.open('', '_blank');
    if (!printWin) {
      showAlert('Please allow popups to print the QR verification slip.');
      return;
    }
    const billExpenses = expenses.filter(e => bill.expenseIds.includes(e.id));
    const firstExp = billExpenses[0];
    const al = allocations.find(a => a.id === firstExp?.allocationId);
    const rangeObj = ranges.find(r => r.id === al?.rangeId);
    const soeObj = soes.find(s => s.id === firstExp?.soeId);

    printWin.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Bill Verification QR Slip - ${bill.billNo}</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 24px; color: #111; background: #fff; }
            .ticket { border: 2px dashed #059669; border-radius: 14px; padding: 24px; max-width: 400px; margin: 0 auto; text-align: center; background: #fcfdfd; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
            .dept { font-size: 11px; font-weight: 800; color: #047857; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 2px; }
            .title { font-size: 16px; font-weight: 800; color: #111827; margin-bottom: 12px; }
            .qr-wrap { background: #fff; padding: 8px; border: 2px solid #e5e7eb; border-radius: 12px; display: inline-block; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
            .qr-img { width: 170px; height: 170px; display: block; }
            .badge { display: inline-block; padding: 3px 12px; border-radius: 9999px; font-size: 11px; font-weight: 800; text-transform: uppercase; margin: 10px 0 6px; }
            .badge-finalized { background: #d1fae5; color: #065f46; border: 1px solid #a7f3d0; }
            .badge-draft { background: #dbeafe; color: #1e40af; border: 1px solid #bfdbfe; }
            .info-table { width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; margin-top: 10px; background: #fff; border-radius: 8px; overflow: hidden; border: 1px solid #e5e7eb; }
            .info-table td { padding: 6px 10px; border-bottom: 1px solid #f3f4f6; }
            .info-table td.label { font-weight: bold; color: #4b5563; width: 42%; }
            .info-table td.val { font-weight: 600; color: #111827; }
            .footer { font-size: 9.5px; color: #6b7280; margin-top: 14px; line-height: 1.4; border-top: 1px solid #e5e7eb; padding-top: 8px; }
            @media print {
              body { padding: 0; }
            }
          </style>
        </head>
        <body>
          <div class="ticket">
            <div class="dept">H.P. Forest Department â€¢ Rajgarh Division</div>
            <div class="title">Bill Verification QR Pass</div>
            <div class="qr-wrap">
              <img src="${qrDataUrl}" class="qr-img" alt="QR Code" />
            </div>
            <div>
              <span class="badge badge-${bill.status}">${bill.status}</span>
            </div>
            <div style="font-size: 10.5px; font-weight: bold; color: #047857; margin-bottom: 6px;">Scan with any mobile camera to verify details</div>
            <table class="info-table">
              <tr>
                <td class="label">Bill Number:</td>
                <td class="val" style="font-family: monospace; font-size: 13px; color: #047857;">${bill.billNo}</td>
              </tr>
              <tr>
                <td class="label">Bill Date:</td>
                <td class="val">${bill.billDate ? bill.billDate.split('-').reverse().join('/') : 'N/A'}</td>
              </tr>
              <tr>
                <td class="label">Range / Office:</td>
                <td class="val">${rangeObj?.name || 'Division'}</td>
              </tr>
              <tr>
                <td class="label">SOE Head:</td>
                <td class="val">${soeObj?.name || 'N/A'}</td>
              </tr>
              <tr>
                <td class="label">Total Amount:</td>
                <td class="val" style="font-size: 13px; color: #059669; font-weight: 800;">â‚¹${bill.totalAmount.toLocaleString('en-IN')}</td>
              </tr>
              <tr>
                <td class="label">Voucher Count:</td>
                <td class="val">${billExpenses.length} entries</td>
              </tr>
            </table>
            <div class="footer">
              Forest Budget Control System â€¢ Live Field Verification<br/>
              Printed on: ${new Date().toLocaleString('en-IN')}
            </div>
          </div>
          <script>
            window.onload = function() {
              window.focus();
              window.print();
            };
          </script>
        </body>
      </html>
    `);
    printWin.document.close();
  };

  const handleUserRoleChange = async (userId: string, newRole: 'admin' | 'deo' | 'approver' | 'Sarahan' | 'Narag' | 'Habban' | 'Division' | 'Rajgarh') => {
    try {
      await updateDoc(doc(db, 'users', userId), { role: newRole, updatedAt: Date.now() });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'users');
    }
  };

  const handleDeleteUser = (userId: string) => {
    showConfirm('Delete this user access?', async () => {
      try {
        await deleteDoc(doc(db, 'users', userId));
      } catch (error) {
        handleFirestoreError(error, OperationType.DELETE, 'users');
      }
    });
  };

  const handleCreateNewUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newUserEmail || !newUserPassword) return;
    
    let emailToUse = newUserEmail.trim();
    // If user enters an 8-digit number or any ID without @, append the default domain
    if (!emailToUse.includes('@')) {
      emailToUse = `${emailToUse}@rajgarhforest.app`;
    }

    try {
      // Initialize a secondary app to create user without logging out the admin
      const secondaryApp = getApps().find(app => app.name === "Secondary") || initializeApp(firebaseConfig, "Secondary");
      const secondaryAuth = getAuth(secondaryApp);
      
      const userCredential = await createUserWithEmailAndPassword(secondaryAuth, emailToUse, newUserPassword);
      
      // Add user to firestore
      try {
        await setDoc(doc(db, 'users', userCredential.user.uid), {
          email: emailToUse,
          role: newUserRole,
          password: newUserPassword
        });
      } catch (firestoreError) {
        handleFirestoreError(firestoreError, OperationType.CREATE, 'users');
      }
      
      // Sign out the secondary app
      await secondaryAuth.signOut();
      
      setNewUserEmail('');
      setNewUserPassword('');
      setNewUserRole('deo');
      showAlert('User created successfully!');
    } catch (error: any) {
      if (error.code === 'auth/email-already-in-use') {
        showAlert(`Error: This User ID / Email already exists in the system. If you deleted them previously, they still exist in the authentication database. You cannot recreate them with the same ID.`);
      } else {
        showAlert(`Error creating user: ${error.message}`);
      }
    }
  };

  const handleUpdatePassword = async (userId: string) => {
    if (!newPasswordInput) return;
    try {
      await updateDoc(doc(db, 'users', userId), {
        password: newPasswordInput,
        updatedAt: Date.now()
      });
      setEditingPasswordId(null);
      setNewPasswordInput('');
      showAlert('Password updated in system records. Note: This does not change the actual login password in Firebase Auth. The user should use the forgot password link if they cannot log in.');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'users');
    }
  };

  const handleResetPassword = (email: string) => {
    showConfirm(`Send password reset email to ${email}?`, async () => {
      try {
        await sendPasswordResetEmail(auth, email);
        showAlert('Password reset email sent!');
      } catch (error: any) {
        showAlert(`Error sending reset email: ${error.message}`);
      }
    });
  };

  const handleUpdateMaxSessions = async (userId: string, maxSessions: number) => {
    try {
      await updateDoc(doc(db, 'users', userId), { maxSessions, updatedAt: Date.now() });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'users');
    }
  };

  const handleClearSessions = async (userId: string) => {
    try {
      await updateDoc(doc(db, 'users', userId), { activeSessions: [], updatedAt: Date.now() });
      showAlert('All active sessions cleared for this user.');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'users');
    }
  };

  const handleUnsyncExpense = async (expenseId: string) => {
    showConfirm('Remove memo sync link from this expenditure entry?', async () => {
      try {
        await updateDoc(doc(db, 'expenditures', expenseId), {
          syncedMemoId: null,
          syncedMemoNo: null,
          updatedAt: Date.now()
        });
        showAlert('Memo sync link removed successfully.');
      } catch (error) {
        handleFirestoreError(error, OperationType.UPDATE, `expenditures/${expenseId}`);
      }
    });
  };

  const handleClearAllMemoSyncs = async () => {
    showConfirm('Are you sure you want to delete/clear all memo sync references across all expenditure entries?', async () => {
      try {
        const syncedExpSnap = await getDocs(query(collection(db, 'expenditures'), where('syncedMemoId', '!=', null)));
        if (syncedExpSnap.empty) {
          showAlert('No synced expenditure entries found.');
          return;
        }
        const batch = writeBatch(db);
        syncedExpSnap.docs.forEach(d => {
          batch.update(doc(db, 'expenditures', d.id), {
            syncedMemoId: null,
            syncedMemoNo: null,
            updatedAt: Date.now()
          });
        });
        await batch.commit();
        showAlert(`Cleared sync references from ${syncedExpSnap.size} expenditure entries.`);
      } catch (error) {
        handleFirestoreError(error, OperationType.UPDATE, 'expenditures');
      }
    });
  };

  const handleToggleFeatureLock = async (feature: 'Allocation' | 'Expenditure' | 'Access' | 'Memo' | 'MemoSync', target: string) => {
    const existingLock = featureLocks.find(l => l.feature === feature && l.target === target);
    try {
      if (existingLock) {
        await updateDoc(doc(db, 'featureLocks', existingLock.id), {
          isLocked: !existingLock.isLocked,
          updatedBy: user?.email || 'Admin',
          updatedAt: Date.now()
        });
      } else {
        await addDoc(collection(db, 'featureLocks'), {
          feature,
          target,
          isLocked: true,
          updatedBy: user?.email || 'Admin',
          updatedAt: Date.now()
        });
      }
      logAuditAction('Feature Lock Toggled', `Feature: ${feature}, Target: ${target}, State: ${existingLock ? (!existingLock.isLocked ? 'Locked' : 'Unlocked') : 'Locked'}`);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'featureLocks');
    }
  };

  const handleToggleUserStatus = async (userId: string, isDisabled: boolean) => {
    try {
      await updateDoc(doc(db, 'users', userId), {
        isDisabled: !isDisabled,
        updatedAt: Date.now()
      });
      logAuditAction('User Status Changed', `User ID: ${userId}, Status: ${!isDisabled ? 'Disabled' : 'Enabled'}`);
      showAlert(`User ${!isDisabled ? 'disabled' : 'enabled'} successfully.`);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `users/${userId}`);
    }
  };

  const renderAuditLogTab = () => {
    // Distinct analytics for security audit overview
    const uniqueIps = new Set(auditLogs.map(l => l.ipAddress).filter(Boolean)).size;
    const uniqueLocations = new Set(auditLogs.map(l => l.ipLocation || l.ipCity).filter(Boolean)).size;
    const uniqueUsers = new Set(auditLogs.map(l => l.userEmail || l.userName).filter(Boolean)).size;
    const mobileCount = auditLogs.filter(l => (l.deviceType || "").includes("Mobile") || (l.deviceInfo || "").toLowerCase().includes("mobile")).length;

    const filteredLogs = auditLogs.filter(log => {
      // Role filter
      if (auditRoleFilter !== "all") {
        const roleLower = (log.userRole || "").toLowerCase();
        const rangeLower = (log.userRange || "").toLowerCase();
        const target = auditRoleFilter.toLowerCase();
        if (target === "admin" && roleLower !== "admin") return false;
        if (target === "deo" && roleLower !== "deo") return false;
        if (target === "da" && roleLower !== "da" && roleLower !== "approver") return false;
        if (["sarahan", "narag", "habban", "division", "rajgarh"].includes(target)) {
          if (!roleLower.includes(target) && !rangeLower.includes(target)) return false;
        }
      }

      // Action category filter
      if (auditActionFilter !== "all") {
        const actLower = (log.action || "").toLowerCase();
        if (auditActionFilter === "login" && !actLower.includes("login") && !actLower.includes("session") && !actLower.includes("access")) return false;
        if (auditActionFilter === "budget" && !actLower.includes("budget") && !actLower.includes("allocation") && !actLower.includes("fy")) return false;
        if (auditActionFilter === "expense" && !actLower.includes("expenditure") && !actLower.includes("bill") && !actLower.includes("soe")) return false;
        if (auditActionFilter === "memo" && !actLower.includes("memo")) return false;
        if (auditActionFilter === "security" && !actLower.includes("lock") && !actLower.includes("user") && !actLower.includes("status")) return false;
      }

      // Search query across all fields including hardware and IP location
      if (!auditSearchTerm.trim()) return true;
      const q = auditSearchTerm.toLowerCase();
      return (
        (log.action && log.action.toLowerCase().includes(q)) ||
        (log.userName && log.userName.toLowerCase().includes(q)) ||
        (log.userEmail && log.userEmail.toLowerCase().includes(q)) ||
        (log.userRole && log.userRole.toLowerCase().includes(q)) ||
        (log.userRange && log.userRange.toLowerCase().includes(q)) ||
        (log.ipAddress && log.ipAddress.toLowerCase().includes(q)) ||
        (log.ipLocation && log.ipLocation.toLowerCase().includes(q)) ||
        (log.ipCity && log.ipCity.toLowerCase().includes(q)) ||
        (log.ipRegion && log.ipRegion.toLowerCase().includes(q)) ||
        (log.ipCountry && log.ipCountry.toLowerCase().includes(q)) ||
        (log.ipIsp && log.ipIsp.toLowerCase().includes(q)) ||
        (log.screenResolution && log.screenResolution.toLowerCase().includes(q)) ||
        (log.browserVersion && log.browserVersion.toLowerCase().includes(q)) ||
        (log.browser && log.browser.toLowerCase().includes(q)) ||
        (log.os && log.os.toLowerCase().includes(q)) ||
        (log.deviceType && log.deviceType.toLowerCase().includes(q)) ||
        (log.deviceInfo && log.deviceInfo.toLowerCase().includes(q)) ||
        (log.hardwareSummary && log.hardwareSummary.toLowerCase().includes(q)) ||
        (log.details && log.details.toLowerCase().includes(q))
      );
    });

    const exportAuditLogsCsv = () => {
      try {
        const headers = [
          "Timestamp (IST)",
          "User Name",
          "User Email",
          "Role",
          "Range",
          "Public IP Address",
          "Public IP Location",
          "ISP / Network",
          "Screen Resolution",
          "Browser & Version",
          "Operating System",
          "Device Category",
          "Hardware Details",
          "Action Performed",
          "Action Details"
        ];

        const rows = filteredLogs.map(log => {
          const dt = log.timestamp ? new Date(log.timestamp).toLocaleString("en-IN") : "N/A";
          return [
            `"${dt}"`,
            `"${(log.userName || "User").replace(/"/g, '""')}"`,
            `"${(log.userEmail || "").replace(/"/g, '""')}"`,
            `"${(log.userRole || "").replace(/"/g, '""')}"`,
            `"${(log.userRange || "").replace(/"/g, '""')}"`,
            `"${(log.ipAddress || "").replace(/"/g, '""')}"`,
            `"${(log.ipLocation || log.ipCity || "").replace(/"/g, '""')}"`,
            `"${(log.ipIsp || "").replace(/"/g, '""')}"`,
            `"${(log.screenResolution || "").replace(/"/g, '""')}"`,
            `"${(log.browserVersion || log.browser || "").replace(/"/g, '""')}"`,
            `"${(log.os || "").replace(/"/g, '""')}"`,
            `"${(log.deviceType || "").replace(/"/g, '""')}"`,
            `"${(log.hardwareSummary || log.deviceInfo || "").replace(/"/g, '""')}"`,
            `"${(log.action || "").replace(/"/g, '""')}"`,
            `"${(log.details || "").replace(/"/g, '""')}"`
          ].join(",");
        });

        const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows].join("\n");
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `Rajgarh_Forest_Security_Audit_Log_${new Date().toISOString().split("T")[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      } catch (err) {
        console.error("Export CSV error:", err);
      }
    };

    const currentHw = getDeviceHardwareProfile();

    return (
      <div className="bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-gray-100 space-y-5">
        {/* Header with Admin Privilege Badge */}
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b pb-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-emerald-700 text-white flex items-center justify-center shadow-sm">
                <Shield className="h-5 w-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="text-lg font-bold text-gray-900">
                    System Security, Device & IP Audit Trail
                  </h3>
                  <span className="bg-amber-100 text-amber-900 border border-amber-300 text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-full flex items-center gap-1 shadow-2xs">
                    <Lock className="w-2.5 h-2.5" /> Admin Only
                  </span>
                </div>
                <p className="text-xs text-gray-500 mt-0.5">
                  Real-time security telemetry recording hardware profiles, screen resolutions, browser versions, and public IP locations across all user sessions (Admin, DEO, DA, and Ranges) to identify unauthorized access.
                </p>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2 self-stretch sm:self-auto">
            <button
              onClick={() => {
                setIsRefreshingLogs(true);
                fetchClientNetworkLocation().then(prof => {
                  if (prof) {
                    setClientNetwork(prof);
                    setClientIpAddress(prof.ip || "Direct Connection");
                  }
                  setTimeout(() => setIsRefreshingLogs(false), 500);
                });
              }}
              className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg border border-slate-300 transition-colors"
              title="Refresh Network & Logs"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isRefreshingLogs ? "animate-spin" : ""}`} />
              <span className="hidden sm:inline">Refresh</span>
            </button>
            <button
              onClick={exportAuditLogsCsv}
              className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors"
              title="Export CSV Audit Trail"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export CSV</span>
            </button>
          </div>
        </div>

        {/* Live Admin Device & Geolocation Telemetry Diagnostics Banner */}
        <div className="bg-linear-to-r from-slate-900 via-slate-800 to-slate-900 text-white rounded-xl p-4 shadow-md border border-slate-700">
          <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border-b border-slate-700/80 pb-3 mb-3">
            <div className="flex items-center gap-2.5">
              <div className="relative flex items-center justify-center">
                <span className="absolute inline-flex h-3 w-3 animate-ping rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex h-2.5 w-2.5 rounded-full bg-emerald-500"></span>
              </div>
              <span className="font-bold text-xs tracking-wide uppercase text-slate-200 flex items-center gap-1.5">
                <Radio className="w-3.5 h-3.5 text-emerald-400" /> Active Session Hardware & IP Telemetry
              </span>
              <span className="text-[10px] bg-slate-700/80 text-emerald-300 font-mono px-2 py-0.5 rounded border border-slate-600">
                Logged in as Admin: {user?.email}
              </span>
            </div>
            <div className="flex items-center gap-3 text-xs text-slate-300">
              <span className="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1 rounded-md border border-slate-700">
                <Wifi className="w-3.5 h-3.5 text-emerald-400" />
                <span className="font-medium">{isNetworkOnline ? "Online" : "Offline"}</span>
              </span>
              <span className="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1 rounded-md border border-slate-700">
                <Clock className="w-3.5 h-3.5 text-blue-400" />
                <span className="font-mono">{new Date().toLocaleTimeString("en-IN")}</span>
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            {/* Public IP & Location */}
            <div className="bg-slate-800/80 border border-slate-700/60 rounded-lg p-2.5">
              <div className="text-[11px] font-semibold text-slate-400 flex items-center gap-1 mb-1">
                <Globe className="w-3.5 h-3.5 text-emerald-400" /> Public IP & Location
              </div>
              <div className="font-mono font-bold text-emerald-300 text-sm">
                {clientIpAddress}
              </div>
              <div className="text-[11px] text-slate-300 flex items-center gap-1 mt-0.5 truncate" title={clientNetwork?.locationString || "Himachal Pradesh, India"}>
                <MapPin className="w-3 h-3 text-amber-400 shrink-0" />
                <span className="truncate">{clientNetwork?.locationString || "Himachal Pradesh, India"}</span>
              </div>
              {clientNetwork?.isp && (
                <div className="text-[10px] text-slate-400 mt-0.5 truncate" title={clientNetwork.isp}>
                  ISP: {clientNetwork.isp}
                </div>
              )}
            </div>

            {/* Display & Screen Hardware */}
            <div className="bg-slate-800/80 border border-slate-700/60 rounded-lg p-2.5">
              <div className="text-[11px] font-semibold text-slate-400 flex items-center gap-1 mb-1">
                <Monitor className="w-3.5 h-3.5 text-cyan-400" /> Screen & Display Engine
              </div>
              <div className="font-bold text-cyan-200 text-sm">
                {currentHw.screenResolution}
              </div>
              <div className="text-[11px] text-slate-300 mt-0.5 flex items-center gap-1">
                <Layers className="w-3 h-3 text-slate-400 shrink-0" />
                <span>{currentHw.pixelRatio} â€¢ {currentHw.colorDepth} â€¢ {currentHw.orientation}</span>
              </div>
              <div className="text-[10px] text-slate-400 mt-0.5">
                Viewport: {currentHw.viewportSize}
              </div>
            </div>

            {/* CPU & Memory Hardware */}
            <div className="bg-slate-800/80 border border-slate-700/60 rounded-lg p-2.5">
              <div className="text-[11px] font-semibold text-slate-400 flex items-center gap-1 mb-1">
                <Cpu className="w-3.5 h-3.5 text-purple-400" /> Processor & Hardware
              </div>
              <div className="font-bold text-purple-200 text-sm">
                {currentHw.hardwareConcurrency !== "N/A" ? currentHw.hardwareConcurrency : "Multi-Core CPU"}
              </div>
              <div className="text-[11px] text-slate-300 mt-0.5 flex items-center gap-1">
                <HardDrive className="w-3 h-3 text-slate-400 shrink-0" />
                <span>{currentHw.deviceMemory !== "N/A" ? currentHw.deviceMemory : "Standard RAM"} â€¢ {currentHw.touchSupport}</span>
              </div>
              <div className="text-[10px] text-slate-400 mt-0.5 truncate" title={currentHw.gpuRenderer}>
                GPU: {currentHw.gpuRenderer}
              </div>
            </div>

            {/* Browser & Operating System */}
            <div className="bg-slate-800/80 border border-slate-700/60 rounded-lg p-2.5">
              <div className="text-[11px] font-semibold text-slate-400 flex items-center gap-1 mb-1">
                <Laptop className="w-3.5 h-3.5 text-amber-400" /> Platform & Browser
              </div>
              <div className="font-bold text-amber-200 text-sm truncate" title={currentHw.browserVersion ? `${currentHw.browser} ${currentHw.browserVersion}` : currentHw.browser}>
                {currentHw.browser} {currentHw.browserVersion ? currentHw.browserVersion.split(".")[0] : ""}
              </div>
              <div className="text-[11px] text-slate-300 mt-0.5 truncate">
                {currentHw.os} ({currentHw.deviceType})
              </div>
              <div className="text-[10px] text-slate-400 mt-0.5 truncate">
                TZ: {currentHw.timeZone}
              </div>
            </div>
          </div>
        </div>

        {/* Security Telemetry KPI Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
            <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Total Audit Records</div>
            <div className="text-xl font-extrabold text-slate-900 mt-1">{auditLogs.length}</div>
            <div className="text-[10px] text-slate-500 mt-0.5">{filteredLogs.length} matching filters</div>
          </div>
          <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3">
            <div className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider">Distinct Public IPs</div>
            <div className="text-xl font-extrabold text-emerald-900 mt-1">{uniqueIps}</div>
            <div className="text-[10px] text-emerald-700 mt-0.5">Recorded connection addresses</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-xl p-3">
            <div className="text-[11px] font-bold text-blue-800 uppercase tracking-wider">Monitored Locations</div>
            <div className="text-xl font-extrabold text-blue-900 mt-1">{uniqueLocations}</div>
            <div className="text-[10px] text-blue-700 mt-0.5">City / Regional access points</div>
          </div>
          <div className="bg-purple-50 border border-purple-200 rounded-xl p-3">
            <div className="text-[11px] font-bold text-purple-800 uppercase tracking-wider">Active Users Tracked</div>
            <div className="text-xl font-extrabold text-purple-900 mt-1">{uniqueUsers}</div>
            <div className="text-[10px] text-purple-700 mt-0.5">{mobileCount} mobile / tablet entries</div>
          </div>
        </div>

        {/* Filter Controls Bar */}
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-gray-50 p-3 rounded-xl border border-gray-200">
          <div className="relative flex-1">
            <Search className="w-4 h-4 absolute left-3 top-2.5 text-gray-400" />
            <input
              type="text"
              placeholder="Search by User, IP, Location (e.g. Solan), Resolution (e.g. 1920x1080), Browser, Action..."
              value={auditSearchTerm}
              onChange={(e) => setAuditSearchTerm(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-gray-300 rounded-lg focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none shadow-2xs"
            />
            {auditSearchTerm && (
              <button
                onClick={() => setAuditSearchTerm("")}
                className="absolute right-2.5 top-2 text-gray-400 hover:text-gray-600 text-xs"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Filter by Role */}
            <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-lg border border-gray-300 text-xs">
              <User className="w-3.5 h-3.5 text-gray-500" />
              <select
                value={auditRoleFilter}
                onChange={(e) => setAuditRoleFilter(e.target.value)}
                className="bg-transparent font-medium text-gray-700 outline-none text-xs cursor-pointer"
              >
                <option value="all">All Roles & Ranges</option>
                <option value="admin">Admin</option>
                <option value="deo">DEO</option>
                <option value="da">DA / Approver</option>
                <option value="sarahan">Sarahan Range</option>
                <option value="narag">Narag Range</option>
                <option value="habban">Habban Range</option>
                <option value="division">Division</option>
                <option value="raxœì}írG’àÿyŠ2V3 l ü¦%š¤–%[»ú
‘Ï¬Va5€&ÑVínˆäpq±p{÷o#.âÞb_gžàá*³ªºë» (Ñ¶0cè®Ï¬¬Ì¬¬üøå<HGÃ7Á/ð…¼	&çáþZ2Í£drø¢|ö×²0¹ú|m}<üƒòìzíkò4Šó0%ý+r4€ÖÈã Ï“ôŠ|½6W[ AdÙË`4Îâð’Dy8ÎºƒpœÓîéŸw/Fô9™^v7ÉôŠ>J“Ùd»ñ9é'éúÃ?Ýó4¸ên­¯“<¼Ì»—YÃ˜ŸÔïEw«·CFø/VÃFvÖ×dÍ¨Î@¡=%äcÏÂƒë`6Œr6oÖÏÜ(™L Ú×­°MIæGzµVØËƒô<Ì{ØpÛlF?…Ož“l¤lä,™äÝq8Œfci:ßR˜$³<Ž&aw’LB 2˜¥Y’v§I ohéó§`HÂgÜâ¸qxÇbµO¯¦aæÂ$£vœœG“Æásø“‘?‘“0ËàõÑ`@¿Õn¦?RX5á_hˆŽ(P¦~+áå4œdaãð	|¡+2KChëQÇõ[‡ã¤qø‚þ‹
&ƒÅF‘…tE¢üªqxÂ¿uÈódðZû1SgKÞmê| ocØÂˆŽ´ÓsrôãPÙ¸ú¦M>†éYœ\t/»Á,OlÛq“¢žØ°—1ÉFÁ–ßÔ6ç~Ž})ûòlF5Ïòeyãƒ$Žƒ),˜:ã|C:yªí™,¦„©»±¾¾vŸ“öv
î¡~E_}}FaÁŽå.¦Ý-rAÛgm3ŠÖ8ü‡ýµ|T³îÖnãð4‡YŒ§¤õìä´½@ííûC@Øƒ…¾H×›”9$t=þ$xCíª;tÔ¯gý8g¯iýç|#.ÐÄîvãð8üBòC/(aƒ}˜&t>‹€€.§KO>† `MºÝ·t§—ï`,y¨Ýó¦¶ì§tkŽÃ<½²µAŸ¥ú®µ ñ~ÞO†Wr_t/FÃ°{Eø'r^Ÿ![	‡tSg½8œœÓAuò´,SJÍyÂã!¡›îdL®ÌÕYolÊ3–8Ï¶u³`{'£(Œ‡Gq˜æêÆ¿OÙñ}2æ4eÜïnJnY93orªH°³páV.[ØB6n¾L²mú„¯IÃÝíù`D®’YJ ³ÞþÚ´NÏ‚T)` ã¼»AqöÐO“`N¢É9k>ƒ”vE»“$¥ RAo)PâI7bÅ (Æ4Ï†Y„´ÉžeÉÓV‹òæ‰(Õ¾D1åÚÒé€2¶œ)Õ<¡äõ€Ð*½¼ XÉ$¼ ÇômKyÑîå	…jÑ9¶á¤ûìe£cíƒðö¯âp4²Q’æŽµt Ê1ùG—fà3‡ù7^®5¾ûƒåuæ¤P1âq'0'`¸€ið×}úKg	ß9AòçÉEí!,F”þýßI£ÁÀA_>²°Õ¶5‘–h¢Mñlf-.=µ¡ËÛ e¨F»í «1Ë`Ü§³Ùá³d¿”Ód¶dîhñœ„qºÆ‡ÎÒ.õ#¦Û5âa1fñ[Zñhó†ã¦˜›Ï2çÐgÓ!Ýñšbcî^¨é4¥bÔóîSI±˜4þfŒ¿o:Ý!¥…žù¤T¼ýè~MeïÎ—Ã(Ioé¦IVNHÓÅß7^Ý«ÉÀ9`”èëv:K§q9^þS1RcÌ6ÅJ”½Húeœ QZ‚Ã— *åøYI6=©ð³ÉYb¥@òÔyU7e‹2<#äuÂJ.5œWµÒì4¤Gµ‰…©ÁÿáÕÁ5t¡#djŠø2‚cÌ^q(Øag8TG@¨á¨‘¤¦2¡ìc¨‹€NYˆŸÏ“IÂž¿ÝØ@óGE¾!s;/÷wUÈ5Öæ	*P²i0€ÓÿELs±èeRÔ}ì®9÷ŸM¥É…f´èØ´qÏrè’z;ž¶iëx^ò©†~À2f#*»|èºEQÞ$úä‘oF‡f×¡§¾vO×T|ÅáÐ¢j²$¬ 2è:Ì&MÀ²¨Ø”ÇáA1Â'czÔ™û¦¤ÅÙt¦†ªÖV‡ H©øõûh‚Ê®>0£BuSçº"CK?ð“{vjÏ> xB†ckè‘±AE^‰#ÿýÀÆdo¡Ž†a"ºAF_t‚¿˜bÀ¢5`I-SdbJ*ºgò`Ñ¿C^¬,Ç‰¬‚œM¬gŸùûz˜‹³*·¡wýû³l”$äOr0!ÞZÕvs,TPÕå»ÖAÖ¼Dö{š†ákºTÒ¶‰¤þ­IËj†³®$b•dÌ¢Þõn‹m,Ç¾“~èåbwäOCZ­q7Æ|pOHêK mtäê†Ü_UÊ²šv¦üaÓôh8LéA·Ëq”†ƒœ<{íÙ3Õ»æ¦\MÓÓò»P®87ˆÌå¢©PI
BMGùÎð‡hFAL^§•<Gòl2Œ‚†—®ì¿¦ti˜x!ÓÆ›àDÁ®o0…®ïùY6]ŽÐ=ÐkÄdi¤”S´£zW@5-”HÙÀVBp¿b_g3›ŽW/ç£lôx'‹ 1(ÚŠsjåNÆAšOG‰Î¥´a,¼•ìêN¥ëÉ$Ê¿àâÔ"Ýz±°b{'ŒÚ¾:iÌÉßÿÇÿe¸Ýg×ÓLÚ÷ü)–ç$·¬`ø·kàee”ýçÁ—“€.yìzf{xÙ€Š@ú;‰g…êõh–'ÕÀâ¤eÄ/¢'“Á,MÃÉàªŠÚZ™Mã°X]Kó5FæÅ²ëRÅ×ÑéJÇ+·{³Þ™ó!“¡Jã’±Uˆºw]jüê7˜ªÙ#ÆÒª òLÜŠ+s¯#Lèïñ´»©ðà!»­ž0/ø)õê]«»Çþ,Ï“‰gÉäq>\·„iÎ³I6¥"/šD<OÎAméÝtÒÀ8‚™ôð—Y–GgWª•SiÜ¤÷…fRºë`ÏŒ#/ª—[¨Rh¶ë=Ä¹ÁvVÀFtm¨lÈáC¡‡#y
&Ò¥þ³×¤¸w7ã%õO®t‘G-ñç(¼¨f–mGUÛí+|,êôy[¿¥ÕlåÖðò_5äA¥¸ß”G÷÷a"£J¸“ã(8Ÿ$)äE2¤'Ùâç:R·€ÎE‰6º¤hD+…T ¨§
Övéwúw˜&S¦R¸¯Üë¾ÝÜXG`gã½iw›GWÌ8 ˜ÐãÅQzü:£Çø;œ¥8ÃîÆŽ.!ëƒ-lÑß”-”bÂmÆÁe—=(ºEÃa8qÙ;A…=ò¬½cÓ„`ç˜¶t™ÔhK·“´ŽHKvó¤›’³4KÊQPž-Q¾ âög“Æ	÷Ãü‚ÊHêzqË°ƒ*È‰ÍªÖgÓ¡mÑkkÙæó\Q´è3òa_[‹e(ïÙ¤<…T*oˆº;`Ûä *N^¶ïäqû£-ÏEF6öÏ
èdu¦í”°%,7½ƒb¶j›·p¹Hâ.[òMƒsthŠEžï¤…ßÏ	5Ó“	0~*lIô:¥Á	­
%z“ä¢ÕvY™8¸·Ó¬Æº€®ÇN£Z°˜P:c›S—å n ­šFk·š÷m\Üä¹V0üEÝ‘p4¶ï·´™‹Ë´ž{r@‹jâ4yÁašW&W°àtÅÍ!9RòÃ–0»´R\q5ìÒ¦Ê6®SPÔ¢n•Z¸½›M§a:²Öoðbs÷"‚1 eÜ¢Êìê[N‰¸Â"Ûá¶(”{<OéN‡ 3:º åÏmœÀ–Ó†¿lÓ)“YÏ±Å<;óIr0·”…²ç1«²Þ¼€Æû|•N-xýë9­væì·0^|ö¥±AaWès±¿âi	¿ov¿Cx½ÞÝ ^®‘×a:ŽÐYáf£˜¬©7¡Þ{ua4bBM»Ü¼P{š¤”I[òÅæ4?± /cQÃGewü•-	‡ÚmœS‚5ÍQ‹Ä9Cn°ˆÂÓÀKÔjˆÎnK¤:ô	¥vf¿·àb|pSÕ+<Pœ&<Öv¼åZ²àª¶ªYÔ
 Æ ÄLáUV|½„{fFJ°E‘Î.B1ÉFv¸°ŸñkJ9¥šiÍud[ÚU;²ÅÇŠÅú·ößry§š+”}6TÙg«¶ìcUcà±Úí)±—ú;¢’¸—àâ¶KKæXøk˜‡F—‰ÂŠ¨ÉÀ…5¥ËÙ…þ¡Ö• ¦DˆÎŽ’S	-WÄO=f“üjÏ§Z·Z@Üèê–Û™«-›>˜ooj±Z1Á2|¯ý„…“m*·P6¡±ôŽæWX+'nbÝ8|vòšÊÃ¯Òó`ý-…
Çìä†¥ˆšº¯ã(›ÆÁ9¡C‰Î"î|Çýß®ž.bæ€l°.(o[°)F·?to…jC¸Å	ô	ÒÒ`5úY­jìª%–dŸž¯£ËÎpt^€á4Ýêùq
ý±î Š½õK¡Ö‹¢G„E7·»ý(ÿ€üñë‰d†r«à¶YÕ 0_Ìâ<ê>NÒð7 PNG™±Ì­‚S±ö8žäÁdHaüÛ¢'ó	àXx±‡Ù‡<™‚Úïño ’§Él0bLãV¡˜C?'TŠIÒ\hEºøðsñÖÎeßSšù}LGÑ £ÌŸ#…H7Q¦~\(ÖÏ§31üOaÿûç¢çcè3e_°¡Op„ò‰ú¯¨x`”„“«Œ
ÂeŒòdrþ.¿;ÿy0Båñ¹ï™ò_Ó¡œ%é¸C^ ¼m=×Í${ß1÷6è¦Žo·vUÈí¿Š&CºŸ\‘uûäR`ÙŸÈ“Ë`nÔ~3€–n–šÖóŽ‚Ì‘ZÖô+ÐÁÕÑ9Ü¬/¦úòzð”«Œ^€×ç]6fnS±”67n‰õé®B 4ì§að¡Ä1aqÎðk¥bÒn7÷ro¨ÅuTÖ‡N+˜§I’×µ ä¾Ý ÷C­ß–qïXX#Ö·ó³YóYñCVGj^æ,ôŽLô”w¡ùë¸LGY$õcÈƒÎ¹oX\uÛ6W¼ÛB©ª˜RiÖUTCDÎÕ47^i›rCD0ô[Î&ëqœÐ=XšîZ æ5Åò>Ò@’ž£óü»?°8 )J‰@^“àœ.âc‚ èUó°£]K€@
‹q­¬¶ÖL(ä²šçtŸ®_a½:uKU2ãV PZIÊù0£Jþ÷:=ŸÃ†Ÿa|½“«É€<7‰fp)•$`ÿ~÷>ë7F†íI²–c"%Ý<bø*×}§[˜ÓÉ¨4HÐéŠn),÷@Á¯:é˜V&û£mcQû
·Ñî v\×V¦À²ÿ&<KÃlôøÂbÎh½
ÖPWNŠÊVIZ±†ÚmkuŒ“ïwŽwÂFDåwò”.08F„…ÍH@Œ	óq*L–2¬£áŒÒbú¸iŠ¦2=màSaÐZå!`<Ü“µý°N»ÊäØÁÕØ{¬üçir~ëg:QÙ®Æz¼ª-ÈÀÝlá=`"^…èšue
®+Î>ODcÀ°sC/Ì¯¦ |i2OŽ(ÞÄ!6'K¶Õ„!ÁˆšÒ<“~7½‰¯ßSÞ¹¥úÔ•]o¤ßÓ !¡…G´´Æb‰21ž'ðƒ’‡¤‰$‘.€¬~–Ñ·JfÏ·Ö×›dë¥¡~˜uà)Ô0F3¯ÃÁblðßÿë?	ÿ‰ùûýo¸dc!ÐÊÐmì¼"åŽå8ðÓˆÒî¥C@JÎÈ€"m4FõÝ4¸
Ã	Æ`%ÚAÚ’FtaöÜG#GC´Åˆ¯(©Ñ°˜n›4ª 5æÆ›â¹‡K¹æ&»ö3¼‡AzÇb­2oèíÁ™p²#I†N}kÅ)YÃDåe{0¡î·V¨÷OÓ m:D(Ë  €6’Š¦¸¦”s†p=ê·½5‘Ò<ùH„~sC\+ìˆ™nþN(¾Ÿ‹Ú¥§÷9ÉüY9¦¬—%ã°Ã°ã±©Ø`Aó¸ùì1Œ=2h!ên–ìóðÆ >Œ8cY™­I>ºKò”UC ñÿþÏÿúŸ„ýF<‚b{øü?1®þ…ôâ“²(. âÁX±@;€AxUPË¢˜Û¿"7Oa-™Ð>öã4¥Ë&KdÇ	0ï6ÃX\¿mži@÷ì–—ôë9|ù!è÷Ù#žc¾Syäïæ;±Œ£C—k|e`i _ƒmˆ²CuÙMdOHMŸ=Ú,z±kØüä	>’¨¢?‡kŸFªŒ8V–Ëwj=ÎëŽ¨¬žSIYÒŸ«°à§;ðQiOÍxÈFlW&1‚]Iz­bìC²œÍÚ¦&Ô({dÿÇIì+­)^¦„Ìr´ÄC·2ÚíÜolŠ¹á’ï×™)?•~ÅÎ6‘”»þD7ªZß¦9Õ=¨yŒq¶ÉËðWªjb.é<™õÇQ^ˆÅX…Ö€
s;¥¤Gz8Þã÷4¹`†ìª¶'šLg9QÀÈÈ¶¡¾˜ÆT6Ñ½¦èÊDž“VØ;ï‘Í­íÝoï·AqNbZ]žŽgÂÆË\Ì”¶,</¥Òþ<Š7ª8‘ž.€•´¥á¿ÍèaH&kõ 3¥]]Ð|Ðym/£BAªQáÀ‚@Æç˜–âàzwî„»ùñÍœßjÏ
k³#(Î®™¤<;MªÐr%aØCÔÚ³øh Øëáñ“Wµ
ñYjUàbKã©U	œÆ!þ©U	BCö·V!&Ñ™ðoµªqA«Èjf«dK’Ä¾ß2¤{‡[”nB\°&O5Cdpò¨ŒLçDûk@Š5âàZ&™+TçUòåS²eS2s)	v¥øÜ•.;ú!ÁÌ˜#.¸#ö³ckž_=$Û×4náŠ@:®ô"¸)Î¬É„¼•™ÑÕ‚Õ¸ŒráòiéYTdI?´oÄ¢’j	ñ(3+¯ùä:<ôþ¬¡¡^±]8YÓjá¾4sóàÈztäÄ±„£åCêW±	“7rÞ½æ§X’Ï†xfƒ™x¢gÖŒ'êsÝ²ˆêÇ)¢©Ÿ’S€–ç¾:æ¬4P±ü(«´¡K#[.º›Û¾˜eŠpbòë‹¡»¼Ï‡³?ÌÃî˜N¦À êNÛË¬Ôô™ŠM–³u÷ìaãtO¿"tÚ2 ¢òDß(nCñ±°«Íu_Ý€“A/ -’/®í¾àVÇv-®A£ò„€}öpó¨YOUÀ„®ù2É)¯Ë›0Îæ×ücèå©V„Úo)éš(L7¡+xþÑï‡j©ä?¦L(<Ã@=*,×ª›‚t“Ù/÷w×ÿñÊõÉ˜Îª*v5bÛPûÏx[Ó4üˆ¼õšôz=øÕ!æ{ä+øÉ`Þöâ¿Fª0ð“¢H•;A×å«|z£^$)ùÃÕèÓø!†¤<Èî‘ÆÉ(¹(ŸøfçÃXÎBèÈWggŽØV Ž2bKoýQÆ«èž'Ð¸c»Û*8mæo"ÄØÎØâÃYü¬)	ŒÝ`éœ¥ñÝË^"Këóì]ìkÒ¤§_ü+¶øCèÚá;S¶Ã7®m‡¯…Ž~Ý{¥fØ)-dÖ$‚óÐEûŽŠþc¿«·Êã¿£b=5€£rµ:ÀQ±ŽZÀQµ®zÀQ½Žš ¨jÍ©ÌÞÜdyDn&lOf,Á!nS|:hl¸^—ø±SäbÓ’â0
¬Õ±o×žFñïEÙßÖÓ ÍÂgCr¯Þ†Ý]â–Üm3²!ül§83Î¸%’¤€1â0ß£+àQ_ õ¡HØë1pú¶4d²$+¥gâ¦'V¹»GŠësîîeŸ®GjóMÄ)¹V‰MÖãš¬(8è†tAˆ]»)r¸}›êªkùÃ[3sæ–ÀàªâÊv©U¬ýrHÞp0ºÉAR+©ÎádèÙÞÕÕVÖu£j=ÿá’Û’Tò†JA;¬4•çÇ0Pf	Y$Âgß–jPí¯HÓjœ:ëô|…·xbnÖ…v£ˆXëžz)¨%ž‘Úsû9bÇrŽpO½8~›SwÍ“¢ˆÝö-9Ï:h'€Èz‚y9_£ÿfÂ:ðæ,ˆ³:V,z#¸>°©3»…À®j! Üésk¦f…iÀî¦YèÃ/¬Za@ÂäÕNš–ÄØcÌY‹ÆYÀ¢¤ÞÁÑm&“Û”2o;ƒ»&`#¯f’ UL2"ÿkqÿý÷B	Qq\ÎÓ†Ü¢¯1D_Ù@D¸É@0øâ^läøå5lÌ`È»Üö²°ã–TÌK¢´%£‡×¸—³˜ŽÑÊE2˜ž$9	†p¶¥\ªåY.m’5š‚\tÓCc´rÐÚ

Õ]Ão¦¾Æ6_ £²ç d0[ÇÀÄÕ±eëìÇA?ŒµkéGªÚ6ëË±QZˆÒ3)SF AQ‡ ö‚Y­`¦Ý­C}Á=ìu8„?Å³‰Iƒl—'F½êk	óÎÕb!QìM‡ºÁBEÔólã°Û%Nld¤Ûuj¡ö9í|J~ô<xm×WK,©”¸JbI…ÄÒêˆ)#–VEà\¶:+Ê"ÛpÍ‰šŸÚîlÕq2{N¼½åûÒ?fXçž«í´¿Ð4`—;fQqñl™ÄLžÄŒMB\ú’–Ðh¶W7!»êÈrîª$¯Üú¸°Fv9¾˜4oÈe®ƒë¯jÒ½Zv·`ÎBÿ5;Äl¸®·€#Y­Õí…KÅîl™÷lS,4&Æu\|fi–¤]Ê°Á¶ÌÚA6Uì«Ï¢ÉÐj_-À°²¶@äaak-IòÚõ¬"¹kJî®â’¡‘EB7ùúögh³øU|Œ\†¾r!‹>ÁÈ¥ÊÜ—ç#µ[Ïíè;òSoTÉãîw»SeüF·ê¦øy÷êÍVGö(Õw«Ü²{»~æÊ=X~Ÿ;Óî£´â-YF;³–æ¯W¼!—›ÙçÝ‰K®†êð­oBåíÝÞ†<²Áïw+" n{;b<WY|y[q¹™}þí¸äŠ`ÔÜyúv,Þ|ö­h»Ýcþ®l¦èpÀ´ªHxÐGchl7xÑi•¶—ŠŽ;³Ž4:ëUmsÒW5ºVÏØLÊN¬½Ïæ¿jºÕZdéê‹QïœÈ@¼ØkßSÁ]†%rš…n:’ìVÏ»<õzWœËÇKiiãéÆ3(Æ!LlÉº°o.Ê
€@‘K[,š-—‰Û¦òÑâuœ%òŒóbáËWúy	¯$=@Ú"^nE›y»•Í˜^oå»¼ßœq-^[Eh£~ã!˜ÍÙ«VuÎJ–®Ï8àÒÕ™]à<ººqGËÁâùænÉ¼„/ž»Áæ'>ºŒgß©p}»¯/$ÿ¹¸¶ÿ\%ÖY¬¹6•Òå«³‰ò	©(^$¥‚¢I©#Üjig
ƒn šG@.‹J>›mš{VÃke§½>š ~ï•,Àá èëÅ%[úLkL0§¨bÙÓ]÷TŸ=zþä˜^=þgúÕï¼àwÄYB†ofÜ›±Íýèjapiá‹Vr#õ{{É~8õðk	7—Ì_.!‹AÐ©³+à£³_Yà”L$9SÂ+OËK£“}áØ	¹á*½=«ç"ÝöûZx•PTq¯L–ql¾1aö“cËUö
	©ÞøîÒGAçn›~qžÒ­YI®èŽ)¨ØÌGÅ`?½\;jÞYjfÌæi…¿öç"^ŒBÝ¢dyJ/dcÏòquw©#W(ªâkû;8iÊ‘Üß„ö5+B¸3õ+1L.(U†¯ŸÂ{t„Ø#bi‡}ˆé?È)>{ÀÇÞ¾+þ@¥iJ|EixÅâÇ†C¥¼xh”W´‡b8:ØB¿dtH­fLQ2Ó°4ëg%Ï®X´=ú…‹®gÐÜ™®CxúWIx-ŠÆh‡½,ÌŸR4<‰þ¶6¶Û–W­æ(Œ?†y4 öàkS)xÝzïá7']òô¯äÞ5ãü}‡llÃEþ'¦R4zÿ•Îcc÷;¡" U§zåh_*±oHêÄº£—Æß8yõ„ñ~ÉÉl<Ò«<æ7ä¾Tî§€Ó-ÚFGQ‹ÂQk¼Õpå]G*»dOA1ùmFÅ¥ü¯{eÇ›òëâS
È›J½QrñvÞ)Å¹zœ‡J)}ðòÐ0µv~‡E¯‘0T÷È·2ãøu0R´Ý# w´¹ž”£8~ºU:ý
¼uúÏúÝ dñfs‡¾ÿÐwØ´±ÇHÒÏ%è ®z”îæGb1öx¡‡žB½rA)+ÚÚ=Ìe”·³¡oêl|»£êåµ(Ø`ZŠª»Äúíòñüö0û˜Ï„Yw/¢4V@uüÞ]¿5ÚfÁo†~ü^ÿ‚ßâ.qð1,HªÏ¡—†Ö¥µö¯Ù7k”O5n¶ç½éðì=ojÎi©Êäž‡Ãó0å¬î7Ç}NX÷Ÿ ·%Ó75·‡Añf×ÖÈ£YCÂ‡d6Í”ÉÇØÆ÷øbn#ð…Ø"Ã5ET¼K£«RŽÈÞY>È“<ˆËA‚s#8IwØMÈüBøz®ÏzTª.jƒdÃ2ÎæÒCq(„‹¨zOÉÊ#Z‘E¾ç‹˜¡‡€XD¬ÙcžÉ>‰¼~ØBÿÍ“ÔWXê“= èVå‘B Mj¼Ì•mŒaÖ?Òš±Ž¤(fk	îhÝš@Ö¾·Õ-£“”#<*¸1<yŠ¡ZHÐá¡Ýa´Òr›(œZGã §Ë‘=Å¥E‰V¢S­¯R$éü-_9Üyƒ‘4*[)¸2ñ5…‹È6ñÀÝ–ªhJ,&4F¿;åªFV.*/p®,YÑ&µM[“ÖÕhß•š	iÍ@&øJ]©¶ˆBwëjQFNNB §a:VE±×à-² ºèz•;|gÖfäð…áñª¨¸›=DoZÊºyÎzôEzØ¶ +’$ƒI»÷KMZMÒl3ä5»/"íÿmŠ•:%"tŠ|çêè¨.ÎÖøÞEŸçŒ&ƒx6Á"¢€(\©,àµLÝe–©Ì©LH²‡uZÉ6I’”…Qr9/é%é“`0²ñyDá3Ý¦@< Ïˆ%/&îúXlùe\Rq:¯Y«T°d?YPq”íÕáe€ØVN&ËÂRÃsÛc‚½báªR3ÎÎZ)Ávb³‚ÁS6Äš_kû–·1·-*—1¨XÃÊ¿A¨×‡µ±S?„Ð=nýÍ»÷®éFdÎqï=*Áoiíwú"@–-§#Å¼‹X¼^HÄ »¦`tÕÓ"ß µa	Ÿ$*Ù¶D.Hã™!Kr^NmB4§’vJjâ#m¥EDB¨â;“öÕ¡¦Ð´NOÝÔÔGKkQÒÕÒÑ•PÑ•ìæÛÛË·J?kSÏÑÎß¬ëBXÛõ)f)ä;É¥IÁìõ¦³lÔRé*\©íán‡oå]Xj:XžãN•nñV‹½gßQpÇš“s9iÊÜ2}™Trª”æáäUÿºÔ,Ž@¦L»Ýƒ¢­VÐ!ý6Ã´RJñvðq2žiØê—oðTô8³0ÊñçírT@Ù8ëæ¦˜ƒ<Ð‚¬·P¿C¹ÌðR3æ†U+Ú9$»ë5u»Rçp§QÐ©Þ}ÐÛYZ¿KKž
½cKÒGÚÔÀYé<A×:ï‘{×8uÁ9y+ž”ýŽkÆÄ”|}oÑ~wA‹ö­½ïS`ÉDfîo2Å‹ õð"œ­…Î»»ãf8é>{ÙlƒÚnóþz9:P¾qtN›nâqSÁ`cÄ´*ü¿Dií¨À°-?gÙ0€´Uoß6Oz/1£\tÃ_IÔh½ÉÚðHÖ:ŠgìB¾=
b§ÌÞ¼{g²zÀÕg@ Ã^¯Ç`ƒTá¾‘Š+÷ iB»wN§!òo·|Ù×_ÚÔL°{ÒÙdB-†x@¬ë¢Êøò	ÈHdÝ€ÜÀ4¤Mò¶¹¡Áƒrit ­u›ü©u2ÎÚïŠ±½“†/-CI1`˜XúÌ R¦³ðTvÓyCQAiQ+˜òwÆùõ«²ž.Nk‹ÒåÍJ_=7ò¿¤,GÃ:1–óVi‹ÃŠO™ÒÍ¶ÊR°À.Ê¤‹ï½lG”„u›pañ1LQ‹Áxë‹¥Ýmªí¤…Ô'áFshøå_´ú`ÊCdÓÑ¼~òòøÙËïÝC¨‰KRõw2mã,Ircc Ÿ¾:=z~³Ý U¸PµwVµm<÷†âÚ¯ Ù²ÐFÉ«~(ÿh}®;ÂÞÎò·„›Û”Íˆ´{BÎƒì7„úä]ÍÃ-#ÿgÙæI<OŠ”°=Žge^RMŽŸ¢a>¢ÀØV[#ïüZ.±µ£—ØTZeüZitû¾^ekñ*Û•cßÚÔëìTwã¿Ð%Úå®)^¬Ð~acìIrÙZ·h5ÊevKú3Ü’þÌ.C~ú×ŸÅUhy³üví_ÿõá¿Þû÷Æþá»î˜Ÿ\R@ÁÙ2ÍÏnPu‘¤úIòßkãèþé¤÷\JîüOMÓ#~ŠæOF!øÁ¹ zÈày«a˜Ù¯µ5r4’S€—ÖÂðMrP•{„nèãÖ[»á’Ì{D=@[TÄÎîá½ =#ƒM,Ù~(=Ãô<|Lw@ÖÚ ">þ_[6)[§¸•Æ!ïy”¤Ñßè0‚¸Ü‘2I àíZdpˆ¾ï„†6,i<EÕ^H9˜M«»Ú”áèC¤T$+E	1”Re¯«)…iÓ¹M4y‚?¥ï³$Ö%óçœtÓ£LzÞ§¥ž>}²ÿÓŒ1‹ÏK2…Æ2Nªòd$Ðø’Uê'yžŒ+‹!¹´”²]¡Ž|ÝaÛhˆ{QˆàÒ*vS±3ð ­§œÔyñ*tð¸ˆk¹6˜kZC“¹KÀSÆÒ½ öÆEÌA	øÛb-Øodp$—,&;ÊkUJ¨_~µùd<Í¯ÄVc×²¸Ò¬¢}\ –xã¡ôƒŠ­Âà\rJçÖÕ“*zäæË¡›nöò¬äZj›sØ‹É3PHÍí“M:xúÏžô”žµ¬ªNpÄ
FTZ2Ó?ùYxoÿ´dAJ_%òX–Œ¨h§.R;³q!MÐ¹PÙãBLÈÊ‚8,~8ÚÊ±üPÕa?.æãa=>ÆSíèý[)b55¬A	kQÁJ
XŒW¢v&“Q¬|,Fg0Ö•WLMöòëæÞV »˜‰
´:¬ÄÃHnÂFVÌDa!uˆÝö}yÖa\ÚôgggˆnÁE•§¤Þeœ]ö.Ò(a‰ò|ÄëÅIŸ¥Ñ¯­·¬¡w zç¤'˜Nãˆi§Ö>N†½dN.Ç1äÿò¬›œEƒžGg@‹{Ù4¥DùÓ8îá_YkgÖ£¬ÝvÈBÆÐ0ÿIõ_¢iyNµÿ†ïaªÿtBËGEŠè=éBCex{Ê9 ¨øšÏŽñàÕøƒLàxA‡oð`Žßð®ßòËW|Z^ž7ß}göäËn~xµ[êá’·#ýº‘A Sév„àk~‘X¶ï1«-Lj%ûWð”§_búL–Ëû`gƒ>‹åòžÚ=ýÒïªƒòº[¾2w¢†Í„f/¡µÅ›z×¶àÂO€ry~ò—Þ,â¬$ÁÏyò3ßÊ¨ÙTŒþX{œ}T[Ãf ½Aö±Åû+*Ò]rIØj
Ò’NÑœÂ,6{DÛ(}‡—Ssû÷mÅ¶9³AMygùB‰Õ8Â A?µmC{P7Maô[ TwLÈ“üñ[Ž°îG*q†Zåpñ…d„¿ùhŸAô³÷ÿpïZ~4ÏîL+O§ã_÷¢¸ê|ú–fü+ŽýXÖ;Ðc‹ÍÚQ–z«G@SÅT¤…îÇkÇO^‘d_‰{`¦¾ü¨ås–UQZ„ä0v”nÉ8ƒaXºÏ[ ÜËlx"‚½	a¢	9}óW	ÉÂ!Òe:k^pDpiÒ´\ç†Óè·#Y¥˜–©˜F‰#°¢?{‰–+ÊS>Û!°ö2åc69M¯ôŠbz70£>µÓ¦[&gz×¶LÕ…å rÐXòã!ö$U+1‘¾ù9cHÆ‘‘5æ:>Ósv0dâ´tNÂ”‚çäƒ–n õG¿a
(¼*£€^~2ˆ‚øgÿ>»úùÞuéø4ïÑ¶íbÉ SYhD·]1vræ8Ó§rŽ•Ÿ«òŠlåo0ôÖªÏÛ77ô×ù<ÆCÍš*a^‡Ã™†ÿ‰¿¦WÃ Zp¸ùxµ\½Ú·É´hÃz•>U²?•&’1ªX„=V.0üèùóW?=9þ™.ë‰äPÊ¯))@Ÿ¢ n*>.Yi7§Ž"¥¥ZY1+h3‘3û»ÀòB«YBÒ¨k{¥÷+¯¤Ù¹õ-¯¬w8£€Ýót«3ŸéˆôH¡­¬Í,~dÈãõø#~¾.Ò"Ì¤#ÐŒ‰ÖY˜ü°m[š³õD¡Beû|A"\ŠP˜*|¥Ø‰hS¥)–Â…>­‚h·{—(† ¼w)ù%EÂ½‚¢h’üûã–mìb¾ô:s	á{’;™SÒ¾Ú“üÄ|’48\ÎrÒ*ïÉ?¤þöÊ¯å[Ý¶Rü€# ¾W~-ßö™)…ôD"(ÍüÊpÕ¢Œ²‘Á•`SœO´FQ"db&]‚hÃ%s8}{næz3ÐX$¯Ê´6óÒ¼Aÿ !Pÿ¨å¼£¤åteQ2ƒ,HM"øïËG,WÖõD!øÕäg¤Œ°.å*Z-LØ‹¸ÒªÙà:·‚Íp:¢ûú=ÖèÞ» ›m>ûù{U*Co¯ fK÷üÔwÌP›Õ¤¡ I{FVÏN^ºF~¶<m¶ß®¿ñ’´îkPZ,WƒÚâ”kQ\l±&Õå@B^â-T‹³¨Ë²Wøôx¬à×5=¸ ÂöªÒñœüÙ„RÅhYŸ7¥gÎ¡ÓzÞœ­â¸ÓÍHr4gˆÆý:²ê(­;#é¥9rŽEú{Ÿ'B}Ü¬‹àçbZGÀÒ:xZÓ_Ã‚´TŠÓÈ®±62"ãIŠ'}î†CˆR@½¿vù„ˆ¿ê‰¶MÂ¡r˜[vÇQ·ÔEÚp‹4¬éÙ$ú·YÈTæÀÒ4¸ê¥É¸äõ„Júfïe¦ÀTD=h›§\´­o[:c‚ÍR±¸tVŠ0ËôW„NXdz²Ü´Ô$¥ð
‹ô›¸ºCÁ¨Ð±Sî"-³Ô‘Ï©Hmo_O™.·#­•4,îÈ†à¤Ž°<jðvZ¨
þ ÆŠ)O€ÂGî¸ñæq,XZÀZ…o¤Åƒ`]ÜËÕ8¾dÏ%².Ù@²äœë–«*‘ð…BC˜~K†‘)×ûæqdÊõ_E #VKFC”…”8à)S1Æ¤\
ŠƒÅÊ¥Úvàšu[ÓsŸŠ¥á€îkVÒ‡µ5rD«E
Ei©e 8å(]ñ‚?‚czÐ+WÎ›}i¥d©–º,møõLéDô$/%ëLY>k20\O¥Ú¶^%øà¥a+ÈèÁô>¥Ÿî‹ÝãcÂÌ"Úúò,éÑg7ÇAÏKž×;®œ=
Ë²O£Ë=ò8ˆ³ÆŒR¦Ä;(ýHSº9ã+˜ÜÚÑ'T …Y!g½R.\²Mƒ(¥¥ªU!uJýÔ³ÉÓ’aY8¶´”ò9-Ü#eÇ7pOJÿ•Þ3mJÛ¦û•.|£RuÂ+Ýp±y´+ÆZªZiôóŽ L…‚Lû!ûc¶Lå“R'<WÔ,ÒäŸ(Jë3ëTSIm“ÊNx.ezá:ªÃ·kt*cêëà<š`A³MpâÊ
ù}2_‡)<E(u7ÀŠNØa—èÂHk–ºÚè§l Ì°¢³bí•PËâ€Æªañ.Ùh“¯-uˆTÊV@_­‘ts-ì 
;¢ÖC]å7³CrZOhÌ²Ërž*å~Ð­)0¸3›¹QËëí”iè_dwâ0Ö)NJéÓQø3Á²Š)E*á}GÇÑ´ÇOÐ|ïÚúÅG)ÿèyz²i0»WZ¶ knsÌj>Hb2îáwØ–ô»’]ì—Y–GgWÝ~˜_„”%B¶1=‘µu#GÙ6ëpÃLc¤UOCJË!G+&‚Ý‹îxèÊ}fËEäoÒžxˆJ£m2¢ÿý,‰g”¿€íjwÌ\»k›,=°DÒÅ¯Àƒ(ìñ•œh2jÙzŒ&ÓYnyÁì¢YTv{ u4!%1Ýa>p~/C@Pž%r2\D™¾×³7F¤×ú¹ÍJ>™<Ak×-¦+¿¦æo´º­gÓ`ªíïÊBbƒ>±äÌƒM?î> Ó”®ÇôªÈ6$’îDŽ‚øœ”	±³l/™åq4	»“dòG ¹ÆTÅë~‡¶ÄÒWÙ€dYMKXx|ìë¯'Orf«“fdK\gI˜‚arA§n<Ú¤
¤1,¶5¬˜í
ào…‘BZ²Š1'£äâ|nh}•é*RÆ:H)´² ‘	ç{×Fÿ<Ga‰GAeÌÒÓ*ˆt…ZvBž³¢VfÂ}Öí%ÂD|Ì®´ÖÙí?…ÓdòãÔÑ$äeŽ“‹‰£T½Ü|0P4F2ÑC!S²2·¢’“JÉU[/1¬±N®êÒ+k:45Ã'E#±ÎúÖ³,R‚Ò‚˜êÆúüpcLi›`½®¬·¹3?ÜÜY¼Þíog‰þºóÃ?GôDÅ3W%Šr¸®‰4²ìYHo³$)‘’h´d\5ô6[o*I‚7üWHŒ!Ûz‰MÃžªÔ™üõÒ™ÜÕHoj’¢ÈtÌg­eý„œŸ¸m)4êîÉáŒ^ì·iï¼´½LvG¸Ô‚y¥ð±l#³(ÖÖ³‹ö»»ô4A~ êóÿÒã!~é‚Z£R¯=©”ÞD8!ðÀ*~¸WþÜ"ñ¹ôs—Ÿà?)i—`À9°`»¬çH¨´o•Äø»8è‡±‚;$M¤_^×Ò/Ë™£ÊôŒW•ž!ð0»¿†-:ûs±/ñQØ˜¢Xuçnò±3!]Ãå¤ÒjGÜ«Ü­SÜ|7›ân¾—·×øF¾ŸÆ	¯3ÕìŒS´9V§¢¼E_}Ñ>f6Ù©]ôß*Eÿbƒ»8­øxL)­Ñ8¤´‹y
gn6(>×ú}^y+(ÚÅvÙ\¬~6?¤ÿ-;ÓvÙ¹fùÖ‰øŸjK0µÎ­ì	†¥«ßû]ñe'àbÙv7d¨½N
ÍÒJ\·’?[Îz×d‰oûðb€,‡âþÜ{¯œ3•¯³¶»Çyûw´ÇqcÜÎÇ¦oaó½¬ïñ/ûZÛ×Ìf¨þ¾.Œ§o¶¯a¯YoñWºá}#Sl©_ÀÝ29Ð‡TÞ‰û)‡oœ’é:ž£#É‚¡Ê~Áß‡Š4)
h¹kFãJ³†Ì×ÞüD(ˆÄ­@A’VOKbgˆ:¿{ÂWšÖ¦}Š×Ë\¶ròg{ÁM¾ÐÅÐÅÂˆ‚ÒÕWiùUe÷us,I²b|!Ê5¤RÉ`àvdÓ’HÞ‚€*S`ƒ<ÿÞ¨±´–‹PäÍ•ŽÉ¤Ÿ*[^”ö_ö¯„`þ¸pŽÞcÞ:¦¼+gê ™ ¯ÿf_8GÎñêÉí0Œä6´•‰¡ªü•ÓúWO ðz¨‰ÏN×mud+í/TÿÕ¯š_‰/®Ij~õ½<nŸë£‡gÅ.\òöñpœ_?’Ig|¾7Hâ.˜Nuw™10–'CÇÝ´×"}4#
ß¶·°Œ%/sùÕ°OùÓlz—ˆX?›ÍzUÛrc@-v&˜þÆ7Ìž£xtŸ>BÙ`£ÙØiU¼8û‹ÃÐÃYãMH@øâ9ñÕn„#Þ:ðÕýbA[aœ±ET³FaŸÑ/í3º—âKSÃºK79WÖfzé°AÁæ`3ø`—ÅTîË"aãÐHûæ2Žôõˆ6ÇÖÞ¾Püý?þûZs4Ñµç¾~ Rõ j, ï†séƒ¶Øx¤%;°šÛÈI`–úøÌp^	,û”ßÙÉ]^VHìG…$ï`EPtÙÌÛ·ç²±›} ðgqrÑ½ìB~,Ý?í	-Çœ&øt´§4/¦ôhh?™®+9Ø'Z¦§š}"6áÚMû¹b[§±ª“,Ø$:Nlåâv~ ·†â¤7Ák-}Î1ÔßM>?$çFB«+»5<òcˆ›%iwšDØ®$H&¤¥½\ìÞ³	~ÓëÁL¾,ÈÝYvJþ² weAŠ“ü—%¹+K"_&~Y–»³,¯ž|Y¤½Òáüó…ŸËØ¨Ê3ï”tB¼C£Ò"ØGFŸ¦¦wõ\²ŸCŠb‹Ÿ§F‚Å_O“‹‰p‡˜þCì€ƒ
Øh.CEw]]X”}@vÊ®1Ü =‡æ–ãØM›fþ8·Ó6÷k¸¥Æ™Aõí4^˜*ÞÒØ%ƒ›÷ ëEoÉm­÷cåó(©…¢Ò¶«qÍ¼ÇÛœCEÃ
ýºO5.Uãë»^öAÙÈ1!mËu”Fuå¨Ñv¯M¤º@$ñÉ4˜\olÌU¨Üg³ãÒ–jáBÝ—‰’ ’ƒPhÍ&ÃÞ‚37ý­\ÈC,ösH0ïÐ±•ËêÐ¶ÕâýùÍPËPÎ¶DX®ö-"ôÒ—·;¢…uò·;œÅ4Û‹ ;}
ˆiÑ^k¥@#íwè¶„¶À8€ÝØRŒ±C²an’…O]dœ\wâ†UP‘Æ!¡ ¨Þu­XbäBik¯1úÙ8šÈÕìÑÇ´èhô<–œ‘kíéœÐ¹¥h4j»p“°Æb÷®–¥7OÒ%6ÆCgÓ
.!ïïá`µGFàÁP˜ŽÛ6l¥¼ß(³8Î«¢7PAÑú^2TŠ¢+U3ÜßònFnógA S´—)^¶¬<JàÉNGBý6„	lý,*vS	ÈsñK^ÎÆvó0“¶ÓþÙi‹ä€D€­öš$ªKËG«oÝ úá¼µ»dSnKy³MŽ<—Õ”-¹é®f}hDœS?î]Â>xHäCñšCzv«í´èQ#<]Pqg$DJ¨´mrOßjLjÜ&-Œˆv+¾ËC6{ÎªI|<Æàñ[€Xmfæ`ù#”Ý€ÌÑ]Y"&Ð»o¦weý;AøÞ ¨°úhgÆ#	JÊ;-ãÞÚš+Æ-ÄgÍ(d)$Á M²Þòt@ÀeÏ1ÁÜ`ùIã4!ýPX¸„C3ºl8|tu¦`oÂô~†Q‡L0DÖ!¤þž‹ öÖôeö&™6[ºxV®!äjyÞB‚³6E­ Ì©eqÀpÀ ÐqJ“"‚Èê-3+ÖÂÊÔoE²W´|u!ËW°üÖÂ•ß<Æ¸=„/ðÈ¨óÎYœGëfKÆu«Îëâ­’›•ûNm‰-oIu÷oÍ–X9µ%g¬ò%â“/b\ƒð9¥ÝÓÄ<$¯Ø½?„13J±*µcÁÀVä~ºÔ™ÔlQœ‡?ÀÓO0gIQA9DNÑ$‡‚Ž	‘*q¢¯×·jpêõ2l/MA5k~ôSha˜`ÈÙ+­­†y—_½MÑ„„Ev@[`wmÜß”Ñ‡-ŠÉÀ‚JöEÞ²5.w‘òUr°1ß”¬¥LGÒÝ—¢t÷m‘¹…2´å^ºÚ¤Ü™$µÉ[âyaäË(àry©¨3ûdUzFurŽGiŠŽÔúx-ùÅº¨B %OÎŒpÞP$”vü©‡û½ØÔžÊ/²ƒ¾„|¨ŽQ¸Þ–|ÂQ[f%¼–)©z±ùíÖðÒBÔµzš¤§£(c‰mÒ	g:ï,	Xu†¢>$.ûÒÐ9¹ø)ÊG MB¯0¿Úþ¼Ø”@PèÜ@yÚÖèŒÎs1Xà§3|©˜ÁÁKn¥.£*ÊÈï9s”)›zÀž·ìH¦œ|$ÎÂ³ÌœX‹ª{Ãâï™xÐ½ks,ó÷’ÓD”aá=tùêPýÒM©=ÝÂ»ƒD¡¹±Œwmq>ÅÊØBYg Ç²þ5YhA>éN(öê@mÅrvü¾~¨,þ³,Ô·±W	Ä‚+ÞˆVÀ±[ý÷ËÅt‘%Íp ÑÆãe ô:3.áúVB0ä½µÇ•i’ÞŽÂš5T°-îß+&ìl&î®fÙÏ•…K¬¬.ZÀÞU´™‘%‰¤¼çlëYô‚Ó%Ù¡E{w—*Ó§#¨šÿ“/Z»‹G]öpP¹î&ñw,¹hë®Ñ~uÆƒQåŒMJí˜±hë®jeÆçtlCçœKêP1Y©™›nêrRßC£,›{~Øs9½y¡ív»¨‹>êg`×˜““Ùx¤W…úsju»’î&àE9ç¢£•=jYÚÕÑt_‘Ÿqýn¦@Ø“0²Žæõe°Jí««µå5°ŽW …uURì•*Té¼ù•
ñBÅxF‘3TJJÌüÄÛFF:–õï$OU|`Åíùc­÷”0·Q¦P­"o*íÌž5 R]µ¨Wæ¯Ó”EeË€¥ØS‡3½Ðì(êŸ´—AÒì ü›3þýbËƒðŠ²<œ®Šî¸Rª?NH½=¬×È¼mdÑ$8êJßâ Ðå0{ÞV5Âbûÿ e‹ûA´iÜ#“3Â3ÀñÌõ„)Q0©\8#xMÈiz%%†‡ðC»ðÂG”$quŸ%sÜ¶¦’Ä[õTÊ'S/3{\Z‡'‚ƒÙ`ö7>®4Äµb³x6¡ƒëÈ*ÚªxeM!<H¯ø,éãTÍà;âQö}œôƒøÇó3ÏèŸ7IÌIÃq4iÂBªÏ‡abyÊg‘6e0Ã<ˆâp(¯èMòþáÅ¦X	¤58hRNjÝ1þØäþÐGø¦´—”rzËUdlKW*dt‘9b‹ÊÖDz”Þ7Š‰5u$ÆÒQ˜BG§åAÃŠÃ 	=è«8³ÈÚII/—(Òñ^ED˜€@Eý&	êYïˆXÚñÓàòXÄa·´ŽlãâŽŽ¾îSúŒË”•a¢Bó²%ìeyÏ2†Ãi×”,è<K®òU¨ÅW¤[A„­½L&¡³%Í!,•TçË¾0’±ÍPI¾ÚÖ†ïìVÖÌ´ˆÃO{áiêy/ØÇ1ÚËC–\žN³ùrí¨ùÞ¢ÐdEpW‹X ºE~uÎ­¤,´©Ú^Ê1×hNkoÉ”á«J~³Tá¥QÁ2o+•%¤å+Œ¨|z|ÒD»R+¢ã[×.x”ä£f[¨jöùó÷'§þ~Y‰Š¾¥Žë$çcr^qÚ’Ù>Šâ˜¼L*óÚ¾®B|ôúè%üáÃ´Ô9£C§EZl‹ï´imòßÈË5dâËË0'¯ƒhX;§­FWCX+}€ÂþáÕûP½ÏXAF÷jVJÌ!m´m41 ‘ö—7‚öeÓrøNéž/à[i¥C–«WzÃzØƒ_&±TÐ›‡.‚#·QEV:éªÒJ)}è¤¥©—ü=þ¨(>aY¥âÊ³fW«qžåFå™Q#CÒuùZ¢!$[fGX‚yùÊ}ŸåµŠùn©UxJ¶ÄpI—8’¨Ë;ÅXeŽæü¤÷GÛ5‚úŽ©²JpI
a‹gãÛ_mkíû21å³T"FÛ(>EöÅ"Åëo>ãê!\?ëâBPþU%^¬ˆE¥'¾Œ¥ôÌ‹Ä©‚Ø»j”*G8*›WÜŽÍãÚáwÍW%‡‘ËœyÄ£yÿÈë*·€güf¥g<ëjÒ]nú6ÃóU¸È+âT+ô¹œ¬FºÒÛ»‘œ¥7¶œÄ¥·2A{¿%ø§­I¯ŸIn &¶DÄoCÓ¥Ð‚šØÂuXe`æ4'tCRí6WÓ¸¡/!ä®°ûåäãU`)Ñz…¨ÊWÕG©d’È˜÷_-æßú êœV49Ž
8ù
ª´ˆ³ñrýI.Æ»ÂÅØwÜ¹#ÒVŸ`P–˜öÑÙXÈ§ŸáT/‘2À…Fd÷·úòÙ<ùt¯çU2Ÿð£®@#Eäô`Q¨â('wŠìvÃEè X"RÄÎ­Š°GÈm ±bÒŠE(ùQº–÷ˆ·sý¨ŠWÅs÷†§k–V?B…Â-8@«>kõC4ˆÜ‚Ã\\“¶Ø$n'è†õ§æ‚«O\*½]‰èeDæ¦?îM»»V­€SŒ¬Rº‚­A¢£]¢ë%p,lÜ“ä"¦$¤I÷ƒ´;Š†ºvÁî'®'«’@Ìú§A¿ÕÌ˜X³^’"Ù“z	‰®š
]¦ˆq¥ë…ŠéÓ
 Žn6¢Øô¡»^øü³!±k"1,ôû7¼ûÕÝ Üýy9ìŒZ&¹LI
9)ÔFD K……ƒŒ†±V•Ô’QêºÒ-ý][Ç(?û2•f"äO²O 9F3ì.‡çaz×Šê³/wcø)ÊBòÇTc)®[õÌžŽ£Q†¦A6áÒ:£zIÇA–£"w’S’[si?áâjãûä‹lÓê¿À1‘ÇlLµ´÷º°ä¤Ç›nq¼¹kûPÜgßŽÊSåN¬¾MAÁ¡@R~»9úÝÏŒ%.ÑÇ¸(Ý2.Jãs¶&Y8ŽðÌo¿£²\u>âðQ>in4«ž^;$‚Èc96	·^÷tD$üua“â"Wñ­¦.$¶j\ûâ" ÌhäU÷‚²vðíŠ{HUS\1aj#ºñÖË{HõjÊyÅn¿—,j×ž5ÂWßU:ã™3„1`kýÏquxÍ¹3ƒ˜µ¶rôãÔÕé^Q.R¥LEý®Ú‡òv…õàçËšç³*ðm©©)TX?·)TMãð=¼T_}Â®šµoXwtûeGÜ´:¦+Ý<NÓ	é>Ü2R‡.cqc
’.‹bÜ°â’ÕB2·µÈJÍk=r+Èf³'Yº	Dû—hZÂB	z›àÁs?lÿåÙë… »o„è3pÓsšµs$÷RÅD/J}*ÞaškÅÖeè­k‘ê,ˆ~½QÁ•\×0ún …w£CLÎÆoh3 ß„Ç>•<˜ð÷«'X¤ˆ·HHr8üä¾Dw6»˜r8CfØÍ ÃÁ¥'úñÔÞÂË¬ôfíù±4Éã³£æ”.Iþj²ï™æmÆcä¿sfïU2õo0KaŒ8Tìí-;rÿ:vÙ•l¾ìˆO¹#¸üø©÷ÄALzSymXil9fWÆãÇ+¹`)
u#ðÜâß,ŽðK wóÄ¦s²µg¤ñŽÏ¥4ÞÛ¸4»6mN`Ñ3ºÙ?íŽöÔ…¨û´séªs[¾ãtØüº¢‹ÎF;¶•®0ß­Úª¼¢·×Qˆê¼g”ÊØÍYU{|ÓGF”b›®þ,Ñg´ã˜¥ÒQwÛcµú&Ì¦	8lÑ„Îë"æ£ƒÆÆúú”NÀ½>ÿåŽ>^ÌÖ]„²³ñÚtpýÖ[†"›GP\APJÚ0Jèž/ŠL[É¼S·ƒÒ»¿FeÀPwŒyø¼›ûM&à[=ŒfÙÁõÖNEéd–—¥wÖ+JOƒáŽðhr‡´xEiXŽ¯8{…ß²CDô0ŽÁ—0>hüÃp°µ³½c‚î¬³±Þp£¢ÎþÅ/oÓ$‰óh
a+(Å¤p£ò
_X®…—÷ÿÿ¾wÏM›’÷ÎŽÅ¼û`Í²å\ö}Ëå—ÈÓ6äcñÐ¨Jg1w†{”÷è‚Z£ã	nÑÎÏ3)lÏ—Á^Ï»¿"I½¯·O7{AåWNž–†'é¨õ…C?gŠC*\®ßWÙÎÈ|ï$ä'ÛÍ91$“ñP’L¶TAe—#?ü'ÙiVŠïç.!É¿8è‡±"SqþC¥ãÒŽ.^P$ÃƒÐþ¶èè-czÔp‘l$X"'†ââïbGôÐ6‚÷”
†ú­µ¸Fc‚åGE”0Jƒ”J+,H÷#†5;DqÃßÓb¨>úóº¶ìcúÂ9Œ·!Í—rfw9.…ÚÙŸs×%˜«‘¯J3Ö\ãl½sU½žM¢›…¬ðs*	âAND«èÅ”ÍÅšgóCú_Ñ®c›R2´À†ÿTÈÏÿ+Ç~žuÅè/ðÜÀÿ/8oÁy¶´u‘þ¤pòÂ¸;™O™Äc×XcšñÒlŸ"†Ð¸O¥°¡”Æ%KÆ¡ðqóFOƒm*gZqiTŠˆm¿í}Œ[`õû˜]õ>æûUßÇ_ön±w™ÇcÝ½[øGÞdïÂŽ²VZá¦vJ	$õ¢ú}8e\%?mpQ©åèF
)Z•ÊÉ×‡‡4!
`¹cFÁÊHS™»µùïš¸	2°ròV¤@_1+	™!ªüŽ‰ŸxT[&)+Ò¬˜´Ù^¨)»¾Ð¼Ehž^ÐÕSb¹*½ÝM©«<$¶J
¼à¹uÈ’Ò…ïê%J)×ÀªÅJ™¶„÷÷Cg¥õ«OkË¹Í>½µ¼0ò~!Åw“?p¹})±ðëä"]1ñW‡ÇDo-{iö…#¸8Â«'«gÉêõƒ‰E9ø; ÷¯žÔ¦òÉU|+¥í¶:rÖà/”ÿå·Í­Ä×µ4)õSUß6ß1ÆZ)	2ÈvÜV3`»uÚw—#Uû¸:LÒ|Ìlµì¬¢
dßËÙ*y›ƒ»)ŒÅm•dpºZ|Nci*Ç³†œŸ[dY>S$+Û*V¡’yœŸ4EÌõejCzÖ},‚µW5äÛža]ŽE}X°Í¦ÇÔh¥·L‡û'Wí›•ù¤!d û2÷e_MZ—U5ÿìí€(ý7ègI<ËCiÁ$4¡CYÛdfÎ`¤F˜_=v¯Ø«2ÛÌdÔ×{4™ÎrOBrºu™‰‘ÏºiÂ]š0=hðIIY‰¨ìÇÍB(™ð·ÄéC3ÖÒi˜Žý&“6õ\k¡¥ ùÂ*î> ÓTDWXFÁ|“öúÁ|>æD`ôN»|)„ö*lëÅ'ò88á.¸ð.Ù¦{dÛ‹à>gö~qÃ7+ÖÔ2g“‚(¨¶kº5ÛnsN{Íp2´žU®M¯bãçKZkÉ"ºÌ¦ž[ØË‰žŽ¯Õlúk›½ª‚zzv‰¨5ÎÔb¿ËÁ÷ä˜&ðè¾‡Ùš~ÁCö_þ)ŽòoB:mglÑçÝeß‹Zˆ‚KbÉK\_œðzn–[ªa­LÞQ4ô80Êj‚‹“Obð™ð:ÂÞ×’i`¶ž2ìÊžßÆ×Õ§™¹ÃˆÍèóh(¼Ól±Wkw]	o±œE€B1_TöÅÖBlec$¿Â•-w÷¬^Éså,ÃŠ@îÿVXs›ö•ÁY¨tá¿s«`®bUaÇ ŸDµÀR }ÜU—âÇŒ2ÎÊ%X¢\ˆz4ÙÆCB…­V«®g×š×A«L;I¾&”l@ÈØ§Ñe8¤"ÂüïAP[ÿc³b±«evÃ)D¸žŠŠÈÜøWZ¬{šÑ #%}žTüMeù¸†]37À=€?çG“…Y§„ùâÝÒÂÑ•ZÇs
pð©ô´ñnÛÕ¬eÛó?hOVÞêzík{ÞtŒ¸C¾^3"f¬&äÏ¢/­~Ìv.îÆý…Xí`›f)Åªî4‰ð™‘ôeZºn‘îè´µc{ô9éLÃbÏI¬§J—“²™m<´ÒR<u,ê•|j&‡Â“¶ÝÙ†MŽ=¡'zS°R·‰b>ˆöÎ¡_k@u)êw»ÑV“i– ¯o2}&ÓàEoÊþ+‰³¬(ôƒ’B…02œ¥,"1h“îál^$ý(y˜P¦yHštEÖI6Þ»èî2íÁE·ø]ÆßEÃa8Hš~´NåÊm«k) ë« íJHj¬RïXh3]ùàWÚ•¦c!í£[ÿÜZËõ&rÞ¡l;˜e{É,£	\Ÿ„üÈ©èµü!ùœÖPQz•”•"O-E¥‹pš›€ÑOóù*ÝÛÒ„Â>d;ÏÍšZÑJÇùå4£ÕºQ,SgÍ>	PD¶x)ôk­w·8‘X,£‡ìGÃ|a•=‡ƒyº#¸ÂåÔNe×¢”s:^0-®W´ªëJÙ5=Yý|¹4mêÙ4×G¾¿5…¤Z"uœð\™
ý="ì´ešNgBN`Mì|ÿÀ£È|`OÊ¹åIÊY¦âôRxW¾«â­Ž®¤œ˜¼I.@Zir!’tºeMÖ:Ï-)Ps­ŽI#wÖ×¶lyjþÿ   ÿÿì}ÛRÇ–èû|E‰­3téæ¢‹e,`c!yˆ°%m=ÇG¡te5]=U ÷&bb~á¼ž§ù‹ùý%'×Êû=«éF’í
‡EWå=W®\÷¥t<ðmÃx6C¯Ê ÓcuÙÚÝ`N´¹ŒÂækØX ´…ŽÀL^$‘)ÈaŒÏÐØLírÆÍ{tåhPžVúøêâ¸ üÆ`tX_Úá©¾–˜T¬eH£pwãÌ!Ø”“ê»b÷óéø¢‹ÀPëk&1Ÿ}œ‘"Ù\&ÂÃ[ªxÇ)2²þ&ÆÏ¹ŠŸËâ’LžÏ¯Æ å®ŠýÁfÆÉþ µÖ/1¸—‚`V¤-}/~ÍØøÃ(Ra¢³d™*Ý¤a®­œãæÛ«,ã»Ô	Åv¼5ôNßÏ²ëµßghŠ‡G{Š²áeM£.bD¡ÄCÓÜÄeêÒ´8·^˜0âN.¼HIº©ôî˜QÒI~cÐný¹7ƒ[yÕ`jN>r+5ÿH coZÎKæ_1OÒNå³™‰Oý8w½CfýÐ¢²U>¥}ÇˆƒBLSèm`ÔS2¤±í¸E1{,?ºP1ŒþÎ±|N.n‘Rì–2l>}©w	©miòXÛ9v¢Û7b§š”qq*Ý­žÉØ?v	OÂRr}Ï0'¿›8ŸŠhV¼%÷I>,{íï½úÐ¨ÐBùêêNMë¢gÛñÈ5IzË¹/TÎ1?“*ë…k6‰atÍ$çö§„	^—j
¿ÉÔâ¥"1™ˆ-Ð4“<dj$"i%	Z;,DºT#r˜dŠHˆÙï™"ûÚ°˜ éüÖ(ÛÊ$ìÈ/…\ UXOÑ¬²äÁŠà7ä`Ù =A78öÌÕ
¹¼Ûê[ö>,>Cª'³%¶«¯Ñ©¼QºÌ'î¥’q9º¢¦½v_”ëÕ–ÏØÚÛ(bÉÛvÁÍíN-˜Ç$¤i³g[Ö×R{±ˆEõ1’>»·üpÊ‡e€Ü›ùØI›ïOÅ	 n‹rR;tvvñV³ÐŸªwpîõ.âA}ªÙ+ÁQf‚*è¹Ýpª48˜ÞC#$€ÚýJ)QÌ4,&Î°P»i`u)ÆàD<Š1¤„‘\JArTZ›Ü­×¿íDá¯6×¹Ê|?‡‰2Q¦6Éˆ•ö¼¦‰æLR“¾%×æ 

y™ûƒF±¶‘ªEejšDÍähOSð¤µ˜ÅkºTÒd„ˆÏ¨Ÿo1ÈN˜;wlÃ+ì·<	Úb¶Ô(ñŽrzÂ­†ò§É×
i—(lHQ›;¨;p†Éœ™ÇÙ´Ïï= [Ä±#ÅM{+]£mz8¯Ùsf0ßÅJŸÊ„×5?¿¯(í…ÎOfÎÏœLzá	-kÚ‹],È¼Ûöšøâ×y™ùÂÓÎÔ7½û9˜û::ûìL~qŒŸÐìûO1ý…'Ýüž[š Ã3«0<1†ç³3†'ºÍQ³`,”bf
Ï‚ÍƒáùœL„áI¶e…ç6‡>áÌ¦˜c¹ÔýüüL‡áùò–<zÊÂV“‹4+†çîM‹Ó–nV÷>§©±s¢­RMŽ±î˜Ó~b¦Ç´ÔüÍáùMq9"bŒSdºª1sdmf0I¦½|"³dÚùç`š|ë‘ÌÅ<ù6£¸åyŽpAfÊóâ¢M•ç9ÖÅ›+Ïs´‹6YæcM(öùš.Ãó™˜/Ã3/fxRöo>¦ÌðÌÝ_„I³ºL³›5§/Rœ²‰™8Ãs3gE|´ØçhîŒ#KYÆð*FLŸY¿ù3-ÐÞ5d6íR -Ø$Zôá™ÅBL£á™£y4<^.tÑYM¥)é7“Æ¶ú€Û™Kcwg2ÏÌ¦éEBI}NæÓþuX¸<g¡\=rœfÕØüÝšVc—w"ýI‘ýÌÛÌ×ÿË2µÆeˆ\è	òžiÏÌ¦×ð,Æüš·<“	6<-Í°á™³)¶:ŠÅšcÃ³p“ì`'ó3Ëu3?ÓìàdæhžOÐDž$3mxnkªMÛø„æÚ@“íE Ýl{A#heº½¨Uhi¾íÆ›pÃóÇ2ãö/þïÐ”ûóœìBÌ¹Ó§ú%›tã,SÄ]áÛ9&¥žÙÄÇ—"OüâL½qÐÑ•ÅÓ
3%c€óv§Qiám5à°½ÎEv„žr\Ñ».@#Ó'(ü™ð.¥%mð )œ+›‚æ#Ijå<o&È«&u5lÏÙ)^çÇx…¬köô
o –ïÞ,)ÙžŽ-ÝìI/ŠÁÓÕq‹žÈˆc©.èN<£¸¤†Ì5E§Ó\œoBâÃ•¬Æ©Ô—\@_A®H> •CÆ[XÒ3ÌÐÊÞª*
o”™Ë¤dº¸íòk	.f\zå\/tñ9ã\},4—åf¿¸íÚëI/f\ü#:–y-|0yt{K¹Ï1Åmíâ@º {·¥mä¨PÜù©F@¯ÏO;”|Ò!pÞüÓ®ÃÅQïóätÿdý+Ìfƒå„Çç1 í2þ<†$î§ÐpžÆ¾ËÚÅÎÇ6jU};:ÃÔ…º‰F¦óì—Êr?EÇˆ ?AÇ\rû)æ,¶î}^öÌ_†î¼Fé°ÿ\š—¤p^cä¿¢C<šÕÚ6O ”ÁÆÓ‹Í`ZÿùDs`0T.«<FQÜ5rB‰Ü	0}_zà…§–ÃYŠ|j’_õD_EØu˜_íñ×n¡™5ã…öSùÒà›oA!{\š‰Gº–mej^cÐ‹÷”XdÄú{Œu`½±º\Õ;"ìâ·T-¼ºš}2æL
ÂTè&;ºÎþU˜V…îó #ÆèL¨È›j6³7Å19¶O<o+ÙTµÓæ/…A7:%%ó+&	¬Pû9dÓß¾Ën¶ÉòLoØðiçZ‹Á®2‡-°{$ë;Q^öiÁÎU5.¡Í»ŸT£ ”:Êúå Aë1a dû¬vqµñÕÆVm0ºÝâ9¨KÞ}äzü¾¶CÿâJ-ƒs7¯ÕÀ’_¬~õk½~Mµ”fMN´A;ýþ ø&ÿõ4¯Ï²¦I&Bl€ÂFþØÌ:¼Ößÿž)‰®µNÎó	Ù€†å{†³"Î]ç^­&ûf;‘¿”ñ¸JA´C¸iØPqìoKâ›M‘¿½MñráQÉMÄ¡åþ‘É’Áq¡-e'­Æð[—µÑý–›— šº§oŽ@8'ù°)´’C#¯¹j ÅÏ|û¡ºD„h–Ç;|zFµŽ¢ecu)nè ðz‚¢M‚·šä›NðLTî#rB}Nºt9»k|W”šºý_«rÔYÎ>E05;—N([Ù[x+tVäÖ¯ˆ]{çëf;S3¾;ŒuDúšôËÑñðbP4e%aÈJU¾Ník
 i_•aâ<¯?íH¨/*‹…FMl=&õEA¿Ýpˆt\
}Bç<ÏÏìÀ†^ôD7Bð¸6?r×ðl—ü@|+m²ÀÊP…ž š_kÆ„äWH“§ ÙZÊm€ë¸‹Ôl_ˆÅ1u+8ã“`Ý‰J#Ž;™¶!šÀVw°Q~5“«„6Ì¯j­å…ÐiØgÝ×Ž}»Õ½ÿ;[ëÔþ'ã4~( õ÷÷§Ò?ºwâÜ`«7ï•Ó†Q¥Uß’Êïô±>Á©Ñÿ¢Ÿíµ¤hY×úg‹¾]Ó¿32÷í;åísÚÖûFÛÙW€…òsÂÜMÊì’þ/«Ë,èÄ>¥«)íŒ‚Ø¦ˆ“´éf†MÝ×gÒF_r
$âæà‰'5øb]hVÇ»^¬ëÇ¹	wžøvØvçQ§x6ËÎÇþNV8m]ÿI;qÉ˜U’ô´j£,DwýñEsÖQ±ë Ÿ
‡þR1¤‚¤hŠßÔ"Í$Ÿ\4ô+ýIt¬7:]–¸T™0ÿÍqâ8F¡Ý]£ˆò¶Ðñ“%yuô+Ùeç¬ÑfÞíCÑN'_ÉŽºÄ$>D1ä³ê|œ×EçH~AV"ç‚h«{ßåÃ4¸ Ë^j¢‡£tz%ÂÿÀD BHáß5¹ÜÁ'ÉR’5!å,OÆ‘RÓ-Ä·ó)á(‹CR1=Ä#åýAª–)#ˆ¶õ"óül4YžîÞÉ\õ«\u[Ô¾ÚAu9Vù€Šô^ï½ÐEˆªÜöª÷ ÿH½Ó£…¹ã‚I38ÞU£®´•Q:"{¢D’îuga}n,VÂ›™¶ÄLr™7™ªZÍØGQ9!nºÂn‹¬Ò¾í]jûQš"YODÚ' MÕcN¬ˆ	V²ÓýÁ•Ãã\K“ÒÒ7ZAûXj‹b† l1*.³=ÒF'Gü×íŸ“Ãòœ°ÆYO~<2?vÆÚô»©3Ób‹!dã·V=,ðˆ¾lí[Ý‘Èã6„ªfX2‡¶ÙÒè( ŽXìÚGÕ'
Ð‘W£­ëTn[˜Ó§GtÔõùrkæ™ÔÂ[Ñ@"d{j
;kÆ¨ªÏó!‰¹´ý–µ%ïêw~7ìd£4¯sÆš%+29 ‹s[R<v©¤f1¨›ÁœN¬¸{Ó#vNÏbŸáÜ†­9÷4ì0…Ùæ³f›KÛý—UÄ˜¥Ms€•æÖœË8‰ ÌbÚO¶5jÝG+Hˆ¶n¼¥õ‘_´<
ÛAdö7„,\ßÌöG„È‡åo”Žre§-Ö6Ò·Á:ÈƒÙVz=¬®nÝ®À·\y®’âçÙífQ¶Û-ÐéBF§,m÷æÜ‰Ëta–^œ¾fòK©ÓEíÖ0dûÀÎÆW›*»ÑÎÍT¡)é	ÔI<%0ËæMñ+µÝØB²†ó¿¨­Ù·åo½Í °L¶X`“=Ö¥Â¡‡üÞQùU›z~››vëú{ï¥k%U¸?U–x'[®QNði lÐã0âlfü6-³¯²YŒ«Bw÷÷ 
2/ñw¿ËIg¹·Üí×YTÏQ‘Ñêr—Ît®#™ÍÜ¢#Û² v=˜Ð9C·ÔÝ¶j3Ó÷nÇiÏ<¾B|Ý[#ÿhÑ•îÇÃ“XH…Ç^E+<)„¥R#3ÁOÂÁ©hßº@Y¬uæ¦¹q%µ•¯‹!°¢>ûMDê'Å+œ*£GQGÆ¢ef ª»¸òØàB&>îsu›NNªjÒ’—ÒÂ£ð#6‘Ç.Á#I³²Ü¸™‰±@Vxq4˜åäùÉ©BÍï‘ßñÀ(F‡'ábèDòÅß^#X§èÃñR;ˆúáÓ¥œaód;KX{ÿÍ³TÆßÂ€£’¡r z•¾5;e6	¦¶¨›ªrÐ qª6¦?)G½³^s7X¦+Kœ‘6ÙOÍ]& ~áQóÆ Ãô·¯«UÉh^—#MTuFÎAv	ÿ39Eª—Ñ%êOÏ6ìÀ’CŠ+›â¼tú©/m3£M“Y¶?]=ÛÐÚ¶½s•Ñ¬~½7–¶ ‘]W5Ùòæì¨ÊëA¿ß×üjÃ»Ê7ãå,|KeVµ‡ ÷g<í•²AÀðVã^SŠìMqB–ìŒ,¹$ÿýP–£ìu~ª§³³ ƒg¨‚¸‰3DG´[xñþ ïg¿‘{.Iy6£þê,†Ùûãbw<f3Ò¯ðAÙÀQl!k„ÈÖ’Òz1¦6z}Q“Ù£f3;ÎÏ
ÔÅ5×£ct=€¤[Ê€Ñ lPà[Bc’á ‚ÐÇ±šI½Ú†ŒÍëpl³C3(±Dýƒü‹Â-êTÌ§…Ž*êdB24H(6‘»ì}cªç¨GìÃ&çI¯4pl§ž]j¼ƒO—Êø^esç¬—h·½¬†£­ÉÕO¾­?Ñ°Ê9‰h hD·ÚÏÕíOcrI@Äž~{åçŒ|^¾i¡ôž?%„ó“v¹6Ïó«Þ%è¹UíGwöË§åùiÖÔÇ[K«Ãê´êG§Kä>œl-é— ŽjIGúëI74
+‹ÆZ¡=ÆZÉËdû{^×U-ÒµM³¢ÏlºiŠ¶fr=,útÆÃm  ‹Úò·V¹™Èó!ˆÑä <’Õßéã`~(›	Xñ’=ï,³}Ýo³c?Ÿþ@Ž9˜ú:¦@þgß[lFLhÞ`ëÖ5ó€_a†Zî›ÐÝµ¹»8E³´ýzX ïÜ”§£Œ p2²üø¸hDO'åˆÐƒ%Áhçùˆ wX%‚Ò`@jxªÑÁÅÑy9á¨¯ÕàpóPê¢že:„J¸Õ0Î‘ÒUFXQÅ"?¨Ðe)MÝÐcbø<¹ÈNBªCí A.2Þ¨'ŸÍ×”:Y_ÚÞßËV³çç„ä{ºŠMY`2E›ÇWr'Úëâ?.J2OëK€H'
½Ú©+Ïá¢|<Á¡–¦±èŸö³½ÝõÙtÚú+ö«s;»­‚”
mõ"IÉt¨¦F4{2Ñ­#Iñöú5©w	Na³ìô˜Užu·yç-6œWi¹çÿøÏÿvþ÷¹m÷‚÷{&szñ¿8õ—"¯³Î‹_º>™™ÛJ‚I@­›E‘	éá¿\bQˆ˜¾°Ù‹_E˜pÃÑÔ™6—/~I€H—h–Cé‹_LøtÉß óìM©pcjQÀ*	s*lªWPÒ&¨Û2=¹~5¦–èèöuM}ÚÜÁ2žVX”*ŠhÑ¾]ìç6ÿãé*-lwi{‡°ÁÕâÝ ä%ž\ã+Ü½ËrÇ›×ÔÏ%Âc¤
qA)õ¶×§O§(îX >:ËŒâ¢sB9ÁO„ÏÏ±@Î!6Dè®šŸËa³î; >.4˜¹á­|X‘~ÀÕP¸?©ËóŽ³œš¬&ýÖ–ÿºÜõMTëæýý©ü}ó×šºÈž Á
yíÞ§N€’o
²ŒòÒ-·>ß‘Œ^6Àû´šd¼ïú]).œ×)6H,›£±ñRKûY‚õèL=;þä:š´—}§½Sîýízmã>s1€ËØ+e¯ßÓUh¾­¼ÄeÀ&Üãñf–%‹Ì{Tš‹c¼Ó¤3Ûge1xPž@p 	ÂÓ•õ„ gÈ¾¦Þ	RÉeþhŽ0–àÐu¦°A}<Ã7Ý¬lÈ–úVeÈ‰o²aDàJÉ;ÂÓýŒáM”Â…PFÈr”Ò¼×	ë
a’'€Sk2©üR@A\LnB^ãà˜Hd!ûÕÅ$b¨í„]<ƒÀiˆkÏ(®¥C#gÒc´k5‰ õ5Ù\.Œi¡AŽ†ªX“¿:9‰þö5z°“©ƒð÷u]M
Œ,’}—Fäƒ*žÞ+›—äf¯ê¯FXÝÐÇ¸­Ÿ%9^½IÕ«3LºCA6åc™+94I	úÃÈóAà‡"dï<„9Â	M±azp¤eÆâÁ÷é‰2¯’ÝaQOë’ÐÇ¦Í.&[R2Èe›3B}~è­Y®nž‡½ý7²þ £(Gb¬EÖÙ‡qŽ¨ iÄ¶ñ‡ª™tí‘{².4Tn:j_˜Õ]‚Ú"kU—Eƒû±8¯9yAvmEµ3[AÙúëüº %óºÈ ŠÉÖ”ä&^gMþ‘`Å“_†ëŒ ™Añ±<&x„à¢ì²$ýÈ­>¬Ô_SrŒÔB5¹M¢­K7]Vò¥†Êæ˜.P1èÛ‹’¨Àt½2ù=‘Hæ›µµÕ'ÈÊm0WËç-¬”4‡¡lE©È«sYÊ+†ß4òbàÐ”)ÜvÉ÷ÈÊfèêY“±…ÓÚôÔ<ƒ€?hö ²H×b\'¼¸ÈßÁóŠ¡²Õ|…(ªLO[hðÚÉ
4Àé'gNY®ÌÐÕì­GiåfEaL·ä0?ê,ïq¥£Í> ŸÌ#Œt–ž“F@Ì4;GÉUf¿!4ñùsL’W:—åh ©  ßÿ[9˜œeÛÙúÚÆC«*+H.³jH¨÷Îô€à'žgùÇ²ª7³åæ¼ª&gËŠ#}4ªÞ87³«žÀ.®ñòs×4xuêðo°³åCNƒG8æÐCO¡i›.ÔU„ÝD8ÏÄ–¤ À.½1¶y¯*ˆCªRŠÒ~”ê‰TOXa¬}1ÍÔ‡D@˜¡#tÂMD>ðWäÎ2µe³”+9Í;nÿ²(r1<)©”ÇZaîk	?A;<cðoªNeG¢Æ¿lú¥›Té ˜PªB½W)—lÛ/t‰ô«Ñ0x$„OÊð9ßš%nR‚æíÕ²…L.RÈ+Ó›I~ç˜·[^ç¶²‚ eF1BÄ¿¸þWD&ätƒìªñä¾hÊÒËq@—6·Æ0¤ì"ðc>¡ò/·¯Ü ÐÈn°°8Ãÿë¸Ç&¸±ÙNŒt~—­;Á*‡É{"[t€¸±U+)´“^àåäñUœæœˆ¡q%!ã‹*jvÎí–(Ò^4#ª<&Š” r)B¤_ãuu#£øìx§/ÜÐñÐ\èü¤)}=%4Mx¿ß÷•H”S|>Óm7tb…GÕË	2®åùu ªûüª'œ¢]M›’¢~Q!íá“lxJþ}¬p^í
}ÜçÚÍÿÚQXEdÒ-=Ž)£Xl”xÈ\—U›Ø-Êäh[æ×NdY~™—²˜‡B¬Þ#B'G¿ ‰”QÑY~Mër¾8ÀŸÍ
×`™¯ï¹•|ËÄS¶h¢mÙ!ÓÙÐ!FÑÙšiÂMµÔë,:3¹°ŒÄéNsm«Ê<Ûš¨ŠJ0Z„'Ñp;^„ñ"<­îýGÎˆü†ÞÝaµødMÈDf¥ü¶ÆEmè@lŒæÖâ’2IVŒð8ndŸ5£¼y(¢£,àgaßˆã6IÊ|bŒ•˜·0€t?b/žcGšÞ5²Ohj?+‚¥§µ„ïfÇj­tvw†ÛZ…«¡[Ž[ ÀDŠQÕu8X‚8+—Wûcìm)¨«e'i˜–u*n\7¼}$(ªÿu®ÎvädûXäÀ'ªxµU€—$ŽC‚p›Ô#H÷ÀÔ™oŸpÉýM	tªd‚4dzÿþ÷Lª
w¸Ëî_—»o×Þy¹ÿÀß>QåIª'©eèõHÏ”£	ÐÙHÎâv¡].{ã+Ôøàî?§áUà®Q%­/mûšôÙ½X(Žé?gAnëÆš)—47Åµ±•ŠX
Ç8™BšˆÛHáW;¯.çYòa´]Ž;a‚¦y™,Oí`#¾ POtæH"üó#pœ™í¼FVu-ûpf@‘kwB˜$vAÇ•¡¨Zq…°m¯Ü§P±‡ã9W~,F›Š Ày¬<ðl(Gï
­½£Î½sö—u	¹œ#E*†éX-[¹<…èp÷˜@Ýc[ôî4}1@É¿mÓ÷÷§|†@=A¶X$›˜á˜_ÜP™GvH05­,+‚z5)ë¦G–Æ„#Ó_G`õp²¤4 xŸ•Ãa‡]fdX|–±àÐúÅ¢ÎBå ¢äG½ûJÑ÷nF‚ù<šîM4êõjà”‚QMœRV€<f"•œUw§4ñŒ
3ÀÛ¯³î-x+eœr„L;œ(Ü×
¸:¬¸…Kˆƒ“Ã8hä(¨nM>Âô¾"¨	ŽØuÉA%õ$"NdÞv;‹Ù¥ÉÞÝpê¦~(t—ÇÕÈ“ÀÂËŒ<ãÁRFùy˜´qßŒŽð,±}|xXzzxPŽúMu^tðœ>üC&ÿëï€Þ."Å­¬"ÿìÕÕÂ•ê»eWáKÿÏñDÁUÈ¨©T ØsŽ¿ë½}¼öÁ¼_	þ¦P@ß{ñÔ¡FŸ€*wà+eÞ
^
¶ñC‘,|m JÁ ²v]Úú€ùš@¯d=bøÄá’™à|÷˜À·qlp{|€m<;+>ÖÕ$
¦¬
¹Y§,I)wq¦¼z/‚Ç]Å¤4_tÜ?Ô÷¬e*Åí-|­xJë]0Ó$|í:elærí6=t/ÒÉ3Uê†ŸýÁx¦:zRK`G`0¿‹?ˆ¾$zE÷I¤ˆè£ÑE²}/µ j¥‘G¢x2‘$j$’J¢¼£UnKc‰†Ò(-úÌ…Þ¢‡ê¢~¶©¯X*ª>Ê•¸Õ‡LmhLL¡vFå¸:¥±ÕGk6Öõ¢Tx|'vrB¸U)>`aœæÏÊéøy”–‰¦ýÖCé$È1`À®¬f4co}‹u>Í)Þ©Ñ‘§£æËjRž0ES³ÌRË©FZ®l_«	˜HéF›Äžp!8¬Æ<®‰6PŠ°³5¶ìÖxÖ@—-§7šþIy^éï|I0\ÊSÅ j»wêAB/­á©G¡C”ø¦ï%$»à¿¿UdåÈC_[Ï@°F4ü4Ž‹KìóáÜL4t; F½ƒò|<,¨á`­n¦QgyE%(N®í÷Û)¹AÉÎÁÙ&%ÑSvÓl"ë «üÆÚÆ£ÞÆãîòÍ;­*„Ý^\kï;å A‡~ß+†Å¤è,‹˜tB+)¥Uc®æ#\Køÿ’p"ÏÅI~1œüLMC
yyîôAíZ3P¨ìKv‰4Ý(%8ï%}“™›-€&<“É¶¯h«À%JÆu?E­¹µØË4¡*9C* xäfM‚,ªƒÍ#šX9{I_Òž6³ÎÇ|ˆS#ÿ¶Ë$J*ø {J $:öE ]¨6p#WÇ$h, ‰Pt¶¾m „î$	-«Ã	KŠ¤Š.GuBKA‚`£^	ž­§ÅöžÇµù¦iý-a­F`P&=4hwFKØ{×
8hÒ¬4àÀ²pÐúqàà)², ÑÛ£mæ‡deLHÿEÔüÄ+F›ÖJjK§.WR€žf#Nz,žôlŸâ@ïR•»í¸;= |ƒZ
®C¢¤IƒƒÒò88´šÌˆœZ‡.‰8 üa¦†	#H Ë*ö_Ø¤ãmòÏ(l‘îµGg2DÍ8Î¤«Õñø}@® æ]óÝõþ ³zŽ¶»àiC*ú_Sý])v&ëRÜÃWQgE­\%é0ª— #[Ç¥í¯‚q´W·F±~èñ¢XyÄ‘ölw%{ùúg‘¶@ŒËbŒ+³R&!]ž›RG“2abóR¯ûb^Í\ÿ;¢GvÚQzÞyvE§€ÍŸ«QÃ¾zK.Q?c£Ö&Z™]1¥Éû·÷§"AãH>EŽÆG¤G„Ž´i8Ñ1 613…9ëÞ{µ…v)iæ´CsÚŸYv'¢²›×™cêìŒw<éÚãÑ)ˆ$‚Ÿød"‚WH #$ÞpQ¼! s±˜6`¶r[Æ…j˜,0Ð‘¥èo£”‚ïµÛA+Ö#+Ø’ádw}§¯³Ø“ïE%ÞîpC>ˆ/~YÔÏÀ¡«Ô·£%ÕmÎòú<ÏG¿>y¼ö×S‚Pæ>RïG½–÷‰RÅ¸S´ŒÉñkEæãU`÷_y~»T½ï°¦¤†6¼(†Çü9×µfb8££¥ †=×1›£eLßHVAËo-@Ö¤z—–ùj‹ãŽn—V)Ž•*ˆ'™EÑ¶¯
º¬B~„­ `úuôÞß˜5†Ã*ÙpwÔA>«#–¬7Þ#¸Nr°qŽ•“µdoSt½™Iø	V…mäU9b¡ö±ÊiJž´KSEkéÜ·LýžÂ‚«øÑu}´~Œvã\åÚ$ÞžÞSñMáFÇØ²¶ú–ªÔo{EƒoßÝ(™Èð³“IJ°—šžiòQþ}ÃäÇ[SúïÍÞ(=g/ô½÷ßû”WVwr‰^ÈÐÁþ`kªü¸Wõ–4hGï?B¸œ”5i[ù
¿q[V½O[±±&ÄÞš{µá¯-?«Ó
êÕsŒç£O¯žã+M±krñò„X.Þ¾|ÖÍX÷€–ÝÓ£yn„v3$NÓ}28›ãhDŽß*òÖBÉçž)ÔEñe«ÙO£r¢jˆÂb~×äú¥8„Þ¾5¯5Ž¬¦kœRrŠ¬œ¼¢ `þÝHî°sÌDÊ›œ*%#õ7¨¦fEFÀ™i½#E“Õi¬Ò×˜Â54êEôÆ,œ`„ùNqeo¶Ù¿N¼Ÿ\aNWÂB&÷6Z´¦€L ¦ÎÒ¥®ôƒš”û;F×¦:ÜøG›ŸC,I•zÉG±Ž ÷]þM¸²2)4¥}Œš÷…@×2¢ô&m€„[d@®,[6ó4Mzë"nãš¸àz“l<§0S;Â¼	ŽÆÈ-Æ+`;[sÄ'’YÝ?N6!švW›ÙèbËE­D%Ø‡
CÞ8hœ@yþâOB¦T¶’Ôl{peý¤N‚ºG365õ‘QÝ[ç´\2ý"R ›6Ê!øi–o¡a¬NÈNúô”´ƒ4ÑNÂØŠ›®ç“+ Ò¥ÃÊ©sà–ð”Õ/«è€CJó¼è\HÉQâ)%»ªMHÀÃñ×óXÁd’ ËVW!;s¯Ç`ÿ–)ƒÈ:Õhx-lXŽ> =\Ñ ™¢Š’¹nÏ½›ôÜªmoeÀùpžÅ³ä€'„âÂ^ôeGû
µR"1Wð”©÷´”©>àêŽãâ¸ètš‹ó•ŒÅ²!üüWÐž„•l-TÝƒÅ½	t#¹sñ¥Å9oTG4Oc ^“ì»PÁ õ»Újß‘y–/ Ô‰Â®cüYßÉ,•h–Ëö"…R3/zl8oQõ°‚jBðfégHEo½N#$”£Ì¢®ã_ÂíHA·Ý’ý-2&)qrÌñÙw!$W£ W£W`ÿGÅD!HÙ¢ÓºYO…7ˆÒFØ1'ÕõæzY4£|ÇùîßY"D`=+S·$}ïOù$ž"9¹,pè×kŠ%>ñ„Éå±ù3Ñ†)—0ï¥ðd¬ƒŸÐ?”Ð"hv3ƒ!k{4É5U‘Æ6æv¡gtJ¹H%: ¥º·ûÿAˆõõtÇÿJºßßY@HÛFaåöŠæ¸.©Y¡dÄw‹õ G,J œ9—A<Â"h Oy(©ü¨.ò=HóÑæLm·áÿˆLzË\ï½ÈTkpJœfüÃ¼øÒDVtúÞ
x‚‡S½3GŸ^öNÊ	9ÇÂDù~äyëO!1Ì<O²Üõðª(À~?agÝÇ¦Â¢{¨õyä?yâÀã‰•”câ,y™bwÔ$­;šÕÁå3×ŽÓõëK¿'XY~Ì²p˜åMKÈOÆKóî[˜Ð'àv­‹/¤&%ö(Ï€Ÿø2€8 ú±•
Ê\‚Ž±¶w:·”ø3aª§wp0Ó1‹G3Ó•RWMÔ/Ë„†?¾ÝÞE ‹ãÙ{¿¢ÏÖôßYl¦´¯+g•ñXY)Ê![OŒƒéŠáùS_ÌŸ˜ÎXö¿1}ªµ6/eÈap=ƒÑÜä„­	8,Ù=)U3¦‰¦˜ß‘R•‡[¥W±³€Ð_;¾1ëv¹TÝàÂ®ÈÊ»F½”øÊšÒ›VË:ä~ë:ïòÐèÀiÜ
)š‘qúÈ;ªR ù•×dªeŸÁêÝG­u¢Çñ•‡Z‹'D•Ô‰J
8‘ðånz0)%·•¾¥4;¤ç …mJÂyUúˆçRû¯©è|y=IÔ‡’T)ãðÌ³À\Û@10+®Ž‹bÐdùÇ¼¤áT°ü½å¶ô¹3;r!L-'¬;ÉŒC\J³&Q4÷.[Ö˜Ó‘}ZÍT´C³1–³½uI]]Ü¼ád55'ÁXW5Sp›è­ÈÄ#d‡Ëfïù«ŽUB\¤¼s ²§Åè<†Yãe$”{êÆ,N4ÊílqÁ¾ Ÿ4¯7÷‹ræ@ÆÎüiLô»‚]lBôê;g¤G—k€l~¿ßW›\Éx››™~›¬d\Nà¾’IY+ý­	KÙ«ª BG’xL6çÑPCè<PBèpš—ÇçwQµnßôdƒrŒ¾óû±Á3mæïÄ†#ò8²£¾.€ æ
À¤æ
ÀLM þýƒ,å<b +l´¥¢èž¯ uµ®òq–ëþ!ÀŸÃÊ<€ÉyÐ-,þ{zÉ-Gà^óm@ÐÏC²StÈp!«.„æRuiöÅ'Ö³[çgIi=zð˜ûƒ×åBŽy'¡5ßX¹¿I}áQÞü1€j=×[P=’s½
õ³n¡‚/úÜëÞv±;O•ª‰›/·¯>Ý„ ÉÕÖÁ‘ßþ ewó„|fs2O˜gMšÐþE‚8(mfÝþÚYµ¹0Ë]JæŠ°«y3ÝÕïdAWÃÅU¡3é˜zŸ	BÉ´ã²éõs2A–»zJ»Ò‰>“¤lJR¦ðr_.hŒÇ_Ø>’Rì…Q[Q—¸u,~ÊÕw82±¾¡ ª!W:Å<Ý'î¡™vÅ¾ê!`Wo±w”\­|‘Ûm­¸2ZB×aöùaÙL–p¡Åg‡A`Ì`g„`+Ú¯™_µ·¦=ÒeÃ^­Ð¿×ÌX¿F†"îâíˆÿkC¯2Ù2¶[Âƒs¸œ‹ä,šë@Nƒº’¡RÇ-¿)x¬Ú´<yð:é¦Ý`š¼`©<•ÃaãˆEƒíx®0á¾R¿#eÏêÂgà"áÂvæÚ "Þ2ÒÅ@Î8¿.ŠO:¬ç;€÷zk7|tŸ	àÜîº8/Î«»¿.°W×öJþdŽÇ¤7t#»«ä[ƒŽñîo3Ä{hŒâf3O‚ƒ}õx3,š%l'{az8È'jh“=úÓwyþÏsKö–»ýº€DÇE§Ûÿµ*GåÕå.µ°• äTjÎÓx®³½b’—ÃÆí€åèuÆâÖÁãl+£Ø€Š‘ÇPeûC‚íþt,MpÑ:Û{)ý¯3,’ƒ1ÈK4˜Ò¾”'Íñ³j ŸÛ°¹¨¯•ŽØIð8—4óÙÇ/j9€».jë²¡û õ9;Yš˜GÁžuÙ‰šÛŒ0Nˆf¶•#jÑMºÔvÝkÒ¤uî6oúÆ!)R‘´ž$ŒÅ„\D›(Yy£pL"”ÒƒM ÉLÝk&,šñ– W]5ª’‚[“ÑŽíg¹¹ÔÒv‡Táçù¦›h‚Ÿ b±"‘Þcë[V§S ³„³Æ]ÔxN,Ñ8;fmGés8ÖýÒ7w¼P>Åî®B—–ê«Ü\‚&[ÍÔ¨4à•±ˆ+|ŠTSèôøuj+áè7ùp'üF¶c¸Û8¢š½mX±lÐ5¡.¬8,@{€öê®#n!}ä-šƒ»•‡© ®þàºðÌóbýs“ÜME©‚Ù" ¹jÞm$rà5ß,X„81q‹ðHÆ·•¿#†}›æƒ ÒØížÚÎ—Ók	þ¨­ÂR]5NØ%td/-½ÞšëÑq1 Þ•\=AjÎéuéÉß÷¨VÛ*GÃrTô|¹çu·U.´ˆd7m¸ajÕLMeŒJSœSuöìÒÈ4O“^´+ðáT ð{™Cdb‚÷¸ÏQ.âùM	–‚!RÖüP·®ŸF0SÆËwXÜ(a¬¥i_ãÙ¹?Šp“£/Ä6_ÔMU÷Æ„y' ®ÈÝ@ÖkœBH«t‹Èˆ³ÊÑ‡pQù°Î›³<ÈdÎàÿVRx«z$/ <…œÅ”ã€|<®+‚/Þž¨µÃnÏIÒfˆ&Ò‚–Á³‹^cÊ¦s€ó](¾Ðaùþ;}þ£‘v˜(sËK™©³+õ±¿.qv)}ìÙoð±w“YFW®l†ø}Õy|ZÎŸ¨? ªû_ò&±â`æÕw–ÚNÄÂÜŸÒ…}r¿¼Mwu‚ÈüC?½CŠGI

S@`¶!¯¨0w‡Ðrnð5¾m—½Ìö÷ü²Î÷ÁP7ïºíešeº­BíB¼Hcã¥D¨Ð¢má?Œ²?‚JG}æáº? DþèxxAH2q}Ää›Ø(™ò/¤Œ…_}ø¹iqö/+icž>Õ´@°qÚº­¬Ò$³f<Njê$‰ãŸDÅS/«[P:­ÅAq6-J¥’?’ô)rdCÈ³jüº®Æù)*IñLI)HHçðoo:sÈä, ÀŠþ’ 
Ím.'ÄÁ¼-Bîë„CõÇRÛn\y‚#Ð‡qïQsÿ1dÿøÏÿÎp!$âÏeq™ýíMösQ‹<ÂÙk²j3æ‹~ú·¥ã	œ™KMHÏÞZ„@d¡TÔ9DCdÅ2MÏIºhFèay½™O5üô2Ö!²iRÚr÷‚XS¡7m™ÜÏk¡8k“Z?æ“³>?ÅÄä-ó†úžE¡œ6“C_ß7“Ôfpø{{n‚AãÌÖì“l¯CÿŠõï»8DDº×ù5zpÛ-“¬ÑÒ(_Æž>¤®/æ\Lúà·ÉT½è¯µàÌÄíªÑe^Bµö q•¹ -œ E6kœ‚CXÃ¤àGí¢Mó5ƒ,2ÿårJ4¤¨øo
 È¢Ö%›j¿æ™µˆ"kŒÇÉîdõÿÕ½Up©¦8/©T¨‡Kè·Ù×4%ÂTèŸ2¼s·«>¾¨ÇÃ¢åÂp˜u6æ»øtöŸrýñâØÚ¢ËÏ/þûs=s…éÂ˜u5…üIu´ ‘ðx‚´°Ç
áëèz‡n€%ˆ
ä%¹Ó‘ý6‹1Ê.ñÅ¡òÙ+3ÉuæpvI»òTÑý.¿Ö¡OC4ñÅÓß©Ù_~¬„¹ 4ÓwŽŒô¹•Žý‚Ð{ `ˆÌ<¤èmÑœ—æJTåƒµsU‡ÚEªËÁlÍ.ô¯	fj,y4Ä%nhmÆ†¢|¶{¾ó³/~™ÛÊöQ³Û"Z²«i|ôŒs33¶t‡[fþu¹ûvJg®2³•4JgÉUÿÜ³—«õ%-,Ò¢3áð/æ1ð™:óaÝe~!ø%¸•äKâ&Tß£,…¥¤XR¸ÐÍ8‚£k*j›ds²ãIý©Œ‹µgŽÌSeu5Û==­‹S0ñP2rˆ¤ÒŽ%aì\Ì‚œ½i[ÅàEUT<›Ã®¶2Ø†s-!PÉH¨f{4¦e$Èµ¼–Œ…þgmXÃéq¬æ\Ÿ#£3“ØCt–ðg=A*Ó¦á§áD)‘¨w(è²XÔ]GÚ	Ð'4Ó4ãjý¢fî«ŠY:´ßžð  øÂ4fÞ\
Æ	PóÜú¶&øsO-kIlµé¡‰ã¯"y‰f¢ˆ§HÂ:I
!6‰ ló2³xQ{ÆsOÎo`nÇ_¶9_, ŒuFdYIò`Fj¸SëÌ aë¦¦Ý	SšŒsÞD=ëtúhš¹)iÿ¤@þ¤@þ¤@Ü-þIüIüIÀó'"Vò¦@ÌÄj‰4HDtö)2 µôXïìDJª†„EÜ™:
5f;UGb²Ž?PºÛ»ö,–­Cñ_\º0hc‚?%]Ç3±Nì›»¢+Ó›1¤ùŸÓ"eLWÿE‹À³{1©zùÇ"Û«ó“	K­–ýËªk0Ó[xi8Ã^¨qþ1*ë1ºZ6d˜lÀJú/ŒvrOÙ‘@öÓ1Y€ÖQ5š[}’àšÒž:\jL'3ŒIöò:éèÚÎo‚ÚÎ§ÏÎŠãÏÊúxXøœf´nó)Q×¥gt(ÁÛ|R}àÐÂöó°$7ÖNö^…¨ÂQpÕ’ƒ,‡¬wžÊÔ€X­NÖ¯¾V«z¯½á1K×p>ÆˆÉjšÁª0W¥ýÙ°È9Äã©òÝcÚäï¨„Ná	KlãÍóaoÃ7Vÿã )ð.dÈÜ1àæá¢!¦e¾Aàô|úŽÁ­ðÁÀÔu,Ö $PÔ#Tg'uu®y³U½I‘áÏ9!f+ÛøOÏícC˜-°Ø–Þ”h|ç>¼ 6Ã<Þw	ÞyÆ-#ˆYœ]¯ŸÁÈ7 S+ø“]–äA÷²,$ÐŒŽµŠ	b\1K™òî@øóY²DŒ•‹BÝ\’«×Ü¹EXüãÿýßŒSY´¿8fÐ?ºOâ,,mÂÉé³¼>®ê	5³’1	Ã?3Ú5”5ßðpFÖ­lš2©ŸãÁÎÃ ³í„?Þú”·ª§Ót—SÃá4ìtÁByšî ßv¿È2IÂ²ÜmŽ@ þ:
±™XñÀCvÞ¾‹T:~X®ôx®gŠ/Ÿæ\»¡:÷}Ít§ lÍ<Ðq½§9u„v:b$÷¿nèoMIØëàùŠ¹ÙÆpÃB€°jœöø±äÃ„õøPëMóÚW°;á¨¼—}Û[-njuÓ	éë0¿=÷9ß«ÞÕÔVîvÛï÷ìôb×–Ä¥/Àï8KohÎÒ>|ìKŠ§ÆÚ gïZyñ„ÛXkÑâ—¶YP$o”xÑ‹?¢GFë$ìËOT.
æ†®`í|[¬ÜÈ5Ã°STOƒ<))ÃQÙù3¥ú~Q@xbšz|ŒYÇ»Vtï÷ÅÕ¸Óí{÷!ì×à†ü¹¿ß‹ ßcô›Å’Ë:V@±n8¸˜ª7 ¸,¥Å?€ý˜.ò!Ûh'ØhfZÖ âÃþÐUä©qµD•Ï›¢r¬<ÉF¥E…øÊ§Ò©3³I!ÓèÀe9’Á ‡„ðMÅ§! ·G"¤ºäokê“†Úà2Org3Ä!®n/×ï÷}]­dÓŒµÃ¹žþ]ð¸U#Ã¤Å(¼bO¦†d‘496%o:2J&‹äÆtÈuf[MÍæAé* ;cœ[rœäÊÃcÁ¯Ý GˆèÐš±!EZ~Í2Ê>¡¾€ AžzT]ê„dÙ ›É¨ÂŽ"l/Y>ºÞ="`ð`íÓÑö›ê¼ ËDö¢nŠÃ*ŸtÆ}¾¨'Ï¶³kä‰ÌMïgãÑL‘jñžR7i¯aÓ×ÅÇfûÆ^Ï0œF¨ø.ÆšaÒeAmÈXË‡{ËÝ.¶Ü_4gìMJË÷Œ¦Ý-ÇfÆçF.u25ÖÆ ¸zu’8H“T‡ß[g“«öãÞ¯dëÑFbBu%„ÜKùïßÚËIß¶ZRÞ‡¿‹[®kÚ˜îdm™rZŽf!ah8%Eq¦5T®E"ëŒËKZ¾µpÚ0Âuk*þô—ÐD^â€«	“\@g£j?)]›FÙ~"J5GxÞS'I½sZ—ƒþ°†L*,Gæ?ºýlëëO¡b€‘=Å(¬D‹çW[ÓQq™Áè:Ýþ¤Ú?xÅ`¸oÕ!úVÝ„Ú©‹ÿ¸(	Í*Ã8^n@:“µF¶8EÄcÕÂ)ävÐßYµ’üß!Ä1CŒÎ?þëºQÈKÄ·•C|9 S5›ˆGx‰DŠr ÜJ‘ÍAÙTºŠa2ÀNm˜DÄ Ï:$·ÖŽ„\\õîÈBböÁ)DâöhÞ‚üŸ• ôŒšñJsddÛO·dã³2s`æB¬ÆÉUuM)[¸ f@[\ÙËmØ91w¶ÊÉlAŒ1ˆj5Jp­á•UÓ÷Ñë0»?-†oöG_” åYfJ#XD^¦)|Þl€¼Œjÿ,S—Qí˜spÄ‚Jˆ9æÁÞ;Q³+öÑA˜LÛ†ÆwKj»Où7.1že,®M'$XüQM.ë|·™²Fî±Èðž«ÆŒ”(ô»lŽaÅI¥DÉ.©Œk(ôÍ¯›7ÚœiÀ¨‚†aÃèzHU/_¯d8°˜¹E–Ã¦ð´ì˜6¬`T Înô[àÛ„0‘,ZöÏ§“o	iñ!šä'Îþq¡•^ìŸ)Äâàîj¥´xÈ­F<Œ`åÞ‚C·88VÁðH›C šÅŽÁÃàÎY†%À +E+Jú˜`DO[Ž ÅËô&gE>ˆ†ƒ›Ôîè=Ú©Ü©"Ë´:±¤ÔedDV_äV¯/m’ãütur6[C˜&“µDùÔ”¶H™:À.aŸNŽªÁu¬³i'Í’1Ë†ó%ý¾®@å—•:4dm›ßó8·k^;ù *Kr¨Ia±ü*jòg~|ÖiÆi+'_$J¦Aƒp×Fè““x8Ñ8B¬ÜNs·jŒ¬;4¶ÓŠhCüHkK“¯¶p€iõÔ}ôSn€
qÒ\yœ¸™„Å¯´¸ÄØØ¿d„mLa{á‰³¾©Ó¡Wš2%¦¦n7'vZ|ÓÚ˜ã´bŒ:–¡·}º@#7AX;¦ƒ±† ò4°ciÚò†‚þwU5,òQGÂ?¸Jô	ÇzÞ‰‘ ôI‡æ¼=0+hUÙò¼ §ìw; Îg€aß»f“¿ñÙxb‡Úƒa¤ŒE¾O‹cŒm¶ô£Î9µ9-ªº2ÌžÑQÂ}™%?Oã¤<	wZ¢^ÖkŒÒ4qª•4þóÿõ?ÖV $	é¥?ƒ-g
egHuRÃ#«ˆ
/>Þ³‹“™üI@"{÷``2sdÍ7V!<“H>í¿\(\´ˆb®>4¾ÿ2 Ã`êÌhØœe3×ß‡½ÂÍ~ô­Ádo˜ì “pÙj%ìÕ½ÖO¾¨œô:†ô½pqÎj:º¸ÃüøÃíÖ‚Š3½:-žuÞ~Ðê¥¬+ø6$Ç"Ÿ÷r>dK„'n8×k7p81F¹‡Š´Pµ4¢Ïœî÷˜¯3+Ô¨|“¤PÉÎQìòz÷%ËÏ“dt†½'ˆyé£š¢%çÎ.ù(f(—ÐkÞ€aþCÎÆO0ýgdú°¹<4µéîwÏöž¯o<xø"–È—=úû9ïñ£§-ê¼½‹æH¡Ü9 ó™r‚i«|\€ü=o`Î€¼±±ð »ëÿçÑïšo•Y£eÚYì÷æ¦—Q“V!er^hbŠh-¯‹ÜËÌ~T6€û¸U§aÓ)[õo`È°SñÝO¶ïÔàÿ5a8*‚9ª5Šd¿ßÌrNæ¡uuÙlM7¼¦/žsá=Ï‡§«v·m;JoG„å…å•&‘)¢<=®Ò2âKH6»!˜þI9*Ž>œÉÐ!=†ö~o×ñ’¥F®ýwN0Û9Ü>Ú-%ÀˆÈô¦ KôÓè³î†ÂÅÁ“¢ƒ×C Tèµá… †ØÏ¯­ìƒbÈy#÷¬Âº3Î“<ä™ËÈŽ“‡À{"Yµzpšlr–“Ïu‘*ÆROàhå£kšt÷(§)<Y^æŒ twéÅY<¼Àq(¾€§ò›çÏ³Ÿ^~·ÿÃÏ÷²Ý×¯ß¼ú™üqø*{ýüåÞþËï=‡ÎˆÀsîGCž±døÅØ¾¨4Ä4±‡‡„§»Æ#â9?Ž£å?‡A?†kF±váyêlºn!þ‘Tç…iÌÁŸð låI*î7¼*Æï€D¾4†‡¿^ÀnÂ™—FFóï@È ‘V<†ÂA˜'ø))v4óE¶‹ÄzûˆmOÎÊb8À°í¢v…ƒtx.b
»eóC¶4.xõ ¥Çûâ -YÇB€š.Ó´ßè
›’ö»{gg€,³üî"AyòHCpŽVÌ¬å˜²	¢î‹zÒY:„HÚjxmò3zqp-‘;Eíeû8Ú?*%HNV)"ô¡Ê• ›cèìÊ(©ÀžïdìÒÂ°à®[K¹‘v–²Ílé§%ÀÙï,­d%ÖîOèVÍðÃ;l7VBP¸£`êMµàJv’›"´•-P»ïäÃÞºÑB"(HP$1dßH®m¨²×töÊZ/ù¦ÀX>:`1ëÞ­1V7Žœæ…fàSv¼’EBV9-Ì°Âwn ÂOHq#ÒØ«­†žábÃ& éD éá)ñ­:É9,™]	—„rN¥´_$7Y“ÃÑ0áA	 åÙ)- ÓlDV °z„¬zÃÈæV2ÞÞ¦áU·’ñ¤à ²’É`ÿô·¡Ÿ€ñ ˆÀL=%”SÂ99rZ0Þ[K.ÉeµKúeÛcá«¦,à;†oi¬0UiBÄ£jh<ª†Æ£ŠÅž
Eò6îZq¿ç­†æ­&MhýÀ'Í‡OZŽ›¾#ˆÞs`	Œ¯'q”êþ¡àœƒÅ|!]Âß¼`]…h7‘Ð-“[D\æÏÐnùù6ôP±È'¦$È2O{ïg9Ì¶ù±QÚŽœ0šÅŸY+c‘\w'ÚV4¡³L5¤³~pqÔ[Ìy×NàÜ®7ýXë§þ‹8âÊ‚§s-Ž¸Író:Ó3M5Z~:ã”È/,HÇÜ:óq–¿l^ÀÍšû"ÁW7Î4ÜXk¸£r7ù¯§y}	Šf’í0nHe”–É›Ùâ`õéªK/êÛšû-YI$ÀÛ±
| 5ãºÀPG)‹úÜ«Þ‘~%Šökˆ¬ÒP¹fLÔÇu‘C´×-”s±ýê\h¹O±\1øîÚ9 *XK;ýÚ!DÓÛ0Ål2Q*
“ŠbI›­9³¹zúU°'ø®õÃ1¾n”’xð…ûÐ—‡TéÁ‚	»„“Ý„jUÎ)j ¹ª$u¯-%vqV÷ž¿•¦UCñÊ‚A8y›:TÇŠb:þÆ•Õñjì_N/†yMa{©FÃkÚÕä¬(ë¬º	=4ôG®`ÿ¦ê0‚“‚–wúåÀ)Öåþ$ø|vSTQSCò¹Õa™_BV²ÆLA˜°¶îcPŠ$ÖvhG:Ž.æÆn.ŒŒr©ÂFŠ	rª‹I†(çÔ«+t… oá%´_N–	ŽÎv1Pªcvè$PÞS2>ã(ÃäºX[6¶}s>nŽe¼Ñ±„™Õaªhô½~˜ÑÍEm«ÓtÂ£Øá÷‡’Û‘êŠÝ4KrªZí£„S<(&ªî=æ“	ei|m­ så!\w™B‡›$e["µ¥–ºØZ¿Â¯r”%7R>:.óá/À½±™åÆ‹_œUî™sîŸåM|¹Mtî™áIY7ªyˆ>1B}”þÖìGÀôÝRQ^8µä´zïT(x°:Ü¡)•-wA><Óž&¥…˜öàQëÛ•Hi¶§Ã9dŽÌKÿÜâiÐýt¸#~æp—Œ¡GJâfìDüvôbÉ’ZtÃåY¼þÛÕ)KïFaµYGÊGW
>Ãœ´¼Ï|bÚK×ìÔ±^8ßæ:ƒÇÐÊnM˜Î>¤RëpäeäÊyFò(ƒ,¿ìØÊçRÄoDžì†ìàñ„VaUÃÚ=
Âó|Bà­f^V¤&"2¾êGô-ÁÿGìot¸ºä¶ørtv-½¤ãèz„xÉ‹ ÝæAn¬×Ñð^wbG[5¼ÍEã‡¬šŠ¶PÅ"=@Rp£}„3h]Ž™¾2{`ø<Ø>sÛ3ºÈE¶â£>ÍSŽ/Èß"²sçY˜‡Ë¹>)äðy§’ßGs7Á”|²ðH£Ž&;GýšzÐ47Ý™Ú¶vMa#ý ¼§Áü;.±ýQ>ãX\hÍ‰JÖáÞDó$¤&Yõ]ÍÍ¬ÁÍ\ÉX8üýÁçáÓØTl›=×þx´>ÙÁÁ</ú¶ÎÚ5%Ö=Mr~Pž‡Åù(?%8Îê2Ì{Ù!*R1¥ýù­V¦ŠëMJ;¿¬–W2ÀnÒ2|CG´™Að`jçµ€Î¬œ¦Ó÷R ‘E“m’I¡–ÍX„Ål§\@Š1G5ß¯×Ö ©W®Izñ
=aÎº._[-ˆIn®{§®®„}Þÿó ë=r&ëâ#A+p­ÊQgyu¹Kƒ©:fxUéO˜Úç¿ÓõÜÌ`D“Õ”	7¦Í$žë½$›©%“Rcƒ54X=¤¼’<­°SæKîUY Çû”ð¯#ž<ìåêîò~nBÛˆ——²‰¯ž†[ØHÎÒø×óBV»™ÓjËÙ+K®˜$6Æ‘å÷ÄÛwÖ>¤e OðŒåž¶ƒ«ž‚€ÐX¬·›ì9•Å´ÊIpž_õÎz62°ÿ<V—d´ùÅ¤2àÆ4{-·œ’Ê ÇœC9H³—/| +‚Å•á q(`#õçª‘P…8x“Ö‚€MŸ¤ÐB¸ô¨@õ™µçÔ1Zõtã™ä5G3®’Àq‡¸äVºÑpöq¬/FH>a8åX¦x@×Ûdr­ñ?î6<ÅY/ãGþ†˜ÐDÎ©ðôpÙ8™„ôãúÑ'tˆê€AŸL–Ã¾$Ê¨bæÞòq~Ó\ƒÜ‘Š°é€×„A’£¸£“EoW¬¾yæéó!Ù÷„&…!Œæb‡¡Å«&E!ñåšŽ¯w4Ï0}¢ñI"Žê!¿Ž9g2\h
×ªÜh4ðxm¡î[XÍW­Ê‡S	ˆÜ:Žì¼©<6*yÃ_XcJe" ›:Í¥Î=eÍh¼®‹’p˜pÎ/½‡åoÌ{GÉÉ½Îsr+ÁƒÃàúbQ€{ýÝ)«ñõÏEM®e9ÿö&ï4£gmƒŽÒ)‰Ôœ.¡ç¿½¡ªÐ“¹§Ên<rdMwíq÷ñ/¾Ô|x¨sê;åzZbu†¿/Fd0äjüçìç²¸ÌÈv=«Ã¹Î~¬ëXÐí+)ÃÓÒºço56æö›ÎÌ#ò‘Ap$ƒ{MZ
Ý¢!ÜêØçw60Sy†Ú@ —þ|  Óª;L¡‡×Ã/¦UòŠ[†2+×ÅlÝÇÂOÞG¢ÑÐTær«hÞt×j",¡ffÃfp:øûÐÂmP±KXÈÓ~vøæ—ÞúÆƒnR¼´4Áy qf‰‘<S[±çÌÄLÑ|6ÞÉ´Ê= ½vµÑÖðÚ|KúÿA4I’Ìg Ë‹
›|0èS#ÇgYÎ.œ‘4%’±£×‰ÁÛtì'ª|Ÿ¼hÍ<ÿ  ÿÿì}‰rÛHÒæ«Tkü·¨¶¨ƒ’/­[½ò¡ní´e¯¥žùÿí˜C$d!L\ ´¤_£ˆ‰}†ÿ	ö-öuæI¶².TêAYî1"fÚ"Á:³2³òø²;â‰ªM@ÐSq­ç†ÍZMªþÖÆ£•ýÚ†-±ïº†6’ëŒr#æŠR:RðœÅY4ŽSh±I½ÊSL­R{ð²ÀGš2ÖMü"+N»W(ëT8$ÃmB½œ6Š­Æ*p…ŒSn˜¦v¹fT¡þ+P°55#Ý(á€àñÄƒ+6j²dÙ¢1rÂ°+u…üÈKt6GU8Ë™Ë_(Íôn666Ô–­ñåë‹Æ%;!ÔÐš©²Ú,eƒKÙßŽ¸txœ ŸíãÓá‘cÔ‹FŒz!Ç¨4F½h£^„Å¨“ù8âÔÙm‰ìÜ2H‘„Jh÷˜<NÞ¾#S ŠâVnˆ%ÙGåÊN£Ïª³ç²%|L}:&3ÑÃaŽ6K¥80“{êR¹õ½%¤6Áã;Š•I4Þ·Y“õ¶íáÀÑœÀ!­X@W¸	.¶ÑÓpp×p$¤„P&Ñ‹‡¯>Ñ„%°v µ¤]Aç²N/ãÈy\Kªç¨ÝS}{ü…ûIçaxÄ;â‚ËN¢‚Ý‰Så’]&9»Þs˜Ü¨fALàKœÇÌc»ˆsÜþÚv	x¢óãAq/xx.zÐ	\~'QùÎîîö˜k‰÷Éò ·¼
(D÷Ò×q/_„"òªy+*fÄBi+Î^ïXè*™1Dòþ1OdÄõòæ“>Ì‹Éá|<>Bz<˜¨hØÞï·>_üxªè·Vo}‘|¸“l”ö¯åp¾á¼¬òI¿ùx|–Ž°žg]U%+Î‘èN?4¥‡Úá= (·GS–Œ†Å›rˆá7;çüP½JÿágÒñÕæ&ÂÐûÂS	J´grN¤dÂYI_écó®mÇæ¨»ÈrtöVšÈôÎb½,!«ÒÝáò+./ÓÒ=Ïåç\ÖÏâÂÇÎáHºå•·àvWÑË‡ÂsCs7yr[YÉ3¹ÑIâx°íø
j²_diyh×¸»U#P€ÔæÓØþ%0M[„$K_&õRi­»ïlS‡$ƒ]kZÝÔ2‘ÏgP³˜Kq á#Ðÿ‡bÀÃªæ5aÍ’tÉÚÍÒVE£¤ŸŸH7‚oî±®8¯ôô%ïêïøG4˜ù?I{ìÓ„þëo\ï`¥²y@;êï£ hN"I$}„l!¦o;hÞÁöÍ¥ …g|xüoäÂAY•ÓÁOD
ú4k±p6Ù›ég#„HóWlo!ùœ­.Íñáš2{eýŽ/ÿðÚ:ûìoÞîüÅE}	J¸x Äþªñ»†âIRà¾¹+ì71!{Ô°ñG<¹?¡<Ò—4P„2S81WX!<ÞÄý¦HƒÇÄ×ó‚ìÑ<ìO’îS^ÙôSË=Ù-ÃéÓ©±€yš£¤ºâåëGî+qÏ…lD>ˆ?)( d¥)plÜÙ´Ù)Âì¬W¦fqQ¢rzˆ•–õTSä©½H"VXêTÈxäˆó}_BWh·A›h	¹¡ed |‡ŸOd Ñ=ÄsyoÑßáV½Wç¤-Ü¯84¹ºÅd`Ž÷K±ÅûÙ%uàhê§©cØ!ð¿´œ¤(aQêÈpÄÒaÉÇ[qAº3Ì:Éª2¶¨î$Ý9Quÿ8·Îñ>K.HŠ¹þ€Ö¦<Îú½™¤÷]|’¶¹ér=›¯l<™³}»eÕÁ/ÌÃn›JPŽ±’º=ž“Ú¤Ï@ˆ2ÌvGóaÚë•ó	Ö¨‹s ‡`#lqm­çºº³sb«‰ÂÇ±–¶’¢‡„×µbÞšuQš ·gV@oo¤ñfå«×o{æ
M,w·N÷­Í‹4š[rW˜¾7´e˜i´“Kì(Ã'±_åý‚$Ð
Ò|ü<ïWi2&ÿæ‰lƒ+ˆß•Ô-žª6ÑsÀå9Â
ˆRƒ•\TNöÈ¿À­”tü¹3àágåUo9i"
36ZFUšËæS[ÂfB›5ûë_ÕÓµU¹‚ë
7eÕ=ÆÚæcg‘)à‚£¿‹ÝË*':Ë¢[m):ëÔ#inK6äù‡rÒ!:¹.q“~UÌ)ÔÂæÎVsñÍ»¶Kßæ p‡Qòzë^|U~Âl§™‘+¬P™N2š¥|¬è=øY|Sy¾y±ëøvf!Êl¶)ò|Çš½Æ*÷5yš”ó‚ÕÍå%`®€’…x[†°-x3Þ½:,IùÅÏòÍðx°øAt’ù9þ‚ƒ³–¡•@.—|Œgù”TA$-\Ó·±¦„xYB¬â®“.%	`ÛI2ã·Zó†cífíCj_ñŒbÉ2˜G]R§¦÷,ÛR5´mêÑc/` n–Ûä¬ÌÇs|.ÇéyEXU•ÏúÛ›Ô'fÜAHÙ¨ä³ãLÆå©Yª¢IZÐ¦lAIE¢Yâ,¡,GyŽ^AetÇï™÷›’]‰#˜ëÆmŠù‹Þ@xywhðÏéõ«ür”cÉêÛl|J)Îðêk  rghLckZ“ž7µÕSÀ¼ÖïhÔ³6æ3]—Ãˆþ/?®rÃÉfãþS4+”"åBéx´µùØÆ§±²_êÒ•³F†HS“œLíðÃó|8/÷0£gÓ´?E?}ŸoéI4Øé³¥,vB„DYaÛ’^w„×†ì$§…‚%Èo×ûøÚâ™@Biè÷­¿u
¬ƒJ˜¨¼F½9SHMåÑø,H:˜‰ê]†cK÷ÎTt¬Ö	kŒ„K4UeLLnáBt#Jžë´H'ö²iº&9$íÉœ!FêÒ‚›4Šè J–‰ $ï‰»:¯‡Â–´…0uath·Äè5°íÞc‡y¿]dˆÆ†o‘›ñúú+Ø	ÐëÆy2ZÊnh…­%iì>ða¶ÝñûÎ6ÄŽag®jcêrãÚyÑì:× hXÁ4½<a¦7›!•´×ÿÜ¥cÌIUõWù°7‚ÿ­£iˆ¡|¬£vuÞ«G±Î~9:¨ö¢ÅÆ4¿Äsq%»i-à°nD²7u7EPQÍj”²ÇµÃV`˜Õõ¯š Ê›Þm]9Ý×ôÔøø&)>¡¤D¯èÞáQ²ïÉ1[µµïp.¸{|þ>=/ÒòâåðÒrñ L®fñ½Ý&¿Äó–’o”ºò]Õz4Tˆ)öø¢þ9­ö(µ÷¯Vîñ‹CiÜãÒ¤ “›h¨o~ûª)UwDÍ‹#Ó´bOd–QfOtX—z" î¦NóÛ×è—4uÌÔ( Lg»o„}ù*Z Ì	‡ÿWª™L±~ïš'Ñ †î¨’¶· Ob5|¾–ä(RÜX*Î®µòÂ¦W½|¤ê°Igl¦Z”Æma°u1!c/”W´7ÔÌ˜Ûµ^hÝÇYr¦í?ÖoÓ°®ÂðÕ«s¡Hm
6LZþžÑÔi×p…”êáR€]Ù‚m1žÔFj{†ƒ—]ì4ÜægÀžLì)^½ƒ­ÂºbA²®Ñ÷è`8$E ðE%/*Ìä/v,Ù]ôœA®ì«¡	”ÐÆ×ÑÑáÉKân_GïŽ©'üç“S4J«$—ÄzÏ4Î?~Ä+–MéyÛ„Ê¯ÿ·“û›·ö²Hfn¨á®á¹!þê"¿$ÿ®È¦Õ›|”Œ{¥"÷íþð´À®i ¢ÓúD	o t+Ñz5I¶ á¶*dêxô# ùmâÿ¦Ÿ³ô²½¹´‹1"&{V¾{u¸ð>‰]âø³b‡xý‡€Ýáa7KÚ!Á”1°ÄÞ§Ýù+f¨Ým°Ý‰ý!ŸÜÂjIzŠûól¬êmŒò¡9ÕÕµQÎ˜Õf^ÖMP±<„VÉ^êåÕóènœµé-Sá<Ü=¢%à!•g8¦–DÏxpX4­œ±âCÝb2i)—Ð8Ÿ´k×²©¬;*&©‰wayY”ºK¦Ÿ„Ü~E%¨ýW¦ñJe?·Ü¸—Cð(Þ~0ÏÅ•#„Ç­hÁ9ôô5Z I%ù47*:4/&åÍœ°7l-†›VŠVGË‰†³êëÑ˜`ü¼EñàG¡Ÿ N±&è~ÓIô¼>œõ×–zŒ’1æ;™v\)IbÈÔ½* tÂ^ù~oãýv’¢½¾·3¶2¿"æ¬&Éi0|œÇy…Ê2ûˆïƒž®ãh‹nO°$N*X{GâÃ@ª0n?þ:hÆx–›lr¾,«Æ€á–qì<½u•QÏÁåÅGcO­ÞÐÊ>†‚¨VC—S÷É,fç™›¼MçÑP‡‡ÆaQN´Óæ:8ÌÀa)„cqå!µÚüÿŠlð‹¬ô@U(S-²¬†Bô®Lú+€`çÃT¤åL‰5nšÕï·›)-R–H_Y‹[ÒºÈ]W‡i…|1‹„„u%ÝÈ3æ::Ùç(+Ãôˆ‰öØñÌšÙMOÉtc‰nî€è‚´­²ò’	ŠÐ)OŠ!+¼‰¸æ‡zxR`„y¢·Óñµ'ÉSÄçó€‡’¬¤:ªW)n×kq„×,;$‡H?^ÜÕ©BŸuA®­gƒ§»ÎÑ:	Ø³ªh·›»[v"7U»ƒþá^ck›¨›3OfB¿ì”ˆV2Q¨¯¨÷¶]} ië3$2…zÚé|E«¦ùÖÚïÝÌ×Æ+í—=î”gå¡¨æ»ð¢ð0Kº$QªBB-+–D´pi}`j[Fþ,Œƒêº‚2Z¥¤y‘FétEß6(9²•0JK­,k žÁ+§&ØÒÝ™‘ç»»1Qbt\®©‡Ö—_šz¹í°ˆµÐGü9X¡1Ä‘hXx#O%©J|_dé2y¨äÓ*›¤v&5¦ëµ¯ˆû$¤î­Ý"e«ˆÛŒAKa‚‡ÅÑãûÖó~#Óµú÷./ XõÑ«–E¢^Ì³ñÏA|ÁpšuÓÜÇŸÐ*zUG%AÊµÏ¼F¾rävîy×©Š<Ç›t’“³q4â•æTLRÈè­Np+®ÓœÆWŸµ®9ÅÔ²Ü- ùÝr¨_O½‚¥%1‡x„äÕEV"²Ú-Ò Íž8?ñ3X<˜¥$…•%”Ý#Å¾®/‰}ÎúDË;õ­ÆK¦`XèÐ(r
c"”žÄýMV­üJ Âóâ7¬—Ø¤¼é&µ_ÂI(‰ô¸SÎú”óIÜ­Ú4 Y#F ýóé–œÐn7­Á´c a<pÖ~T(ÍÂB‡Whq= ¼äó¬ÑÌÇ)VÂÝ&Ì‰#‚J,ãúAê·gudÄo<(?×˜@‰di‰.’A´Eš'/Ôzû17ð“ +²Â*üÙ5f‹©È$ Æ ¬Äl¸Ê:‰›$:WÙ üz³ù	ëÆ„ø"™bfŽ~ØlÊ/ú¦yFâ‰¤`<•Îë K\^7Á€äðtÓ‡ªÅÑNd‘]Äþ¨‰%if*Ñ1A‹ÿü½O?bšIQ,^h9¬Ž;-$–S<¼	VrÕ¡ŽÓ
‚{É“ù¸ÊføøRêmo­£ÁÖCþ÷/'pò5‚×B9>6Ur…F€/Gã0çc†Ýn™s\ô"·vq[/u¡>MA Šæ !¬9ò/„©±k>Ý¤ã|C™ý,þq~‹zÊ§4ÛÐŠäïÑ­ºÕð©£ð}Šõ„oG)Ø['£Þ¡d×RZ–D†Ðx±êñšìËd:LÇ|:ÖFŒÚG†Ùü›$›¢Aÿe>žO¦èç"í!.j¸ñ$ÃWÄSðeÉ Í‘·Æ÷¤?©cÏ¤tÑÑüšžWøz0/;…¾d8½'ÚY3¥9Ü7~½dÙ"K¥ÇqR‰+‹6ø¬†È.Zhv†øÐ¤¸<üï"i†šÛ äZÀé=7ž—/³b86¡+Ä-·þ„V·7ñÓ=æ¡t`rÀ_P©‚ŽÓKU„X³¤ºŸgª±\òÊþX–2¿ñ9`Ð•8¤f«<¦èWóÙ FÀë÷2Ÿe¸Q2+‡^E~:$ïù^?šž;d’é"°«9®žm5Ó"š î‚íábÇþÅè€e|6 ·JÆŒà¿.òŠJè¦á1PüRæ<N¸#WòÊyÈŸã]ºöü`ÓmxãöA	#JÉàO†æâ7 ¥Òhé5?‡ ÛJ_xš@šÛ"ÂSö¦S)ì §
. U-Â"ëÞœWY9LŠQ}Pëí¶ozwÄ m›–â™†à>Æw1àä€êx•ŽÜ¡`aË psŠ9(ð]"¼	<Õ§§p)1‹Ã°2Ÿã«f>ý¸Ï’RE?Ä$H¿Û@ŽÊôó0òEI‘BÑ¼E3ÐÔÓÑú|N€<“}–ÃW®c½"¾ðî^&S8 ´£u€nLF#2Ö÷YŠåAŠÊä3¹†PKh’U \ºò4ÝŒá¶þn>ãq¾ ÃXyv.¬	šÑïÏà{X;£×n²ÂM!}Äóþœ¸Öü¬Ÿ½÷‡åül~Œÿ·é(X9?k§ã'[+L•S$õÒUïö‹k¬#}8»F¬/Ü~ >öd„¥ƒÊíÙó³¡$Ë«ötKfâ5ubízó‰12Ü[/ä˜ õlã¯æá56¿šNû?¿Xuâ°úî²ð8ð¬xc½a´ïÓIR|*}Kn'7›Ígdgz›ÑÏy}.´,*Y*i‚‡s†„p.àû4)óéZqÍöÖ%µ=«èZ·ÕI’«ýŽ±p±5Í/Ñ$¸0G  ¨#ö¯!èGB ®œÙ„¾'Fj"1W„DÅÚÔ0E¤Vê]
®—yIÈp¥æ4g’_´àÞñ(áÕkš‹¨_f(z%ùäòkõWl<QžLØªëÛç,©…!ß8¤¢fú£	Eá_ñÉÄƒqZT§E†—Ïa Íµj$Êå“B°|
šÝD¯^¿]½ÝëZ6zíF ÖÎ8E
Cd[qøžØ=‹`4LçžÊG‹SÓ/áÌ–Û†|³Ë4Dú Ç..€Zà\ø7Åè¬®ðŠ½Å®m±Ëc|¤éŠ€šÌuM€vgŸ)ÅÉ§”¸=§é0-Ë¤¸Fõú• "ó; ¸$bYxSxÿà8Î/×¡…i”ÔÅã¥—Æ<S×â÷`ŽÐIò9e7 †óy4í!·Å°À´ÌŽ†dæ³Œ(îGD.|4¹çRùB¥P„%N&?ûLß¢>wy[,-ºRoÝ»"¯(¡†H£H^“±·<âþy€þ¢O³	Dî} T5Búe¥Ù’
ßMoß¢ÞIržRùJlfÓ´‚35Ë/ñåó*ù˜®‘kåkË š\ZÌÅæù¼_Ãb˜1Xv Qú9¦Vã<Yï1sŠ¸¾¼éMxô´„ØìW2fn_`ˆ’ù]Ç—¢û0^jC}‰¥Qý“Œ—•¿½Ü›lLåø'µùÔäT9¨•kæ4èÌD¦Djq„[Q†bÐ.ì•6-ÎßˆòN9µŠBõyrNªt†¶÷Ð¯$’ÇÇ¼Ì§çÙÇyA½”¡Ü»á&-'{zúãŽ\ç0ÖŠñì`F½H~œ)2Æ]dE¯¼OÏynŒ'#™t1%Ñ$+û=²ÿYhéhÍÏT‚Òå¼GÚ…Üà!qÏñ
ËÞëaÝš ¦ãõ‰P%’/Äéf[¹e(ÛÄOû4¯ !‹€’ëÕ´¯EsŸIî9¢Ib~¸+Â‚Ì€€’T@50îVå¨ÞÈ?Ž)EÕ	²B~þ$È?6é’Š#S¬”l¢wi‘å£{Cƒ>æ&Ñ ™T nMˆj1ÔØLö=˜â`kðtØ­§›ðO×TÚRtk:6™ÕPÒWIã,ïòdxçì'AšF]´Í£Q<qyáÞèÍë”çÍ†#-äNÞ‚·ŠÿUžþ$¶´Ý]Ð.óŒ´ÒÙI»ý>	!ØD
¡~ßŸ¶áPó¢À—7ú#	íÚcŽÅÂv4MWòï½:»„Œ•¤éN±|J–pŒ•CØê;Ú¤kZÆDú+ç’ó
òe:„þ®ÁçÁ"ƒ¿Ü(Ù‡ÄWØxËŽˆB:!¬€öàf:;ÀËþ¹OþÌ€-¸hèëå~q@%nÎtD•ˆvÆR“A@Nuã ¾ÃV9F½9 é,®DG±¾¼•úÒäKwpK»cv1=O-®BÙìÈÁç$CðãI^Û…Ål9A0åQ²ÕÃ;Î%'ÔØÆÏæ#¼[(áñê%Ë-gZÎ&e—Aj†¬žæÔµ!¿>äù9Ô¡¼ýóÿEÖó1Áì¡þŸÿ‡¿„E !Î¬.w0KÎå£ãÕµ[ôw€„òôw%üáÿøU³©Ü_Ïo}ë²8«ÿšµÃNx=„ ƒþ"/²ÊSmWüª«wûÉ"ê`ü„V{2’ÐI@¡ffš¼»¶ WË#^âëô
*ETÁVy%•]V‰ŸÇˆªN„•±ºF‡F_®€&¬Þ¾E'I‘àUZÙ—þ@½“÷¿ÀQYš<Æ­|¤’¢ÞñûŸÛ7÷KrvÆHÿz¿¼8nß «€G[äåðzïÿGË1Š¼÷·ççÙ0]Ù×>@½WG	jÚÇhýÒÕk65œù…ÚÊN‘Ó'ôtŸÃléý‰“{“:™s€–b#9ÍQOH° ËHç6öS¿wÐÆ”ÙO[[Õ_½<ûë¿q´´¢“äëô/¨N@ÕYàEoòi–3R>^A6°ùüÙ-„6E£„vÌš‡wØ‰óLªþ6	khõOx,Þ…2±$tú·R+>â¸®ñäg0TÍÌÎ`Vgž˜ç¸1ðz3ÖÒ\¾«Á«|<NŠ“ìãT%þYÐ¬EjÁ—'J1dâgÖ82sUò ÐSqAË$÷®£èÚVj)©>I°çÉ´äØóê›2¸ý‰¸¿}M ~{Ó1(çõkä¢Ÿ¨÷ê›]kÛ]	4Â¿Øûµ›^¶µFª¼JÆäëˆ‹·>~÷¤ƒ½òÆð7¹.T¦)ì„Ïìc ØPZîmÛ ´g8E¼Ü˜€ß‚?žnéDŸ:@øÈÃ•›¡Åhóð¿àìx˜‹wvMö¾§-/sPÀ-ûÈMkË¤¿ù¬«=¶b’6÷: r4d¿C#ò££¢ùó×¬ºÈ¦|ã{†“^p>ppŽ‡N’¬\ã,'4ÅýÆíÈ	Ü_·EÆ¥Ü‘7|u'”8Ì]9nÝƒÕbn_ÂœØW„ä×F¼ùE¢h²t×áÚ3­×ÍDJKÉhŸy&ÇŸ. ­:’÷~'@äv¸5N•k¤ý¿Èú¦Ã= ÈºÞs¨,-gn>˜×ÿSkZe:M}m„¿ðŸ(å&ÝXFZ½ýàu--º/©B¯æDi·¯7ôTY|mö,ŽP±ž°Ñof!€LŠµ®[}QÍz¡‹¤\ö(`4zŸ_"z;<a9Ã¡éŒ&+4« U—ÙÍe¿M¿oá†Œ n=°7o*' —X/`íénT9=l½K
Ö-}±„ß¤#W]#;|[“m_a®u…¾ûñG4Ó£,ÃÂSLySu«‚ŸŸ‰Í5XÿdüC´½ˆS±¥1]6k³Þ=  ©I?¾F˜!µ‘ªgåá=¦MÉ„·¢cc'HªÓfâøÀJØ`ä§×3Ø ÎÎ{+BÈDWe!bFí‚&—À½ên
Ox…°Òx…Z~
òë4®t¿\µöç‘A‹Ûu…õ¶aáåùiÿf]ã!kkáuYÀàk26šäJ@}]úÜ,hÌýfýfÕž¦ê¬qÜŽl²ÓÎ-ž±ìÇü1?¤¥øI-G8¸Ïl>öTwdÃ±ÇÆY^/ÌiÓÑwèíg<Àq°ÜÛ×vŒ9Ï{Å<Gú>›[¥siÜÀ0ƒë3«Áuçn®O‚Bo’+$"8‰#ùT”ÅqéÃ¼`7]òòR·ö[Z½Éìæ~j¤ -·©)P)ùfså£joó»G{¢õ†Õ<—ÞˆýøZì¯÷g/ÞV¸Wf#È¦ÄšÐá†`%6û˜{wdÂmkàüÊç¼þ%¾¾ßËl}Ãq^lá	ÔÐs|÷
©ú²»½¥›}k¯ò3ïÕH,µ×À—­¦Îf»³)?¶(¹Ë‰D/¨?sUÌÊ¡á6¡;Š7¬¿JQÎß$ÕÅFrVêþú/¢'†—ox/~pFB­¢z(­ÖÃ(øc3RC‹fTâ:ÌVÅøzlaˆ›ÌÝãßB·¤6Í4¯ìúŒ*¤4·Æx]½žùS}R²Ae€„:K.Õ<N ?G‚n'ÉUokÝH¢©ÁvÒíð¶déeËÛÅ1h‹šiƒ¤–_š=ûáEž—¬îJÆ}¶RàhZVÉ,P®3¿r>®JôkVVc°A×0Ø4%3æjØvLü’úíŒ\®¬€W²G ñòÛ˜ŒÕÇk,=§——3™?½>˜F,º¸à•ÚÜ<¡w"¶}]ÝY)çøŸs@ÙdMÏÓ3©Ã€€
¾"®„¼Hþòg5æÛle°òÑÐt+ÑøÁåiêz:¢ê)Ð¿ðjñB{)ó„¹“àC`‰=yZ.öIù3lådà	/)V¶J3_ä—o¸WñU‘Ïðižö¾+M{õ“¡•š,£¡XÇºìÏNV7*Ÿ“Ô}+bï7­(Ö$Áôõ¹È§¿Í4&Âk-c=“¿ó
~b~Ë¿5¬*¹m«¿€Û[0"ˆˆŽû3f,È SœRsH)Ò¡·„<¢º!8¢¬hø½98Q”òe26¬&Á¢ï3ÀtuR¸9]}ê8!ôÕ6Õ›à	°Õþ»P)xè&z‰¥6Rda–O#aÁ2ö¨~¡’–ú×_äW¬)! Ü¼Â:®[)é.ÛQJbÓê}Gæûl,ÝMë$A­‹8!5˜”Æ6ðE`Ò[“Dåšw”È.Åª"`|îEsBÁšB^g¾±ÂZ85ƒ{>]%ßâHhRIHc°À_“«I¶Ž†Cr‘=Î×ÑÑáÉËutŠEÊË|”®£wÇë`ccÃMê†Èqÿ)šð×lHÙÄeJˆTÂÌyŽñI“u±iuödLùsy¡e–¯åÖakOÎÊ|<Ç³§çÔ
SåXLlPŸ¨9äbxM>`"¢Ÿ~Æò¤¤‰Èî˜¸P€u(HDÇhM<"oó‚zˆB´bí‹ìã…gñULIÞË©¦úÎ@×u/ŒWCÓEy°ï<Dnû-,N¡l1lj:*ê7œ¢K(OQQmŽYh­`erÊò›žY218{<Lâ)²Š˜àE^ƒËN}M™ŒÃùÿû"ðD›åÐˆ!£÷·ýµ†šµÍä«¿päØâ7Œ,¢}š†ûV1·¬1Cä­Å³;`bnŒü~ðê|ÐHÂV5·Ô{pC%i:ôXÇ÷ž¹¬}èvºÕíE’!¥>=Î×>taöEíçÉ¸ôåÀc‹€¨¬XR#Œæ33ƒIâ./ayÙÍ­Ëï(28z’«þEÿÑãša]÷¡PÂïe˜O]ó„”PVæLÃqI„¾ôä©íYÃælÐ1	úÉ1+¹D5Âœ|xrkÅÌµ‘ìy²ªuîó‹tÎCz[§±pôI2'ãñ5D@é(¨(ÂJ´ÓzPAƒ
ðÍ‡E¶!dØ~ ãëÍ5>•u°º•"ŽöG¤—€ŒfÙÈ§ÖmÍÈ5Ö·Eê²ÞÆTðÞÙùîhÄ¢yC: U!Ö ë Å#ü@³Oö¼ÿàf}ëeøü‰²…AÓ˜>Zˆ/ä*­»šÝ,Ïãþ¨‰š~Ò#/À»¦°2Øvè~„pxlà†[l•ûÛh’Mû—}¿ËÆÞ¨µ(&O-ˆhÛÙáÌÃ”Ì …Ã<ðõƒYiRÎ‹k0S„Xš=ƒ]?"ñI´8ú†c­¡SÂõ|½.¨……ç´¸ÞÓgzNÙŒ£9˜Às#q×%ìÉ æ[Aþ0ï­ìh^ò³6‹Of¶'Oñ~-½ÄÖ–x´<ÆÚ+”oÕ­£Ÿ›«úó€\ÔòøäïÏž)RÒë1e"JâcQÖ”0¡ö¼{ÿ`óå¯j l9`Û	µÓje·¢ˆhü¢ÒŽÿùÿÛö§`¯tv^‰´€ô¬ãÍƒÕEŽÛÃ¿`åÚâO]ô)h¿NõÏß.•˜Ú"Kçà~0 ØÑ¨HË2v«ô£iå®ŠùR
"Ü×ž4ÆÞ¹ÁKbMÞ4ñ'Î|ÄŸú&æ•Ÿt_›gïŠ|–|$v×ž×u ?·¦ðæ¼>‰úQ®XWPØøšHg=ä–»Éˆæ¡¯zF¸v“²Š• …ˆë•|ÁŠ:êeì‰ªh—]è]Œ>jh}•ÑiGá×ðmÀ ðƒax+4~H½ñòK‘Ôî$$Z2IÖo2©€CtëKó°¸EÑ\ìÅÔvœW Í#
•ïvhÅ„Ku—;@Ôd#P="üé7CŒ¿–Õ%{“Ø¬éójƒr@çQVIQ5¦ênzNÊÝ»|M5È‰à“ŠØZpŸð¯î²˜úµÿ&+ò+!™=¡s¡Éô“°ò‚CŠÿüFì®54Ôc
¼É—ƒö3ïqœóS¿†’"å[å(–óŠ›à}šbäN‚üï Ö`áP@K¶½8FYÐÐDW}h¤ÜnÍJÆtK“göÎ2Ÿ×'‚üE|61N+ÒØ3´‰~›U,\àgªÎßY¾¥qÄE5OX¼+";Iƒ]8=Ï‰$ŸžÌÏ&YÅ«ÎcmümñÛ*à¦ë[Gº¶±@9ã«u‰rgÎFp‘r»UØíòZ¼èM0 »b’=€^Úç:„Wû…'>ä8<ÁÂƒ€%ÙÉÌÇš%dSƒ¶ˆYÌ ÆT;•0jO ¯?+£Ëb!ÿ{žžYñŠ"?Þ¼ÈóqšL{2MxX:Ð­å©ómS•‰×ü§qešµ›I'Bõçù$ñðoù^Ø,7!)kJðãóú°ë—êãÜ@¯]a¨…B‹Ë4ÊqrÌàú«îë”Có™+î?ãdv ,r!û²øT~ã}ò¸ÿ5y_Glí Ä®hålì×‹27^ì_ÓŸ{Ï×$ÔUÂ.žë·Àj²&Fð>‘€¾'	$÷àkS Ÿ}IŠhöBvéÅO}pY!	¾L9œ-*¥òoƒ‹ßÜaáâ¼…8>®qòú¤úÖ(˜“×®ë?O÷¤ýHòè¼¶¦FøñB„(¤Ä7:\„ÌÆýc¨<Ùß½Þ#øðMs°Lá›æÐxbØtgü÷4"ÔÊ‡åFâÇüýAÙ±ÂdD—ÿ*5†ÎHñ]2mMï cV ÿ8þKÌzkqdXóï%b-‡îI~åRJö~UúÂÏ'§ŒÈP"È%ã ÄÜoúÂ½Öî¡åöçÒSõÞÊWñ/á«MûíÖààÅËW¯·;»‡ÛÿËeð»²gÛ.D7¼¾š¥ÓQVA4CýÌÏúÉçÃH¡iíAAˆAAØ*F@å™^u&–ÜÖ@#á"´®ŸÞ
vÒÇÉ#dŠ¾+y2<
-¦¨­+l (QÛÄ½µK0kRÈ™ ÑA%ê½ÇÜ´ì´‹:¹‚À³&jØ¿Ä;¼S:˜ãó3LÆÃ9œ{°œf’¶í™‰@3¶E:ÚÒ®"1·_ |Fì?¡˜Ü±¤ JYOÁ]ü¸=üvõ €_‰÷â3ß²kÞÒMŠ¤äm©½~Y©¥%¨lãþŽ•­êRTÂP§0Zªp¤?|æ—†Œ]gÓ)îU–ÅîYøn¯¥i`K;v¨%r¢4«ÝªMþ+apÕíø±”GlCó±8Ã>:_&×%z;#ØA#ô3
¼SÖ´i‘'°°TNÆ³*‰rQ	4Px‡ä0êbZ/†'CCÉ¼–u6”T˜‚t·DÙ8³qzŠ{ò"¼…ÖAb„6Ñ‹3Bó:Ä7/ÿh–”âFù4&sNåÉöAßØÍž}A”qþÒê-­—ê_€@¸ÛE”£:W^¯˜ªyTu;™Oh]«Æ’éh>L{½r}¦;ô}ÞHˆD[G[wQ•ÂmÌ`LŒÉKDTú¹›G5XiVx¡çë¶z&Û^?¶^±#é!²Yžl©J§¨DºMÿ#>5ò†T«6,(ç4žÚëŠÖÖœð’*~ ÚjÀm(¸¬&é
:»Ñ‡–¤,O²	Yˆà\¸˜Ä×ø¤W‘»B“^B51L­¦®°ôC—4Ð.»|µ múƒHnê¥!aÇRcÔæ².a@‹BÐŽ‡æ3†ady™°Û7©e•l3_eÍC´Ýä#zvbÁÒ,Ö1ø4M™9Î7BÃö‚îêl„!7Rú„ßKéÃ®¥8sÇ9É.
;z&í‰ÞT¬ë•>Mí_ú[ÛlñÉÜ›Ÿñòi;ÔJ±`¬jÇ˜J#.“ô	`ö©Å®Ãðè.Ã_óâ–Då°Èhí¡Môn^Ìò2½§ÄcíŒÜEc]ü1–'iqðÒås`1'y6F—x‘ƒDÉ7âÜ%ñËÆéû ¦A‘ô)«töãÊvËSDgÞÙ9’šëê$áEw"Ü×ÏBG¦¦¥XÀ?§×P'.Ê>¤ŸÒkbX}šêjH©þ¤³‚€ò¿JÏ“ù¸ŠšiÈõf‡ÿ<lïƒàgîŠû›d¸©íŽo\Ê>Ä^µn>0 NRûBÂ™±Ÿ	ÍÌˆÐ£]ª„Å%Cðåì•`%ê?{`6Q`r„§O7ØÀt¼‘Fô	»ÀYÆ
Þ°fÜ‘k+ ŠÚs "H™µ1›wã¹YŸL]lE»žQÈîÇ@µ·¦‘ž¢PÉÚ`’uf iàŠ°‚9ktÕî"0EDÌKw\ÅÆOBW¥Êª1ž†d	ûe0øTë:(¢… ê I‚~hv'…TòsVŸVqZûlïVßË~&¿5œo-&°û´QL`V].šC#ˆñåçu”®ˆ
rîƒÑÕ	¶zùÀÔ[Ivqž¡õ™*×ó#jÂ27q<ÐCÉ@uÇ3ÍÅ±ësqˆ@.Òúæî–d0èX¢ÂÜ¶G8÷90„“_¹®^ö±æù§Lcè!ÚŽÂêR8I¤cDT±¨¢@ÓŸ¹y‘`ÝJn:o8‹˜XÔ2@Ÿ#Éª…²j¶Ð©?ÞRæÊ*“2ZÇ^ÿœcÐ1è³1kx%@rÅ‹®ÁìÆá¢u3I×ÑE$|ÌÊ&a<ž?¢ž=!{Ìã™c»µW[>îôÞ1n­A nªªˆ—!Yu o»ØBdŒ-ÐsÀ$IÂ¯¹2¿/¸?P°e¦
$¬Ž¬Ý ‘b7§º(—C:ƒ5#õ‹ãût‚÷K¡Ä˜B0M33ãê…ŸÈ”AßA1´>MFBtb2-^N‹¤¼¸s‚	–/ú3í,‚G×|AT«Á–à¼ZARk®4µæ–UTäHy—û²À–"ç÷"‚>ŒúŠ\ŽQ—˜ÞèIþÜm(™Là´»ºã†X½šußv”ºoòEF%Ž>‰B‰6ÇJKT÷(*@Ê[1Nh¬¤dœ¸’Ó[úuZ ßê"+iõ¶Àšl¦2quB}¥¡€Å+ûÄ_¡GBÐ3\r\>‚0@+e«à
@ Ïò9$(­ÏHÈˆCH`Éõá_Áy©Q"bÙ=Ë¯R·eÄí£‚|­‰’æF™#£®uM_«MÑ×äf}Á4štÜY&‘UöJnÉùÛYZÈmŒ73ÞKG,ëƒîìÑirÕ:ß¼n¢öœ²Nb°Ÿ­¼Õ{|m¬5;6;M©ˆ‰½*ÖF¤2¸÷Wöé²¡£M¼pè¿£íë ¹LI¤ZÙ?˜Í03JÎÆ)ãOXîl­ommu..	¾8çc:qÞ4dïSÒ‘¿·ÐóP§¯Nð‰ÜÑ‰¬?ZÂ™ó´àç’¾NU^¬?éç,½Dßµ…Þšð³ŠG"ó`;êÝ2àÃ…¯07½^@ä
-n[åú+–E™Žó¤jfA‚)xËJÂkî‚xäõvk	V]èfá¶Ø>øjKðhhð(¸%<€ƒ	´Å†÷’ˆ=hîÔÓço¾KØ+2ÿíµ5¬c®!Aý~,+Þ-Œ<¨Wü¢Üé ºÓiZ½¯>–¸W©7ÜÙê‹…è³¡ý7·¿0°2±‘·‹Êxµe%^ˆZŠb¼ËD%Ãß]``ìzû›¸1O›ñÕŸ‹¼,ñ5¨Šºïû*æÁµoõ.J´tTÕÍîq½>TÁëmÿÛZÄL~¶>¬"£î-÷u±gób6Nåõæ¾7èdÅY‘‹NÉtÍåbolÑÓ
½?ýù¤ƒ—s‹ãM}ŒÏß£•1ß¸$ØíZÏmÞ	prÐûJI
ó¸öQV‚ö:ªq–&é$?ÉÓ£	ŽE>y1á{ÇÑô<ß Äp èg$î¥ùVVž^ÏðV¼¾¦éÈ{U1À	q&p¾¥Y„šöphÕ.J•qo¯·bã’ÉKf“¿öØŒJa¸=Ã%9Ÿ%Ã¬‚ë’¿Pc|H=hV½é»{ô99YÌ‚gÐ„þ»DÉç$“ÚY2†€:ˆûÂ¬ãÞˆxñ0/^Î‹ÓÁÊrbŸ 1žgÓÔå¹6ÃMÉMÝ£Âz}K@rb)OÞ¾F/Ø"öÞ$W{Ý¬ãÚ‡ "ì1—;¿Æ«~…¾ûñG4NO/fÌÞç—ÄÉNmòô3ÌF JO$„@ù|ŒN	ð|Ê¦™¿uÝå	X@•ÎÐÎEq ó*Ñ)¡í3¼QJð-¬×xý†Â¸B·÷Ý°Ä*}íøbP-ÝÌçC
Èq[û/v­X&øßEb0YÐL<ÈÎ_ý£ép<‡ÝÈ¦„¾PQú.Kj‡{ëùæÅ®ÓŽai”Ë÷åÔhÒ„zÑüÉ e1îeb]¡ør§m²ŽÎˆ'AQ!Tan¢fÖÚ;rý§ÞÑ€Gá[dëÂu½ -p¥Äüº|ÃÑ‡x3O†>¯½ªðfÂyO+Lå%·µà„†a„Ïq*{š–g&B	»²™Ý3SsJ€Æ±òújÞ^,F`•]&ø“˜Qá+Ph˜OðõÜDêðÂ³˜; 0ØGg_in6Éþ3u#}Ä %QBžOï/Ô(\è‡~YØ¶.©ëÐ¸‚ìu9""é¯èUÂâN’»Â°6+"v›™œ=.=¯” [Öß0“™õ÷yu‘&#ïE³*i.á‰X]#ÊŸˆA/¾PfÉ§jP	m´XÙÿÓóÍê¢U›¼¹þûöÎqÁÔY•Ôn~E”¬è )Rbš×õa9‹/ µÿ´ŸŒŠÅ¢°& ´«uÞ¥ë|2?CÜO—Sán9ZÒ;:}øóÉi“0­UÃð^[­:Z»•} Ëÿ\ŒGÇ­[“OÐeûñÊþYŸñ;…ÏÁÀKžWgùèZ©oˆý#˜Q47IA"N:ÉŸ	^»¬TŽPºAMÝ’èOó.)G³Ä˜Û¢U¬‘bêGû!ôâq-Áè^‚8>É¹Âë(—Ô¾iÞ!òqÔt§N·š	Aæ¡³ÕÈEÇ&ùó¨–?øäHÉB•WÒYû¬ÛjÁpÍ4”°
»±1ÅoÝFä•ð¼„£Å
¾%èÁN þ˜ušüºŽa>ñúÙÜ$Q¿hœç ¡	!˜“¬,ŽÑÀ–è1¡íP«ÌTO.7Î3|
g0¬À¶b§uVƒ˜AòP¾ò9Åá­Î6Vè¨¨ÛS?'ý}HoÀŸ>ˆmþcRÐUeålœ(å†xTÀF%ˆ'+­ðO/Wû«¡=¾ÁÁèã8ÉR’_˜‘""e
“¾ac’ÖBXðîúœ=Rç±k¶4ž¨JQq<ÒÈ…¨+´ˆ.³ór(ÑoLoÉ£^-t´T=E×hZ±ËßëDœ¶"#¶U5b~ã²å¥áZòæ·tÃG4OhQdÍ—GióÚXirü‘ª©­Œ±2FæŠÞ†¢šµÅ@ò
|\ ?¢Vm0ã–Ëüúðûå£Û¿!‚á¿zŸQªç,ãÖQïž¼ÆÚ[Â~OHòU«Æc@fl¦vÁÛ2|Ú³!¾fãu§$³dI©ÐÝp½h^wÃ.AmŽwÌvû¢Eà,ÃcrÅ–Eì ­¡¨ØÆ#ƒNEö$z‹{¨yöè”úGàë¸#îAØ’ÚÙÏ'¬7üo½búŠ<Î_æ€’ÇÙ˜˜‡¶‹ÿˆQÜíÁ´!
<SnÙ((,ÀêÞ©(‹ë¼`½Ûãâï´¤ÐoDª´7ÂtT_jî
=kúï–tçwÛkÚ_‚"«"x`lÛ~K?¸½„ñòúh
á=0ZÆð:eY‘é¡Fu°$îÇÌDÆzxŸ_†ÿ4Jâuéâ3lüâËE@¨Éâ°È'×1:#†ß@†¿$Òè
½!ž6–`ð»Tàñšwn½'ÄçâóËœçyåu¹Ûs…›‚æpüˆº‰¹‰ð¶èµáæ‰ì2â{Õ½@‘‚¯1¾S“¶Öq¤É8`Š•“ôÊªVÃx rh¡ÐBòñ	V¶~¼Ù½u¨·uX^´2ËBÊúwªŽ+É^ú‚µVÄùIý’
·)Àiw¦þR–‘Ãu°Š˜í<dœà‹¯¥:ÓeÞã(ºÛ	óÓ=¸Ý¿s‰ës¥àŽ|ò¿q]-c»Ý–ˆ&&ÏcÓéÌUEâ1…ªÕ~{scXe
/,O]â8g[†ÉSåŽä ôøðmžÏöM'Ÿã#p•ŠIK‡eŠí¯â÷WD¶|’Ì Y‡©·½µŽ[Ë~YsÆ½•·ÖÐÄÍHhŸ…LÒá@á7ù(;²èo2[¼­;¸qíÎ®HÄy™Vý-ôŸÄéc]ž–ÔíÃeTä3¸gë
éôådˆNÍÓw]ß“i6FÏ“Qz4*>ÎÂ!Òø²ÿäjÌÀ¶ûùâo5Ú§æAy	ÿT‡(‚8/2Lu¼îi£ÊÃ'aãèý’&ðûÈ*‹+ZiwVÙ]ÁòÌûUšŒ	÷–¢xþÚ.ŠM™Ø9i~|Œ 
{v‡½%†ÍDæ²¹-s 0ø¸ßJp*´—CÍo7,NØuúk‰'Rk;‹ß¨±øèÑÈàf©i²®Á2Þ–wòÏüzÉƒÏiä&þÂ¹ÅüÅNàuºš¸ùXG”“Ðy  [ÙtõE¤—l<2&˜Ä”…¥5‡-`à¢'7CI6ÁètŒˆE4L2WV|°„Þ§ç{èüÕ-ú;wõ¾Áëy±ùiRìÑ_“àoÚØÀYêz—Y>
rß†Áñu”DÝåÑ·0ÆZcrnÎ“qV~*(õFfÙÂ¼H?èâi‡¼m_‚µ3çF´k´:8o²©)ý%,U–µAŽóë«¬’“xÂñ;ƒ+œ~½äI:2±4`²ãm×¼îÉ8/ŒÕA[þïÅ#PÏÛî«’œÌ'“¤¸F/’VÊ©”^'#&èI?²œ‰×G9OXL/µt	]ì.ÝH\e$&€ß‰KlˆÙ­C€#Žl­É|éÌX6½e€ÓÈ‹À„—ýÙ•œ%J`é…kÏx&èà.›8ê¦îžî”@¯ããñèïµ%žå"'I—Ë%ÒÍKu¨$ì%.Âvúê©¶§jâúæÓ-éö*Š1ô': Oòþi6ÛC/)¼¸ˆîÀ* J¦×,¿\†êœÂ dSa]õ®S0H©W}¡H$/ñ2%Ù´…}Õ¬Ö>Šý „&cCÁi5ÛVt·Æ4\Í–4ÈMßì&™›¾Â$sicF¬ÃÛáîY=9ô%þ½Ù	(18ãÛÚM#ù1ÉûQúh“îïð²?ØÕk#Ô}Æ¥‰‡OïQszíÓÈýÝYðÌŸj³m“]<éÁ 9i9ßœUZxQ¤É§ùl	SW0Ôq´NKo¹öÞÜŒ6ÙêÁ{aZƒèlvkojN»:1ÎžÈŠÀý‘¦»“z˜»8(ïŽ¿³ÜwxºÊ‡gY9ðuÛ]çÁ×-·Ë…¯ßM’ð2R„ï(A8jµ»Mê82é×
]ðÁRGÕQuÏs ú¸Žt åó"QÁ;W[
ÂµçèH ÒRX^o]¿©5æ<îALî—i°›˜Xaî3b%Dô—6ä£ä°+¥ùçˆ/È¬ÈAšŒâÇ¹œu¥5‰½Îµi– »ÿùø¾é6¦¸Œ£Öa
½©WS=¥N~fŽO^Šœ£’êIï‹¯¾3½¾ÕFÜeª}G‹°Ð|—œtOm+X¯UŸŒNf_Bâ½¯9ù^81Ÿ’¢©Ö×”$'{Öb²ä”Þ‚Ê§?QÊ§KÂµe¯Pú[@tT/]›WtR¶ôÛÐúÛO¶¶:YŒ¥ãÖæÕzM¢e*K¯ŒÀ¢<I9óu½z*¯„ÕšD[V§WÆÝvåZìq«¾b€/ZuÒBQÁP^k@2Â;ØxŒv²»%TXB	„Ìà%^–œÁÇ„‡ÇÙ€õÓ‚¨°nG§´<“V#&_éSÖ€& ‘(
 ×hÓeÎñeùÀ®.yÛân‘}©óoÍÞz´µ¼{L÷ˆ-Œ›#Z¹êŸžhŒãzqØ2ÁŽèÝ(TŸJ^fÞè®ÌG4wŽÇu€§-¶<Zøl(jýx`BB5ë' ‚7¶ÑÛXáá&¥?H]¢b™ˆTò×©°°3ŒÜ+¬ÝÇI‹_íáæ##£4E4CÎn\[Ñ:Ž	. 1¢à0X[4j7ØFßŠ‡[!¿úÒG³cèxQî- XåÅÛÑd;Xx¢É§;xÒ\zZ¢e9, ž >êƒma]ú¡[È‹Að-ðÄAr1¥ó»„rQg}‡p.JÇwéO7{‡Ð.»Q&p'¼ËŽÞ¥Õ-éçb^£ _È8Ó½•™¸_Z$®Â³˜î¦î„ƒ©w2ÆÕ[ 4LgëÜP™Q×óWüí&‡aqÐ˜–“¯Ácô…Pâ¿¨~‚× 4Ã^sƒÍÐ—Â‚ï)²Å!î¶5²Åµœ9¸#dBÆïÐO4Z86R0W mŸ\ä—P2”}8 è2ƒ Vœp–ÇðSV·Ï»!÷±Ð»TÜíi™Àò¡»êï’RqôžÊ—$5x¬™ìK·•dôÍÛ¿ù$wXNïöîaÈ}Æ{ƒñ¼àÃrüÆø¨¼»Gkþðàõ“_+_Éy…6Ñ	© ¿Ö²"°'9øÿ  ÿÿì}mwÛ8²æ_A§sbùŽ%¿%éÄ“—ã×´Ï¤m·íî{{{{oh‰¶8‘D)Åöh|Îý»÷Óþ´ûK¶ª  	€ D'ÎLô!±D B¡Pxªžrp°9¸b
ø‰~&òÒ9 B"`n§P¯—„zC£UWÜX˜p«J®mm·½ŒÔî J[frj»aáü?ÕŠÊWE9(™8\­¥Š§Û–ÐG-eÏA¾ÊÀñŸž›Õ‘Ç°Iý3ì•¸ë­é5ì¯ïT4=k'°à.@‰¦ü.9çyòùN0g*?û¥ò„ÖƒÌüs	EJºjöêÆsE0”$Dezá¾
áG½y…ƒ«zö„Ùþ‘RB-:öŽ5ËzÉ¾fáŠuJ;™“`Â~›˜f)/xìè*èÃ+°] '[Ïp?ÚÖ®¢¶\mÁím|ãvž¿„aº;CŠ9ØGYˆ€%ÖÀtjY1/ØøLæÂŠY\$kä&Z<·Ñ«þÓ’a-÷&7Úw¹_®w?"QÑuÔ-‚îÎ¬ôê „çP¼A¢õ­ëóÊÌo¨%zBbZ³Ëh M{ô]Úóö¨šW«ý§ŽnrÅ€[¢<½9øm‹ÍÒpvAaüV¡gÖ_Ã¶W®	˜›†£n¸7ÃX'¨·ªxÝuíš´+ûCC¸ŽÒ9¹î-á,‹Ûƒ³ŸeFŠ
Ë´î¶³JçVÃÞƒ2÷P?ý‰~Å$lËNd=Ù‹<Óéˆœ»)ëö“xâ+¸6ÜbÒµ¶¾¶Žÿlt:åúá«Óð†¯¿{]÷ø/§J³¹:­’N×:àÜ…˜fbžÍ'ß”	BA¡\dÁö·"@â„HšB@òaå™Šˆ°LÏ¤íþýN:	&Ó”bË«Äå’åìÆIÒN¬TP7»äQR7Ñ¡õk6MÃä4ðÂ¥ 7ŒFKx† ÿÞcúõ;åUœq|^Ç
8T»gÂÙ;Q¯Ê¹ FüaP³óy_˜¶«ÏÂsø­ëÄ¶xøñ0^P×(¸¡@O'lé?5‘Œœ¸¹¶äQÃ–*VµZ”i¾\ë®¾ð®3ÁZ•ò-Ô&¯5‹£|æY­©éÜ Ó´·ÌPQUèÝ·8UyÐÖ9~£Ñe\éµ˜[fó&u2áö^ø¥—\$­#Uš?äÊŒý-˜Åc,ïcÚ~z€gõM,ïf™©*ºFÂ«;°`¤H T)E­žÈ=Èl­ÔÃåÑÞN³
6È±Á¬ŽÂEÖ„Qöó8a§Áè*d˜Ùy‹üíìö®è9ÿñgØÃ}ŠÂk†wÑ_0fð÷öC`áïêÙ¨T¥TVT$+AšÕY_ßœ4ÜR6«5K~ž>¦èìCÖü‹[LåÛ	Ä÷[BÁþøóÒúó6eÿÅQg#€=°W[êåüT/)'{ïv––ï>Ô“E6ón”tÅì¬ºxªïóãÏõ¤>(^gØáò·s»–Ó—6ò¥CdcŸM1öòE´‘ç?ZÆ]^¼¯Qß¿-ƒ­ã@ièÕ×ù,CÿÙšMÙHó¯CýHjÎ„«Mì °]þ63ƒ’"‘èvØ.4ð"dc¨» p+×Ã¸èÔg“ÞØ×õ-ä –tü}£÷½uG\¬b\+¬²½ýã-v³%ù†üõ–Ø$¸úŽ¯o™î¢e-èÁSG0Zÿ¶úmIR–$êÐ¯tUú¶¹×#’ùºeI{«Æui<MÆeÈÅwŸ•	\æ ãžO¤¼‹Îu@îÌ2Ênãi"°'ŒÏíÀF{ÝÃÑE|ã>Ñ|9€ØÄËl§ñpÖ#ß-Î+Oø½g3ÍYg§|P)8
Ã^Êj¹°¨ñ5FÒ·#uï-G,,Ò¡³5{)7övÐÖãÎâ§^ÞÉâ{Ž9–Ç'âÂÆÚÁ¬è–z*¶Ô$4jp£]¦úüà‹Ž_7GaC†~"·TÓá%ÏV¡SäÁ®—x‹˜á„LuÙóBg+nvr×ãZ—{ÙÞ¬Öâcä¬+•ò¥FËŸ<É;Iœ–½—Ú"È‰Ô!hÍ[Æÿè¤0“ÖR{i¹«\˜¤ak¹ó×8µ–V—–¹­ÂþÁˆF‰JB¥;`Îôø9ˆÜ®Íô¸‹©ô™¯†7µy8Ò“¡ ýJcî‹Ê™!CE=`2~P…*pd±qŸ"µ&0Ê­ÔÇîhR†|0çƒÚíƒ¬°3X¯Á
{ÂÎŽ÷ç=s)ŸAü œ$!©%¹(0` F<”©5äÕø¤åï´•¢-@¿QrÚæºJcçÁz†ePñ$æO´
éB¡Âìf*¿Z¼ñ™8”¥{Ñ€{4›î“Í<ÞWÚh>QÒycµ¯ÈOdZ3‹"Å|ä¯W#ëŒ×ÈøÄs”IûLÚàÑ\LIý“õE­eO’ IþÌÎcº0‰·§“~œD¾eÞÛ=€íÔ_¯‚¤ï%IþS3óÔàöè :˜qƒv'ù°3H)þ*/©*Îi1ƒ¨" æ¡Þ4à„Šôˆèä’ù®æ>y5²êÙR¸{ïs~)36.¦ñÕæ@ ² ´gH¹«yé®Fêïµ´è	ô³ÀÛÑ´LÁé¼,/×Lû«…HYÄ[ÍáÉløT­Îì»|ÂÑ<ÆÊNx-?¨T”;‡Aò1e{<‡/’Îì‰Ïù®¾ÅÂ¿ã7Yè¼Úf“%\lKÊ&ß[uÜV?ó¥P·µ¨ZQ9à¦ùæ~n”Ç‚ ƒW¿Œz±-u„ÖÈæÑSRS©R5‰EôRK“¡® zªyVM ‰Z'è«¦<hÅ®ª¿o Ö)ZGmfzç‹nëMæU>!Ä6¾äïÒ’ùnX”{î0óìçG³’2ºóIÐÒ÷±Ÿ&VÜ¼(â–¤ŠÓ}DòQÉršYâcÅvéeáSþ²j¹®c»y­¼=“Éy”3Õ¥³óûù•³AÐZ3ƒd¨'ÀüHýžu©Mëë¼óW×†Æþëhíã§´³“¤:»ÔÁ\%X…Ë˜mºçzà`Stv–Ù‡¢»7ü¼b÷H®“^®~R9c é¯QxM¦CÏ¹e…š'ÐÔ')B'–Pë3¬{jâ“„'fX !_=~åˆ *¯ŽwÁ?ÉÜg•%‡Ñ (Êèdï a"8Š&IYþÌ,FY¸Üžè~:	]ÂzêáW‚MOIÊ*YX³’ ¥÷#Ž¨Ø³ZöŽwÙ=JN8°‰(²ÛHñhpë	†ŸµêÅE)UÑZæ_ ÞÖ²§åNï[+ûç<Y?Óêßa•òžW®™5	ƒ>³è—?³ÊsÝ(>ž&vˆXw:„6°V§w—YeK`…Ù1H–5òTÿxïÌÑÜusÚf!Ìzÿyè%@4Y³ÐÞUØToy'+³G§Yê9Y¿€q£yÏ»(•ìS´·1¶©ÖaÎ¼³kžmrá*;%ˆ&Œã$fÝ$¤Ô*·pþB¯JWMIÃR<kåÌ¤`áds”3ÿxåö=0Âæ•§]é¤îZ({à>B¿sÊûÛaÐaÅ/ïß³íÝ¿°_ÎÏ¶ØñÑûßØ§(p `˜óè‰”]÷a£Dãr£^q½>:>g·á„qðj …1ù.áj$¬Ä…ˆÆïŠÐpþ[ÝóWêÒÏ½zst<G'ÂqJÉYÄóCœLÕÅ@>P=”i¢ÂÏÒÙG_=©ŒøyéœDzà.BXšB-RˆÃ¯ShzÓk¿é(×@ù±Æ\?tìsV€§²ÈZ%N"õP=ÆÞÅop†cgÍ7µJ$zHÍ$-’˜ÁIPk&Ì>üêˆ-Q
Í™t(gÔÛÆÄ£xRˆ³ëT¤#oÙ÷
Q©ÎŠ¤•Ió&*t}žw?Îwÿ9µ½÷ÓáØF¸7ÙÞ=?<>:Ûâ!7O$<‰ñ¹ì=ÅšÚÏ•PÀ; %äE™øè}ù[²è’¡xå¶D`øYƒ&½LmŠyé=ŒævîV¶Sì ¶~*ÜiÜ4eK~~°võ=·H¨òŠ¯<‚Í¿ÌZ‰'JKYÿÕ€º:e’¾¡Âw¨U—‡ÄOkáGOJ7 õoü1¿9¡/³ b®åêMç9£¿ðÌm®åjàŸÎ#•¬“ù•¢C³”|Iþ2s¾*»N´7X†ôs&‘qO3ÙŽ>U¦4uô\óY¾½OÒR	u§t­pÚ/<k¸¡OèÆ'Ï³2‹UÉU§$gÔI4LÞ;-yÞþU,‘Â…#§S‡ÑhàF:áJ\†'“9†ð2/9Ùæñà‰òuœ÷ÿßÿýVQ¿b›ÝHPbdr‘ .Í¡¥exÕnîLMÝKÐµ2õÝr·vïbn»ÙU'•½E“Â¦5<RBÊOÍ<ˆòÃƒ¬sb!Ãüf!¶Ý‚`ôÈp¨~¶”J×Ë•n¬•2
Y`“¾•V¥WÌ?Â/£%xË‰ÄóÊ<~ÉGð2NynQé½$;§0mU_æ£†+.pd®tµÄFµ|,ÅÞX¢L­tFÁï¢ðó;«sñkUÝ‹“Z*Tæƒ=ËÄFÖÚ[ô¢Þz¯gß‰¼¹u?¹ö2ë®yÌoE%Úë­,¸~*»n(á‡ÖÍ>ŽZ‚©y#2JÆò&tö(ž`WÄ×aÏ¯`_å#TR|uá¥me«-Ï5A—Qb¼òMxÎ¿…ù{çaí­äxQ?„]9kÕŸÒ¼™ŸR×œÎ÷<mÕ-™‘u¸<gñ×f¦ìæ—œ²<%—Õ¦'jìÆµaßTy‹ƒÐÎQì(Õå+)b0Æí&žÏ@ÞƒÏ0ûy^¡"§N‘Ô+fÇ‚3{ãù‚o8€&ÖnJ-HÁ¢@²V<zs‚ŽýaÓ¿i4gˆÀ«ñhý€ÅÓ	9ô.‘˜ô2þ=/‘C˜]…#Øí´HóºJv«‘Ç|¿ñgÃZ™…ëª­I© §ƒA;%ž=ŽñÎaŠ„üæ¬”êÂ4Ãs
Ib…›ã2ó5ç@ TÁ9ªöš)å¹´V¿5ßÓIÂñ Ä®µú¿NŽÿgúo«Ñ
[¢4>âwø²MWŽ//£nvI«æ9
dMtJ}$¬yKÚôÍ¤ƒñøn‘Ù¡œ	G}ræK‰‡Ä¨íY Þì.îÒd{–÷º;’Ô×žª9€´î2TpÚ¹Ïá,¡^ä9¡‡ÓV¡öeÃÈ}aàºPGt
!¥á¤½ÆþŽ°ˆµ?øqcÐý¸úÃ3rNõ’xŒfvB<KÖCñ#rÒá’ad\'ý¨×K–VnS…F“è’µÂÎ$H®¢üõkv`±O ¢súmùAÍ»»…‰JKÌ®LJª¢LËd®yÂ:äo¹n?½H&——Ÿú˜è·ðÏÂ;ÛØr<`7†‚LtÎŒ	ÝuËÎã±P@çq<¸Öf?òÚ"¡®Œ6sñEGq{L7ªÆ2Y‡’F×Iùå¤Å]Œ¿Ñ68«Å´IvÊ1óÒ‹¦”Ýr‘Ä¹H¯ú›¥…9_”±‘úÉY[Y]~%dKÓZv90±˜C†B©†ÓðrK[“$‡û;	“(î/g‰Ý¬+v;ž¹«æg{£ë<ì¶xx¿ë³À¶OlÌ#®ö—™LÉåÿN}Üóˆ\lj™Ã¢®<ËVªõe ´§	RTÅÓ™vÙ+¯€Å¼Ê3dm’GJÉŠ›X}Ëª±úÑ*øÒýäE²l:ÄÊ©x9¥ìès²[ÖÊ{TeØÏ›¥¦X”ôÒ}¦<¯˜›–¶LÛŽ€×(ÔºR˜`uE kyò¢Ê£XÇ¹ë%¯sÔJ1²íÖÌ†[¬*bP}$©zmÉeh]‘¡Ü†˜K‚D ¿3ýbÔ¼›ùVëÂoŽ­âÐ-%‚)çÑhÀÊôu‹Å½GqVª§ª°K?¥ä'LŽðËûPG>Ë›5®rÁÊÚî¬jZç:BYEéKËH×SJ|##ÝÛ˜úŠ”½XMmdÌ¡Ð€2â×÷!ÒÉüÊ¨¸å¸no mÏ5[bxƒÎZ×¯–˜œÂ‚Püõ5÷€×ÜÝAœZÌë0ý‡ÞwÏ±ë›dûH¸µf÷F7‰R\\
ØY?ý|äÈhI”vÉË²N”éÃÞ–™Çþ™N™Ó{r:ô¾qkŠ%j»Ï:­§¡á/]VqÎ,;Cë’X>ÙØ4æ¡ÍÏ6TÿŽ"až	ÃôÓEß,ép½Ûbrß”S¦z¥«â-Õ=-6òøbž\#ýïë›ã›?*B{Ï¶(ŠhˆNNN…?ŽOÙÁ/G{ìtÿýþöÙ¾«ÁîOö”Ve:O|‡Qœ+Yn4ÒC%\³¥Is!-/?w`ñ%ÓD»Á»¦œô®‹ÆVVeZ³S°ê^õÈ³¦nØ@£pÒ†jœ…9g9
á£7yô ù…Ý‘„ÕûEGÛ0Û3æm$ÇƒŒA¸ÈˆüÇfEjé¤–§´p(×ŽÙ0¾°{¢õÓx©aIÑ¬°|þ[ôF:QŽ\p?vN:¨öÏÎÙÞþÉöéùOûGç–.íoØZ½^jõ¶I¶ûf 2`8u]%r>»ópƒ8ÁWæ—’6ËÏ?mâÑ_·¼‹ÝžÑ*h ˜ü¬MÍÛ‹0{¦‚ä*i“ŽÂ\èNØYÊp
zÇÃ¬¨,*Êž+‚â(…›#¬ðà¼–`WŸÉXÖíÜÔ“@››|
Ø–dçÉ)Š£¸S¡¯Ìºjâlyß˜Ž0Ü›ƒùÎmxæÄGT)ÜÊàjWm·àÜ)þäbßéHö¹^Ö-ygAˆßÃ®‹³&=açñò—_Â&ŽþNâk'¹<nö¬2jÀ¯XåtCÈ©Ó€t7§û b‹œüÃ¢²Œ´g]Ö™GÕ.…Îñ`4’?•fŸƒ¶•Ì†bû…N+é:®Ñj6¼B´¦»æ-›ñ07ÀD æØÂX9pø]›
æÊ¦Ö Å¨ôÓÑÔv"Ìçºhék;ç{i¶cî°°‘~Â‹àjHç«¸4ZN]Ì8!~“‹¤G¿1gâÁä­Ç–«wË"C„~]Ðìà£¿#){lº~÷‡L/a[gº%ãuuSlšwâž-‰ _I=‚(Ö¢†4ÿ¬Ú¢Œ¨ãšÖÿû‘xÙ$
°Ø«¸NôÄL“~0áç“‚Ž4JûaŠgKÃà#%æ";xÂcnÇ‚I+“³ð&ìN)„d­—ÙÍšy8ý“[„}LÙ*DÁE+Òu•ÄT*+H@*Ÿ eã?Ü'A4€¶Óa/ÂA|m”ûëÓßâ)’ý×Ü²ó›¥Ô¹1KƒqÁXPPHÖn?äÈ,è¿1reO¦‰Ly¦öR½i.+o[ÞA­s*ø(œ°ÓówgŒ£ß¶ÊeæÀ®êR—©ùAB+P.}P¡W}^Ï*;Oàå¦0âaû]‰—0NÃ¸Jáy‡§*ó½á	#< ©˜Ô9÷ü>‘	åND[÷„°œ“ÃÏoªg®=˜µMe~ø¡¹‹ÉÅA8ˆXëJýcJëe¾ÚØWcè¬ çØO’Âzg€°¿Ô¶-F{'k7œÜ>˜¿'³DÝ—Ê¼n¿ —tØí¯V'ý¦ÊF£ö5ú¾ÖÈùÅÙÛØv¯—6XÑ£7ç Ý» ÒºŒ>²ín—f¿Ó&»Gµó¸rà:á>†à0wØ¯ñÌ†ä¾ßh/ìMEøÖaç<¸abïÎÎ—ï§>7'*]Þ·¤ë7áÑ›“mLõ/pxä~®&vW¤SS¼š\ ½£TŠ/ê½ÂÄ•Þ¢‚GSãâ,@×L´œoØZE®k	Ã`Üj^JnWXÔ»±@yÕG:G(&ü4•Õ‰ nŠÛÊ¼JÍl^ä‚Àk–+34qI†2þìc[]eÛ°8ƒIK8˜ƒ8FÏ“ºŠÓó|â1ü]ÂlMS´~£ QJüóèíðh –>Óò
­äëxçŸÖGÂYc¢gè±C:qw¢ÇZ«W¨ÇªÊw®4y¯+¥è¿S-#\ò;ð«›Ök¸zœ+HZ>ÔÝ²Š‰ú#¼ŽÒsoK—ÚKãlÄÛ—?h|o_Ï`nÜiÖ€i¥êEòjÒ«§*…ÃE;‡À@¯ƒQ9qØ0óÔ©9t¼™ŒÞÌrQ¸ó'bÆ¢ õ¢9~^<íÒða!¼—Þ#‡¢ƒ­I=^ò5³.3Ì’ó¢ó_™ÝAù jú‡Oo3ÔKQ®¨.ºL»Š"ñ­©ù.ÙI¾óÑÌi.ôhƒìå÷ôÊ0Û;Ï|…IZ:½2i‘Òþ¦’Òð¢“7'ïž—g×Ô^nƒ¥ŸVXz(,°	Ä°´BáÃÚ%ý-§OÚX¨“F&
ÖE«OGxÙ,ýÔùÄ_è(F/äïµŸîþ`ÂóXN¸‰{jìrˆõÂ´›Dc‘Þ„Ü£ú¯™[Ô¿ë|5~–½“LyêŸ'ÎÌ	'Õ‹`>G]tuÒÁÐN4ÔJ›[ýF÷§×üf¸°âëN]ß¡µ0ò´´y¾>h‹ËÅW#Õ™ei"åWŒ÷ëíu°³Ðš¢æžoÑš€W¬M—Îÿ	4Ã›Vl´yMp¥ªüÉ½ÇòóN>²ö¢kÁ¯Õ )ÿÂ“ÎítYÐ¼I7ø—50JÆéZãÝF[r¥†AU,Ï£Õí¶ç,s]ä›gã¤Ïç9M¡âsîg\®.þqnÖx¿ªiõªªbÉxp}þzöòNÕ§ÚnØ¼‚íA>2 Ñ{–„]œ˜ÊÂÝ3î^qd7!ïŸÝ7xÇ“¹ÎL€Ô"°bRFÁU/(ýû´Ð¿Y¦®%r4™Ý×îLrŽo±¶»ÓÝƒRšÙz‹šì&·±þgƒ÷ò:zR‹öÎWö=tœjë,Ò‹--ÈŸÔs­M¿j:,­zÇáoÃoˆ0®§*Î3ìZ	.âlý3bqRDá¦úbh"'€(?ÕÍ>	néÜ¹%Û¿ì†–ìB¾•E+?«0NÇxº>£zï(ª¯QtÎN<™ÄCv]‚Fœ‰<¡mR9›Ã±ç@'íõçzZÌƒ<ç`xA¾ÅièÆ†Óˆ+–ÕÇãqwˆA«’…uò3A
]9±ŒgS—Ð~ÒšX¯:5†Áì(·à	çÎâ–9Šb“9?v¼“Ý”G3Kwò÷œÃG5BSNM“Œ0¨6‘tf¢ôÉ=Ùlêø¦M	¬7ôüu8)dî;MEdr9{‡-êužœøVÉu‡¡25šÔ"•Î0Ò¦ÇpÁì…1¬Ç®ý G¯"¥S3<„á›7ÄW¶¼ž[v¡)v÷ÈÆ«Éq!e[»¿LÌXúIqÜ-·0,,ûN óÛQ—£¬iÓ~|×ð&~¹ÑW::,åÀÛPsà=_›+Þf)Þ­	²è—{óÊ5šnŽç±{ñóØ£hˆöðeÐÛÑ¨´¦a[cCÊÈOY!T1“œ3ÚØ‘IÎ3šØÂ£ûê4¼k¥¿{]FüÔšWÇ~$àÎû"#£;(öNI6ô†„y_‡€C`O
çØK¦Ê¥{…áfQšNAäD(–ít;AŠ£á8N&Êí†úúþlUÆUŠ¸¬„ÏŠÓ»uÒÐ¨‹‹]ÃS,¼ÐÈ‚¸`Ž‹©6Êiêçb0vÜÄÏ(m6}k¶™¹½Ü»e«°3’nßs®ÊÒ¦-k‚\¢žaò6YéÈ`	7‹¶2[i<˜NxBa$³ˆÇ’Ò"óYY&e4OÍ9Ë¹xa!æUòáöaš†ÉëG¢m·|bÅ+°¼Œ&ýUŒwX7+æ$!Æ:K’¯OXb¯gˆÞãEž‡ÉÐl*€x÷q›–e@ÿI{.Ë~Ú¡b-&Gé>´_²qÂÓÈäéfLËÃ¦ÙÜ
‹¶³Ýiº…¾(:É¾äé&Ö074EÁq”¥,º1r½eÝž„ç®~¥æmÞì2ÀìÉ®NaGŠŠA"5]h;ð­L~U6¢õÞ3¤ß´ï	Í	“0ýþ£eþÎò1×š|Åîzp¤OƒO‚-JZ\üÃp´hkÌ€>œ8’ž¦ ¦ÍGlÐ£EtÅnãi"ÏUB}D¦Ôÿ í–S–G£`DWpB²Çjþrdžµ$ó¯·‚ÙŽŸl8[² ³XJßÌ‰Ð1¤lÎ?„ÓRž˜žW&³LŽÇD`•æâÖDOó,» æö°´Éì,Írxší¦.š‰Uyc|Ð‹g§åÔ	h/º4i©ùž¨Gt»£Väš›Ÿ’ÿgMwýPÖ Z?Ì)\×½P$âm²PÁª7ñz_^¨	ëÅ?WÝK™ª,ô”™‹Hí#?àQ¡Ù>Çï>ÝQ|ðaêððü>ã*~Í²§R?<U]‘î ûÛ#Ïa"Ø?øa‚,Góg—9	è™Šnp§Eªê%{ºŸŠþ°…(Û`3îÀþ¡ƒÑa‘S {"ÿQœD- ™ñ‘/x=2ÏÈª{Ü}ƒýÀ€”@ÓtÇ–äqÙíCS¨TéGø[¡vþ1“
ÙêM«ïzŽ^¦‰èÃ+cªÔfþ AÆob•@úXú,HoDç·Ê=+:ñÎä&ÒÉ(’Sý”‡ìÚEFÙx…}‘Ó9^z\ÄK{¦åñæÓú¡l]ðTŸ“d:ê¢²ånÂß7žzcÔ¨íóüC”xË>´¶Ww1Dá’?VÚÔèFŒ`†¤rpl@f—jÄ+x
gµ ¹ÁÞ¸˜/™ ðÝprƒîæCÌ*‹Î‡TÏŸL¯àÎ«¢¥CwŸØãáö‰o¦Z3RÐ‡½-&¶A+Œ[Îâþ…Ý-W„ïqv™ï(ßCîåa Ë•£¦í—l}Pý¸"/·ðk&>†ŒäåáŒUÒ¡ŠÈ,üúvßØÏ.ï-q…ÿÿç»³ó¥åßÿ`¯ÙïTÕÈìé!ÕUÌú…æ-‹ðW«¼˜Â¯¿ÿ±ÜIãaØ
qÄÃo|º*a¾t£~ƒ¥ç9Ô'ò¶¹ŸÊnË^âÅ+ï€?Ìý
âahUUÌ_ ÿZÕþg^/€b*Æ`¹8ºñ4íóö-‡Ø­¬Åpù¨*J™™y´«P¨ß¬m‚ùLO>æÃ`¯·I`SÁD,óKIíÿY$x‘ßÙç¼`†^ÆÉ~Ðís»v+µÓ½íQD½n¿‹f@gãÖ~ê@_Q¹›¿Åå:‰½~œG‹âh1¿“Ý ÷•…rã=î8sþ©œø!æ£ç#–øÑ„’Ï;_zÖL|E¸öy>*]Ø;¥8NTS†èÚg/ùt£×ã®Ê{|ÊQú]td=ý i0^ÂÌ_¾¢ë%{´¦ê*nÇ¹f¥&†çˆ¾õálÚí†iŠ¾©[–’Á¤&Lïªì;÷[éô,¾‚9fÐ’wŠ—üü©ˆq™§{KÊD®Þht¼
*E÷ç:·Uü¼óæ·6o?l‡¥“bëóAO< ›LEÉv¥ýŠòñå ÜO³3ÔÆá[ 	v˜àƒº°çOœ¸êæ†Îc9Ä’ƒ¬Ž¡CŠP}Âb)Oq «“¥µÍ‰Ïú|¥U¥^ü¤g¶^0ªxMïÎ{B—‘”~9ÆR:»7“”ö¦dm-Øí/ASºIéÂ¥f(³‹ŸÔI¸9BÛf˜ßá'''©/#)?;%Øˆ‰9"ì­Ì5ÌþQ8äßÒHÊ›Á¢-pêë$ÊT…z+Ý‰½y²wP	Þy lQžv™‡¼®K˜+¢ÉXöJ%îBÍÓ’_ù`Z‰¿ê¨¨…‘Jsàùò“Í„t-ä•"qO„_Ut_–á©AõeÆß`äl×ÛHÕeYü£ÞëGJævNîõHëÎœóK U‹ Öœó+O žKˆ)¯…ó«±—=5/!:–¹XFœøRDâ¥+`aƒn–Ï“$è…ißF¦b"³!hº5ÒÐF"ðßÿõ´×Žd@ØXiª…ÏŠŒI.« Œ¢OyŽÀâ þŽœò¨ŠÜÈ`PØ E	¸J¢Ãp¦¦|?‘}*”}–R@Å8B“‘…dÜU¥ÊÑžºÕ•eÙ²_n“²n­"Ç¹¬j‹õ&xZ¶Zét¸ÂüÖ°?±,ŒÌêÅiVÜç¶Í´žã²¾"’ðÇ'ØÂDzÖ38Òaoù-÷KAJ5Šï³	a•µÌ[Ä]Æq¥ûä[o&ÛºoÍêîÎ ì‘i=£b³gZ·#¸E¼²%mòÜêzbÎï+ÒG×)—/¾îŒâë$?"î¢Æ*xÄ<Œ#-e*ô‹ÿeMØªÈ°ÊÎŽ÷,\9’‡²O`ŒºÓATd<¯3 *P•`
¯Šç†4	á¢®2ì99œÆ_93•T3Ïè}3öÉçÍ0Ç¦ÔÓù1äÍ8;„´Ÿ:ˆ’ð>’;(1×¬YüÓ«0^^V	Ûd6qdSUÙ°Ê½† {éÌ›¿x€/¨/žßâñö°È0û¢ÞrP¬GI)””Æ²œ”–WµYpJ±ã‘Þ:KmF½êìÖeßye¾8ßu¥º®SŸASW ´üÓ…^ÍyÞšl™/á(•K©·³Éå—šn–M-ï„Û¹YÀ”ÈJðÊL×h÷xu‰+ym’+&]_…MIñG,u££®C}eB;.’Êš˜%¥n¼n¿Œ¬³–þO¡`@œ¢·äÖ õª²9¨>£%õR‚UÜ¼oYWéäÜRÜ»V«DzOŽ§"4P¡àž¸’ÝÂ÷ù“hÉ²w=£p"sU•*k¢ü·oY+ëù¶Þ¸åjl?Ó_ÉáI­À?nnJP¯¥³V¼)i¹½Óäz§ I³Lon~¯Y‡g8ü2yn•
Æ´|Ï‹+ÃùÞÙkSËziE£|“íÎÄÄðmØxšŒa¹mïÎÎYÞ>^hSM$=@ÁÁÐB9‘å÷…»R•ù†Z\!l*„çMÄ¾"’ƒ*&1ÆÉp¡ÌÕAbéŸ@yq`U¦(¼|)/•ÜU³¼,•€lÊMãºc®%Vµb»Á˜–ñ¿‡ENZÕSÚ0Ò‚NØÑUõ²^‘¥ÖöÒwV6wÇÞÒ»¶*s­ê¥)¦©u‡$iIjŸ-”¤vSq¥*nîÌÁéÌ Y‘W´4wË	rŒñ½û[í¯\Õ!Î¤£N÷†M”¬éF}½œ³áÂ(Ù)íŽy„„BÁr’àØD;·[uR0¨q¹ÏÑÕÊatÅ¶Ó4e0B§Ú^0	Èí‹¤ºI0‰m	'êZ¯Ö÷ü5L¢ËÞó	’¢õã$J~iSJM|ovÀWÖÊR)€‘Û’ø°¢r¹&Î)S„J&1[J™…>w‚-žÐ²ª|<W2}}‹Gå(—þ³ý\ŠŸxð,ìÈtüßÿõÿdbðªâµøêÏpe2ÜÉ¨ŒÇ»O€WrDøF£	G‹G8†1õ¾ÌkWÞó!Œ¡ÆÑd†ñi ™hó#S• :qæô Ÿa`D=ˆ¹mxöemäü©†ÎO&÷2öT>GWem9Aé±azáý¿aÏ¿aÏ© s^ºÏ>/Îê¶üµAz,´ñpMU`è—p0Q)÷ý ×Èó²=,}màtþÖß0é_5&âW>‚ß€èl. ºy=ÿA×4ùºÂ×œóJ8ú7ÈùÃƒœí…=™¯¼!Èy}@ Ñ?üü2·Ô†‡€ÌU³ºÝëÁh¿
Ö0Ìd¨o°È`ôQB…óùØ\Çn	øOƒ…"È`±Bó²H?ÜxW­tèçö-TâÚL»™÷0‚•ßä›X® :Îà©*.ž°³qØ¥s ¬¶àôE`"x£ÞüOïr÷A;ð½'¼i¦¥á•7Œ2ÝÔôbñÈDäml¶ºBð	¢£¹ö¹×Êù•ÕcpåüªŸ€µ•N“[TW(Pn>m•ª.çô]Ý–”á÷úž¤Ÿ<ïpŽî…¨ñ†àRÑeÚ•Ýù°ª¶žûA4ÏÈ—ý¹*Ï™¶©Ü0¯xMÉ‘ÏÜý|h&`/<QVšŠ [m¨…pñ†W°¾ájçãí«ƒC°CÚ%5f…Ê(B5‰G˜ÿƒ[ö)
ä>fgÚCgý.<—Ävv~„Æ+@"‹Äªâ'«O,Û8Ø9vN:Ía3œ?Uœ)‹%†Lböó4‚]1ØÜ«û½H&8j†˜%’pÌ°ÍßÛ?&º¬íÞ0éèŽ™º`ÑÃ|O€'NQJ´N¥P
¦C\è,j½³¨{?„*›–ðçÆ~âÜŠeÅVxàYŸZKå\¥ºâ[øHkØ›ïìJ9{(_IÂ½hÄçû{ñÿÞ4¡H1£1oÎ-ß‹ ›Û“V¹ËJÑà„q{úûk;±º±l¼Ú™F„ºnèž…ÉSë}¦Ó+ôaá~ZÕ Ç»¡«¬î¾w(kT^®ªNâ$;Ðb‡{‹©Ï¼•¬=çYç=~ŒI+5»3jú¿•Yó(ÿ^]Âú
ý~6½Fy0t|
Õî*`ŸŸe®c“³Ø07sÆ bB5ª8QdUT)œk*óRÛ89-PCˆïå³µjQ›ÒÈÙ)r×”=Õ¿“7Hé|ó
\J/µýÅø†#®8±è"»ŸÆîZ1&Ë6¬]dÎ¥v×9—M9Ö"›ãÞ	K‹…ËÂ¬­j€O)ˆÔø~Öí…<}ájhqn[	Ù6¼„[=<8£óŠîŒdvÖÑ±¾wÜØIkªû­yI3âèªÌCdi¡a;ê»­¬Ô"ƒà"hjkK‚¥²6©à^Ñ×ñÕâ*kÑ
˜õÑ›sÍ/j¬ñÒü4¦‚qtR\%ýIGË¬ï¨ÆŸv®:ìüô7˜ýO_l Oê'ç§ "7ž¶×ÖÖÍM/ŸM)ØTÌ&ÍWšâ’c"-Í0œziJ5VRSþs¡2Ô9ÒZ0¿Áà£†«f+Ô“}MF¬"ÔíÆºpìsj¡~È¦£èoÓE°™ ¿=a™¢|Ó:éG)Ï‰R@÷gâ>ær~ÝÇîÄ@œœ”)viÇ(ÒœbÚ—steä5Ùðû‘×4ŠðµŠÌeËÖ³f­IØàa/×šrøªŸ‹l&Ts bsžêE1Ú)Y½•JÅ‰7u\Þi‚0YT(…/ Ø9ð½¹|…6½®)µÌóÆ‡èm1«5Gn9ð´ñlüçÁY´Kgä#‚é¼„
ÿ3ûîR£k‹ÏnÇIŒTÆÜ›ºdõ®i`zto{i9¯?(Ø«\
DÁKŽÊ÷"tBßéõ›ìå7/¹zã4ìÆ£n4ˆÈ¤¨_h¹ÊxBÝ¼;‹E£§ÁJFé5³´­;ñÂ°·E'ñ5þ]•öé®³(7 ÉÏóôT€³L[\Ý—#ðÇ¥iÈës\ÄñGÆ»ØìÖ)¶NÔ%VLzÍ«–  /Yp‘Æ †˜ö&,å ÃV7¸Ž!7@›þ [ñ–_RÏDm®°ùMNÍò¿¸eJ–™Æg‡{+¬¶ÞuÛétœVì€ú7§÷7bßžô³aÕµtÐ~ÉÆIÃç@‘¡„ÌÌ¹ŒKnášÞÞèryèLHQ3N”÷Ê¥­ïÒâO³–%4»:ª¼>ž•Z€Äð:7µ|¼¸ÍBžø¥L1K¦´q
ð6ø£‰¹#I´Ü¾Ñ3¾%rÈ|JâÑ/cK}ðNò„’[î2m¥½I?qÞûÈž0àïn	0ð §+Òu•%\É†¨ * ÒË’˜-v¦	é^`(C<MµHAöøa«™V{DöoÀobL!yjvlè•jŒÎNÀâ‘Ô!ªgŽ12zŒŠ] ª7Â]üÆf—·Â|‰
ýRº‚óì×ŠÛW#Ê0µÑ~L®ýfÐ\÷aQ‰€éöÜt°™"ú£Mg“h·˜¢—*rk¯£ý˜Ýdƒ+åësaòQŠ„<Í¶”¹É<I¶Íû0—£Qœ˜œ¸|å{è©õ+jõ¢ØÐ€šCB
Ö’(~`ËÍb7™¤q0c`œi%®p¬çÓM©–vû ­˜¸þ&>ÿ›ö#ÑäV\™^lë?ÄôŒždtæFR%ÆXÍ²îpÒYa¹1O‹È»ôÑ£7ùÊ«U~Íz^ÃýmüæUKÉVe©WMïä€¥wofè'åZaZ|ô›:ÀnPrÏhä].PÍÊ®Ò’ðþ3
,_‰ww8.ŽK¬þMDÙU³,1Û“dV»»ñpœ„ýÖúO"Ç¦*âÜÊâ°“Š¥EÎP¹D„¤ð…Ðü˜‰$ÆÍÏD*¶á™(f\q&þóÌ>zAïÙÇï^pöátøÎ¤E›ž–®Œ­"2ÛóÒæ²×Å6ñ¶‚…Z1»]ÙEJsz;K5B-P=ŠŠ0ÝäªItŠòBÐËjÅ\Ñ rd/ÍÎdü¯ žäTn\AI¥Ñ¬ŠÊUQÉ\øzÕ’hwämäHå4¬œLh6}ÓZók-Qõ£­&®²xM¹ú2Ý´¨~TÛAêRÎ-NbñMaZì¹éEûÞ”¦¢Ä6íTíXR_¦T:ß_[žeï®(Ìô³iLÃ9Ó¾)Ó¯A™
"!1f–º2åIµiªÔ|cÃê[o7ó9Ïmào:Ý¦Ó‘Ø¬iU7í'‹KN²¯Acï{ëéXš}AYSûmpµÕm<¥i?×C±Â‘äéÌå[d&‘˜Ó›/Îô-‘]˜ Vä£ûÈJGi¯®Ÿg?e	ÄQ›ÏPS"A=›Ç³‚ñ;áí™èq‹`9àe¬:M“Çg›LGaÈ´‹ü­}cÎÍã:L+,;£Ç7–ü†A9B”L'Ë» QFN[£®Œ™¡TÏ2Á¸4ç’NVÑ¤.ÕZLnÏ^CŸøukiµô˜w¹œk¥îÞ/êmû¨êéVø;ºÎ¤Ü…ÞDÞž`€xêFzS&ù®êÎ^gê¾¯B\MÃP•ü¬~â3(Áš’É’ÑìY!wŽÄºØ³¦é±™›™n­bóu>Z‘†Êùì¥Éž¸9yÝÏQ s–°CGç*Ac#äêžµ¶9kØs°ôú–I:µDŠ‰m+ÈœÆš6ÌÎ÷kšç*‰-ZÉôÝnDú1ÕB\µbŸSNÅT’év.Éê@ëþ-µ†°O6&ÛKÁcËqPËÒ³Çƒ0I’Õ¶Dq K'œ$‹@©%™ñÇèÞÂMµÑ£s‡¿†b-Cû¬ö¸è„ ÙzÞ$|Yò%¥Þ­zÑ6û9s.`ÏÊÕYŠÓàmæ<¨,öØÖì[vù^³›µ³;y+üÕ’Ó#Í 	X //+ŽjxKðû{lKT"¾ÛÙ/Õaûnç¢Bù'•$~ø_XD‰µß0k´°óÏe ð Ÿa*•ù/<Psÿt—ßÐá’*ˆm6gu‡øò¡Fð­tÉ“'pYpô}‡ñ9IøWi.-ƒÞJ&­V°rAÛõ,el@ŒÑË«prCØÌ·ó‹Å‹–7µª0‘•G,6™ÊåG5ÊªHÕøê4„‘ê$Áå,á)©®öã™ìž»w¶í2ÏÊt8‚¹ FÀvÖÁì4¾6diVj.™fdå>[[Ý\Sì‡‡´”Ém³Kó¦]™³­\„ÎF.è£«“¿UäxÔ3¶“O»ZÉ ñŽ™\¼«Ÿ˜ãUyòÆb>Ê²PÌQt#šo%ð¶š¿·Vê$3í¢¨»™¶gP…†ëê¢nÓÝ©q‚ÝŠ³:uNë™¦ŠÉ,e\y0¥k»6ip§²ã/öé<-
Z‹&¦éiê×LBk5Iä+ ,±òÏN:D“ÖR–”$„JRL;ÎÃU°zÑ+ë‘¿ÒR¡¯¢²PñÈ}+)YCEÅ?³¢±eÏ¡âñByâ ±¦&	ÊŽJ8%E›Ÿ9'Æ(VÓ'÷ÍÏjÚ‘b}ÞÎ¯÷*X‘«WºÇ:r÷Âœ‚YÌÿŠ™í©h²
ãý–}øþqáÇ»8ŸÚóN¨zz»²íÔƒ+ì\¿Õ_g<ë»Ÿµ‚*w®NÂpûiå«UÝ5W`N&g,´ÊÂÓÊ*lòðRõ‹BJ¶§áwAñÅL)útü›ðT.Ñ:­¡ßÂ éPS/šBD-.!Ï÷yÃßÆ¤9É·Å#ZIämO[ÿÊ.N~ï²ÌæY—ßní±óãóí÷NîîßBG?µ`Ñã…*oàôÀV¿MX9ÀKÕ˜Ëz‘T%•jY9ègŽ’)úßÇWôXžœ$+ƒn€ë¹N~IÁ¬£2ôD&–§4	ïû)W!j:=Q¿ýžÐûœ1J­Q´G¬t¶ëgq6xv5¿¶ü»A|AÛ²0™˜88¼@¿w¢ôùª$%z¦ã,Ã±3	¤øú÷öïHÑ9.e±fB)gA¿´DÀé°”ñW	x7ƒï"¤’ÕéM”5ú(Æ„`rÇlÌ^" &å<ÛÙØqüP>–+Œæ#þ1\3ßXž.”¦`ŒÌïå6[ÈâyŠ}EæÀöOabXH—ŠÑÌ_XNMâ¶ZiöLZ¥G‘œ<·•yoÒ/…Þ«&CÌûø/Fuý¹©…ÚG—Q24i¡.¿ôMk0É)AïUÉ¡ÁÍj}=´«ŒŸ@êx>l]¤IÞâÚÈ•a½žNªÝ©Uz©Y/ƒCè>°õµŠ´Ûs)¢E²zu˜ã/^x†m¹ò@ÖíøySDj³Û{•à<on¯û Š“½b•mX$F¬äh08é]Î¹>üPo}(ÑvÔ]0
äFZö5db×³ôÕYJLŒ›´_šuâÁó¤£P*â»8-:xoÅ¾ƒéŸ£Xá6µÝ´Ç%\—=Ž*„cÿ« R¯\Ò	€,³×¬w§¸%GGtÄþ€oÐ—;úŸìô“ð/tê48Ÿ’	¿àÉØ÷ÿùØ1´tž"Û·t×÷.?ØJÏ^=‰è;d·z-¬×ý*]ì4ûr¨„CÐK•EW/}›úü_¼ÞgÎ\&±Œf"qNÄŒÙ¹=ìµ–PpÛ §íˆnÍ¤ìÇóŸÞàwq¯MÒÅ«€uþ@³kOü{4‚©cGjÁ¸%·ÎgCqÊj|ÛcÄlêzìŽu) ½\mºö-ÏrE€ÚìUøU`.þkŸÁÄ%_súÒ3uRŽzÍ^*uÎ©[X*¯ÛãhNNZ=:¤tÊv“¦éÿËé{´âáñÂäà{«¼ô¹vT¿jw»ñ{$Ö!«I5w‚÷‚“â9ö¾Õ¨\t¯d¶Úë9€‰ÐP%×À¦Qz`dëŠÚÔ!iÒ}=ûP2P`œï¾ŸÄñà"H^¯}¸3ÕPŽîèóÿ”12ç}”û¤h4Â‹ŽT1Í¥Ñ]¾¥4F^®œ*}{¤—H6Ÿ‚ó¸cËîGÌ<}õÂÔÄê¢!%’¢Œ6€Þû i<‹YÂØƒSá?Ÿr¢Á›ŠÏ'ì×0ÉŽ**öçð<nÏËŠ…Û3â-Ya£#É_’,gš=ýóéŸKá#U¨_² Å•Ã^Ú‰FÝÁú»v¢ÄYX5Q’Nö³¦•ÿûÚå{á}_«Çþ&,ï­?6TL1 P+Â5Þƒ5²^€5
#kÆ(ñIRØš<ž	ËA6¿'ÑU4º+_“>A”ÞR9·8r¯ÏÂQdç—ÓÃÝx8ŽG¸¥¢q¸ƒMŒÒ#¢Ïß‘£ûf*Ä‡R½;HÎjH‹[×;¯»z6o¤cT\ƒÿ\uÿX‰úÎ eÊ‰©ltm »([nÍ±v|2óYúc`­&$§c?žN0³‰B‰•_xzvêx…LžýÅd¤ž¼Rý%åèb‡ÃJTñsB*Î{‰wÂìì¾£Å|¢;ò›ó]¼Tø©ÒÑÔæ9ã©í5ÝŽ¤ö¶X][tµ`Ï>Pˆ³Î³¨}–ÂlóóñŒ´†ˆÓ dè¸`ý=ì-È	L" ©“Ý #x—œa0{”ÊqŽxß
X¦Ã¨óuWF _wÆâŽaÖRtšŽok{müâÖýEÅÆÃ¼I«Ëõ¥o!~>å;­Âz)Å}f¥)Ù('a¥ÍÄü»·Kºb÷`Tî™‘–ã~Ôe»AÒóRò†ê¡úçNÕðÌá‚öõäÔv­Å
¼ÒkÉX{DJŒò9‹•¦Æ‰uØÂóƒÛ¦á•-mˆ2sÕ¸ñÁO0˜€Ê“ƒ€PR°B•	¯½yß„Ý÷ð/TŽiÞF—TYæŽmí*ôfÆ]jDoH,ñ!M‘Â[tC?ØA-\¿Ôµ8*ý½.Wû–-Q¢°\Ù¨müO6Mv658î#P­„0œRâ¬‹Øu4éƒƒ:Ó:TèFÉNÀºËøT=M"KöˆÌP°ô²Ÿ'ânyu’Ôï5I­myÐg¢øp+áý–¶D!¼P|	&	2]àpûŒö[6¤1abtKL8ô¦n>Uv“`4Üò]çÊDJÐ>’pv£Éí
 >á‹ü
”Ý#©Cƒo/{ìOB#¿¦uµtŸ‡9õüOá$èzuù©ñB¡ô;æ_ŸæY>‘…WÃ®Æùmrz[†ÇqJêJ¤b5lEÔÈ¤4Û%Q+Š=õN!é`¦fÁÒ$ùÞ6Æj,…£öáÑ’=Ù<ilî£+Ižxª’ºýh (}Q!øÉµ†8pžë„ù¡÷«HáÒlŸf4Þ²[K‘B)/íù‡ßugÇûä†øb½ÇeJß5(q.Ý¾%˜HRÛ_¿Gø_â|ô¼Û2¨H‹¹nK‹ùƒ5ÀG4™›>å–o;ÆN nò~¸œþÒf"+˜¶[³ã¬û–ë"±dÁ÷é‘ÖPÈˆUUX3Š4èêPä!Ž4O_I?•À²žïØtûl»Kï>˜Dƒp7“kW•òêzä#%PéÜWö‹Îð¬û>í4Ñú™^æíƒ–ÕMÒÇQØûù”fŸî>Ò·4Òqä äÖÆnÉR¥‰ìR?¿ÂŽ-Ï¯z/g…×B¢J˜ýú¦G¥¥&2K¬Ø~‹5h“ôb±Æ¥]úþu–¼äÈ@Wqä]_¿rÄ:Û¡2SvÎ'´Ÿ5žûûçµçSN‚~„¿«pZæ…õ.SB>30B}çöSŠÅ&oÁdËâ¸¤Ž³Í¾ªæcû}de?9z·ÆªÉñ'Øü“ ÷müóOôŽïðs¸ÐÙ ? °-`f'¸uC^qÿa/;ƒy³(Íìgša¯†Ì—o.yn#Þ¿ÿ4Ã>€wsÛ®Qñ5£>‡œÔŽýÛ¢ëØS8°™ƒ6ú^k”¡º[.ó½æ1w§2’^±»#W‘f‚žu<ç_5"|t•âÄø·õ5ÓÁõ«ØC¾ßèïàÎ¿8È‡§-8¸Íû¶ea°¦î½Ä´JÿøüJ¾’üw™Õ“š(×†‹”Þª1¤Ï‹"ÒÇxDTBúl¢=/?÷øÉ#»ð¤”êõÀýlÖÃý<õÇý¼ÃÛF¨þ)œ'ú¹0yÝÄ©0þÝòW…ÿŸ#­Ê½»aÝŸPGåÇ{ò ÷þ7S®•±`Õ²¦P¹ŸX]e»˜
lšØß¦ar‹ã$L¡s-Ï`„€Øõ£tÃ3 `õ_:I8†IžM0ã¨þÏ§þ(¼æ°>tÏxà¨·5»[Éã@È™±"ªj<Ð ;§F’ú§Œ2Ý-.î*¾ 
ß\3 Ä‰õ½!ÄçÃxÜXNííÌ¯š¬FÊƒEpÒ	+§oQ‹Þ3x®„ƒïIP€ŽÀÖåF.¦ªDQ¹^Ç†äE*‚I‹¯¥H®„õ¹òY)~Ætª•gmº5lµ&Ìê²‹ÒL“¾ém¬ùËŠb»ˆùðEÙbjB§á%y÷Z—q„à¼('’•–D:ŽFv7ülgZ”+0ën¨a~2Ä*ƒ”£Q.¹\âKA[ÂPx:N=0B‘ÈƒFæ;óÐ|GMñ£ÆŽ/&Z±pcÜÇ@a"µÌA%i¤4­^i-T$Yý”úç<‰À /R“Ü¸ÿ´ŠGœåŸ†]°ÙQ<Õ¯óôé"¸RL_š§­©ê#s,¦4$qbŒ#ìÖ÷V¦˜‘‚=ùÛ4žü¹¸C¼ã?wØ	¨.Ðr‰3é‡R'Nª æ`#BŒ¨T3ðf ªe•@$ÙHl~4ií&ålÒzYQ}è0yýˆ3®1ëÓØþœzhyËŸ;ÄÆØ±É&ÒÁ_‹·t~AG†XË9(ßØ\7Ê	©z É(ðv5ð¸òEkÚzÁŽÚs|ú}Ö êã¦zi(:“$¶–Ýq¿ÆÍ‡µ(Wx23¨ƒº?ö¶ûY²Ù!¡NGžÉ[l»FâolUvS»¾)Sg±t®•ææàv=÷'K/ÜNÐ»
­¹×ºlö¶Fe`Ÿ}yÔÃy¬§æ®ëØègëBBwRMª—¬a-N‰âÈË¢¯YAôí¥í$Þé±ØÆË2z¶¸jd‚»Ÿ} 
Yögå’}¢ûG¾óê¬ÞdÌ ³—b@ $óuW<3‘¹Ë°’×Û,šCòÏ±4ƒ¶±Øa¹VÉÃaAwËmaí•½BEK«¢§Žw·ß¿ÿm¿?üû{ä»zxp¸¿G’µwº}pÎvß¿g­Ã#vrzüîtÿìlÙœózW÷öØyÑÔGIÅVØÆcÖ+Òª¥EghhŠ›õ£2Ñ=õ~¾l”›=Bä{+'HnÙ_Â[Ä¿'Q×Î7R½Ãés…LÅ¿ÃHç@ÀËÖø]V¶ c*Œ›o¯Ü#`ƒÎÌ°°³RÄ{8Iu<Õ¿@ägèI[€D­ÕñÆZÝYpÎûÔ­åŸÅƒx¯>è>÷
B0Vgs‘½ôÕÆƒx>èþÖI'>_§g§ËnP=ï£ÝƒÞÁVN¢ÉôRÀ)ÎÃÆýõvµO8ˆ±¶¦:¹F8ï°{²€NC0€>¦èPÂÐF›ñÃU["nvp^YÏeªauoª?üãÇ6.ù˜ð²pP´`…:]¾ý*º¶ÃÔIŸn¬V‹	zôFë&'
Ý9h&’_Ê Ñ}%@”½`ZúZ®Y<Oç™Ó¾t;&ŠždÎ#Ö'ëÕ7ækÍTÄ”`ºC}…éòœ#¦E«c3ñCÜÖÃ°œ´ÐmÂæ<}QÄåTEáÞŠ‚=dûVþ‘%3 ‘…¶wŽ>BÈÁˆ$Œ+,êÝTú±ÆÌ/Cê²2°Ì\ìÆÁm(K"&È±Z]­(ËæÇx9Á#êpxá»²îËæ…=Ç5ÏtÄÃ=ê3B%"µ¯Û<F{=; ]sLv{aÖ (/òAçZmÒÄºúý:’ý‰­W(PgMóš“ùgNVM÷:®´xþwS{QK£¡™‘<R¦ÀÈ<Â5x‚³~H+ì¦Fß®Y§yÙÍ|Î2Ïƒ°Á¼^n6¯ö–}8AÍ°ÅÏÆœõ’78Oj‘^’l1S~òè¯wõ»Éêˆ¬S)¥xûl|}5òã¦œ¬^óÎð‘"Vêª[æäœÛì>»Ó=÷q<íÉž0¤n‡•`8ö6äD""yÌmµ¬õxã^Ég™0.5)€²²*øÜlÜ‹®ñv*—aŽºa}›¯ó¬‰ó'{m<ahÏD€R'^¤\T×VØ‹eß_pË° ¶Ÿßª¿‡|S>¸Í¤È¯›ÞÌFá5¹ÉZËÖ‰zo»O=.¤Jë;rD¦d2oR$3’O8-Õq€á„ä=JX­âšøbþ°"*ý£œðãES2Ì,à©€t$Ðbj”H„V%¿.´Æ½Ê‚í¶þå~£ å@gqŽ‹ùnôœgÝ;/ÇˆÂ©|ºfTÜÞvä0$ã“ÄãéØÎ”Ú“·ÓÝtÆ8ƒà¢t{€éHnBZ¤~Ñ·µ–2;áMŠstE?zšÞÎ-í£‘lömgJ{^­°8„Á¨åWèZ„Ü´‡=r)ÎñX§6
Löåe[ãˆ*—G»ùW*Âã²×4[	3ñXí‘Ïé]"ºlJ)&4À1¿ûˆ‚îñ‚µºL‰Ó[âÉÜÚ¤¨Tx¢Ÿ8¹m.VïùüY¹Ÿ­¹#ôþ?   ÿÿ @Rºxœì=érÛFšÿómmjDÆ"%ÙV&‘-©³®¯åÌî”Ëµ†È–ˆ1p Ð’†Ãª­}”}‹}y’ý¾¯ô	‚’|ÍDU‰	 »ÑýÝºwv6¿aÖßƒqú²¤ª^$S~°qr68Ÿ¤5ge1ÏÇ|<¸s‘±ir18LÇì|p:Ï26|ÏªI2.ÎéñIQŽy)ÿð)/“:îì°ºLòê´(§âWZ§E>H²Ì›†?‘ÓŒ_0˜È´TuRÖì,™î:z]ïfƒ»zþ4ájR¦ùûÁûv‘V³’'ãËçÉlÆÇìˆmÂš“é	L}§Ì/jyùûM¶OÏÕ¢tuÛ,ß-CÓb,ð²3^Ö¯Ë4ÉÏ2n.ø€:|¿Á¶á¥~­xùhÂGï#m–!@l$:@ˆ€;ØeÓ4¼úD!;Mîš}hÙÙ;-òzpRdcˆ³2¹ü`É`i~6¨òùYdÄ D6_&—œ3y›¥9ûS1/ÙÃÑ0Xì¯øYZÕ¼äãÍ“»‘%Í¼]TÆ:¾‡uLëÁî:óïEšÂÛªY’ÇFÂ¿×“´b?%ù{µZöbŽ”Çàv"W[êÕ²$‡ÿ>$i–œ é  .P‰ìÚ{PÕe‘Ÿ.æ@:¯€¾8.“ýíoŒnýÞ¤N%>Þ\>Ø–}úC„:%9§%ÕÙ%öÄ÷<¹˜ñ|œÖó’W4‹ç|ZT@	%{
\6ŒC`»}ÀíÇ€ÞyR™`;¹d4# j]4Ði Ð `šÌXcóHšÍˆþêB ¡Éž=fÛŒ†Yð²³	zxš¼çøš)ÏëŠ§õ¤˜×Œf‹ÏgY:JP,óˆÙêÊ qÌ,(6ÂòDÞ^)Â“î®¸;Üc B«æ?ØÛÑ’ÙÓ¢êXÂˆ.w˜äÑ.r_ðó›ÝÝÙÅ[WJ©—ï°9°m9J€¶A%Þ£Ä:Oáí‡\aBà<æ50[µ¿–XežÑœ^Nx}Îy.uØ+]v†{1‹˜óD‚	gNKšòq:ŸnŠb³ý6Œ{CF¤÷ÆáBÓý¼'Ùã¤N†Šèö0—Lyã—*[:\\Ó"/\‚Rúÿx›a yÏ@³Å‹é+bÛHqUÓ!Rü‹ÅÅ¸ûkB×³§ÇØ£b|Ân0%F3ˆü‡5ˆ<=­F8Û/¶n^…â;¾«_ÍËKzßï~Qˆ7Œõ0òZNq5V>2^LXµ ¤)Aý÷‰ý—2=Kó$¦ÄU(>Ï˜HñÌŸ›‘k{Œ KÈ¨ ÏlºO¿ËâÜ±‚?žäúÝ‰ÀûÁÉ¼®‹<¸úr†.-5Ø¶(òG@wï½>;8d¯‡)±—ƒ±ž,Mø›ÓýóA2Ã´Á=$—;¨%AèBNŠ ‰Ô;Ê©¤+ð(MI9)†þ±\i°ÿëôß¯FIÆ?î±Ñ¼¬Šr0+R„chÙaòx”ä#ž‘. r·º:Dmxê‚)WpD2o%D½.Ðaîu’#4®DäžF¤Òßß›¸4Bkv ¥.ePe:^UfDG~‘—È2`†„aÞ‘ Ã=6ÜÃ C›Ð?|þð%ÃØï„4kÐqšŠ»…7NCWã÷¯ˆPÖÃÿ/ÿ¶Å~…9”U]ãõçƒ+¨07½[ýûúbÙïÁÈúr±ýÅØ+^ÏËœ¢Šƒ òáÑ4)ßWŒÐÉzŸüBa‹‡ãiš÷ÙwÛÍDSªé®ñïšbž^€°Ks  y 'x“ÛßÃoøw\3¸3/	‰«xuöWt’<¢ŸšÂª’1ý;ž—	aÝwÖÝfgB·"ìÙ¸Õž¹‰E¼ü+(Îp|nt×5·P~ÝUó:a³“ÁÝõ‚Æ†õôû’‰a'âaaPŽ™€]¢¼ëúcÝ£´'Þ¸VœVòñ™Í`7qõ,âÈ„h&ÿãÖ!Þ§ÿ½(–ìïÿý¿Q¾-K3H·ùJCo.2æÝŽê­U:Ë×WÏÃëŠë+‹Äw5ƒHh°pO«';nÛ©»J0q æ{QVþ¬ƒðúO›	ö€	ö‚LS0aC+ƒãùDþ%‹˜]tV¦àYÀÿp!ò2®ßÛAiîK…˜ÄŒ*c¼ößìî`üÒæ“¬½G¼¨A=œbÀ§ÅIëÄÒ2¡-ÊÊØßÿçÿÏ“z2$0ôD¨©×Ê`5Îô¯_ÕgU™¬s'±:ê³Ó‡[?hÙ× oÏz›<<{±öâ# ‰ñäaå9 u²ý’—i1^+Ÿ¦>~X‰R§ó'žƒŸHd]W_ x(—ó$zâÕÑ0ãùØ‹HlK&ï^lqIõïsåÚ8}Yrç]D•2ºvƒ1YrÂ3ËD˜³–tˆ
4!(fi˜Ô¯xRÁ?r¾û(Ð:šgî9/“™öQƒ…7AÜnR¾O¥<Á0¡hñ4­ÀnM6·ÂþXÌGÚÛì§#&$bVö:þå	´øšcè?ÅœëH€%Úg>›¥ÈÙ£9¦ó¶Oð£bv‰o¬àQ¬ó3^R'0Ó?¿e6ž€Í3ò/Ë¢Üñv8Mf½ÞŒðDöÂUÜè÷üò`!Æˆ¹¿Wð´‘f`ŸyU£ÛZ°Ø‹qo)ØæþŠr~®ì@:¾žÂ<ŽØ»oö½åûV­öÛgâgü%-Ø‚‡ÃÈì·˜œþ¾9Áe?ö¢e‡ÈÃìÌ•H…phÑñ]tlÑ°óD“€£²
¥(·…°mú4¡K×$¼Vâ6k%¾ÖøÃšˆ€„V˜y@HJž¬#œ×Ív|ge ßÈb–¿cÏ€äK!#ªˆV.9¦%aÌïâZ,&¿k	†ÀtÊâ¼:XÜaéC’ÍùATKNõ¡1Ag¤_åðtd7>„œñzH³–½w–'€^l<¹ÌÐ•¨FezyÂ…Ð­PŒôë°Êä/óTÖ˜`+ªîèñáÙ½„Q*ª×À¸†ÐS#Ìñ¡;-
J~N.¦)¦€#±ñAê*©¤N8´ÇaªùÉ4­Q³ôaÝ!žòƒ’Ây	y,wí$ƒY„éå–~$g}4¯öÑéc\(§i‡ó:Ks>È‹œëhCÉ³ä‚ýy‡\@?f2-ƒ\òzÂ	ê‚AÎQõÇi(¾DÍÊb
“Ë±À©.¤OxŽW—ìCÊÏ	)°úšÝà/2—+ãþD²«‹É	16«Ñcˆ­Còçs èX¨6>X¤Õ1ÖågMÿU‘Šž6[¥£‚"8œ"‹ë¤€¿Hs=*òÓ´œ"JFïŸe;([aîÎ­ÒY™G­hZ¼k’WË^„¶X?ã$¼¾×[]Ä1K‚ïÚ±_ô>åÌò¢ÆÅç|2Æñoßè{ß6sò?N&'4èò 7P±VáMðÿ±Ð3ÙArxˆx0º5(@Q].½éÑ‰ÇëeVD÷Zo97‚9¡—¨öhF<NR”lbš²òó@Ø;ÿsg
ŠŽÿ–jl#	íKÉë.‰ _óqqç+L 5Œt,O.ª¾?úÇcÒHò'Òî+Oýx«úl‰Ÿ+YP_UH²y§D)4ÿ3AnkËµuù-Ô-Í†¿å€âÀù\ ßÒ¸kf›;'™›¨­Ò§neUHe•Æ¥/«Vl’TvÉ1YÉ1>""^hæýef¥V¨¸jH†ú	iM60›çˆü
¯ëBˆõ¸LNký	«èÓ*úúŠ¢.Rzo±dŒ¡†Á5XðÉx¼=æ¯¹£¾…Z+4£’rfŽëÍÉ5FÙ—”’{Á ·20*?ž¢DÝ¶ˆ{Ê©by/5€D¶ì¡;¡`˜×ÉÅªÞ’&¶1øÖKÄ@«_:fS±
é$‚‚@þhM¬‘l¢Â,‡‘Ã1œ/Z¿¨ôœo%ël÷è3¤æ\2”›“m~KÎ‰¿/<9'åñ¶ž#Æš”*$‰ý’swW$ç¢L¼NjîŠ¬¶~Zî¸ÆïŸÏ'—d*ÀÂÙLÂ—ìŒZÛ,"ÿ–;ZI¤ÜTåHNnÈ–ÒôpKåÜ„ÂÚë¥Ú$ß|®M;¦Ÿ9Ùö‘ø€àï¦ÜÄ‡êr×‚i•Ê=TBMÙx¿%ÖÖë˜©)a°ço	µ«$ÔÜ× ´Î/EÅçW“;»ÊªZsf†Ä3˜«dÍd×ïvÓGÌ›Âõg&4q[Cû«„™¤ºÆ‘ûÄ™2ãþ	†.<?çB’>J ¨E—Å(-¯@¡C³Šü´Úß¼Ý±5ªÑ˜W„ Ñæñüä¡w¯0þ4Ÿ¡£- Š%¢¦F€n ~Š&oêÕÇi*u	³Ã¨ÌùpÃ‰‚-6š¤ÙÆØÖ?ß–'l¼õS’¡ðQ7ôÖ-ÏÆ[zÛ¹šçÇ´Ú-ú-×u)®ŽõBÕ‚‹¯Ä¸øSbý›å>(ÃË>¹ZÂz# ŠïEp,/ÞÂ²`
dðô67ÉQíi&ª½¼hi/1s){<Ô—mïhÖ¤^dÞiëYè¥íëh®ffÜhédŠ´	Ã#{RÏ§Ö­–¾¥B/t’¨¶Z$@¹!³3„úY4–dé_ùZPÇWü#ZÈæc¨±…Þ¶;jêk:›Jªl
CFGq
–Þø°§Bgè½½Í^Td6KÈE/Na‰ê•#"piÂˆONOávÏôøÓSÖó¤¯Ã–óo.dÁLÚUTiÑ›CG‚8\j¬˜$Á–÷­÷Wöôú·,ÈUx§åÞ/'†§Ã÷ü²ê¹+è«bÊ{ðŒÂ0Þ“Šxõ4xËn°^ø}f;2ÿ hÝ™÷@KxÄÚM$ÂC”¿.ÕX(<üÿrË–37Œ«€|}«	ó‘¬[çää.™žHr—DÊzÆNc0lvÙo¥[ôì ³iôÛÄ¢G¾Gè™„ ÜP?C`¬q¯—HQ¯K†é˜^b.ºoÓ§ðe•ð8°Æ?–P±„Ì‘©uÀ<íð4IÌáªì×ÀÂ“IR&»ð1 &ƒŠW.«j–U©e‰1îG† Lù9âíEò×ŽÙñ/O6ï7ÞàýáÙYÉÏïj!7}’ò2)GàÊ#U`4_ãÌ@äEo>~Z”b56®2°o-lVP¢Ëtag5Ô@&€47õ •d\;ÉnjÞóo»Nsº†xo5YÓ}uðYcøí>ž8 ~>Ý×A\¦g À~*ö«@ÅO@Ou$éç4ÜSCm7è­àŠâN‡.Íá_)>9îUK6éÒìoÇkåX°jv›õÄ:Ž†2¼C9ec{-ºS/7äÑÈ÷z×št•Õªé–·!‚¯'ø0(ŸŒñ¤_-¸Ä@D ®vÖBNÔ¢[u·+êö8¸B5¾C6ý¯Î—ÆÖfOoå1}{‡ªóŠì‘Í’ÿ™Znú3kœ¡£aDG1bGÉh>Lé¨ŒêrÎMgßÚ#,WLn'Îã’åôx¢MÝ ¤xÔL õ¢êŽ+0,geFõ-Ž
4€÷Ê!LwìõsñÊ¾b<«¸ñœTè=8ÁÆbï0-Ý¸mb¦àMí:¦_ÄÆ“¡!‹/\yÜžSmyå(t,?»±ÏŠUåÇúÅ1+rs„VKÐ"y$b´ŽLÂÖl€?,(¤Y¾]ƒ‹0wô¥To Ô]s‚àš­ßPÎ÷«Öåæfè¡%k"V·P1…ès-ŸÃÏØs§-³6,–|S6YK	„#zÐ	2ôÚÚ: §¦Ä¸’DXõúº‹!CuqõA k¨_²ª“5õ*C†:4hTÔp‡¯‰çÔ0mÉHßltÇ@¾¸¢éº è_°â5Á*šw*µîRÖx8¶à´}u„yZobpÂôÀè£iá£]o)/šCcó£un©b¹CÐç!×!-«º“ÏéŒûfçmÀ%À9©!ûŽ¿áÅýtK²Y­qšÐÿ².¬ŠBB”ÙF=®÷OSt¼Edy£ÄÁÝ£z-Þ¦æ"Äåßrû Ò·†03   %Zðìey%Þla¼ølÛq'}° DùÚÅE„Ze~ä7!º¾Õù0•©è™€ÓTêf>z!Ø™­ý¦áv
 ½É ž³\\Ï†„laÍ-Ã¸D¬è¬'UU›q‹-˜7(%ðÖÌ#¤s ž^ð¼¼¬'T!\1>Õ—hçŸ–¼š°:­e#šp’ùÈîK^ÝÞÓ"ôË¨t:JÊ}¬N=Xl–,\Ô¸I BœªsÎ¦éÙ¤fçINeÖï9Ÿ1Lp ¼ÆT9VªRáÖ²¼s[ù¬oµØg1¬€ÁJ(4ÚÙ2l”ù")C|ØËÒq’þ–\—ÊÙúM	)Aòh%U$Ò¥1GÝËaƒƒÌPR)ë*£Ã)Žu"^¦ä•ý˜7Ç„œ\2
¢-È·’Ï8†WDÅˆõŒ×¿æé_æü§K$;P(ó7o7[i~ŽÙ¿œŸ³c^÷ä¢d\zªè¦ðk-Wû‰h²ÐTŒ3ÙÉ”fôºá$©zMÇFþÁ£d<<’CßÒ#I‰H–÷˜rŠ&¤8aê)MÜw]cy¦#Àª›¶qp¥Ë$`H¬ó’‹žªŸ€ÑÈ: \¹a‹†)¤ùTTPü1åç›˜§	‹ûº§Ù±jôè-#´ë[3L8l„À+nõÌ:¹4J³24F:…âžh2Âò)¨¬jÐû¶BÊ¡VF€Õ€uF—E”³‚„C£Ëþ–*Á ©ÛÖ„†³cc‰½‘ˆU'ãP‰QWãŠ—Â¹9B1 d_u!SÑU]QWíOø]…ÄW†¤üê¨	ÿŸÌkº×0G´I¹¸5!Z6€ðÖàù2¦ãscÌŸ1Fºoˆ5YÌd7>ê¹åGc»Ìç»Ì4ØÅ¢º²^vÞË"Ìgu0ì`	øõ¸µa¡­`ÁJGÖÕoA>vÊÝ\ao=vE~˜•­ËÌv×ÊJ?ãÍå§þA*´MBñWF‰fÛ	0Yô5ÕI<væ¦;íz„&ã"H½E€h[ÕODÂ£Û¢‰†œå®€)GáHVO”ô`®ÒÇ¦‰Ê	ª­º”Õ/l.ýæö°ˆX$è•P_‹R™KlêwC§,JSfj¶ÑGQ’RyW«ŸAIG­”ìm³ÙQ„ó¬;a —Y" eTƒÈ’mÊ©Ì±¥&o)„u ¢“—I=a´3+F	°êÂöíHã¯aÛšFJ5Œ\­2Y’vØrÝMðÚxZuë–‹¿ÐàqñRÄ4cKŽ8 xJ+§+YN¡í¡Áo¾‘ë3c«•»**„mLŽ@Å´”Œo	ÑŽ¸àpel¨2>Œ$y£«3)}6(éõ¿žß\²«Løh´ûpSÃ#o·à­y-¿ yÆêü¯Â ao®…ÒCì'u,C‹ê2ÑF	CƒêWðD B4Têhønu¢ÝªD-VHÚUˆ›ù¯"Së\Txõ‚ÂeÐ:[¶KÒzuÝŒ˜¸š¦êåá¼.M¡~Vœ¥#­ë .m¾&Âä‰Øo–W0ÛobÅ0Ûô	2‘Øã#í(â–G'LeÃFXRÎý*on˜ÉywÌ¸m¦r3°8Ø]¿mc;ÌnÎÊŸ‹HZ®ù1ÃUµ¤Œ‘’'/Ì˜`T%š±pÁ™ælñCÀÐ‹¼@^¸0Þ´[¹ÀÀ˜üOÇ±0z¤CsU}Zëò¬"˜>Úã	€À¬Ü:<™L3SÁ’º M
ýÆˆÑƒ©?J„W©ç
–gi„êY8U+î=Å#8â®›>¤Œ–ÛÃ/H±¬¹ãž¸°*¨¨o4ƒ©Ó·èÆ·ç:ÍOBûì6û¨Ã§ÅèÓ³«ÜBÉù¥ùÌÛëaa>µÓ®öF˜±r[š‰6ï©“Kó{{é6¿‰NJú£»¬i· ág	F‡_ÍæÖ¾/jã÷ƒ…‘8UÅCÑˆ]3‚õµn1#ÙØ8\4LgtÆtþ`°(¹ëP%wá%²·ðƒb1žùŠ…“<ûVžší-„&ßR^Â?òäw9¸ñuûƒmA˜¥ŠO—íoòÝÔù(CXà“ŒãÏŸÇur2PëëGØ©×_zûë˜8ß¶¯Á÷ê
¢NëÆ#K›^¹qxÛýòÚÞ|bÑk¾&“…H¨Ò¬C;dÚPˆI5*ÏœGª°É¨ÉÑÛ:…kûö¶Æ„OPü\K 5"H
ˆ¥ý4 „LiâÉ€0	Ë[d¥„ÅëÎÇöWáVa_[Ü*è¤÷–dý0ãš¬+«ín’uCÌ{#ì+ü)Ù—^`_g;„f3ÁÉŠ:¼¬…¹ùŸœm·¨ã\bÝàÓ°VnPaieÖƒàÖ}ý½Í’A²º&£E¨bµ#…Ó2ïÀ1¤	Šš•"¥qz…TI"R%1¥J"¤Jòé¥J3ÝO)XúÖ-M‡,ãiDÀüH+ÖÒÍ°$EP–|bí>·HÓTò††XKÕ›ñ(©ðc¼YYÌYIî¬>{:pø”ºß€s'6?³â‰éxÌs–æ³yM{ÀÊreµ{àiÊ³qE»Vú<„f³¹ÔQÂnBcmPæÁ†r°7|[o'ÕÖ]Ú»¾µÜ¥{£¤7Bz»ÓL~Ûˆðm§
Eg˜Ñ–fÆMc˜Å-'ÍÖ[©Úã¶€`„G-“”I=?ySsÕBÁCê àQ;ó‹Þ…9
ï*²#ª[M7gá7,°åöÍ¶¬–7Wçâ†ã'ÁJ›]!Î0òIHÜÿHD´°w®±·€pÇ°)êLí"!®È¨¡½ÐiûLp—Š—J96£zFþïOÅ\žåú²oƒ/ÝM½Íøú@d´BüxMŽdÎT‡®Pr¯
øÒô(~áÇê+ß´•ÅˆÓ^2ñ-Æ€+ŒÜt[”Ž8K§H#øa m$•Û
ÑÂô6R¸­û‡(†LÇ¢U–œÆ^±û™%w÷­Þm'ìî„„‰¿åÒoæmŒt'A Rå?ÁAÜ"¬Z9hªVÂð2v>ìm<Ã	·m&£
¨Ò?2*î8Ô»ìè Äy´±p5IOÅ¡œ`z–—áºÀáÆ–¢½ ð¼aó';õB½œ;Þ.ðÝ˜¿õèU”q¥Á¤X«/¦d¦·2¡Ë:“ÀÒo¥©ßè"rK)/­ãUòç³¤œ +sSÜ-.XTs±ÏDŸð¹5á…ù{Ë‡ü+aV®a{µ²¤ã3FEÖÕô’Ú–W²ÎèÄÁ9oT«¸Ô^]ý`Ÿ±¦HÐþ
™yöLmP‡Ì=çs¿»¯[ M3`t–=-‹iÄ@ë…µŒÜm¨4‘Îrw‚H-³ùçÔš®ª¦óº~¬zN÷E7Z‰ê­â&êeÍ¿À§êÏWwžœ¥¾þù—ÿxòø¿€6ŽI‚	IÂ¡  *(Ö5î9(:¨D±›XÞ6æ`%ñô™6Æ¸æ®X ê°pÛ3ÙåYþº¼ìUý-k:{ø¦îÝ.”°ÆO¢6•¿Ç%}1LYwÙÌüÐÛ¥o'ñBM0ÙPWs·Adjqp·¿í `R£NÐ8•ìTV™¢ªì‹²ÁÐ«­Ý40nö53^^akW3wðebµM#¬bc“k»ò¹
.t*n7™j©qÂ«Zçªˆyš¤t¶ïACØ]4EÆÄãH€.ˆi¶Œ=£t©—x‚“<6mqSâä“°b™ 'ÂJPJ^êû½‡’4Ãíû÷éÐBãEîi‚ýe?fYôµ.û½ˆßnÞìæO¯ð¨?¡OýQ½êó«ƒ;ßû¾uÔ»¾¶½ÂÃ^écwó²×ó³×ñ´}_;ôöën¯épQ.wÈéºÝ­Ž·ãY|Ïûú¾wŽ¨óÎ;!ßô*Nx«Þêˆcþ­ ?:ô#æt§Ä]Î˜7ÝÑŸ¾¢G}ã>u›W:³-äY[¾µ'¦W;ÛnaqÜÛV§¹9>vW…ûqâ®-¹h
;”éVâ¡ì•%IÕOõÞÑ0QÄ&Åðàö9mCN9Rejáñ^@ôæ†ƒú3Nþ=¤ÅÄÈRBWþ.3~ýj4’å–5t&ŒßÁ»þ«(Ÿ$£I/‰iQÇ=‘Íã£Ûó½Öéê\rª³·Fûä¶×¨NÜ@§ò÷L´–ly8›Wp_×Ð!'ÉÜÚ\—ý®Øv“êOúv'¿UG¹Û%S”‹«ÄŠSA‘H$z³™EßçrÏ"Kñ¨ Ñ‰ÕOTüñ<É“3Þ$tƒÅ¡ãúc›J¬ÏÈ­Ÿ\íö;ømQuDcÎÊ¸+,\*h×Í1Y×5YÏ9ñ¿×hñQ„°À+Ï-^ïƒ£X²yi"600|©Ä<è ~lp›½«oý%6pXÇuG—ã×d·Ã8äy½Ã+ÚÀ.†,Êô,Í“¬Ùí3æ+ÅÆÁ¥ë‰ÁÂÔoës{ç5mäu]Ë8Zä†3œ’V\«“E#„)«}Ü£-„ÐÍˆ9·1VzÝX¢=Bi4Ô‡‹°w£ÛÌ;Ü;pe-Y!*äPØæqZÍ²DmÌX]Áõë•Ö×ïïóâ<—¡ºÖ÷ÇŽ²ƒímfêz;„úv=Š(6»SëïkÿöyÏäùòŒ‹µ‡Žò±Oâ±¶X‹Í”·Ï”³ˆþr“¾ÄŽw½Ò/qè6.ãAøx&L9 nþIKzåG
.ù‰=“5Òø:°<lVÓ“ëûÐˆ¤Oúñƒ~mÄ´ëî˜DÑÜÈƒµ'ŽAíQ *“lL‹hÎM†…ª{;ÒvÖuÖ9íªÂ,~<ø	¯Ï9Ïƒ¡«S€¯‘H‰ÜŒþs:MkI¶ë>¯Ì'ôø›˜%ìúaAýßXƒŽW;êƒH¿¨Œ²Ä6ü®•é2
bÚ0Nmv¹°xôëÆ¶/p‡¼ž¦½ª–55¬q°âÑ§Ò´t¸‡z{ØÌ^CwUÃÉ}¥êø†”±T¢ï0z"K¶H¸|»Ðˆ‰Šö7F'³É„¤`«W*í.[öÜ±•¦ïûÝßu‰²ÉU¨"È$Wåb¸©¤<¾mÓ(sË¶W×ªÆEK!é4u}˜ºÊC‚•>}„;ñçµN`êí–ì}BØr_Ÿ­¬BÚº¥¾6:Ð=X‘u¬¸ ÖñâÎMëðkï 5û$–àóf?ËÕ'ÁxD®kÅŒá7«"¶úÍMï”â~Œe<Ržfhwd&Í?tqÚ·Õ>®ßÓ=[·±ö	îR;€T¿mýÑiëqq
t?âóÓdžÕ nc*:3b¹.•t"wë‰0h–6Ë#í]njå[´§º›}±‰Œ,2—‡ô#š{él²û‚¶¬’(£•SÅ¸\-TÚ|&PåIvmÀÀWñ	xP¼„iª¯—Eúº9	éª±Ñ/Œ´âÏ¥è¸×'9(TêN±ÚúFŸö½8=%í%/õ"ÇŸòèQº~f4÷¿!åê¼ì¨±¡¬·6÷í×[ÍÍi¨`ö`Ï¦7-ð‰gF;"J,œŠMf€Iw\ÝïKë'£zžd¯•©³uÕ¥µ‘õÞÕ¬=ÜT<Ú—ö¤<¾·t–‡o©ÙëR'Ú†À{2ØmºÃÊx™ŽžYmä G¡%í³Œž9¬1š(Þ«Œqv¡Ûó¤žG<ÍzQ·CQ–íƒ×ûà ÆCÖæÁApA˜«SÁ_!LëjiQi­WCç=¡D€½šžA`BïöÙwÁWßf°ÁýpÇ
B˜@5Msë%Áñ·˜ûâÎ$"DE†ªÄ¿øcZ¥'WhÚ“€ÆsdiÅ’Ä<’‹Þî³×{‡pƒíaMVk˜uƒ«-c¸ÛÞ‹lŠÕH£Óƒ`'é®Æ¦Úäv¾-ßH>f‚z¸„]v5Ö}¸2’cÀÕíÛê}LQ€‘6ÃØvrë÷IÕtŸ~—ÅykŒ‹Ì™{lZÃÿfø?6©7ZÍpkHÂ´ˆf^8§šÚJÒ²^Ž'Å9QèËzUîÙ ›A3ˆ
c¡ï·¢“âÝ¥8]ÕÅ'sÝ™‹kÃ:™5Y„¸Ýwuƒ¦_ÎæÌØŠ‘Éòp *Br?ã“FªãK«L	!Æ÷-)îzá~ý@$ÞÔ­t ²uÕj¥™—Ô˜Aö+_Ü‰€å˜.QƒfFš{çX7pêh^íó:m8È‹\ÝÂ@E õErÞ±+ÄÛ+ô»;ËÃÝXÆÁi|goyxg¯cã=y¯ëÈ`Ùâ<:6ß lÂÿÂÍƒÛÌØqmÔ[÷¶4ôî¡P»+>ÆôØf×vŠ¤£`MÔñL“­·ÛwJfÆi…ù¢ñÁÂTI´·¬Mn¦÷ŠEÛ.ÃÞŽy¿˜%#0º­{ð²ª(øê†1ÏÑÌ.“¼Jî(ÝÁy±)NzOñ`d´¶¹ùØÁÔ£	ÿP‹þÌOkÛÊ¿Ç& Çí¯1‚5VëÖ5`¾r0¿,ù‡´˜W!}u@[Nº¶@$sx:üÖF!ÜC…0¥:XFÆ–zó3‚ºçÎ‹i9ÂœûYX?DÉ Û{ù¤f‹w`ÑÎoîÞÁ|ßdðCKŠN^:ü"ÊÅ)û6 ³Â#Xé Œìd£“WÜSþ¿ºñ{¸QM’qqóÚ	®¸G·v@A¾C…6]Õ÷®}Ë žBÄf`’ˆ2r˜"úÆ(ðZÒäö:Ò¤™ý)V^àá½ÝDÊ+:øÓïjÿ0þ9é®&¯êÐFtúž ÃÄÌ£¤÷bŠàkË|^
Ö4xÄ¸úÜ‹ÄýÜw­QÕ£‰XaÏ}öŠƒ;3¤ÿ¿(Æ\ÕôRC©[*½Ôêkjsy6¸«p8¸ÈIç$áu(5l·ÝeBòïÄœÎÅ;ð‰rñHƒL¾]Ð²–0XÃûÁÎ;ã«ª…€Á(c^Æ‚{<eÅ |ô É/· ¼úí  %Î7Z ½.|Üå}b³®E¹pFQË|6ãå(©8õè=z5Ñb]Ù$5Wæý³<”ÿZ>©õÚÅ;QRc6ïÑ°Sãì¨î$ZBbú¼š$QÖø£½Ä{î%j-‡z$zA Ê6ïü?   ÿÿ NGK|