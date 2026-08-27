import React, { useState, useMemo, useEffect, useRef, useCallback } from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid, LineChart, Line } from 'recharts';
import { IndianRupee, Wallet, TrendingDown, Landmark, Activity, FileText, Map, MapPin, Plus, Trash2, Download, LogOut, User, UserCheck, Shield, FileBarChart, Filter, Search, Menu, Table, Pencil, Edit2, Home, ChevronUp, ChevronDown, TreePine, Check, X, Unlock, RefreshCcw, RefreshCw, Save, Eye, EyeOff, ShieldCheck, Lock, TrendingUp, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, Printer, CornerUpLeft, Calendar, PieChart as PieChartIcon, Maximize2, Minimize2, Bell, MoveHorizontal, PlusCircle, Users, Send, History, Building2, DollarSign, AlertTriangle, CheckCircle, CheckCircle2, ArrowRight, Clock, ArrowUpRight, QrCode, Smartphone, Copy, ExternalLink, Share2, Scan, Undo2, Loader2, Inbox, Globe, Laptop, Wifi, WifiOff } from 'lucide-react';
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
  if (typeof window === 'undefined' || !window.navigator) return 'Web Client';
  const ua = window.navigator.userAgent;
  let os = 'Windows/PC';
  if (/windows phone/i.test(ua)) os = 'Windows Phone';
  else if (/win(dows )?nt 10\.0/i.test(ua)) os = 'Windows 10/11';
  else if (/win(dows )?nt 6\.3/i.test(ua)) os = 'Windows 8.1';
  else if (/win(dows )?nt 6\.1/i.test(ua)) os = 'Windows 7';
  else if (/android/i.test(ua)) os = 'Android';
  else if (/ipad/i.test(ua)) os = 'iPad';
  else if (/iphone|ipod/i.test(ua)) os = 'iPhone';
  else if (/macintosh|mac os x/i.test(ua)) os = 'macOS';
  else if (/linux/i.test(ua)) os = 'Linux';

  let browser = 'Browser';
  if (/edg/i.test(ua)) browser = 'Edge';
  else if (/chrome|crios/i.test(ua) && !/opr|opera/i.test(ua)) browser = 'Chrome';
  else if (/safari/i.test(ua) && !/chrome/i.test(ua)) browser = 'Safari';
  else if (/firefox|fxios/i.test(ua)) browser = 'Firefox';
  else if (/opr|opera/i.test(ua)) browser = 'Opera';

  const isMobile = /mobile|tablet|android|iphone|ipad/i.test(ua);
  return `${os} (${browser}) â€¢ ${isMobile ? 'Mobile' : 'Desktop'}`;
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
  const [clientIpAddress, setClientIpAddress] = useState<string>('Detecting...');
  const [isNetworkOnline, setIsNetworkOnline] = useState<boolean>(typeof navigator !== 'undefined' ? navigator.onLine : true);

  useEffect(() => {
    const handleOnline = () => setIsNetworkOnline(true);
    const handleOffline = () => setIsNetworkOnline(false);
    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    let isMounted = true;
    const fetchClientIp = async () => {
      try {
        const res = await fetch('https://api.ipify.org?format=json', { cache: 'no-store' });
        if (res.ok) {
          const data = await res.json();
          if (data.ip && isMounted) {
            setClientIpAddress(data.ip);
            return;
          }
        }
      } catch (_) {}
      try {
        const res2 = await fetch('https://ipapi.co/json/', { cache: 'no-store' });
        if (res2.ok) {
          const data2 = await res2.json();
          if (data2.ip && isMounted) {
            setClientIpAddress(data2.ip);
            return;
          }
        }
      } catch (_) {}
      if (isMounted) {
        setClientIpAddress('Direct Connection');
      }
    };

    fetchClientIp();
    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      isMounted = false;
    };
  }, []);

  const logAuditAction = async (action: string, details?: string) => {
    try {
      const userRangeNameForLog = userRole && ['Sarahan', 'Narag', 'Habban', 'Division', 'Rajgarh'].includes(userRole) ? userRole : (userRangeName || 'All Ranges');
      const logData = {
        action,
        details: details || '',
        userName: user?.displayName || user?.email?.split('@')[0] || 'User',
        userEmail: user?.email || '',
        userRole: userRole || 'User',
        userRange: userRangeNameForLog,
        ipAddress: clientIpAddress !== 'Detecting...' ? clientIpAddress : (typeof window !== 'undefined' ? window.location.hostname : 'Direct IP'),
        deviceInfo: getDeviceInfo(),
        timestamp: Date.now()
      };
      await addDoc(collection(db, 'auditLogs'), sanitizeFirestoreDoc(logData));
    } catch (err) {
      console.warn("Could not log audit action:", err);
    }
  };

  useEffect(() => {
    if (!user) return;
    if (userRole && !['admin', 'deo', 'DA', 'approver'].includes(userRole)) {
      setAuditLogs([]);
      return;
    }
    const q = query(collection(db, 'auditLogs'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const logs = snapshot.docs.map(d => ({ id: d.id, ...d.data() }) as any);
      logs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
      setAuditLogs(logs.slice(0, 100));
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
      { name: 'Audit Log', icon: <History className="w-3 h-3 sm:w-4 sm:h-4" /> }
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
      { name: 'Audit Log', icon: <History className="w-3 h-3 sm:w-4 sm:h-4" /> }
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
    const filteredLogs = auditLogs.filter(log => {
      if (!auditSearchTerm.trim()) return true;
      const q = auditSearchTerm.toLowerCase();
      return (
        (log.action && log.action.toLowerCase().includes(q)) ||
        (log.userName && log.userName.toLowerCase().includes(q)) ||
        (log.userEmail && log.userEmail.toLowerCase().includes(q)) ||
        (log.userRole && log.userRole.toLowerCase().includes(q)) ||
        (log.userRange && log.userRange.toLowerCase().includes(q)) ||
        (log.ipAddress && log.ipAddress.toLowerCase().includes(q)) ||
        (log.deviceInfo && log.deviceInfo.toLowerCase().includes(q)) ||
        (log.details && log.details.toLowerCase().includes(q))
      );
    });

    return (
      <div className="bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b pb-4">
          <div>
            <h3 className="text-lg font-bold text-gray-900 flex items-center gap-2">
              <History className="h-5 w-5 text-emerald-600" /> System Audit & Security Log
            </h3>
            <p className="text-xs text-gray-500 mt-0.5">
              Real-time activity and security monitor recording user actions, IP addresses, and device logins.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs bg-emerald-100 text-emerald-800 font-bold px-3 py-1 rounded-full border border-emerald-200">
              {filteredLogs.length} Records
            </span>
          </div>
        </div>

        {/* Live Client IP & Device Banner for Admin/DEO */}
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-2 text-slate-700">
            <Globe className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>
              <strong>Current Connection IP:</strong> <code className="bg-white px-2 py-0.5 rounded border border-slate-300 font-mono text-emerald-700 font-bold">{clientIpAddress}</code>
            </span>
            <span className="text-slate-400 hidden sm:inline">â€¢</span>
            <span>
              <strong>Device & Browser:</strong> <span className="text-slate-600 font-medium">{getDeviceInfo()}</span>
            </span>
          </div>
          <div className="flex items-center gap-1.5 text-emerald-700 text-[11px] font-bold">
            <ShieldCheck className="w-4 h-4 text-emerald-600" />
            <span>IP Logging Enabled for Admin & DEO</span>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
          <div className="relative w-full sm:w-96">
            <Search className="w-4 h-4 absolute left-3 top-2.5 text-gray-400" />
            <input
              type="text"
              placeholder="Search by User, IP Address, Device, Action, Details..."
              value={auditSearchTerm}
              onChange={(e) => setAuditSearchTerm(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 text-xs border border-gray-300 rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
            />
          </div>
        </div>

        <div className="overflow-x-auto border border-gray-200 rounded-lg">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200">
                <th className="p-3 w-10 text-center">#</th>
                <th className="p-3 w-36">Date & Time</th>
                <th className="p-3 w-44">User & Account</th>
                <th className="p-3 w-32">Role & Range</th>
                <th className="p-3 w-48">IP Address & Device</th>
                <th className="p-3 w-36">Action</th>
                <th className="p-3">Details</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-200">
              {filteredLogs.length === 0 ? (
                <tr>
                  <td colSpan={7} className="p-8 text-center text-gray-400">
                    <p className="font-medium text-sm">No audit log entries found matching your search.</p>
                  </td>
                </tr>
              ) : (
                filteredLogs.map((log, index) => {
                  const dateStr = log.timestamp ? new Date(log.timestamp).toLocaleString('en-IN', {
                    dateStyle: 'short',
                    timeStyle: 'medium'
                  }) : 'N/A';

                  let badgeColor = 'bg-gray-100 text-gray-800 border-gray-200';
                  const actLower = (log.action || '').toLowerCase();
                  if (actLower.includes('add') || actLower.includes('create')) {
                    badgeColor = 'bg-emerald-50 text-emerald-800 border-emerald-200';
                  } else if (actLower.includes('status') || actLower.includes('update') || actLower.includes('edit')) {
                    badgeColor = 'bg-blue-50 text-blue-800 border-blue-200';
                  } else if (actLower.includes('delete') || actLower.includes('remove') || actLower.includes('lock')) {
                    badgeColor = 'bg-red-50 text-red-800 border-red-200';
                  } else if (actLower.includes('sync') || actLower.includes('memo')) {
                    badgeColor = 'bg-purple-50 text-purple-800 border-purple-200';
                  }

                  return (
                    <tr key={log.id || index} className="hover:bg-gray-50 transition-colors">
                      <td className="p-3 text-center text-gray-400 font-mono text-[11px]">{index + 1}</td>
                      <td className="p-3 text-gray-600 font-mono text-[11px] whitespace-nowrap">{dateStr}</td>
                      <td className="p-3">
                        <div className="font-bold text-gray-900">{log.userName || 'User'}</div>
                        <div className="text-[10px] text-gray-500 font-mono">
                          {log.userEmail || '-'}
                        </div>
                      </td>
                      <td className="p-3">
                        <span className="inline-block px-2 py-0.5 bg-gray-100 text-gray-800 rounded text-[10px] font-bold border border-gray-200">
                          {log.userRole || 'User'}
                        </span>
                        {log.userRange && (
                          <div className="text-[10px] text-emerald-700 font-semibold mt-0.5">
                            {log.userRange}
                          </div>
                        )}
                      </td>
                      <td className="p-3">
                        <div className="flex items-center gap-1.5">
                          <Globe className="w-3 h-3 text-emerald-600 shrink-0" />
                          <span className="font-mono font-bold text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded text-[11px] border border-slate-200">
                            {log.ipAddress || 'Direct IP'}
                          </span>
                        </div>
                        {log.deviceInfo && (
                          <div className="text-[10px] text-slate-500 mt-0.5 font-medium flex items-center gap-1">
                            <Laptop className="w-2.5 h-2.5 text-slate-400 shrink-0" />
                            <span>{log.deviceInfo}</span>
                          </div>
                        )}
                      </td>
                      <td className="p-3">
                        <span className={`inline-block px-2.5 py-1 rounded-md text-[11px] font-bold border ${badgeColor}`}>
                          {log.action}
                        </span>
                      </td>
                      <td className="p-3 text-gray-700 font-sans">{log.details || '-'}</td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const renderUserManagement = () => (
    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
      <h3 className="text-lg font-semibold mb-4 border-b pb-2 flex items-center gap-2">
        <Shield className="h-5 w-5 text-emerald-600" /> User Access Management
      </h3>

      {/* Memo Module & Sync Feature Controls */}
      <div className="mb-8 p-5 bg-gradient-to-r from-emerald-50 to-teal-50 rounded-xl border border-emerald-200 space-y-4">
        <div className="flex items-center justify-between border-b border-emerald-200 pb-3">
          <div>
            <h4 className="text-base font-bold text-emerald-950 flex items-center gap-2">
              <RefreshCw className="w-5 h-5 text-emerald-700" />
              Memo Module & Expenditure Sync Admin Controls
            </h4>
            <p className="text-xs text-emerald-800 mt-0.5">
              Control Memo for Fund sync options and manage lock permissions for individual or all ranges.
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Option 1: Sync Option Toggle */}
          <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm space-y-3">
            <div className="flex items-center justify-between">
              <span className="font-bold text-sm text-gray-800">Memo Sync to Expenditure Option</span>
              <button
                type="button"
                onClick={() => handleToggleFeatureLock('MemoSync', 'global')}
                className={`px-3 py-1 rounded-full text-xs font-extrabold uppercase transition-all ${
                  isMemoSyncEnabled ? 'bg-green-100 text-green-800 border border-green-300' : 'bg-red-100 text-red-800 border border-red-300'
                }`}
              >
                {isMemoSyncEnabled ? 'âœ“ Enabled' : 'âœ• Disabled'}
              </button>
            </div>
            <p className="text-xs text-gray-500">
              When enabled, users can import payee, amount, and scheme details from Memo for Fund directly into Expenditure entries.
            </p>
            <div className="pt-2 border-t">
              <button
                type="button"
                onClick={handleClearAllMemoSyncs}
                className="px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" /> Clear All Synced Memo References
              </button>
            </div>
          </div>

          {/* Option 2: Lock Memo Module for Ranges */}
          <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm space-y-3">
            <div className="flex items-center justify-between">
              <span className="font-bold text-sm text-gray-800">Lock Memo Module for Ranges</span>
              <button
                type="button"
                onClick={() => handleToggleFeatureLock('Memo', 'all')}
                className={`px-3 py-1 rounded-full text-xs font-extrabold uppercase transition-all ${
                  featureLocks.some(l => l.feature === 'Memo' && l.target === 'all' && l.isLocked)
                    ? 'bg-red-100 text-red-800 border border-red-300'
                    : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                }`}
              >
                {featureLocks.some(l => l.feature === 'Memo' && l.target === 'all' && l.isLocked) ? 'ðŸ”’ Locked for All' : 'ðŸ”“ Allowed for All'}
              </button>
            </div>
            <p className="text-xs text-gray-500">
              Toggle Memo creation and editing lock by specific range name:
            </p>
            <div className="flex flex-wrap gap-2 pt-1">
              {['Sarahan', 'Narag', 'Habban', 'Rajgarh', 'Division'].map(rangeName => {
                const isRangeLocked = featureLocks.some(l => l.feature === 'Memo' && l.target === rangeName && l.isLocked);
                return (
                  <button
                    key={rangeName}
                    type="button"
                    onClick={() => handleToggleFeatureLock('Memo', rangeName)}
                    className={`px-2.5 py-1 rounded-lg text-xs font-bold border transition-all flex items-center gap-1 ${
                      isRangeLocked
                        ? 'bg-red-50 text-red-700 border-red-300'
                        : 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'
                    }`}
                  >
                    {isRangeLocked ? <Lock className="w-3 h-3 text-red-600" /> : <Unlock className="w-3 h-3 text-emerald-600" />}
                    <span>{rangeName}</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </div>
      
      <div className="mb-8 p-4 bg-gray-50 rounded-lg border border-gray-200">
        <h4 className="text-md font-medium mb-3">Create New User</h4>
        <form onSubmit={handleCreateNewUser} className="flex flex-col md:flex-row gap-3">
          <input 
            type="text" 
            placeholder="User ID (e.g. 12345678) or Email" 
            value={newUserEmail}
            onChange={(e) => setNewUserEmail(e.target.value)}
            className="p-2 border rounded flex-1"
            required
          />
          <input 
            type="password" 
            placeholder="Password" 
            value={newUserPassword}
            onChange={(e) => setNewUserPassword(e.target.value)}
            className="p-2 border rounded flex-1"
            required
            minLength={6}
          />
          <select 
            value={newUserRole}
            onChange={(e) => setNewUserRole(e.target.value as any)}
            className="p-2 border rounded"
          >
            <option value="admin">Admin</option>
            <option value="deo">DEO</option>
            <option value="approver">DA</option>
            <option value="Sarahan">Sarahan</option>
            <option value="Narag">Narag</option>
            <option value="Habban">Habban</option>
            <option value="Division">Division</option>
            <option value="Rajgarh">Rajgarh</option>
          </select>
          <button type="submit" className="bg-emerald-600 text-white px-4 py-2 rounded hover:bg-emerald-700">
            Create User
          </button>
        </form>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-gray-50 text-gray-600 text-sm">
              <th className="p-3 border-b">User ID</th>
              <th className="p-3 border-b">Email</th>
              <th className="p-3 border-b">Password</th>
              <th className="p-3 border-b">Role</th>
              <th className="p-3 border-b">Max Sessions</th>
              <th className="p-3 border-b">Active</th>
              <th className="p-3 border-b text-right">Actions</th>
            </tr>
          </thead>
          <tbody>
            {users.map(u => (
              <tr key={u.id} className="border-b hover:bg-gray-50">
                <td className="p-3 font-mono text-xs text-gray-500">{u.id}</td>
                <td className="p-3">{u.email}</td>
                <td className="p-3">
                  <div className="flex items-center gap-2">
                    {editingPasswordId === u.id ? (
                      <div className="flex items-center gap-1">
                        <input 
                          type="text" 
                          value={newPasswordInput} 
                          onChange={(e) => setNewPasswordInput(e.target.value)}
                          className="p-1 border rounded text-xs w-24"
                          placeholder="New Pwd"
                        />
                        <button 
                          onClick={() => handleUpdatePassword(u.id)}
                          className="bg-emerald-600 text-white px-2 py-1 rounded text-[10px]"
                        >
                          Update
                        </button>
                        <button 
                          onClick={() => setEditingPasswordId(null)}
                          className="bg-gray-200 text-gray-600 px-2 py-1 rounded text-[10px]"
                        >
                          Cancel
                        </button>
                      </div>
                    ) : (
                      <>
                        <span className="font-mono text-sm">
                          {visiblePasswords[u.id] ? (u.password || 'Not Set') : '********'}
                        </span>
                        {(userRole === 'admin' || user?.email?.toLowerCase() === 'admin@rajgarhforest.app' || user?.email?.toLowerCase() === 'sharmaanuj860@gmail.com') && (
                          <button 
                            onClick={() => setVisiblePasswords(prev => ({ ...prev, [u.id]: !prev[u.id] }))}
                            className="text-gray-400 hover:text-gray-600 p-1 rounded-full hover:bg-gray-100 transition-colors"
                            title={visiblePasswords[u.id] ? "Hide Password" : "Show Password"}
                          >
                            {visiblePasswords[u.id] ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
                          </button>
                        )}
                      </>
                    )}
                  </div>
                </td>
                <td className="p-3">
                  <select 
                    value={u.role} 
                    onChange={(e) => handleUserRoleChange(u.id, e.target.value as 'admin' | 'deo' | 'approver' | 'Sarahan' | 'Narag' | 'Habban' | 'Division' | 'Rajgarh')}
                    className="p-1 border rounded text-sm"
                  >
                    <option value="admin">Admin</option>
                    <option value="deo">DEO</option>
                    <option value="approver">DA</option>
                    <option value="Sarahan">Sarahan</option>
                    <option value="Narag">Narag</option>
                    <option value="Habban">Habban</option>
                    <option value="Division">Division</option>
                    <option value="Rajgarh">Rajgarh</option>
                  </select>
                </td>
                <td className="p-3">
                  <input 
                    type="number" 
                    min="1" 
                    max="999999"
                    value={u.maxSessions || 999999}
                    onChange={(e) => handleUpdateMaxSessions(u.id, parseInt(e.target.value))}
                    className="w-16 p-1 border rounded text-xs"
                  />
                </td>
                <td className="p-3">
                  <div className="flex items-center gap-2">
                    <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${u.activeSessions?.length ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-500'}`}>
                      {u.activeSessions?.length || 0} Active
                    </span>
                    {u.activeSessions?.length ? (
                      <button 
    xœì}ëvÛF’ðÿyŠŽÖ3$'$uw2²$ù–x7±},åËÌz}bl‰H@‚€–4þÙ·Ú×Ù'ùºª»¾ ZrbÎÄ"¾VUWUWWWbâéÓ(þvtÓî£c2¦£ˆ>hœÒ4ãiÚž÷ÃQgñ'Geü£ M_z´‘Ñ«¬—ÐQok‹Œã49È}ÃÍzÛÞv²0‹XØ7	¢ˆ¤b ¾*ÇÞ¦ßÒó„¦ã§—êè.{»dÜÛÝ ›¾š‡›ƒy–ÅS÷û9 Óy¹@q¸9
?ÚÕ7³‘ãi6RÇ5cãâ`
/Æ9èùužfáùuNGä"˜õv6\c:äÃ%Îá˜½ñÌ9¥ÙóQ˜…Ó‹7lD—q2z9â(ä¯ñŠ^æ¥§³yÆ*ÌÄoò¯‘VËSy±pÕ¤¡A4§&á3 "ü•NÈ€õFñ‡¿Þ"»êíÙuo›$ñ|:¢#ý¸1|J3Â¦FäÜœˆö“ˆD‡sM½¥n²a/:	Â¨³°¦~‘×æÔñ™êøÚ=uß<Yá€ÈsÆŠó¬Cv|ògñÅEDJiršÙœó”.aÿ¦ÏÂ4Déæ<ˆRêa5œn>`°¦M²$˜¦Œ¨ãioGq’’7ZgIkp!XÉyÔCCSø¹“cA”ÝÞÚj1& Ué„&ATT—¿•&ä#­ùšZ|pMÕ2køò;èù”w³¦)öh†H"`YVïžÉáY¤ã3ï1Î¼çáÌþ	¸Ø+{–èÏ:E±÷ƒxt]a xòàäòû¡Žt’õ¶wÃãå8Ì(›åCIs½+&±ÆÁ(¾ô,P†q"‡ã]œÑ9§Œ²é$ÄÑˆL0’¿‘Ù€‘9JÖõ$íé4c}Øâáð‡xø›Úú¸·O.Ù*¡°ÉÊVcB	T`€<eÝ'q¤Àf¼«Œyfù*%:×bC~¸qŒHgtž‡CrÎ{IÙì’â)ë‡=bOØZ½ iŸü<¦S±šX#»”ƒé4ÎH0bO(“R](?Bº$	2ÈðF³qÉnú‡›³6ÓY0¤½k¥
2£‚þþA&£üžÄ—î=`£ŒÄä	¼ÔÀpéfÒÆªpuìX:‡Q0 ‘Zp "	H7H3:¶É|6£É0HéÆñ)ƒÝ0#oä]ò–¡  
küp»°:Nykí}˜À=ºá¯é0~$4³yã.cè±Šü…	žS«^›ö3üÒÇ¦b@ãçLcß‘`–¬?_›BH˜ÒßÁEâH
1£ã^8ñ‘‘^ïp“r×¾`ÏÂïh€šºÕ5½£7ŽŸ=íoÛªÌf	ðZVïd‰j§A00ôó/KT}Åj\lãŸ%ª}Ð!ÿ»DÅgáÇ´~6Cñm‰Êoƒ_/‚d¼q,¾”`mS¢­F‘A9Qz#˜×$˜µ í¶[æ‰qþF¯n&LrýàcögÊèzQ6×ŽcE,5XåžY ·ÅIÌkNb®NbÎ'!”Úi³¯ÀÞæ&t¸ÉYÁ>í}W%{å“s÷Ë$˜ÙÌÖ§ßŽ„ÎutóUM¾çS……à…ÚíÖIÄxy `ju‰Ýp)'¼ùÀ´ß]Ô~ûû…þ¼ïÝööìê½*,­Ø­M@K\Sjs¾ß(dÍû5œ'iœô˜Àî±í{|ÉuÓób¦iÿ<œŽÚ #ê‹äèèˆ¨  ù{ËE¾t@ä1S†@MAÑäÊqþ¯iîßØš»¯¸|ÿSC·	ùfý3´ºËÞ.ÃÑþE=î€þ4ÊÙTô	F®T&ü!¢ç£´kŽÍ§òßñB}~5c*_þ°+U…Áït©ÞjŠw»Vo‡¥¶µZÕ–ýËõŽètÿaW&N~ÝKr6Ofõ–¯^«ÍìnWâŠØ€jhÒxÁ¨ÐZ„ÚÛû½O¯§Ã?ôRD ¬{9f4ˆ¼eñå–âj3»ûå¸"F _yærÌßÜùRtîãiš>SØû“#nUå@Â>lÜ}@pà…ç¤]4Øñ*zÎ|ì>ºÊèœGA®1,e¯jt­ï±‡Cš¦n†äìÅz¶ø¬ùV{Ô?6‘QoŸÈA¼ÜëÜƒ]Á}†Íáé8¤Ñ¨‚oòBOÇtYæyŸ§^ïˆÓm*°ì‚ú±Ÿ»1þˆ-<9ƒ|ýˆN/²19&[0e	.,+–E-Ý	ë†TÇaâžm|ïY§léDa8,–<åŸ³ð#…#ý,	‡ÀZÒÃÍñžÕ¶1 X1çŒEô®zÁ<‹¶áC<uœ¼ðKzžÉó-ÆÊ¢`gL.;}6¦Ã†¿K´s­â<M›¤<êòœ*cºWÍN~tºqÌ	ìp3¯T]ˆ’•ës	¸rõŸf£€­òäúÖMœ¬
ÅE‰ÓœÌøZ²áóç>B84ŽåÕ©ã¹W*_DþãNixxáá…Jtr’¹Ñh°„êÞ\;|±Nè(œOÔÃWoÅ9rQ<HJ$#–L“qG8#.ÔÖiKž
¶iž »Ì+ùýöÜ¾iþYmßx8;ëõðùû ‹å{)Òa:¦³~9…oÃ7Š#ÔV¡Œq¶•³ÑÒ~>üS}öòôäÉÏŸñmÀë§ÿÁ¾ú¡ŒWlÔ%0XBú	?beÎ÷“ë¥Ám56¥—äk¬·z’uúYÌYÑS&^¦íÎŠýîQ¬2×2ù©i†(P¤$˜^wë¬
ø˜âWU8W'EÏTèª¤å•IŽë>Hp|‡\Ape^¬¼„{>Öí>¯ÅW)EçÊ
cž:Æ\ÎŽGÙ2RÜsüÞ_þ(ùÜºù×ÉhNoÍ·æ»b+&çbó2.ëéÕæIëÞr3Ëa6KœŽQêçn˜çP÷†)YNžÊÕÙ³xl¹‘X”Ž´wÊükçì4¹.¡ly%oé,N²”Õ`ÇKŒâKÆÕƒÑ›g/à=^„8 )Ri—ö-fŒžƒcïÞ¿gÚ4c¾²4¼ÑŒq.:ÒÊË‡VyÍz(‡3dÃ€%ôkÊ†ÔnEŒ$Óa0£…[?/y~„Æ
Ÿ_Õõš;7m/þ¡(¯ÅCÙë°ŸÒì#ÃÓðŸ´½½×q¼j·Æ4úH³p´º¤ôÚÒÊ]·?<¸Aø-H¼øypÃÇ¸øÐ%Û{ð_^Eü‰˜Í†Dÿ`óØ~øHšÀÔ©‚^ÛÚÆGkìÛŠ9±îè•ñoœ¾~NND¿ät>™Éõ¼æ×ä[¥&lÜÏ€¦Û¬®f…­ÖygÐÊû®RVÉFbêÛ”©KÙ?ŠŽ·wÔ×l/7a„ÚºHÂQK«7Ž/¿ÇÎ[”qœë7ÁÕ
ÌØƒ'	~c%`zíì:¢ŒDo1TÈ7]2¤Qô&Ù°.ºÆ\O‹Ša=Û*›þ>ÞþûgkŸ- ²|³³ÏÞÉØ;ìÚ8R¬Àä9¹ È…rÙé3¾›Hd<î‹BK
õ„2Q´»/{X¨4¨.gB1ÑÔ1Ùþf_·Ëi1°Úš©» ú½âñb}”ýLÌ„{÷.Ã,VÀMú~¸}¼ÍAß*Ëé{ë}Ëgˆâà#ÍY*êO—4y¤Œ’ú	EÁ¶7ÿ+ýz“É©Ö/­Î¢?M-/Õ…ÜttA!ê~wÒçt4C~ÚV\ßþB4×A‰€Ínn’'ó0tˆMµÉGØÆwøâ€-#¸q(U†2™z—Ç×…‘Æ:*dqDÅ árãd@“.?‰ªY3ß,Ðã^iÕymÐl†ó„);™òPn
á j¨ŸSò†Òá˜ULj‰)ÞHÄš}^à¥z'QÔ§@-ìß,NÊêcG}²Gðoø1ÌB9„ šô&D™k× Ã|pb4ãI^ÌÕœÑú-¼|ïªÏÅŠHÁ£!A8Ã“qBY9éWÏYùÒNŠe¢Ýà4:šCGúQ‹­Â§Ú_%ÈÒÅ[9\yÃ±2*W)82)k
‘ÈñÐß–ªhJ"cß½ÉrU#+ŠÃü£+JV´‰è„Ö¼ZÍá»Â2¡àt‚¯tLuXeÆ¦ül]/ÊÙÉ)žqF“‰®BÈµoQ0¤›5téðÈ®ÍÙÐ§ãs4;0u7}Œ"dÞ¬”sñœ÷ÙFô°ÉWžÄqDƒi§ÿkNÛ-Òêpâµ»Ï"ëÿ]NŠÝ‚”º!ts¾÷utLô+Î²º+ïÓœátÍG<"rˆÂÀµÊ^«ÔÍIf•Ê‚ËP¶ñø¡¦Fê&Ù¥I
à€ñ"W2
)ç%ýó8yÇ.¹a,|n:Âä„ì¹‚°TdâÚaå‚QßÁvQ¥%ÖpÓèàó†·J…Ø@ñ“õGÑ^Y„í”dÊ°"•ùµ½a>&X+©ª4ãì¼•¼l÷16+<C¼i)ð¶×¼Œ…oQÆ ‡·Ô~‡P¯kk¥þF¡¶ÝÈû[ôÜ°…È/Ç}°Ö¨ª¿cµß›H²
€.[LGUŠE/µx+×ˆAöMÁêªo4D¾nLâù4S¸dÇÁ…"{†”ì^Í\J´à’nNjÓ#k¥äE4BB¨â;›÷Õá¦Ð´ÉOýÜ´Œ—Öâ¤ÍòÑF¸h#«y}ky­ü³6÷lˆwþŽ`]ÂÆj¬Ï1%ßË.m†l¯?›§ã¶ÎWáHí W;|ëjïhaéàE8ËÓ¥xªÅßóï¨¸cÍéEKå±Žé«¬Rp¥$£#>t×ƒ_ªyT›v§EÛí KNi…–áéàÓx2ÚopWô…°°Ê‰çbTÀÙëÎŽœƒ:Ðœ­·Ñ¾Ã¤ÌèÊpæ¬åí“í‡[5m»Jçp¦Q ÐkÞý[eû.+y&íŽmÅé23beó[ë¢OÜàÔ%ä|Rô{a“S*ë{—õû¬hß¸û>‘LTáþ6…Qüdã>„s\˜²Û:;nÑiïå«VÌv;ßn£ã[…¬é·4
¶FÌªÂÿÒQpÇ†=õ9ßÈÒ`å]ë´ÿ
.»µà þ*ªFûmÚGªÕQ>ãÊðíIÓ¡xóþ½-êV_€ûý>‡r…÷æBÊÜä	þ›f8atIzÅËùÒef‚Õ“Ì§Sh9Ä#âÄ‹®KàËç #‘-np ÓPÉ»Ö¶7Ê•ÉµÖk‰”FôÉxk¿ÏÇö^¾‚†‚c ¸ZúÒbrÆaú–þŠÆn6o(*9-ZñÎÚ¿~UÔ3Õi)=Ñ¬Âñõ}CŽ¯ËKªz4à‰‹œwZ[VbÊŒìtt‘‚ u1!ï§³(d,¬×‚‹4A+—­›-ô éµôv¬
"ÍIø	Áš þcõ«þø‡rë
ŒæÍóWÏ^¾úÎ?„š´¤T¯JÐŽµ0Îã8³ñÙë³“n·”ª’ªÖNSË¦äÜPûå<[UÚ{5Õß ­»:#ìï¯~J¸³ÇÄŒüÇ8'2È}BhNÞ×<œ2ŠVm~GóÉ4ï@[[ÐãXJV~KªÅÁñs8ÊÆ{zk„lC¥Äî¾YbGk•Ëk­Ñ½oÍ*»ËWÙ«ûîŽYg¿º›ò]bîÚêEƒþÛ[ä@Ñ[Èî–ÃªQ(óSÒ_à”ô~úË‹ü"B‹“åw›ÿõ_ÿü×ƒm¿_âŒùùì-S¸ yçU—qòÛ ŽçÚ8º?íÿ,š»øSÓõHì¢EÁÓ1…û yg°/€RxÞÞ°\Š\Š×æ&9ÈÀËèaø6¾¨ª=B7ìqûÛqI•=²‰>-â€fð\€íŠ‘ÇÁ"V|?´ž&4¹ OÙ
HÛÛLÅÇÿhruŠKiBEÏã8	ÿÉ†DÅŠÌ‰I…h×¡ƒÓ¤Æ°”ñäUû”)r0›vVµ­Ã±‡.HéH^Š1b(¥ë^×3`³ cs›ú„xÊÞ§qd.Ææ/ëf[™äbÀJ½xñ|þg0k$"Ÿ5–x¥‚UeãpÚ2Ü%«,4ˆ³,žTCvé(å»ÆÞaÙYdˆkQªàÒ©v3µ3(!VOÛ©‹âUäPàz ®	äÚ`®hƒL>e ÏKïÄ›PR'os\ðß&È`K8	®~àÌ÷„*® Jõ"Ž6ŸOfÙµ\j\â:p˜æmäí#‚ÐæÄ¤fñƒ©­Òá\2rÆç¶ôzÔæŽ‹¡Û×ìÕY©µô6°ç“ç >Rš;$Û;lðìŸå)Ûk9,tœ”ˆ‚¨ðä®¦ðsÈ ÑþY!‚´¾
TâaX¢¼ºrHïÌ%…eÀ”BEK	!§°ø^‘@N`C*Æò}=TGüø„O‰è)<õÄŽÙ¿“#VsÃœ°¬ä€ùxngÍË×+bLãÄ¼.`jŠ—Ï[z;ì&:Ðêˆ’Ar1Ò°YF„Ô nß÷ÕE‡uh3˜ŸŸ#¹—AXì’úWQzÕ¿LÂŒ>ÁÅþHÔ‹âØJ=a_ÛïxCïÁô.XO0›E!·Nm~œŽúñŒN¯&Cò$ÈÒ^|~)ÛÎ÷ÓYÂ˜"Ê§IÔÇ¿ªÕö¬'iºí’¥œ¡a&å;ÕÿgÅ>Õ±uü'¾‡©þû)+km¡o÷•]Làé© `âk½|†' ¯ŸÃ’h Ç:|ƒsü†g¥øV¾âÓâð¼õþ‘Ý°/·#0xøáÑnag„CÞ®òëVÜ¤Û•€oÄAbÑ~‰[mîR«ø¿Â½@upæ!f™Ërqìm°Ìc¹8§öÇ<4Ç³ê 8îVÌýƒ¨á3aøKm‰¦Þw´ð3Üß8ý{ž…QÚâà—,þ…«ïTÒì’~¿ŸÓ³±§éG½5lÚ¦Û¢¿¼"[5 —ÐöF  +¹ÑÍ›Ó„ÅNŸHg­oz5³—O~Þ–/›g4&áLDpV”xÐDðÓX6¬}ÑäN?°¨¾b(^¥ò”ƒÖ=âhA%AÁP«.¾P|ð·íKˆ~öáßÜ¨ø™‰ól:åx/ ŠXÓw4SŽqìÇo*¡Ç‘ÍÛÑP½Û'`©*ÒÆëÇ›Ïž¿&ñ4º–ç& LáZ;fHàA6 \KÆ *žhìxš‡ä°vŒo©4ƒáT¶Ï‚Z ÜË|tA‘ÀÞÒ!áQ8%goÿ¡!_f³'A‡&-Çq.”Æ¾¨& ´T§4Æ8ƒíóÙ+ô\ÑžŠÙŽø€—‰ûËéYrmV”“0»ùXíˆ©Å0Ý"9ÓûŽcªåTXà ©P€ÆÙP9bOJµ‚Ù›_RNd‚yc¾í3û7ãC®AKtJžÐÚR»Ôƒ[A-0ìòÓaD¿€úÿËùõ/nŠ‹O‹>kÛ­–ã	Ó…Æl	1Œñ³ ™Ósœò\×WT/K ·ÏÑ|Þ¹½£¿)ç1jÚÒ™¿AQz¡ÆºL#þ€Æ_óVÇ¡Zð8¸•ÉjµzõÝ&Û£ëUÞ©RïS*ù£
Iô ì±rNá'?üðúççÏ~ah=U®zÁÆ
ðNQµ´;.iá7§"¥å¥Úi>+hîu¸ÿ]àxaÔ, iÕu½2ûU1iwî|+êƒèÍÙ ÍÖ<[êœÆç¶b =Ñøc;ípòx<þD‚_àu²åP†îÒÎÄ¹èÌ]~ø²-Üˆ9>Q©ÐÅ¾@Hˆ¨ ÒUá+ÍOÄ˜"U¦X(æ´r¦-¯ýàØ{Ds5Ä5‘B^2"<È9Š¡Éð?~ÝvÈæ‹[g>%ü@¹NæÕ´¯”{beš4\¸¼å,¨?”þ$Š¯Å[Ó·RþP€#~P|-Þ¸+…òT"(Íï•!ÖÂ”‰‘áµSBNt4AQâb¦‚j#4SrxïöÜîFèíîƒ6à‘ÜÔ]ÐÚÂË¸ú'…€êÜZív”‚~ÉW–e38ÀœÕÄRˆ¾Ê8ãÈZáž¨¿žžâŒ´Öå\y«¹{Þ‚0ºÁ@\ïRp9N‡l]À½7à³-f¿ø ÛBUèä 3|©¡â„ŸþŽ;józÃ„Ú’ŒWàdõòôunk{Ë³VçÝÖ{/Éè¾§År5¸-N¹ÇÅkr]$”%¥…jI-ùž/ø-Ã.¹°»ª²=g 9e\DsZ6g&\éåžsäõž·gë4î½f¤O¹™sBç~“XM’6/#™¥qç‹Ì÷e7êÓf}ê\†>—£Ðš4º•Ö¡Óš÷5DË´8ƒíZ¸Q	Q¤žÌ¹[B´úùµïNˆü«ïhá”Ž´Í,ø²{¶º…-ÒEëB¥áMÏ§áÏ)7™ƒ<I’àºžÄ“6°×S¦éÛ½™õ cïrÑ·¾ãèŒ+6+uÆã",ÑY¡Â¬Ò_:a™é©zÓJ“TÂ+,Óoìë£Ü>Æw¹Ë´ÌSGþÀTjwûfúÈd…¸Iå¨”añht'.ç õ„åÑƒ?ðÝBUð=VL±tÈ1rÏ‰·ˆcÁû3Ž°*b!-ì ëâºX­®¤ñ{.ˆuÅâçŒT·ZU……/Â¾w°b™ß·#Sà¿‰@2*E4KÆ ”†BÊHð†”©c\ ‚Ñ`Žy«TÇM ÂÒ£/k¶ïÓI¢p0ïšüas“œÓ\“Dça’BjH'Î$JO¾`Û£oô
¬Á~s `JÕjÅ ‹ÒÖ=Ð¼ž­ÈžTTòÎ4ô9ûSù„uõT©íêU¶ƒ”m|ÁîóöéýøcïÙ3ÂÝ":&zV¼ÑçvÇÁ›—"<â›Ž6sÉ€7
‹²/Â«ò4ˆ†óÆŒZ¦";ÿH¶8£k˜œÚ±'L!…Y#Ž½J²Í‚0a¥‡ºW¡Pu
ûÔËé‹B`9$¶´’ò9É¯GhÆŽ¯áœ”ý«¼çÖ”ŽËö«•J·eüF¯MÇ=.æÑ?,ka>j/=ZäÐÏ#“@¸	…öcþÇn™é'…Mx¡™Y”É?×ŒÖçÎ©&ŠÙ&Q/áùŒéùÕQ¾=«S•RßáÚmÂ%®4×ðNæšÀS„Ro< Ø”ñÛ†QA.Â€´é¨kŒ~ÆÀ+:Ë—ÐAµ4
h¼ï‘íù«£¡.QJ¹
˜Ø+'×Ò"÷#ªá=ÔÓ~s?$¯÷„qƒY½²œÅ„i¹¿™ÞÜY€Í^¨ÅñvÂ­ì/Š»n¾ëæ;¥®²‡éê[ücYÍ•"Qè¾kÒhÒ;h1ˆ÷óà£ÐÌ<=é,ÒÞµ‘-È™Û³šãˆLFø–%û®eûužfáùuo@³KÊD"d339[·r”íñ·í4FFõ„2^9Šxypì]ö&#_î3W.¢ò&Ý‰‡„r¡%0Ú#cö_0HãhÎäø®övÁÍµ·½¹ÃÓ³H$=ü
2ˆÁ_©‰Ö £–«Çp:›gŽÜ/šGewZGÂq±v´!.Îe˜ “*É³ NN‹¨Ó÷ûî¦Ñ‰ôèÆÜ·¹CÉÇÓ§chíè¦Ímå7¬Óì­Q·ME6î¡ÚyTÂxÛì‰#g|ÔhúQïod–0|Ì®ólC2éPžáNæ(ˆ.H‘k8Oây…SÚ›ÆS*åS=ä?d¬û}ÖO_å’›Ž°ðøØÖßLžäÍV§ÌÈ•¸Î‘$0£ø’MÝ5xôI•DcXni8)ÛÀß	#/„ŒdŒbNÇñå[ußÐþ*5U¤Œõ°"S2ædå #În¬þEŽÂ‚Ž$ª”e¦Ué
ì„"gE­Ì„‡¼ÿÚ(ÂDbÌ¾´ÎÙ>ÓI<ýiæé’Š2ÏâË©§T½Ü|0PtF²ÉCcSª²p’’—K©UÛ¯0¬±É®êò+g:4=Ã'##‰gsé9£¶ §º½µ8ÞÞ"3Ö&8Üä¯+ëíì/Žwö—¯·ÏúÛ_¡¿Þöâøÿ…lÄÔ3_%Frˆ×ŠDÝÖ,¤·Y‘•(I4Úª®z›ã›i’p‡þË5Æ®ÔmK™Í†;U©7ùë•7¹«•ÞÔfE5ˆé™˜µ‘õr~â²eÐ¨»&—„3Þb_3¤=²óVÐ.²@áRæ5r”ÂÇÁ°­Ì¢XÛÌ.:è=d»ñpúkQŸGø—mñKÌ=ÐzÝI¥Ìö Â	 V)¨ð£ƒâç.‰.”ŸÅî þS’vIœvëzž„J‡NML¼‹‚4Ú$i2ýò–‘~YÍU¤g„¼ªl›ÙÃMlÑÛŸO|É&Æ4Ãª?wS™8“šÑNj­våi°.ÝºùÉw«ÕÍÏ¶á{qzoÔói|‹:#0½ÀÊ8CŸc}*ÚK0ôÕWýácg“¹UÿÝBõÏ¸OÒÊOI‚)M¢ml3ÞÅo
§~1(?7æy^q*(ÛÅvéBb?]³ÿò–½i»ÜR³xë%üOµ$¸Yg-k‚Sió‹BR¿µ*¾¬D–k%G†ÚKá4wÑ,¼ÄM/ õƒ±åœgMŽøð®(æð Hã	•çç¥çZ 9õø0íø{\tþ@kÆzÖ86½†5.Ö²¹Æ¿¬kc]sŸ¡úë:wž¾Ýº†µæ<ÅotÁ—Ló¥þÎ–É‘9¤âL¼œs”Sq]gÓót¤x0Tù/”÷% ¢LŠZíšó¸Â­!-koñ…"”Lb-,P²¤æ™`Áì,UçÏø
WÂÚ¼O»µÁ3—5Îþ\/„«Ñ¾x¾˜_ b ôõUx~Uù}Ýž«#A–¬ù†_˜r­TqXnZ0É5(¨*¶Øó+¸\†#ŸWé¸Nú	¹²ãEáø…a&;¿ß‰ˆóô¦9óÖqåm\HèäŠ¼áü›~‘u$Çëçëñ:¬•±eªüÌyýëçKpx3ÔÄóuWÕKû×ÿÂõ«æWÐ‹o’Æ=ú·<Ö/u¬ÑÃ³æ®Üö)‘8Ÿ‡<RYgtq0Œ£¸Nõrgé,K§#ÏÙt©G ÿNeËÞ!2V<ÌGÃeÆŸV«EÄéøÙjÕ«"Å–Ÿj‰3)¤ð¯t¾áþù£oÙ#Ô&tÎ'^¯âååÛß=ŽÞo) ÈóÒ«Û	G¾õÐ«ÿÅ’¾Ò9c—ènÒ?cPøgô®ä—Á®–w—ér®áfvåñAÁæ`1ZôàÖåT¾UUÂc+í›Ï9²¬Gô9vöö€âÿþçoŒ‹&fÂ¢Î¢¬ß%¸T=ˆZ”ãÝö¢n9hË…gAZñkÔÂGN³ÒÇÃ¹X˜¼sS\yiØIÑACPôùÌ»wª³›	}àðçQ|Ù»êA~,ó~
ú:¶=8M¸Ó!Éžñ¼(˜±­¡{3d_]ÉÀ?Ñ1…,1ü±	ßj:Ì4ßzØUídÁ'Ñ³c+6¸·ó7¸ÖAœôc¸†u¸™ïrõWSÙ=$ïBB¯+·7<Êcˆ›ÆIo‡Ø®¢(.¤+ß>ð‰û’Eð»Æwýù‚ûƒ¾Kþ‚û‚|'ÿ%÷%êaâ´Ü´¼~þKh{Å…ó»ŠØ—ñQ{Þ{4(e‡xFeD0pŒ=MìÛuÎ}Éa)Š÷<µ0<þz_vIˆ+Ä¾?Ä78h€*TÌ««KÛƒ²Q}»²7Å&Àö¡™c;vÛ¦ù}œõ´-î5¬©qîP½žÆsWÅ5]q¸i¸Õþ+{‹×…%îGãó(¸…fÒv›g¥ÇõÎc¢ƒaÑrÝ§—nñ‚ü/÷ \ì˜Žã8ÊàºjÔh÷­Mäº%€ˆ£ÓY0=ºÙÞ^èPù–ÏNh[º7„t_ÅZz HÂ 5ŸŽúKÎÜ¾èï”Bnhû`q˜A‚y­@«ÇÚVŠß.nGZ–q¶-ÃruÖHÐ+6¬wDKÛä×;œå,ÛË;{
„é°^¥À"]~¡ÛÚã ö¶aI(1ÆŽÉ¶½H–Þu‘IfvòÈVÉE6ŽaAõnjÅ#_Æ£˜Z{ƒÑÏ&áT­æŽ>fDGcû±øœÜO„Í-A§Q×y„7˜„3†?wu ÞÞA*‡ØO+¸‚¼¿3„ƒÓa¦@#é:–lÛv•Öè~;wÌ4¯«ÞÀeëñ,2-Šaªf¸¾å8¹eÌŸ%ÎÈ^	¤xC8ZE”< “ý®BúØþEnTÜ®çâ—¼šOÜîà(¦,§Ã#²ß‘5È	ZÝ51H>TWÐÇªïÞ¢úñ‘º´{dGmK{³ÇK.ªiKrÇ_ÍùÐŠ8§ü«„p“(†RêY²¢xm¯Gáé’©;c©ò0Fe,“æR“`Òã6aDŒSñ‡"d“µæ¡šä§Ä¹±<å NŸ™…XåÊnÁæØª,øÝ×Kó»¢þ½`|oQUh>Ú™õH’öÎÈ¸·¹é‹qËÀ ñ™À2
Y
I0Lâ4…·"HÙL07œC~@²q“•.tÄ!cG—¥£'×§à
ö– STˆºdŠ!²Ž!õ÷B±w¦G(²7©¼ÙÑÅ;ðr¥ª]òœuiE&H#‹†u€S™”TdîTßh¹[±VÞâ~Ên.hys!Ë›X¾¶på·1îN ¡Ãs:Ä*é¼÷EÑº9ÊD nýòº|k‡äæåé-q”‰–ôëžò­Ý/§·äU¾B|òCŒ¾`¼{V€X„ä•«÷{3ck¨s;<àÕAïg¨N•fóâ"tø†˜NE‚9GŠ
&!2F&\(¨Áá¸©3'özËz«§Þ*Âö²ÑäüÐðæÇ{
m9{Õ¡uô0ïê«w‰dš"Ïè
ìnŒûë£"ú°ä¡y1uXPË¾(ZvÆåÎR¾@6~7%mkÓÑ£t”(ÝWdî'T…¶ÚKÏ˜”?“¤1yG<ï<Œ|\-¯õfŸ¬JÏ¨BÍñ¨LÑ“ºÑ¯#£Ä‹®4 QŠDáÜ	ç-#BeÕ©ï‘{øßËE]ÒB±ò•Bîq°—Õ3
ßÛBNxj«¢DÁ2WÏW 8Ý]9XƒŽ«qr6Sž(Ð¥¡óÞ‘€Õ$`(ZFÄE_9Ç—?‡Ù4@àIx+Æ¯·¿È%0670žv>cÊ\ø™_)fIðBZéhÔUõ½Ž*gS@´cðó¶›È´"YD–9€oQ¿Þ°ü‰Ü€2èÁ=–ÅåÒD˜bá¼òÕeú‘²-¥=óEé
’…ïr>f\ˆÐ°àdâAKPŽ“¥òIWBÎ°›µ“Êùùñ‡ú¡²|ôÏ³<xHß%~<@nˆ¹T¼€ã§úV‹#è"Ošá¢KÆ«@ÌèMa\Àõ‘QB
ä‚¼ŠÄUy’ÙŽ&š”‹-îŠÜË'ìm&î¯æXÏ•…ª¬.šÃÞW4×™¤8çì˜Yô‚³Å¡›€e{÷WjÓg#¨šÿ“Ï[»GG;VâÝfþ”Ë¶îï×g<WÎØæÔžË¶î£Öf|ÁÆ6òÎ¹à“Uš¹í¢.&õ4ÊS°ùç‡=Ó[äö×^¯‡¶è“A
~9O&Ar›¯1§V¯§ØnQTH.6
Y¹Ä,Ëº:™Í¢kòÓKaßM5—$Œ¬cy}4i}õµ¶ºÖÓbVXßXÃ^aBUö›_éÏMŒçŒ8©VRA0¿'Þ±2Òñ¬§Y¢Ó/îÎë<§„¹Cš@µŠ¼©¬3wÖT HuÕ¼^‘¿Î0@D,žbONöôÒ²£™’~
I³»@òo)füþÅÑƒð
ÓŒN‡×ywÂ(Uóg1¤ÞÕkdÑ±²h
M£oþ H‚h–q˜?ïèa¹ü¿W²Å}/[„4nH‘ñ9àNDæzÂ(˜TŽiÂ)9K®•Äpt?Œ/|ÄX’0÷92Ç½•`kiI¼åPÏ”Ìq*÷²³Ç%pD"8˜fÓFCÄŸÅË)\W5ÑvÑÄ«Z
áAr-&àH§[ßëÓï¢xD?¥˜ŸyÎþ¼#‘H:MÂi©?ÑØñTÌ"i©`Ñ,#:R1z›¼x°)1¼-BÊI£;.[â>4§±(Ý%•‚¥å*²v”£*¹¨‚©Ek2=êïòÅ®"šºŠ`éjB¡kòò®äaùfÐ†ô•ïYTë¤bŒWKäDéy¯"L@’¢y’õœgD<íøYpõLÆá§´žlãòŒŽ½0þŒhJ‹0QÔ>l¡ý4²yÊi8¡p\ÁØ‚)³Ô*_Q#¾"[
æ#líU<¥Þ–1‡°ÔR§<ûjÀYn8Â6©–|µcßyÄ­­	„i‡Ÿõ"ÒÔ‹^°Ac¬—Ç<¹<›fëÕæIë‘Ù¢´¤yp_‹X ºEqt.¼”,´‰Þ^"(×jÎhoÅ”áM¥¿]ªðÂ©à™·-‚UÊÒ.+Œ¤|öì´…~¥NBÇ·¾Uð$ÎÆ­Ž¡PÕìó—ïNÏÊûå%*úV:®“X\ŒÉ{ÄéJfû$Œ"ò*®Ìkû&¸¦øèÍÉ+øÃ‡?è©s2A†M‹´·ÿÜ‘ßYÒÞ¿Q–qnÈÕ—W4#o‚pT;§­Á›á?¼•@áÿˆê¨>à¢ ek5-4fÊíXMÌ @äˆÿ Ù¬¾Óº|§`šò\½ÊÖã>üû*Î™…Z ‚ß<ö1µ*¶úØËWµV*Xéc//íj @H½ïñGEñ)Ïz¨×žµzF‹4³jhÏ¬Á’®«ÏÐsù Ùâ(=ÁÜÉ«¬ÜwiV«èH¬–Z…§Ti´-‡KzÄÓÂ]Þ;8F“9š—¸'}8Þ«Ô'¿˜ªÞO	®H®l‰l|‡›ã=£ý²LŒDû¬”ˆÑ5ŠO‘}1Oñú»OÀØ<„ëg]\
ÊŸUâÅŠXTfBá«HIÏ¼Lœ*ˆ½«G©ò„£rÝŠÛwÝ¸ö\„»ØGÍaì»ÎoÄ£{ÿ¸ôªÜ7ãw*oÆC°®cÖ£ïš¾Ëñ¼‰+òš:Õ¦eWNšÑ®Ìön¥g™­¦q™­LÑßoùéj²ôžIn &¶BÄÒ0†ö•Bib7´*ÊÀÍ—Ð-Mµ×j¦¸†¾‚’Û`÷«éÇM`%ÕºÁThåMõ$H*žÆ*ýW«ùkDCCƒPã¨À%_É•–¹l¼ZÊã‡òŠqÙvç~ŒÈØX}‚A9bB¸Gç!Ÿ~|Ö¥zeŒL .5"÷Ýqç]>×M>óÖsA#¤Î'ïQ× ‘"2¶±ÈMqL’ÃuŠt½á"L¬)b­"ÜQ ÒåC(L,Ÿ´v‚*÷(r[Ë’k¤´sÿ¨J©‰çþÏ´,5?BÃ-9@§=«ù!ZLnÉa.oI[në	ºáüi\Á5¶'>“ÞC…‘˜edæ¦?ÌzV¯Ye,ô[ƒDG‰i—À±ðqOãË$˜‘t˜ÄQ4’Þ8™Ö÷=q3Y…’b>8íVÊ=ÀZõ’©7ùgWxàZ’©´eÊW¦](Ÿ>« àè¥cFM¿õ¶ò;ÿ|Hü˜Hïý[·ûõÕ ¯û‹rØ38µL3#˜’rRš9¬ˆ &•J9ƒb&©QØzÊ)ý}Ã‹g”wŽ¦ÂM„üE½Hž¡›GºFÄEttA“û†(1ª;GŒ¸Æðs˜RòŽ©*nÚõÜžž…Ã]ƒ\Ê¥/tF5J'Aš¡!wš1–[µŸ¹Æø>9’]VýqLä)S-ë½©¬¸ Ùö¦—ooîÛ:ÔwçËQ;cª\‰Õ§)¨Ø p@kÂïà7Ç¾—Æ3ã‰‹@õ±Jw­ƒÒè‚ã$¥“÷üî3*ÇQç‹0¢O‚äé8H2«Y-ðôæ1©PDžª±I„÷¢<§#2á¯‡	Û|‘\%·Z¦’<Þ­qì‹H@ÑÈ«îeíàÛçºý&?bÂÔFlámçúÑ”÷ˆÝ}.™+Ô¾5k…9®>«ôÆ3çcÁpý=SÏ» ¼ÖÂ›AÌY[9úiæëô /©žR¶¡Æ}V]FònƒsãW–5¯Ì« lIuIM5 Âû¹Ã j;‡à¡zó=H¿jÞ¾åeÜ5ý—=qÓê¸J4ºx¼®Êy¸c¤[ÆòÎeDº*‰	ÇŠ/DV‹ÈüÞ"Z©÷ÈZˆÍåOÒ¹IBûÏpVÂÒºNðÊà¹Ÿ¶ÿùòÍR=´BôY´Y²›uK$?ª,DqÕ‹qŸŠ@†÷˜ç:©u~ëCR„˜;Ðõ¨
¾äº–Ó÷zxotÉ×³ñúÀ7yãŸ*7˜ð÷ëçX$·È~(z8üw‰6Þ»übŠáŒ¸c7‡ŒØ7‰ðÞ™Hí-o™·Y»y~ìüšrã³«ç”Î)÷ÕÔ»gÆm3#ÿ½7{¯¦éÛxKXJgÄ‘æ‡ènÙ“øóXe÷T³ù²">åŠúã§^_1åMå±a¥°c›]ä‚i8a$ÔáæÞÿ¦Qˆ_z ½—Å.›“«=+wt¡¤ñÞCÔ<tYÃpËîÑíþYw¬§DÝg+G{ê§Çç×çw6Þwå ¨¼Ú óÝ­mÊË{{R4ç½d\Æ•hÎiÚ‹öYaDµ™æÏ‚|ÆûžY ÷öJ¼VßÒtã†ØFNÙ¼.ÃQ6>ÚØÞÚúóãp®/~ù£ç³õá…ÜâC~€7Ý¼+-Ã€Ï#® ¨cm%ô ,ŠÊ[É¢[·ƒâv.Š€¡þóðy¿(B8ÂÝêQ8Onv÷+JÇó¬(½¿UQzŒFl„'Ó‹ˆ²â¥ÿA¯6pö.+?eh‡ˆè4Šà.at´ño£áîþÞ¾3º·ÎöÖàoßnWÔ9ÜdôUZà,Ž£,œAØ
Æ1Ü˜¾"ËÃõ¢òòáÿþçÜàsÛ§äƒ7ƒc>„Òup¸éXr>ÿ¾ÕòÆ+ìiò±”ð¨ÊËbþî.™ì³ÿ$·Æ‹'B¹E?¿’Ia{eìÍl1°ú+’Ô—õöéf/¹|Óìie”$u¾ðØçluH‡ËÍ‡*ßUîïâ™„úd¯µ –f2)šÉ®®¨<Äÿ)~š•*Ä‡…OIòÀ/
4ÒÔb¦ÎÿVyqißT/‘áFèp[ôô–Òˆm5|,–Ì‰¡]ñ÷‰#¶iÃ{Æ©yj-Ñx `õQ%ŒqÇ aÚ
Ò}Ëˆa­.Ñ®á±
ô;ú‹º¾ücß…ó8oCš/mÏ×å„êÞUc®F•ÌXKÇéá&ç«z3Ÿ†ÿ=§¼ðLÄ(^8‘­â-¦t!qž.ŽÙy»žeÊØÐþS?ßü7Ný"+jÃä/éÜ¢ÿ/4ï yŽÚºDš_òÂ¸;i™1IÄ®qÆ4±â¥¹>y¡É€ia#%KO¨¼ãV=6¶‰šiÅgQÉ#¶ý¾×1.æ×1"Úô:ëÕ\Ç_Ön¾vùÇºk7¿y›µ+ÊX©ÁEí•HêGˆêCŽÌáq•ÊyƒŒzH-O7JHÑªTNe=	x(b V;æ¬ˆ4•ú[[ü¡™›d³·<zÃ®`d–ªòfjbâam¤¨ Y[Ð0ks½ÐSv}áyËð<5¼ ¯§"ÄrUz»ÛrWuÈlµxÁvëÑ%•ßæ5J%×@Ój¥Ê[-ÆûÇá³
þêóÚ"„ÂnÓOÆo/¬|‡_XñýdÅyÆD—§/-~\¤3}x\õ6²—¦_$‚O"¼~Þ¼ ˆ›·Æãà€ß¿~^›ËÇ·6ñ5ÊÛ]uÔ¬Á_8ÿÎïš[A#¾	iRê§ª^·Ü±ÆZ		RÈvÜÑ3`ûmÚ÷W"Ußqõ¸¤•	³fÅYDØ~©d«”mé¦	¿W’%éjÉ9C¤éÏrJ~Ö(²Ê\‘œb+ÇB¥ð²8{vºq,c®¯RbÐóØc¬½ª¡²åP² œèXööPâ³YâjÔè)$Ã÷“«ÖMãN>	…d)¿¾\V“ÕCÕÝ?ûûà ÊþiÍ3J  -¸„Æl(›;ÜÍœÔ¿W‡½kþ*‡Ìw-ë=œÎæYIB2¶t¹‹Q™w!³(Ò1CMŽ6Ä¤”¬DL÷n!ŒM”·$ø'3ÞÒM&å.“.õƒÑB[g@%ì>j«¨÷72Kdt…%Ùá¬lÒ¥ƒåÎ`ew.ì‰Àè½~ùJí&|ëå'³ˆJšð\z•ì±5²WJàe—UøûåßœTSËM	¢ û®™Þl[¢ùkÒéÈíxV‰ÿš·nˆKž¯è­¥fˆôØ2[fnak,§f:¾v«U^ÃZìUôÝ³OE­±§–ë]¾§Æ4Gßêq˜}¡é—ÜdÿÝs?ÅSþ-eÓöqÀËnw¹WÄ²¢p%vä%®¯N”;zîåž»ºc­ÊÞ<Q4–RÌ806áŠS™ÆPæÂë	û­‘L³õaWÊ}|}}Ú™;¬ØŒe7òÛi®ØŠÍú]WÂ[¢3P(gR•}9\È¥láA¹WØ"ÄuÏj$(7Wî5^ÿwÂZø´7gh Ðùýµ‚¹
ˆu°Â·AeÕ¨@þøPGÅO)œ•(X¡DD=žƒbã1aÊV»]÷f×fé­"í$ù+alBÆ¾¯èˆ©‹? EmëÏ­
dWëìÖ¥y$¿©¨ÝC‚›rlB‹uw3Æ`¬¥ÏÁJySivWÃnø5À€?ç„Ó6ƒY·„òÖ„8†©-Ä˜WƒOåM›ÒeÛÌæÁY¶³ø“ñ¤™ðV7›uçMÇˆ;ä¯›VÄŒfBþ,ñÒyyÛ-Åý´¿”èbì‘á<aTÕ›Å!>³’¾ÌŠ«[¤7>íìØ}NÙÓðØsÊç®ÒwIÙÎ¿69y)î:–½•|f'‡Â¶û:²‹š<kÂLô¦!°Ò ·ƒj>¨öÞ¡ß@õo)êw»½¡`“[– ¯o<{“Ä³àUo&þ+™³jÈíƒŠA…102š'<"1X“àl~ŒaDE˜PnyLZ#[$\örëÁe/ÿ]Äß‡£B$ÍŠ{´^ãÊºÍµõMn#¤5ViwÌ-©i|(7º¶c)ë£ßþÒÚÈõ&sÞ¡n;œ§ñ<‹Â)\ŸRñôTôZüPîœÖ0Q–)+UžZ†Jã´çŸöó%Lºë²„Â:ä+¯š5­¢•çW³ŒVÛF±Lœ} rˆlñPèsméjñ±Dç‡üÇ†¥”…U.ÙDXÂÓÁ§ö+#¸æ¥¼Ó)ÓòvE§©±®FP‘]³$«_Y.M›f6MÏñQÙŽß™BR/‘xvx¾L…å="ìŒEšNoBNM|ÿ·CæßÜI9wK’r©8K9¼/ßUþ¶Ž¾¤œ˜¼/9@ÚI|)“túuMÞºÈ-jPóaÇæ‘û[›»®<µåh$v>žMÜãA†<^µA¦Çø²ŸŸî–æDkdö¾FŒAJ[ëÌäE3{!ŒBìsG<6Ór9ãš]8…±>¾„)ÛoŒ^NÏ’ë»ž*ðµÄ¤9,ËN>Ý8öãe?¡'÷c¤³y2‹hÉP“ka1_}œEH#‘á-U¾ã4È€™Àþ¦j?çÚPü¿^²É3¥àùÕN¹cúrt@3y9ÂSëWÜKa0ÝÂ—…?Ï–º±ÉÐHq,lM>²AÕºÉÃ\[9ÇÍO®‘X/=ªÂøÒÔ{óa¬¬ý`S2<Ú!Ú†[Ú‰z#
-ÚÉMµCMÉ)Î­S®¸Ó†]½¨“tSé¹¼c^DI'ù7CwÛóçÞ,EåUŠ©9åÈy¬Ôà#£fPŒ½i9— ™bž¤Êk3Ÿú²q½Ãf½çQ¹T>¥}ÇˆK˜¦ÑÛà¨lH3Ûq›bž‰püÜèÂÍ0ú3øœ»¸uZ±—´aË	ðèK½KHmãH“'ÚöØ±kZ±}#v“J§"iNåØê™û‡.ãI¹•\Çæ’²IîS‘ÍæO™<	¦£–×ÿÞ{Zi´€P¾úq§vê¢gÛñØ5Kú’s_«£9—*ëk65ÃèšI:'Xûë„	Þ.Ž)ü.Së·ŠTÙDl‹HNM+ÙCnŒD$KYCJ½Öb)ªv*+È2!á¿gš@Ê®¯E4Õù)Àˆ‘–±ä7Ê®:A« Ï¼Yä¥áÞ ¨ƒaŠþÒ±Whäò¢©¬ï¢÷ˆ~¤?g³e¶£Ãè¢Š)½È'^ÞØ¸]q×^»/¾ëÕ–¯ØúÛ(fÉÛv!ÝíNp[ÐÄ$
×fZ¶·êöb)‹êÇHúìþ¸í‡7rX• ÷ŽÍÃ|ì¤Ínò »-¾“ÂØ¡«ooÕ¹ÝxW½ó€swÔ{ê®f¯GY}ª¡çvÃ‰ëÑÁÍWè„TûrªR‰ƒ›(V–0gX¬=7i`õÂŒ!•x4cF&”JÊ•ÖŠªÉÝþËNT1þjsmÔæ{&*L™Ú$+¼´›š&*˜Ÿd’šõ­†q­SPÙ-sÿrÐ4Öe¬j•65Í¢fîxOÆ¦àÛ¥Ílh^Ó­jž€0%žð{¾tDÎÅuî*—CØïyPfh+1³Õï(§'ÌÑj(_Í}mnm“…ÂÔæê*œá2gæq6ýó{»à‹8s¤¸YÞK×h›/ÎkgöœÜw±Ò]¹ðºæçwãÍK{©óÎÜyáÓK/|JŒëuíÅ.ÖäÞ‹m{]|ñmSn¾ðYÎÕ·~÷¸û::»w.¿8Æ;tûÅþë¸þÂ§¾û/|néŸUÝ€á³W`øÜ;w`øT¢¹Ò-Õq3…ÏšÝƒásŸ\„áSÛ—>·Yô5Öl—a,WŸ÷Ïu>ŸÈ+WY¹×ä:ÝŠáóé]‹ënÕë}NWcçD—º€T×åë~·cÞO•ë1/Õ¼û1|>CdG…£†+2‡j•;²£\’y/wä–Ì;¿®É·I#îÉ·Å§qQnr„krSnrˆëvUnr¬ëwWnr´ëvY–c­Qìþº.Ãçž¸/Ã§)føÔÁ_3®ÌðYº?—fL«»5×RµfSåâŸÛ¸9ã(ªÁ‹ÝGwgY0–C±ÂõYñ»?óËßB-s›v ­Ù%:ïÃ3‹µ¸FÃ§A÷høxw¡ë>XÕUšë‘~7il»ä<àvîÒØÂ§s™†ÏÜ¦9Œ*BIÝ'÷i?ÖnÏYG(Wg	·jlþÓºVc—ŸÄúSÇöÓ´›5ÂÿórµF0Tôöž:Öž•]¯á³÷kÙòJ.ØðYÒ>»b«£X¯;6|Öî’]ÚIsnÙeÝ4çš]:™Ý³áSê¢ŸZnÚð¹­«6oãÝµ×4€%\¶×5‚únÛkÁR®Ûë‚Â’îÛþa|rnøü±Ü¸ýÀÿºrßÏÉ®Å»þT?g—nœesW¹t®²R¯ìâã«cOüì\½qÐ•¯Š§U¾)©01–ì¼ÝiT–¸m5àð½òì=e¹¢+vBá$¦HŸ ðÆ»:(©°xƒ§¨áü¦”ç#©ÕÊ$H3Ü«O³$Ž–‰çì4¯–çÇØ%eÞµzeo žïÞ,u²-Î,ƒÝêI/èèps¶DOlÄU©.8&žrD /I sm·Óùä vI‚¹Õ—	 ¯!W¤P—`Èx‹Kz†Yf´²ÁŠÊ‘ eA@L·¿–àbEÐ+ëz­À—×Rèc¡FÀ_™ýâ¶°×“^¬üKS€/M½¼§Ü}Bq[¿8°€®ÉßmãwThî¼«pñy·C@ÁAîÍïóAï~rºßYÿÊf³ÁJÅã~HÆ÷cH¹|*Î:ÃÓØ²l§ØfÜa+½ªoáGg¸º¢Q·¦“i“ýr[î]tŒú:––Û»˜sa°]sïMù3N¸MÒáÿ¹Ñ”¥°©1J_9Ë‡8XÕÛ¶"ž@YP›O¯/4ƒéýç3ÍÃPÈvYáMqîä„¹sØô}î­guìSYpÕË3ø*Æ®³àê™|ì6šYÃ0h?•h^<‚Ùa<M3u5¯1œç!
¾Rb°ëÏ1Öõ4uÐ‘G½S¶]|Ä…77Éw`c&…Q1eâ‘ê’Á5ù>w­þ¯Ïƒ[à3á&Cl*= oé-ÛÃ×[—Ü¨~ÚòaîÐ-à™’’ù•œºÜ·éïÞ“Å1ÏÍBŸwÎ©•ŽN”9ß#ƒo¦<ìó‚mœ«ê\Â›a²ŸUã ”ŽÚ)ÚúáIë	áå(·í‹ÚtµQŒ—ÕÆVmpº=’9¨CÙ} zR^Û ÿ\¤†¥s($¯ÕÐª_¢~õ½~ÂO)ÍšRiƒ÷§øHðmðëEŒÉ‹˜QSFòp`Sü8 mYë_ÿ"J¢k­“I1¤"ß3¬•|Ýµ¿JÔdßSÐû¦ŒÇU
¢}øB¤aCtèoK•6$‘M±ïÞ¦d¹òQHÄ¡þ‘%K[DB[
&­Æð]G´Ñy$ÝK€M}¥#'g8çA”R­ddä5W¤äz‚w?Ä—ÈÍò(ÓÙ«§l£ÖVNÙD]Î[€:8½ž£i“ñ­ô1î›ÎqMÄî%rÎïœt88;’k<‰K¦þ¯q8m·ã§H¦fçÅ%”#ò.'¼nA:ÝõÝkï}Ý5ã»ÃY'ïO‡I?œ£ùˆ¦m’0d¥ª„Óò5sY¾ªà t$¿1tÔ¨ŸWÎ‘;5	xdÉœòwI‘¡ÐgzÎó`8¶€M%²è¹î„˜ÓðÚ
ú)°†kƒ=”âQá“^†*Å¨ôÕ\üZs&¬`þ(BÒ û/Z«#€p²H]À¶@¤C~­`˜V‚%•F2™·‘7­>ÆF¥hf¢„7,EµÖòZè¢pì³äµo·’û¿3X×…ðŸŒÕø…Ö?<¸)îG÷@œluñAYm¸U]õ«ü^GŒõNMÿÏûéjVt­¿¶ôÛ-ý½Psß½Wž.œÓ¶FØ7Ú&_
&ls—å,³c0H¦ú¿Š/I0Ò•}®WsÝ±)-'+hëÍ‚›º9®MÎ¬~±ÛP(‘ƒo
>©Ñg)×…fu¾ëåº~ž[ƒã6Éoà¶¬ÿu­þ5òÙš\¶û;p=¸þI[qµ9k¡Ò{ØªÍ²ÝõgótÜV¹ë(È…ÅßT©0)^€ó7µHšÙ<åoùwTÑ±Þô¢UðReÂò·ä‰§p1
ýîRÅÄ²§tÄÇÏ@òzð+Ã2s–j3ïô¡h»tÉ ÃI¬PÐ#4C>'³ ¡íAñ·4D[åÄóŽ¦±°ü¥2=¥óV"ü.B
¿'L¸ƒO2P2˜°rÖMÆ]ŒH©­ånç!ÛQÒ3V±~ˆG¾÷«Q#˜¶”ãEqó3ÕlyúõNqUQåêµEíM~v_N£8q“Þ›g/t¢j·½êíö÷ó wz´0w\°âŽŽï*U!­Ge,."{¢D²îõËÂúÜD¬„gbfˆ…å2H	›ªZÍÀc^¹FÜ0¼
{œ×]°öíÛ¥ö=JÓ$ë‰Hû­Fh7ê2ç¬È	ºäâåèÊqMò„%˜IiéˆVØÇK,uÄ¹‹`€·˜ÒKòŒµÑÿuú4;'lkLzÅËùÒð3Ö~à¹ù\Å)v>Cm|dÕÃÏAé#[ô‹DžkCh<j9N›­…„Ë#»ð¨Þ‰vä=Ñ®Œ­ë<Ü¶8§ïÑQ×w—[sÏäÞÊ	$R¶§fîgÃ˜ÆÉ$ˆdLÌãw¢­BV¿÷_Ã®í”æu°sÎXódÅMœÅ9‰­V<åRI­âP·‚;]q7Ò+ü<œ7‹}Žs;öÉ¹§a‡+ÌNî>k¶¹q|ÚW8³,Óp¥Æšs9'1‚YOûµ}–î@²•STD—n|Iï#ÿIl©çQ¹ßDfËÔÂíòrÊ‰ 
ÿÉõ(wPvÞbb3}›¬K	y´¤·Ë«—n7ç·òð\Q$Å7Ù}ÌÒl—`§kæT°qÜk¸—ëÂ*½8ïš,¥®-Ã2ß¹pv¾>P·iÉº¹QtD®z‚vR+˜aú–þÊ}7ŽP­‘û_<MÄ»Ö#o3h,+Z)X`¨“=Ñ¥²C/»÷–+•_/SÏïsSã»çêú¯Ðµ’*<¸Q@ü˜´b¶3|ER ¼q—ÿ?   ÿÿì}érI’æ«„8µM°‹ ¥âˆdS¢T#ÛÅY5[+“’@’Ì€Äd¢ØšÍ+ìßý5o±¯ÓO²áwÆ™ ¨’ª*­»Dä§‡‡‡Ÿ·t}Jåoó‚|K¶q®
íÝsœ8PÐ~É¿{õdXL;«ÝÕµ^•ÓÁ@óSm¬®±ž.µ%‹³™[Tdq[bÇØC“:—âè–:ÛÁRíÃÅü½;ðFù3O>árÝÝ¤ÿèJßÄáI,¦"°WÑ‹W
=RéÈLð“Ž`U´/]²,^:3Â¸’J†¯ó!ä÷ü7ù>	¯p®µ>Ž¢FÆbï,@TŸcËã¹üù¸/Ô9ì
8=/ËiË³”"–ØT-»„ˆ$ÃËrûf¡ƒ…ïN³‚<u©Ðˆ{{ü7à$FG$áÝÈ‰ô‰¾½N°NÕ‡ã¦±ÍÅgj9ÃîÉv–°öñÿFd©ÂßBÀQu r 0»JÏêÖ›W[´M•X€„T³Ÿãîe·îÃFLc‰i“ÿ4Âeæš7™Ó_¾iV9¥­9.Æ†ªê’®rÿiž™]ÆÔ¨?¹Ü¶%‡ŒWÖù¨pÆ©¯ìq§MÉ=ÛŸl\neÛÑ¹Zk6¾ÛÀÀí•½ÙDërVÑ)¯/ÏÊ¬ôz=#®6<«bR/çÎ§TeU{ &°I÷i¯´	‚ßi9éÖÅ 'oòs:d—tÈä’þïÇò¢“ãìÂLggQ†ÈP¸‰²#V-Üx°{Èßé>—d<[Ð~u™ÃœÎ}??˜LxÌ-|PÔ°T»x4ÂèÔÒ·Í×¸ÙèxVÑÞ£e“ô³þeŽ¶¸úzÜÇÐHºE©ÀÊs¼KeLÚdf;ã«DÙÕ¶6¯#°Í†fÐ°Dýƒâ‰vZ4èœÏ€ŽêTB2tHÈwðtÙý¾ižcV9;âLúh3`ã3õìÊ8;øl©üÜ«M.Èã‚ƒÕ“ý¶Wu8ÚŠnýôÙÖc€UõI¢>dˆ<n³Ÿ«ÚŸ&t“ Äž^këŒ>^½iaô®?Âùq»\›£ìS÷
ìÜºõ€“£;ûå“btAêª¿»²1,/ÊÞd|±B÷ÃéîŠÉp)ë(WL¦¿õˆVÃPX9k‰þˆµ’cÈö÷¼ªÊJ¦k›“¼Ç}ºNYŠ¶zz=Ì{”t&Ã}  ‹Úê?[ïiGž)EŒ§'ÅÙŽþ~ócQOÁ‹—Îyg•gè[ûgrÓ˜Ï'?Òe®¾Ž.ÐÿØûï· 6w°-k›¹/¶°†YîûÐÞµÙ»„D³²w<Ìáì\cBY8mYÖïçuìé¼Sy° m”)s‡Q¢,­†ã`(B¯—A9>™Š©`µ¸EGA7”-ªqf™á#œjŒs¤tU+ºzCæ•¶,­¨¶L1O.±“ŠjùÐX@‹LêÉgó“N¶Vö^’ò|DE¾'X”U&S´ÏøZîDûa•ÿû¬ ý´ðˆ¬£P«}"uå9üQ¾Oph¤iÌ{=rx°µ}ŸÐùÁ ­¿ èØ°×/G6»m‚Tm}#IÉt¨§FlÖÔd·Ž$Åw0×Çô»+
[d¦'üãEg[TÞbÂÅ'-çüÿùßÎÿ}iÓ}Çó½»Ûø_HžúKžU¤óâ—5ŸÎÌí%Á5 ÖÎ¢é„Lø/—ZS –¼ø¢Hîvø(šÓéòÅ/	éRÍ
*}ñK“>]ú7HãÀ#{S>¸±uWÄªsÇ)Tú6¤^)I7IÝ¦ùùõë	óDÇ°¯kÓæËxRâ«ÌPÄ^½ÓÅî‰?žl°—í*íèpJ68úQ¾Ô¼Ä“ƒç
÷Þå¹ãÍkê?%ÂÕHâ¢Rm!·OŸMQî± |tV¹ÄÅú„z‚Ÿè9ŸPa®C,ˆÊ]>W=Êf3v@¿\*hpsÃ]ù´¤õ@¨¡lpoZ£Ž³ìšúLÅ­­þeuÍ×Q£š÷ßÌÕï›¿T,DöVÈk÷>µL||“ÓaT›–,¹õúŽdô²	žÊØå”ˆº÷-êw¥¸pnCnBddX£°ÜlÍ—ZúÏR®ßd`úÚñ'×1| ½ÇwV;;½¿Ýêmn¿3;ÐÌ8e¼7FÊ¿'P|[}¨Ëà˜pOà,¬;KV™=ò˜4ïîà-Mæaûä²È‡ƒÀT$$'Ðåé¶~ô€3<³hƒ©:eþlŽ,! tæ0A=\Ã7k¤¨é]5ÆVJtÅ×dL¢5œJé=z¦Í{„óMÔPÁ…JFxä( ¥D¯Ó£+À$O§òCló*V`C*ˆÙô&¤ám,œ&ÙCÊ~=›Fµ´kg8êÚÑ Õµ¬itMzœv­"‘¤¾£“+”Rƒ£<”´àl¨«€]ðëósè8=ßc;í:(«rš#²yšÇô®žß+ê#º³—Õ‡×cü¼aq{?
º¼ºÓ²[LºÃH&åc‘i94éìG#Ï€XŒ/¼Qˆ	
)D†a£Á‘–K€ï³ÕÜJ†y5=­
*7}v1Ù’¾”A/[_RéóCwÓ
-pŸyøÝ¥ã6Šb,§ñU9ÈIç%´sÌIc>?–õt-Ñ™O¶¤ƒ†îÂÍZíƒY= ¬e!:VU‘×Ð¸Wù¨ÄÃÉ:këºŸÙ:êÖ³ëœ¾™U9*¦SSÐhxMêì#å¨Š§¿ ×%e2ƒücÑ§|„ò"rUÐúÆtW–5(ê¯™8F¿B3ÝMP¢­
Ð7ƒ\Vˆ¡‚ŠºÏ(ôìAI4`ºn5Ï{2‘Ì÷››ñ(·ÍCi$-¬”r‡a'Ø’2R™WçªP[ŒØiÔÆ ¨‰Œ†Žntˆ»'td	†zæÎØÊi£{zžAà‚=h,2­X”×É(.úwp½"T¶ž¯ð>c•éi›¼v²Â x£üäÌ)+Œ¦ßõ­ÜG*³)9ÍÎ:«‡Âèhðœ,F:«ÎžÓFIÜ4;€£äzçeMeâÑsL’—:WÅx © €ßÿk1˜^’=²µ¹ýÀú”¿H7³rH¥÷Îì€'NÎòËìcQV;dµ•åôrU‹ad—!Õ7ÖÍâ¦†Ç0K—`k¼úÒ-^[ƒÞü<ÁlýÓâàQŽ9ìÃ(Ú–M“=n¢\dbK2P`•^ŒmQk­+âPªTª´WÊ<q‚æ	ÆÚ‡i>Â¡Þä@ô¸E÷¬¦U£¨@r¥«yß_e.HJ¦åÃö€UXÄZÂO°/þïMÕ©ÍHÔ¹â—¿v“i Ò WJÕR©wãzË¥Ûöë]ªs«E6Éc…:<MÏ·i©›´…Ð`óöhÙJ&—(äÕé-¤¿sôÛ­¯s{YhYã5*Ä¿¸þd&tuƒîªöä¾hCÊ*ÊqB–›¶ðÆhhÙ%ðc6eú/w¬Ü`0Änð°¸Äÿš¼Ç¸±€ÅVŒ
~,—¥;`µÅä]‘ÎŒ-&AÜØ¦•Yˆk/psòÄ*Î3!Ä0\G%ÈøPE›•¿%Æ´·¤Ìˆf#‹"ˆœNJ`íäxÝØ&Œ?€=)Ã7u<ðR‡°:F_Ï†%Ž5¼×ëùÞ–LTH|>×m7•rb…K·Ë©2©Ôúuª{ýê+œ±]Ãš’b~Ñ)íÁc2¼ ÿ>ÒN^ë
»ÜëÚ}þµQX%2é®‰cÊ%ÁŽ‡ÌuyµÉÙb‡}ËüÖ	B²«¬˜Ò‰ÈÁ=°zÏ „N1})£&¢³zÌ¾çâl Ö3T®Á0_ßs)*ø¤#“HÙb¨¶%²C§³mRŒf³m¦	o:¨¥nÇàÑIÔ8À0:§;Íµm*óLk¢)*Ái®DÇE¬ø.œájµï?t"ò7ìî¯ÅÇ›R'²¨Äà÷mllÔˆÍÑÜV\úN’#\ŽÙçÍ¨vÆèØð‹ðoÄv7EvøDŒ1Ó˜·p€t9?b-žeG‹~9®§tžÐÕ~QËV/	ï-ÎÕZÙì>okWÃ.>·`‰£nëp	âG¸¾ÚÏ˜øÁÞÖ‚ºŠÑf’Á´l1uãV#
Ñ'‚¢ùß<uÀÚŽ¬lß$øDW¯¶xI:qHE
nÓj6†tÜœùö±Ðì°ßL@gfA®HÃCïüQ¦Â}²û—Õµ·›ï¼§Ãß>ÖõIz$©åèõÐÌ”c(ÐyKAÎ4âvÐ.WÝÉ'´øàì?’«á>3àn2¥m­ìùŠôù½X,ŽÛ?an[1Ó6iáŠks+ñŽvp1…54‘·Ñ—_Ï¦¾ã¼±¸œkÉÇÑRl9î„	†ä(ûX\Ø`#> ¨ÇæáH1üÑÎLét^ãQu“üxfÀkWBI|ƒŽCÑ´â‚°}¯Ü«Pó‡9W^åãÙŽ¦ p.+=7Œ#ôì
¥½žäãÎ½ÿËÚ„\Á‘2•ˆFÃ¬­‡½\ž@Lº{D©î‘­zwº¾4HÉ?mó÷ßÌEAz‚l±(6qÂ~qBUÙmÁô´²ü´«)]7[Ú04M:jÆ‹c Î#)à€Èþe1Ðã°Ë,€uàg„ƒ_@écT‹:_*€"u¿Ñ^}ïVh$¸ÿÁeØÞd¡Þ(ñ¨N{1j‰ÓÞ•$™Bo.j»ÓŠxÆ”í×Ùò¾x+cŸ¸œzb¬!nËmÂ	^Ü2$Äq’C4ºô°&Ÿ`ú¦¨ŽœuŽä ‹ú‘'*o»ŽÎuØeÉÞÝtê–~uýrìÁ$°ø2ÏXÊ8…E÷Îè€Gà@Æöò°>lõPÁzu9Ê;øVþ¡’ÈñwP(™â.)é?‡U9¸Rs¶ìOC¼Ç²ÿ‹«Áx¢ä*uÔL+ ÇsÁ¿ÿÞ}ûhó’D¿RþÍ¨€Ý÷ò©W%•FŸj{àk­ß_
–ñcž}Ì}e J'Á ³vmÚfƒÅ˜@­t<bø•9Â• ™‘äüù9o"âÜàöü Ëxv™¬Ê1hšº*<Í:uI:ù+½“Ä™òÚQ¼.÷'MÙÉˆEÇùC{Ï&Ñ%n‡já;-ºPyï‚›&=×n±ƒíßáp¹ùÎè†™â™®uÃÇ~0ž¹ÉAÔ’ÜÑæWr‰Ù—b¯!tŸD‰ˆ]†\¤Ê÷Jò«4ñH¾ž,$É/E%ù¾Í£ŸÜVÆ’¥IZìZŠ¼Å.ÔÅ.sm³X)¡°ÔL}–«x«™ÚÐV 4E0M>ØçUð^8Ði…m<Ü´¹®—¥Âå[‘pñ•â­Úkáæiþ¬œ~…ŸÇhi‰hÆoJ'As ~]à5c8{›SÔ0ç³ÌòžŽŽ<ï4¾<*§Å974Õ«<µœþÂØxaM•o|	œH«ÆèÄ	®!8-'×ÄèEOøÚšX~ë>Û`—-—§Mÿ¼øQ%­o±$—ò¸ibÐ­	šß;‹ a›ÖðÂcˆ0…!&|J×÷’Ýðß¿—´ÅØ‰¡oŒg ¬ƒ?Mâê{}8'³A¦P­ÑÞI1šsæ8Xé“Ùøfu](Î¯kã÷Û9ÝAéÌÁÚ¦ob¤ìN³ÒÁPùíÍí‡ÝíGk«7ïŒ2˜ö`0xqmÜï$öü0æÓ¼³*1X‡Ö	}ËøŒ‡šq,á¿+2ˆœòól6œþÌ\Crµyî÷Àí3 (Ô^ì©ã.
i¦SJ°ß+æ$ó0[€&#“é´¯£ 4Jí~ŽVsk°WYBUº†tBñÈÌšDøªI,h
EàÇäˆÝd5íÎÇlˆ]£ÿ¶Ë$J?ðQÖ”@H¬íwA@l ÚÐ7‘ ³€A&ÒÐÙš\Ä4¶¡î¸“D&ì]“NxR$“PL=ª“jx
$}KðL={=aîy{\“ßt­¿%-ðÑhCZ§¡cÏhIrîZKš•Fønƒ8Ø÷qâ)²,1ËceÂ1
ÜNéÈhœ&þ^Ñó¯7Ê´¸VRY¦t¹žBô,q2Ñãë)DÏç)Nô.S¹Û@8±³"&¨õ"aTáZ$Zš4X(-—ƒÃªÉÈ™wèŠÄa„ä‡}˜7\AQ[^é0ÿÒ'ìÑÆat¯?:×!ÎqM¹zPög÷¹‚xtÍÓë—ƒÎ*ØyÛ^ƒHú•ÿÓß'ÍÏdK©{Ä(šGQ+W‰b:\ê¥ÄÈÇqeïÛ .€qëÖ,ÖO=^«–Òž¬“£ãŸ=BÚr\ÁÛp\••2‰éŠÜ”&›T	ãœ—E…Øóq1t¸,üo¬‹ÝXGÙzÙ
6®Fƒûš%¹T=þŒFA˜hQfftÄ&ïß~3—	o@ó)s4Þ8>i2x¢£A¼9²)Íæ¼zïÖš¥¤9ZÒ-i~™ˆÉnYSÔlSçr¼þtÍn)A$‰bÅ'âƒ1Bñ—$!
°P‹æ#·ÛØP.œtÔ[ìwã-ß»ƒñZ[Ž¬qKÎ“Ýß;c¸Çžº?ÈKyw_8ò¾øU^=ƒp„5í{%(éÛú2«FY6žýíñ£Í¿\B•û (½ÏÎº-÷í“ÆžbdLŽo+*¯F»ÿ"òÛµê}‹5%50”áe@1>æÏ¹nã#0Ì¹ÉÙ%cúF:
F~ku³¦Ÿ¯±w¾ÝüF°Ç¦„Ÿä}íä“Ü£hÏ÷	mºú„þ{À‹éÛÑïx~cÞ¯äFx¸ãKPäÓ¶:°d½x:)ÈÆ«
¦cÉÞcìz‡(ú	~
Ó(>Œ…ùCÆ>N3ò¤mš:[K?}«Ôï)Gp?º¶OÊÖûè7.L®uâîéÝ0ÕâÏTnLŽ­¾ÖïÒ—JýYÉo1Púàí»-Ó½qv*I	ÖR±5Mª¿o¸þxwÎþ½a9Àk­²çü†9÷þ}Ÿ•õ™\a2Tðr°;×~ÜÈ­zW9´cô\Î‹Š–­ýÁoÜö «ï§­Ž±MŠ½õéÕ¦¿¶çYS6ø]HP¯Ÿ#ž!<½~Ž·Ãnó¯VˆâíËgý¨‰¥qO#hU=[š§¶Ae·†Æi‘ïàÒg‹a‚¨ö[¯¼µX2ç¹—štgI|dƒü4.¦º…(¬6pE·_ÆCØî[Áë•q"«Ø´—§Lœ¢#'[·Áÿn$7ìw‘ò&§JÉHý=š©yDQpf^í+GÑdsÿèÆëLáj‹"ú^‚1Ë é>ÇƒsÙ›=þo ï#ÌézXÉäžFKÖ””É	¤i³t™+ý¤¦ôþŽÖµùœ)nü­ÍF€%©K/Ùë(s?Ïô†k#“"S6ÐÇ˜{_ˆt-'JoÒH¸EäÊ²ˆhk®¦iwKâ6nÊ®;%“)…yh˜ùžã&Hy4"·4n	t‚=²éÀ'RWó3ÿ8ß4=ì>íñ°å¢Þ‚ü'ˆRJ„¼¡ÿÒ8‡$òþ%®„L©|$™Ûöà“úÉ‚Íˆ '65‹‘ÑÃ[ç´\iÆE>`6¯µEðÓþ¬Þ a¬JÈÎ{l•´£4YNˆÂ˜Š›5Ï# V¥“Ãª	)‘sà”ˆ”ÕG%ùèAim^–‹)9>g|JË®jp	þõ<öb²H@ÈÆdgîÖ“¼þoDké”ãáµ<ha1þ òpÉ@2å'Zæº—‡îÙdëV/{—ÀÉGœY<Ó@øpJ%.\àyOUôR“V
ær‘2õž‘2ÕG\=zâ˜õóN§žÖ	Ç²¡çùo¡\	ëd3ªîáâÞº‘Ü¹¸ÈÒpÎ[ÕËÓ&¨£|Jž†^lAP¿©©ö-™gÙ°?¨í¸Žø³¾•*Y&,WåE
µf^öX!Þ’êaŸIÅ›eŸ¡z¿ëÔRóÇ¤1vX4m`âI¸¥è¶K²ŸEÚ¤4NÎ†9û6„$âª5âªâ
Ìÿ8Ÿj)tö-éêÔá&QV_æôs³¸.‰f”ñ8ßþ»B~geêV¢ï7sÑ‰'(N®JúÝ¦æ‰/ ž0¹<0Ö3&Ú°ä–ã½žÂzhœ'ò´(›Ý!ÐdcŽ¹æ:ÓØcÈ`îz.¡–û„~Ä }î­>:úúºñ¿’ö7~îÌÒ¶ÖŽr‡yÝ¯
æVD%ùÜ:‡z È1‹	ˆ“¹ÐA<ÄWÐAž¡8RùY•gºæ£–‡wpµÝƒÿ"3é®
m¼·Cl#Ó½Á™pJNÄƒeK¢ó÷à	.NEôÎQL}zÕ=/¦tKåøPä­¿€Ä*ó<ûÉs×ÃsNªòþû1_ë¾c*ºGZ_ðŒüÇ™8pÅÎÄZÊ1¹–¼‡b7j’QKê8åÂa®ÝI×#¬¯ü–NÀÚð;ŽÀ<!‹8ó¼)à	ù«‡•{(÷-NèSp»ÆÅ©É„=vfÀGb@ õØF­/ÁÀX;:·[iü¹2ÕS€¬˜%ÐÌL£Ô§:—å¡—o¶(ÔàxæÞƒhÓ³ÕýwÖ1SùWÉ‘³ÞñxYiÆ!ÛNŒÙŠáúÃ^,®˜ÍXÕcv•cæm^(Èa—5ƒÓÜ<	[px²{Rªn‰fœß‘RUÀ­²­Øù‚´_;žqïv5ÌÜàâ.dåƒÆw)øÊ†Ñ›}F:t[s6Þ¡°3Ü
¥šÓvPùÄ;fR`ù•7Uªe…Ï`ÕîÂQkèqòÉ#­Å¢*éDœÌÎÅv¶0Ù!JM+»ËdvHÏÁ^¶%	çVé.œCíß¦¢ý {J¨%©ÒÚáégŽ¹¶A ãd–êçù &ÙÇ¬`pªgøþ½Õ¶ò¹3¹º–Ñ£_Éü„¸’æM¢Yî]¾$¼0g"´Aôh‡ec GÎöÞ%UyEyó¶È$êjW‚³®î¦àp61KQ‰Gèõáó×ëm€¸H¹ç dÏ+Œ£8ŒæKˆ—‘ðÞ7gq²‰Pngëìýdy½E\”32VæOcbî|c“ªW/î\#=²Ü4xtò{½ž^ä:eîs7Y'BOä¾N”®•ý6”¥üV™³?$ÄBÇds.Bç¾¡#d^Ïï’j=§ýf$äãò?Ž®y½ü 6l‘'-ˆúzçŒ°Tæ$µT–dÚ$àß>É²“GŒd¥¶2Ýsòô®6M>Î÷Ö~ä/he™@‘ä2—€Nèÿ­½:-GèÞˆm@ÒÏBºSÈp1«5¡Í”éÒ^ü‰O­g—.Ö’VztáñðoÈ…jó~Bi¾¶Šx“jæ1Þü>€î½Ô]P_’KÝ
Íµn±‚¯zÝ›Ñv±=O×ªÉ/³·>Ó… ÎôÖÂQÏ~ uwË¤|îs²LšçE6©ý«$q†(!mîÝ–þÚyµ½4+BJ–Ê°ËeºËß
É‚­:Æ‹ËÜ<|¤K`ú~&¥¦—-ßè“²ÌUSÚn.ô5EÊúNEÊ”³Ü×+6 càòáÛKR©½ðDÔVÕ%wë<åª;ŒLlED5J§9§ÇÄ=h¦]ñÁW= î*ð»gdræB´ò!·ÛVq­µT®ÃìóÃ¢ž®:èÂÀg‡F f°!ØBû•ùßèþÖ¬F6lX«ýûh³‰õÛÈP$B¼ø¿6õj%±Ýz!Ü8GÈ¹D çh®›œºÒ¦²À#¿)D¬Ú Zž<xôÓµ`š¼`©=œÃaí"ˆ»&	VñRiÂ½¥>¥‘gUîó4ð.éÂOvæÚ¤"ïòˆÒ»¡œIvç¿
éðš?í `Ü1ÔÖ‚nDë¾Â¹Ýv1ÊGåçß.°V×ôªóÉ7ŽW´6#¿«ä]ƒµñóïMˆ÷PåÎÖ\ý	ö$ñÖðhr¼aÙK×ÃA6Õ¡MÙO'îò>üWä–ì®®õªçµÞßÊbÜYÝX]cÞ¶££HW¥<ëšæÓ¬Öî ,7 S4KxOÈ.aÜ€©‘'ðÉ$ûÃ.ÁöÍ|¢\pÑ;Ë;Rñ×_Éú}p9B‡)ãIq^÷Ÿ•óõ)ÝëYu­=p`'ÁåRâó¿«á€3ØÄ¸hŒWÈ‡:wÒ7æìäib1<%šº=va’€fh¶•#êÑM«4fÝëÒdTîvoúÞ¡)Ò™´™$œÅ¤^ÄT›hYy£pt"”Òƒw ÉMÝë&,‹ñK°­®—IàVÎËöÆ3Ü"je¯C?ëùf-Ñ?Äb¯D^ð.[ß°JœNÉÌÖvL„¨‰œX²p¾ÌÚ¶ÒplÆ¥iaî¸¡bŠÝU…6-=WÛ¹ ‚†l•¢2îbƒ˜"Ý:¿N/%Œ~“÷ƒà7ªœF¸Õ4mÃQtä°AÕTº°pX@ö€øìn!»Ô-‹ƒ½UÀÔ €«\®en¬L’»¨¨T°’ëËÏ‹‚D¼›ƒ&n$i|Oû;Ra8¶i9r Ü€Ýá©íb9½±½™Ú
–êSí€]Á@ÖðÐ²í­¾÷óœ]éÖ”æœQ—žlñ½‡ñˆj½¬b<,Æy×—{Þ[5åBKHvË†ÛM«ZÓR“ÒßäçtU]>»jdšgI/²+œÃ™BàŸìa‰‰	Ñã¾@¹Häg4E$\F
†È»Åëúi=ågùÇýˆ
ÆFšöM‘]PÄ£È09vCNcVÕeÕÐÃ;%•Xseî:^;BKeStB[L~,ÆÂEDEäÓ*«/·äA;s	ÿµ’Â[ŸGòÂÈEB@È&“ª¤üâ=•ãv\À™£DZ‚l"´®Œ#¬);Î.wüP|¡ÃŠýwÆüG‘v¸*s×+™MX°+‹±¿Î©pv¥bìùoˆ±w‹„°‘Ë;á ~ßçŸ‡}ëO~?àS÷_ò“Xx˜¹Ee¾1˜ƒoæl`ß‚Þ/«ÁÒ]ž#s£ÿ°GïPâÑ’Â„`
P~Á'ÄfànàÉ¾qn;à7ÉËC¿®óý?!”ÃÍ{n»AfQƒm+×«7ÒŽ‚q`)	uV´]ü‡KögðÑYG¸¾P!ÜÎ¨H&·˜~¥]þ%‡”±ð«ÿ¡;-öþ¨ôˆö±ÃÓ¯Õ-PlÜƒ²n««lŠY.'=u’âñ£ê©£ò’NkuPü˜RÅ%úä	9²ò¬œWå$»` ‘ÏL”‚„Ô°ÿú¦ƒ´°„LÎ’ ,ô—Ðdn¸@RF\÷¶ˆ¸o
vÓOm»ýÉŽÀ.~ê|–û2ùÇþ7ÁIHÄŸ‹üŠüõù9¯darLGmÁ|ÑOþZ¡v<á¤@ûRQÑ³»9”ŠÞ‡(DV,Óô’´‹M„^	Ëëæ1ÕðÓ{°‰MÓ’Ê–# ±fJoV2ÝŸ7C8kÓˆZ¯²ée‰Ÿqbz—Ç@Ã÷žAa\ÔÓS_?ÔÓÔb°ù‡€=7EÐ¸fiþH•×¦Ëë÷m‘î8»Æn»dúº¿OÃØ5›´æÃ¼ƒÉlüíª—ýµVœ9@Ü>Õ¦ÎKšÖî#¯Ò‘Û)Ú¢à)ö°Eq
Na“ÀÚ¡M‹	m‚¬R1ÿåÑj
RTý7Rä¨u	`SíÇœXƒ(¡Xc'pìÜéá	élýµ[KÕù¨`Z¡.,•ßÓ„©Ð>ç|çóŽúdVM†yËÿáä”t¶—;ø¬÷¿æøãÆ±»Ë†_lâ÷—º–Jÿ:ÓÿgbÑÑ”ú'=ÐTÂ“)ÊÂ/„ï¢ãÚV 3ÄKº§3û=Ž1Ê7ñ»cå‹VÒë,	pvÅØòtÕý6¿ÖÐ§!™x6Ç´ÁS=ûË«r@TfzêÈøÇ®[ÙØgTÞÛ W@Ì0óf·Ew^–(Ñ”Ûoç²
•‹R%¾½mVa>MpS›ñäEPÐ¸¡·oŠöØ®ùó¸Ÿ}õÃÜV·–ÝhÉ®B”óÑ3ÖÌÒ˜Ò}á™ù—Õµ·›L;{r]Óž%˜¤Q;K·bøçž=\­7iåd‰î€Î„ÍŸ-£át–stWù¤â7”HàVš_t,‰;˜0{æ°Ö’â›2„.èÆl]]2oØ$Ÿ“}OêO­]¼¼fË<Ÿllƒ‹‹*¿ -#‡LJ üX²1bçb¼`ØvÈÊÊ/Êê¤ÙŒ)P`^ZÌŒ„Ø@-#Aà3;¢1-#Afä X2ýÏË°†3q¬â\#­k&°›è|Ã·œÍLYD©Ì&˜ÁOÃŠÒ¨÷ér,ê5Ž#í$èsƒšYšqNµ~U³ˆUÅ,	Vï¾HxP|a3o.…Æ
ÐóÜz	œÐ²¦ù‹H#kIl´Ù¢‰ó¯<yÉbO¢Œ'Oâ:I!Þ‰ m‹w[ðòë×=][¢€¥-Uær¹€ÖÖ™Ad$%e¨…ùÂZgÛZjÚ°¤É`Î¨k­NŸL³4#íÈÈˆ»Ä?$?$?$¸þ@äH~ÅH3±Z¢QýÐZhz¬{v¢-UgBŽWÜ™:r³Æª#1YÇï(]‡Îí]sËÖ¡ÅÎß]ºphãŠ?-]Ç39Nü™ûCW¦ÞcHó!§!eÌ7þl ðÌ¦e÷$û˜“Ã*;ŸòÔjäÏ®ÆÌ;|à•ã¿¡ãü#*ë9†ZÖ´™¼ÁZú/D;¹§ÍH {C30,Àê¨;Ím<NMé/!5M'wŒIŽòé˜ÖÎïƒÖÎ'Ï.óþ‡gEÕæ¾ ‡½Û|FÔ-JðÁ'Ÿ~>rjáóyZÐkŸ¼×)j€t”]ÕôÍÉ ëçcæ@¬NÇ¯ºÖ?ö0^{Âcž®á|Œ—Õ4‡Ué®Êú³až	ŠÇUåoºÇµÉï¨A§ˆ„%¶óæhØÝöµÕ?ÁØhÆ¼rw„y¸dˆyQCl„=ƒ½cp+~pø]£ƒ5((u)Uä¼*GF4[ÙæÙþ\³à¾²µõÜ¢YÇ¶ô¦DDóÜ+ðDß= |çA˜·Dœ 	výTûF¾™Z!žìª û†—½à@Öj.@šsuÄ-e.v¸wÌgÉq6Ô6
}r)In<ÚtçM °øÇÿý?„5S´rô 7Š†Ob/<mÂÉÙµúr4)«)ÃšY'\Ãð'.»’AQÑö¯q	:nE]Ïh ¥þV ]Üi'üðÖ«¼=Q¥¦‡œ6NÃAÊ³Ùàž{Ãá„(–çns ú¿Ñ„ÍÄO<òhçí»ÈçÁÀ+ôƒ-Ï-¢ÅòÁµÛzpßwâewÃ6ÜÛ{ZPGh¦#NrÿËVñÖLD€¹®¯X˜mŒ7Ü)Q1Ð£š=^•ƒlØXßHµ9Ñ,ß¨½»Žª}Ù7Ý8Õr§Ö'Š¾~€åÍ¹/ø^®fqÜ0r·›~d§—»¶.} ¿Û,½mKûø±/)žŽµ`ÎÞ´òâ	·ñ«e Å¯ìqP$/J¼¬Ž£ÅÙR`h}„‚=…ñÄô¢ànèkÓbåþ‹Pn†±zò¤¥£²‹kÎìb£ xb–z|‚YÇ×,tï÷ù§IºÛt¿Ø¯Áýr¿— ßúÍ±äHÇ[Ë®zÁÕQYüW#°WÙx–ù¼@9)ÄÆ2Ój´®pîU¥QžŽÃh$ª|ŽÜcÅ9—L•ê+wžJ/¥.|L
¹F6Ë…˜N [$ôÜ|¹ëX!"55»sŸ&0T†ÐyÒ=›3¹u{å¸^¯ç«jÌ	/Gœ2<ý»àr+ÇÃ¤E+¼jOn†äHš‚‚Gi
“erFwÄuî[ÍÜæÁék*°»F;wU;é–‡	<&ò¼vC	t¢CcÆ›)!ø,Ë¨/ù„Åvƒîƒ>õ¬ü”×hÐaä®	ŠQ¹"lY6¾>8£dpó8³Ööêr”³a¢sQÕù‹a™M;“ž˜´““=r“^‘¾™õl?\¨"úY¼¦ÔI:cXw&Uþ1å0(bc?AÌ0¬Føð]ìh†I;´ ”¡°VOOV×Ö°äÞdV_ò;)%ßkí.9Ö3Ñ7º©Ó®ñ2ù§×ç‰Má@šôsXðÝ-ÞðjïçplE‰)eô‘Drå¿ýprj'»ÛjHEþ*n9®imú,cËÓPrDY°8†DÃÂxS&çVCÝà0$Š«¶Ö¸Ú¤Õ½P	5\wçòOÿû™È+3aRèbRí¯*×¦I¶¿’¤Ç^à ïù&É¼sQÿXM;Ö£.BK¢@ÇN?ŸúÀø3ª ²gà5F+Ñ×FÙ§Ýù8¿"ÐºÎZoZ¾<y-œ`DlÕ)ÆVÝ„Ê©òŸTf
½ÃO¼Â)€VkÙâ4/´ÕL”ÛQ{fÕJóÿ)Ž;btþñ_ÿo-Jy‰ü6brˆ’¯ `&¢iSñ7‘È«‚ wStsARn*Å°`§6L’ˆ?"è³IÉ­­#¡W³::˜½Až"¸=F¥·ÿ= 0ùZÍÏJK<¨²ŸìªÂ=,á°a5.®êcÊŽ…wp0WÕr›ãÀ’Ÿm”“±ƒAÔÊhH‚›=€?N4VÍßG·CòÍ¼¨9¿y9¦ü¢ -Ï*7’Œ`‰¼ÊRø¼2Ø€xµþY®.ÛQ9í4ûàÀ‚JÀœ÷`ïžhøGO4ÿè€"\¦mGãÏ+j»Où7®1^–d,·Ã&Xü¨&WU6‰ûLY-÷xdx×Æ§º‰”(í»¼aÃI‚¤ÄÄ¡©Œ-kxäû¼ÞÜÑ–,Ë G•2oÆšGdÐUñêö:Á†ÅÜ-É‡uî)ÙÑm®X@'˜Bïè·à·	0‘üébúÏT´øMò?þ~©•mì_(Åbã>ÕrIéî)·9FÀ Ë½-‡vpŒB# 6‡R4ÇŽÁÅàÎY†o€VŠU”Ö1EDO[`à…uzÓË<Dáà¦•½ÇØ1µ=Uf™Ö;¶”ºŒ¶Èª’\BÂê­•½SºœŸlL/+Ódò’Ø95¥,úN°KÇ'Ó³rp«lÞIód$d˜óXÒª² ÚèyYû†AÖ¶ùâs»é°S°²¤€š”#–ßDMÿÌú—z’6bpyòE¢f,Èð‡ˆq­¥=9é'g(ÄšÈ]›2w«Âè¸Caû=i¸€2ä´²4:ùv˜ö>~ÉX!vZç‚(ZüÖÀ%ÆÂþLè±1åØWüè›Ú¶¥i]âfêv}â«Å×­í%v+vPÇwØnŸ®ÐÈš$l¨ÓÉØ`PYÙñ4mYÍHÿiYólÜQôÏ WÙ=±Ž:1„]éÔœµ'f­jSžµ#ä”ùnGÄÙ4lò{WoÒè7Þv¨ÝNAZ[Ôý„¶Xëàl³«±½Ï©Å¨êZ3»ŠöË”q=‰‹>p%¸ÓõòZc"ùòÀP‡¡YÉ8þã¿þßýÍu I*:EåÏ`ÉD“ìZTxdý’¨±!xéôö¶è]\ÌWbO±ŸŸš‡9:æÛë‘èJ¤^Ý)]´@1×/I?„ÁÇ¿êhêšhØâÈÖ;¼Ž
oÖcN&{ÃdDÑe«‘°ÔçÚ\ù rZW 
:µésá:9ëépØà³þ‡Û'€Šs»:Ã'7§?œ´ZE)ã
±ÉXäËÎ'‘l‰pÅçÖÂÖlNì N_¾‡^ia.jé,Ä®%íï±XgþVÐ¢ò}’A…ŒPír|pÄóó$9aí	j^vé®h	¯‹`—lóT—ËèXÐpÿ¡kã'èþ3ÚýNØ]^\†Ùôàé³Ãç[Û÷¼ˆ%2…ËåO„ñ~NÅ{¼Äèj‹¯G÷¢%J(ŸA‚ùÂ9ÁµU].BþA°dBÞÞ>€Î [ÿûáo’šo•Y£eÚEü÷–fW¨I29«>Ô1C”–Uyæ§eî?ª
Ð}Ü«³áÓ©JõO`È±S‹ÝOöï4èÿ˜8JÊ9ÊsE²×ëz¹$÷Ðª¼ªwçÛ^×Ïºð’žçÁ“ämÏFéíHX^^åÙ™#Ë3q•VŠ/Ùì‚ ûçÅ88êp&C‡ôÆýÃÇMž¹òï9ÁlC´å°;ød·0J@‘éMN‡è§ñfÝÁÅÁ•bƒ7! JŒÚ@x! !þó;+E'Ð€²!o¤áž°åÄyRÀCžŽðŒìØy Þ“Éªõ…S“éeFW9— ÆRMaieãk–t÷,c)<y^fÂ (ÝUzy–€è‡ð<¿y~òü”ütôôå?>?$ÇÇo^ÿLÿ8}MŽŸ¾<úÁ³èü@žuçXj%Ó/b3øPÄ4ÃØÃE"Ò]ãñ¬ÇÒò¯Ã €ÇŸÃµ€€Ñ¼¡HÍÆ-t~¤Ÿ‹—æ`Gtxöò¤¾¬Å§ˆß‰|†‡ÿ»€ß„3/BóP÷@É H+Gé‡ Ýü’_š§Ù‡œ °Þ±íÉÉe‘Øµ+ÒáÙˆíõ%øÒ¸èÕC”wìƒ–d-I
èºÌÒ~c(lJÚïµÏ¶¸Ã2Ïï.”'·4DçèÅÌKŽ› }ô`˜WÓÎÊ) iëðÚôg6¤òâàZ1wÆÚe»ŸíŸåT¤+«·Qú0ãJ0Ì1´vJj'0çû„oZîÚµ´i…ì•ŸÆL oô~et4¬ÝŸ0¬šó+Îwøl¬‡¨p_ãÔ;ú‹ëä<Öyh*[°¾ß7˜¿ëf?šˆ 1 )‘ÄX½#¹¦¡$Ç¬÷ÚX¯øºàX>9ànÖ½[s¬µ8sZ›I OUñ:‰@V9=Ìd…÷ÜD…$%œHc{Œ¶1¼ÀÆ†EÜ5¤ÍðÏ/Ð«“þ‘ÁÙ/¡à’ðžÓ(íWÉ-EÁõp&<¨ Pž‰PÚÖÍZfð«GYÀ?ï Üˆ*nˆòvQuëD$€ •u¢ÀþÙo¡ŸÀxPDa¦¯’ûÚ*'9ºZï­å)©	eu@ëåÀí1øª9|Gø–Ú‚©‚LªfxT5Ã£ŠaO…P§üŠÏC­8ßK¦VNCK£VI“MjýÐ'Ë§Oöžp}G½çàˆo&q¼µö»¢sAË¥tEË¢u¢-ÞüUR·Jn%p•?ÃH¸å?·a„ŠÍxd>1-AVsðûþ#G³l±l´²#+ŒeqkÖÊX¤·¶-+šÐ„Ù@¦šßÓZ?™uïf½+piÛ›¹¬ÍUÿU,qmÀS–¹‘GîfYs;33MÕF~ºÆ*QO~_”Ž¹u–Kâ<Ù²ˆ›÷U’5ŽnœœÜXkºcz7Ùß.²ê0äõ”R2®éÇ¨-S?vÈÝÑê“—]Ô«¶mÎ·:J¢ Þîø¨ÑJ3®m$Òð4¯Fn¬zGúUÔ(Ú·Y?d¡rõ˜)¨ûUžÚë.ê¹ø†>ƒofFîS|/<½vv„
^Ò~¯r(ÑÌ2šj6•(•Iy>†¤M7Ù\=uò2X<7êá†_5Ú›¸ð/9æ'`/™Òƒv)'×b2š	t=§üÉÈõIRõÆPbå€³zøü5˜0­Úˆg¨å<æmê0ŠYûkWBTÇ­ð¹˜³ŠÑ(ÖRŽ‡×¬ªée^T¤¼K;4ÔG·`ÿ¤š4‚‚’÷{³bàTë
|9³)?ÑSCŠ¾4>‡a>‚¬ÐmŒ˜‚0alÝ3Ç©I¼$¬ PŽltœ],ƒŒÝ§0ÚrÈ¢
o)&È)gS‚,çÌ«ël„ oá”_LWkEŽÎr(ÕÑ;AtŠ(ïéDïqôÀäÚX[-6>}K^nŽÕ¸cr‰fV‡¹fÑ£òúivÆ&­­N×	aGìZnGf+vË<É©î5ôÃø Nñ$Ÿv˜¹÷|˜M_Q	Ê²øÚVA* \¸AG¸$‘]™ÚÒH]l_î7H9Þ¥;R6îÙðHà‚ÑØÜsãÅ/ÎOî5ûÜ»ÌêÄr7Ù¹§‡çEUëî!fÇ¨ô‘3ùÛðpt°w+Cyî´’³*Ø¾sR¢âÁªpŸ¥T¶ÂEó¦=KJ˜öQë›•Hi>§Ã%dŽÎCÿÜ:Ó`øép_þÎá~3ÆFuh‰›±ùÛQ‹¥KjQÐg‰jÄoW5MuXz5ÚQ›W¤ÝqT¥Àè“‘÷YtÌ¸éêþB¬Ö Ï·OƒËØÊAE=H¥ÖÌË^xD*É£²üòe«œ+£!je×|e—'”ZK¯^îYhŽ²)¥Oôš9*é—ÈÈÄ¨Ÿ±»”ÿŸñ¿1àêJøâ«ÖÙ_™o:–®§AÈ—¼Ðíäæzm¨u?ÆpŒQÃ±9h‚áÐQÓùÏ.'êÆk‘ )x£|¤3(]µ™ÝjÖÀùy°|¶×¨"“ÙŠÏz,O9Þ KdçÌs˜GÈ¹Ù)<á‹JÕyÝÝS
ñ©R„GuÙ9ëU,z€¥¹Y[¨l«`W×972å{©ß0ãêŸí1¶ÅÅÖœŒ¡àÎ Í“ŠštÔ;l4wH“¹N8þËÁ¾¸ç9§ñ®Ø>{®ù‘ôh= ½9ƒ…|^Öm­µ«K¼z–äü¤M†ù«lœ]Pg5uú½êPéœÒ~üÖA+óùõ“ÊÕu‚
ØVÁ;¬E;Àƒ™Wœ×šX9Mçï•E¦Ê&šB#›±„Ål§BAŠ˜£F¼ïw››€AêÕkÒZ¼JFÌÙšã”oŒ`’7Ç‹ßÓGÇ
`Ÿ÷á¿`½K×d•¤lÖãßÊbÜYÝX]c`ªŠ9_ÕêÅ¦×ùol<w´(à²š²!áÄ´Ù“TÁKÝ—T± µäZj,°‚+i‡T[’§¾Ê|)Ð½&ŒxŸÓóëX$;Ú8X½ñÓÏMhqóÒ&ñäõó/p
k%ÂYßÆö|'£]/i´Uïµ!×\ëÆ’ûÄÛwÖ<¤e OˆŒåž¶Á‹õHA`hëí†<gº˜V9	FÙ§îe÷þ6ÿÏóayE[›Í¦e3nÂ²×
¸å”TØ689ƒ4ùÜGº,®ƒÄ¡‚~/ON`	}'oZZ°Ù•-„C40ŸYsÎ£õH7‘IÞ4&	œ7Ä¥ðÒÂ]ØË±šQ|B8åX¦¸@×ßdº­‰?ö6\Å¤KÄÒ†¿óŠÈ„ž—IH?Ð> CAì4ªì|ºŽ%ÑZs÷V—Óñ›åTô˜HM$Š[½]óú™§GC:ï	EÊ BhËÅM‹š„BâË5ïhžavEñI"ê¡¸Ž%g2lhÚ©UÛÑð¼m	¡î]XÏW­åÃªFn-G¾Þ‚Ò.]¼7¬6¥" ›:Ë¥."e›h¼®’ž80á
¬_Pz‹¿óè-'÷–ÈÉ­á	CØ‹å"êï³5þ½ú9¯è¶ ç_ßy/ 3zÆ6(’HÝÉ° èjþëf
Mp‰QsªÍÆCGÖt×5ö>ñä¾#–Ú¢OuA½m§\OK¬Îyãù˜6†n"?ù¡Óõ¬ †sE^•ôè˜³é+Ø§¥wÏ_+,Ì7MšK:#ƒäHwLK
í¢!Þê˜çw613y†Ö@ —ý| Ó!š;Ü ‡Û#.¦UòŠ[C5?®ò!º¹_¼ Ñ0H¦s¹‡7Ýµž•E ¨5³as8*üupL-Ä°béò¢GNßüÒÝÚ¾¿–„·‘–€&ØÎ,5’§këvŸ¹š)šÏÆÛ™V¹TÔ®ÑÚ†7¼Ñß‚ýwM’¤òÀð¢Á&:®Ñå³ªzÎHš‚dì¨µBað6û…*ß#/[»sâi•›€m@å±^(6Õf4šv7{Wö”[[b÷–m¤çJÌ#u¤ä9·g-P8]N©É&›YžÚä*õ;/K|¤1gÝhY¹e÷„y…±NƒCrÜ&Ò)Eb£¶ÙX%®³«	›¥©B2š’	ý•¸±Ù’QS)€à‰øƒ6f²¦Ùb>zÀp«-ê
~ö%‡‹;:»½*‚éÎ\ŽåOb:ó^¯g–ìõ/_¿­_2ê±PSs¦êb³®EüÒá
|.îŸ—î£^Y>ê•î£^1õjõ*ÍGûðSço,Jl8swAjœ$LBû‚Éãäõó4âp9¢(aá5É?ªPt»DT×Ë÷1óZ¢3™»··Y®ùa€š<’—*,ïÝAh\±¥¸¤4‰Î³ãcòž¶#<ÀÃZB8d!°,Ü„ÛXÓp×[à$hH	©Lb,B.>µ¢‘%ð¶ A»èAÒNßÅ’‹.¸©^ .,ŸêÇ_ø2é<oÁƒ¸£ó€_p½¯à`urU…7Ée-&=º>²˜Â¨fILà×X-0ý[\àô·Øb×€'–¾Ü[AP|‘<=}Ð‰¸_þR¼òƒÕ}ÞeÞ¼Ï~“}ñ× …Xþî‚‡ø"—_;8ˆ0 DËehÆ­˜˜·
[	Öú™7]#2wÞßæŠlq¼œ¿~Ò/Êjôb6žô!<TTÌmïí£Í—ïÐRÅn<Ú\½‰yòÑJŠAÞ½ÖÝùú³zZŽºu¿*‡Ã³¬
¸5÷<ï¨QqxÄpø¡+<ÜQ~
ïC¹}9æÁHiX¼¹°ñH~·q.ÕkÔŸ¾&66@FÞäàžŠ(Ñ‘Fè1‘š
z¥=j¶-:¶KVG}Ž(Ç`mµk€\ïÜ®–;ˆªWxñ•Á
ï.Ò2ÜÏ»¹T×í7?‡ÃpËOÑ„ÛËò^n±˜Á=7u9/'NC/«¨E$7ÙM§0èÀ#ÈÉ~YäÄ¡]ÓêV@1âj´¨-&;ñùË 3if‡¾ÎÔP5J+Þù¤ö1‚½Q´9©u¦¯Ï¤b)—ÀÐq áÈÿ}ÙàþTñš´b1\R)b¡XVª,ëÙÇj$ßÜáU	^©KŸÕ·ô#æŽ*Añ'–Çïfì¯wBîà©²…C;éî‘$oA"YKúH™BJÌ {o1o!(=â#bÃcUA/\-2P°ËÎÅ"ØdgRå"ö¯Zl•N!Þç£Ëb|„¤Ì_Ù!oéá^[ç÷ÞE«‹'<îâÝmÃ÷×ôßu$OÒ_èÉ¥šòÄNp“{é6þP¸ w‡Ò•G{È	¥+3ƒ¹ÂÜiž™óá€øÍ¸ ¿7ÿ‰á>õeUŒ?t7ÃÝtt‘1ëw˜g1J`Ã¿¼ºôÚ©÷Ä‘ÈFÆƒÄƒ‚RVº‡@Ç]Œ»WÝ”$ÌÁ|ef#ª …Ø(¹ªáòœÂLÑ·	ÄJJiîq¾èJ­6i=.—Òµ:¢gøÙHÝ!"–÷†üœªwTLÚ­ë•Kƒ÷Aµ”2ÊÃé|ÉMìöõ<À<p,tOP‚–1mÄ_º› (©QZ’âˆ‡ÃâíÍvNºÊ:qT9[4g’ÍœÌ†ºwTzçIŸÞ'$¥\›9*Už`ýÑHÒ€ù®}¶î¹2=»l"€Y°}¿f5À/ÜÍ¶M5¨ˆ@[1†nGÄ¤Úô™QFÙî`ÖÏ;z6¢23‘Pô-èÅ-®“ÍµäX×ptNÛl¢p»-‚¥/¥(Ã!¹D½˜·nY´M††«ÚYÔÅÛYÃ0Þ¢>|þºãÎÐÄcwU¸¯R/2onÍ\ázî(ËÑ0×4V.[°ƒ‚®Äî´ìV@+Ió{àçewšgCü[²mñš¸%BÕFÍp=BY
5TÈ%õhÿ® ·R“Ièý Ã7Â÷ƒ™W£>äXD2el,ªÖ—Ç¾€?Ê„6ûë~š«r«
7äîŠª{D¥ÍGÁ$SÀ…@}—,–Uš,‹Mµ'élPŽd±}<ØPÄêA‡ääº¦EÆE±à®`ž6îoÚƒïžµìm
g#®WEðÒ£Zÿe;Ý«40Bu>*X”ò€d%oÀÎëÊ“Ë§×bôf‹‘¡ˆ×	ŒÙs*r_Ó–çY=«xÞ\±AÔ²NK¦…NÆñá‹Ó/~Ô§hBÛC·Â:YžÓœµf­¹\è2ž\–cÌ‚ˆ%\³·©¤DDZB*â®cHIØv”gô-DkîÆn²¸K­óQÏØ&<X†ò¨+fÔŒ®e_¨&‡¶`E=|âÀÍa›ÕåpF×å0?Ÿ"«š–“îÖÆ6é¢ZV‡)Û&ùÜãŠÐ¬ U± -(3¶`„"±(qPV’²"/!3zà{nýfdÄFâ%4,tâvyƒüÜ, =½;ø?óëÃòjœcÉóÛô>ägxõ9Pž¬nô¨¤5êDC+‰êå=0~/oa1Õ•s8œèÿú  2Ìl6ì>&“ÊHR.…Ž‡›||š
Ëð°¹»
ÖÈiÉéÔž—ýY½CÍ°çÝ1ð(vä]º¾µÚÖà§Ï÷â äAèA+-,Ðã¢¤·<Â[„ì4£…%(N×{ôØé€…aôvóÝ²C› .*W¨7g©™<š®MsQ}Hq¬cIÐÚ¹ˆNÅ:©aMƒ‘mM}SÓ6¦ðæ‚²£†Èq:¤Óö°é:&$ýÁœ)JŒêÒg7±’è: J,ÍDJ“¿88’pvÞ…ÝÑB×¥Òa±	D¥×¶oö¹8ÌÃöÓ…s„=´œŒç×_ÁL€\7,³ÁÌF#±µ¦ m;¢™‹Î…ü~iâÇ°sgµh£ê
ãÚEÑì–.A1·‚q~uÂUo>E*5h¯q?$cÌ0«úaÙïàÿgëD!q”u2çGçÕŠuþåà`ºƒˆ½qyEû
vM’ZÀ`my²Û²Æ7skLT3…RöHl%†™Êe(?²¼[Íí(h¾f«&ÖÀWYõd59dsG[ñ‚?Çe¶ê+?`\×øäM~^åõå³þ•g!ÒF¸LÍò¹_'‡)æ=)ß~+yå—•ëÑ‘q M²Ç§ês–íQ+ï÷–îñW‡Òø‚S“j€La¢¢¾ÅõwªfHTË#j‘!™¦}"×Œr}b@“x§+¡% ÔçÉÓüú9ù—<,™©1 ˜¥Í¾öå«`h) 0st‡ÿ=åLfX¿Ÿ›'1†åQ%+ï–<‰çðù
XR I±5T‚]7Òk˜^jø0ë°Kf´C	=BcÊ‹Œ¶ÐX•LÈYÇSã•ÆfdÌÍZ'5ïã$»ÎóÅ?ª·™[WåxtxàNÙè‚“Ö…¿çTu:Ý5B.¥M÷-OÀ]ƒíQž(%µßI#ÀË.ï[fó3`O.ö¯Žaª¨¬XatÁ5ù9è÷1	=¨”Õ”2ùËûžÊü&zÁ Wö¤Öi‚d¬ðuòòÅÉ34·¯“ãƒ#f	ÿáä”òiVkT†^—³ŠË‹:bÅ˜­·ÈüÚóØ¿ƒÜß=µWU6	C/žü¯.Ë+øãªO_•ƒlØ(µ”}ßoß KüsÍƒzÐïgq¡ÛðÖS$¹ ·W¨À®ÓÖ{” ø˜lÐóE~µ¸ºt7àDŒsV¾¸õ<ÉYø³r†Dþ‡„Ùn7w4C…)g ‰ý’fç_)C]ÞôHÝœ¼óeLKòSZ_d‚`DH§7(ûîP×ÐD}Ví¸¬yR²<BVq®ùòTÆ<6› gízË•8VOXŠ8FhéŽØÏåð‡Ak$ƒsf|P%fƒA•×z
yg±r=“Ê«cÛ$SqÐ*</ËTwÙøƒÜ·ÙêÿÊÕ^-à{ÌÑq#”ÛÅyÝ‹âÍ{wOH»t„p…-X‡1‡>«*)Ç¥SÐaq9Ð©hä„¿`oB0Z´1P,;Z‰Îj¬Æp`‚óþb%œ¡…q‚:¥’`øÍ Ñ‹üpÞ¯=ù$ãŒwrÍ¸‘’ÄØ´ª€Ðíp{ó½Eç;HŠþü2Øî oD~µè³h¤‡Áˆv•SrP×Å=FªnG[Œˆhy’%	R¡Ò;‘7©Â9ýôqRi/7xçb¾¨§Vƒá”ñ68¸zU–ÑÈÂÉGÛ®Z?2¼ÑHû˜²´*5u)uŸLò~q^„ÉÛµYp„kÝÊQÚp½áÎƒÃžD8S1³=ÀWt…_ËLL„råBÀau$B`gý#¬·]Ÿ{„Þ."ÝM—xá®^ˆz—Ó¥Û¤%jŽˆ.ÅÝÑ¸èU8G‡K…¢m‰ÈM6t£÷XÈ@fgoÙG)X9º‡j$Vã’{Ìæ#Òº…vZ'¤]È0¡?)‘ „œŠ á"$?Ò¡m ô“¼¯ý8I‘| 1›\Œ|t!5½Ê0“„^kGxvÚ!½=¸û‰„à¡JúT	¹6¿ß~ü ØÚ ;|VévãÁ¦ŸÈ]NÕa§8×øÛºˆ×¿Ï˜§³)_.•ÈR{¢_Içõbù´qVk.iË”âéRû+KuõWI¿Ÿ§¿>^é?ì	£<OÅ$ß[Šp³dCÒJThPË“%¡®oLí‚ž?·ÆAA9­2Ò¼,ƒ|¼ÒœÞ(Ý³¥'W–×Ïa‚ÕC|áî\(âÝÃ…¸(±µ_®«Š€6_‘z¹Ðˆ- Äc°R}ˆ[¢aÑ‰<ÕvU´}á$²aŠPé³ *ßNjJ×[<#îw)yoý)_F\Û9-„	.îGOÏKTÎû	§•øw\V¬úåá‚I¢žÎŠá€.Ïíö	ÃY,ÔÜžÇ}²
‹Þ”QÑIùÛÆ½(„Q,¹Ÿ{~îPEžãU>*qm¼ˆLóAª&ÙtÈè¬Žh)¡ÕÆ§V÷ZoÅÌ´Ü, üîn¨¿zC‹>/hñéeQíÂ Ý–¸8ñsXÚ˜;	tnVWöÈ.öu=Dý\Èõ	Ë±ù–õ’Åê…«°‡ÒcÇv?/j •*üEYýDåß.ï@ºÉ@ì×p%J"{ gÊI÷¡rã^‰”A;ƒ kd$£`?oêí~Õt» LîÂ[	ÓÐ°°&ÃZ—¼¼„J4³aN…0Â¦)sð ’C8~`þöQGâÄCÊós(‘"¯ÉeV“3ðVcHóø‚’Û„‚®ÀÉŠŠðg×”-æ2’€)ƒŠš²áiYùÐI¼Ø$­c•À¯ó?£[7%Ä§Ù˜2sòç{ÿjRß¸ìNÐŸHsÆ3é\9hxüò–ãˆ‹g	>}tCõº[£yö.bhcIú73“èøFKþ‰¼É/(Íä•‡(nå^èi®žÇù¢æ”6oD…\³©Ã|
Î½xc4N‹	]¾ÌÀC:[›ëd{ó[ñ{B‡8ùâõ †PI—Í4ûD€/‡‹±Ÿû³!Çn÷ô¹÷¢Ðv`»½‡ºT{#0·n— )¬9ò/…©ñc>›¤£²gô¾ƒTÞŽq—Ez‘ü#²Õr%|f(|“SùÚîÛ­ìM†“¡æB
Ùj—Öw"‡k¼õö’ì³lÜÏ‡b–,:µÏŒ²ùWY1&ÛÝgåp6“ªb°CÄVÀ'="ž‚-;ihò¶ÈðbGûÉ{M¡ëÿ  ÿÿì}ëvÛH’æ«d©=%ªË¤$J¾imÕJ²]¥²ì–T5Ó;Û{‘…1Ip Ò’F­sæì+ìüÜ_ûû:ý$›‘7$€¼D‚ $W?ª,L$2#ã_³ù%>ŸQó`Qv6%|Ùt:ÏŒ³fœJ}2ôÙôö.e÷™.•ž†I%©,Úà³jJ 0´ÈôŒð IpyôßYTO5·AÈ5€Ó{ùa4Ï’l02¡+„­
oý‘¬nöñó=–©tàr _p©BŽâË²±VI;u?Ï,ç>hpÉ+»¤²TÄÏƒ.Gàš½ò”¢_Ï§#€¨ßA:Mè ì­zûé€Ýû6KÇpûáäÜ!“L†Àv%põb£^QpWl®8ð/Aÿ}*ã“1 ½%Ð2fÿwWTš@7=HñË™ó(’\-*çA"IwéÚKðÏt€M·ãMú%ü)'ƒ?6ZŠ_J¥7ÑÒë~Æ; ›J_¸ê@š›*ÃS¦s)ì §
Î U-À#ëÞœ×I>ˆ²aqPëíöozwÄ m›–âEÁ}Dm1àä€êxÝ©àaK psB9(ð]&¼<×§'`†ä”ÅÑX™/©©™N>í
‚äÔÇÑ)	òïzÊQ…~Ž#_e14Í [4M=öÈŸÓ9òŒ†pôE_þ˜êõñEw÷2šÀázÐÑpÈÞD<û,¦ò &yô…™áô'3.myš,c¸€­˜FtžûÀ`+OÎ•w "Sþý|ëagô«CW¸ù ìá¼?e¡5?ë÷ýn9¿x?ãÿu2LûVÎ/ÆiÀøÙÖê3K9LÙè*v{ÿšêHÏ®É#ë·yŒ½ayÆÞÌÙTïgCIÖWíù†ÎÄê¤Úõú3cf¸·_È0@#êØæ_ ÍÃml~5žtÚ_uâ°úlY¸xV¼¶Þ0ÛãxeŸsß’[AÃ™e³þâ‰*[3Õs^œ‹J•.•*‚Gr†>4„p.àqåéd‡¬¸ÞöÖ%µ=«èZ·­’¤Tûs‘bk’^’q:pa!Ž@@q!Æü_Ð”@]9a²‰|ÏœÔLb®(‰Jµ©ALX¯Ô»\iEÈ`RKš3É/Þð‹îxðêÔÝE<.3POeõäúmÅWb>AžBØº(ïÛ—$*„!4Pß8¤bÅô{Š*¾â“‰{£8›f	]>‡€×@4¨‘)—7N
¡òb(<iv¼~ó~õv§mIX{j;±Æ•ä¡rD6‡ÇâÀîX£áu¨|´5ýÎì¹­É7»L#ìüxƒá¨Î…_qSL•Õe^±·˜ùX¡-a<†gº±G1p@“û¯è	Ðìì¥8ú³°ç$Äye×¤X¿Tdi*#‘ÊZÀ›º ûvà(½|#L‚¤.˜î,7{þ|¦¶ÅïÞœ
 “èK,, óy8‚ï!µ;kÍ°Àµ,Ž†ææ³6Œé—ÂÿŽ‰RøTäžKåÃJ¡ OœN~v›Lß@QŸ»=­ –7])¶îC–Î8¡b¤ŽQ$èN¯ñÈÛñ†þ|€çEŸ&cÈÜûÈ©z„tóZ³E3j5šî¾%“è<æò•ùÌ&ñÎÔ4½¤;”ÎgÑ§x™•oD.ƒpi)›§ó|tGH`ÆPÙA†ñ—d[ól	¼ÇÌ•(âúò¦3–Ù{bÒb?°_Í™Å¸14|m`Jæw¨ˆ/›Dûi¼Ü‡z@¥QÏg‰Ì—Õ¿½Ý›îL•ø'…ûÔ—äTÙ/•ëÍyÒ{¹Ç¬¢Å,- YÚ+Z¿!çrj”…ê‹äœÌâ)ÙÜ!¿°L™sNÎ“OóŒG)±Ü»&ÍÇ;ÕòÇ-½Ï!ÖJðl4£^¤>Î”ãn2(²WŽãsYã©Hf˜°l’•ÝÛÿO"µt¸æg*¨r9ï‘v¡Â<î9nÕºcwÐ­jU¼>•ªÄê…$Ýl–¬ŒÒ6ÉÓ>Ig€CEÀèýjš÷¢yÈ¤vŽ…è€C²\§?ÞaAe ¢%PÌ»Q;ªwúCZQµB„¢‘Ÿ¿ò÷MzoYÇ‘	UJÖÉ‡8KÒáƒ¡AsÓh½t nLˆåB¨±^ì»7ÿÄþFÿ)è°Ï×áŸ®WiJÑéØäVW@I_%‹ºË“Á}g?	ò‚Huñ1‡áÄå-„{W¾JyÞj86Cî”#x»ØÉ_¥1ò'¡­íî‚vEd¤‘Î^+ÚívY
Á:)Qévýe»5Ï2j¼ñih×wl(¶Ë¡éôhºŠÔÙed\*šnõë§d	Ç¸tão²r ùšæ¡‘ÿÊ¹ä²ƒ|`‡¿«ñyðÈÐ/{¹øÅ
kwÙQØC+àOp3ƒ*; Öÿs—ýÍ€-¸hèëå~1¢·d:ªKD3g©É! —º‚s€Ú°3ÈãÑtW¢ƒX_ÚH}©ó¥;°ÒîX§]LO ÄSˆ+,›‚Ùû%#H~<I¿0°˜'¦>«a’Ã ÃbzG©„º ßøÙ|Hw‹D’!_=µåBËYçì¥f8Á*áª¿RòëCžŸC¯ÎÉßþã?Ù:@b>%˜ò·ÿõÿè—°,ÅYô¥f-¸|x´ºvKþ
ƒÐ¾ÿ.‡?ü¿¸j2ÑŸ'×Áó[ßº,Îê¿fí°^)Àà£¿H³dæé¶«~Õ€Õ»ãd}0~$«Ih p73/Þ][0ªå/a“u¦zUITÁVy%•]V©Ÿ‡ˆªV„•±»F‹F_­@EX¿}ON¢,¢«´²«ýA:'Ç?#à¨,CÑQ>ñÙ?Içèø§æÃýÉ	ò“ÎÏûGÍðøˆ²^çø¿5œ£ª{~žâ•ÝÊ¤óúð7ÔÐ>Fë—®^·)ÖqŠŠ59”œ"gLèù"1!‡ÛÒú&÷ÆE1gŸ,ÅGrš’Ž’`ÏHë>öStÐÆ”ÅO{Õ_¼Uûë·8zÑY	rBuú}®ÓïquxÑ»t’€çŒµ/!ØbþÂ
áC±Â(¥‹á!GÅvâ<ãYw“¥€Õ*ýOd.Ý…*±4tþw©'VxÆqÑãÉŸÏ¬`¨ê3˜žÁ[yržÃæ ûÍX[sùLƒ×éhe'É§I™øTA‹¹k„Oœ9dêgÖ<2Hs-åA* §ãB¥’Ü»‚¢_¨¥¬û$Ãžg&h.±çËwêàö'Ê8~ÿ†Aývj®cPÎ‹ÛØ7L;?)ÛÕË®…µm¯Ç…á_ìš}íf'Hc»2È,E#öu€á]ˆ´Â=ia¯¼9üu®ÝÃq;ásû(6”·{Û4 íN‘l7¦à·àçU¢O |òpÕfTr´eúº:.!¤Ôâ]³½ïT–W¸à–]â&µeÒ‡ß}ÖÖ[1Ië{€Åì76#?8+Z^ÿ”Ì.’‰ÜøŽá¤g’ìÓ©³¢+×8‹F/q¿qrûëöÈ¸”;v‡¯ïD)s[Ï[÷`µ˜Ç×0gúöÂ¥ùµo}‘jš¬Ù:R{æýº…Hi(­à3/ôüÓ¤UKRàÁïˆÜ÷ À©rmBÚÿ¬¿ÒhZÜ†¬ë=Hei9ûpóÑ¼þœZÓªÐi
³þ¢&>SÊMº±Ž(´zûÑZZt;¸ÂMs¦‡4Û×›ÕRYj6{G©XÏÄ‚T-3Ì“b]Õ­îU³^¨Çâ%ýMŽÓKÂ­ÃQ3Œ-·49#Ó ê
ß¡¹í·é÷Í!ÜˆÄM¢¶âæ-å ã¼ïE¿G ÖžïÆ,å‡­sc)Áºå7æð›xèêkd‡o+j²aìëCÊµ®Èw¯^‘	¸}`žbª›*ªXKøÉá•ØXˆõ†Éÿ@6	*6t¦+Çfáv¢»4d=éG×„Ò!”6òFõ¢=¼Çµ©9‚èV´ììIàÚ,@¹R	kŒüôz
$Ù¹souA•èeYH„S;ãÅåCc¯ºäÂÓ ^¡¼´ ^Qn?õu®f_®ZŸç‘A‹ûu•÷¶æá•õin]ã!kêáuyÀákr6šä
¢¿.¿ntæ~óÆ~óÆV®ºê\á¸-ùdÇ¦eZ¼`ÙOå	qHKó“BŽHpŸé|äéî(¦‰ñÇ†y^/"Êiãáwäý:ÁQ°ÒÛ­,ìˆržã’{Ž/ôCv·jçÒ¸8‡ë«Ãuën®ÏP‚‚wÑQœ,|ªÚâ¸ôÛ4–.»y©[{ÏžVo1»ù9RÐ†ÛÕ„TJ¾ù\å¬šûüÐ^€h%½Álro…@À~|-þ×‡³ïgô©ÂGL˜7¡Å¡Jlò)õîÈXúÖ ø•ÎeÿKj¾?Ïlaá8[¸yIMq¯*ŒÝÍªÛ·ˆ*¿ðÚ¸p•ÄÚx5ÜyÝkê¶=ŸòS2[N- ÙçñÌEH´äVÆ¦Û`w”nXw•£¿‹f½è,¯ÆëïEOÄ:—wx	?8cBmÆõPÞ­GP:ðÇz¦F%›±”×aö*†÷Û“Ä!nŠpCH|ÿ%†nYošI:c°ëS®òÚ£¹z=õ—úÄlƒò"€(u–Õ2O ='ŠnÇÑUgã±5‘¤¢ÛI·EkÉÒ+–·cÐ5Ó9È=¿¼z<öƒ‹4ÍEß-ú0l5=>¬5Àá$ŸE°@»FÊüòùh–“_’|æptƒOSsc>ãŽMåÇ¤7•¿2ãÊ
x¥ûx¯´Æt¬ÞP8^cë¹j{9“ûÓƒ©å¢[€Û^©-Ìƒµ‰Äöµe³rÎñ§9 lŠ¡¿—å™<`À@_³PBš!Š¿<Â¹œóíK¶2xùxjº•hüàò¼t=rõè_Eµd£½XDÂÜEðXbO…Vƒ…‹ý¥üÁ1D;¸ð-eàª´m†ÖÌéå;U|¥Szš'ïrÓÇ^ýdAèÄRO‹Ó°XäºüÏNV ×:Ÿ“¬"6úVãï7­(Õ$Áõõ%K'¿N+LDöZ¦z¦¼ç5üÄ|—kDWrÛ<V†°·b,L1wŸ2f*È RœS3¦{ ·…\¢ºa8¢¢iøƒ98Q”óe67ª&Á¢²è3À´uR¤;½0úÜqBø­Mº7Á…ðÕþ³P9xè:9 R›
)¶0Ë§\²Œ=ë’_¸¤åñõýôÊ¡€Õ%€›Ï¨ŽëVJÚ«vÔŠØ*ý¾kŒ}>QîVyH^qÂz0•ëQC`ÜYÓDåšw–Ä.Åfb~îEsBÁš¾…º:É|s…µpjüuKeŒÔŠc©I9#ÂÀ‰®ÆÉc²70Cö(}Lßž<&§T¤¤Ãø1ù°wô˜Ã ôz=7©J GÝçdšÁ®kÕº‹ËTYJ3×9†MÍJ´ÕÙÑ1ôÏKÈ«|-V‡¬=:ËÓÑœ¾ý(>ç^˜YJÅÄzŸt™šÃÃköÝø•'9/DvOÄÄ…Þ!”ˆÐñŒ:D&Þæô‘ÑªµÏ’OžÅ/crhò^yN+ªït]÷Âx5´ª(GÇÎ1rÛïaq
e‹ƒd½¢£’pÃ	¹„ö3î£M)K ­¼LNY~Ó1K&g_!“x
ì"¦x‘×á²U˜)ã!r8Ñ_$Ã!}Ñz;$2ÈèÝM¯¡zo3ÝôW[þ†±Ep¢OÝqß(çVfÈ¼µøa¶‘¹&æ&HÀ‡¨ÎÇ
éBúÁÊ#£æ¶B:n¸$‡Š‹üÞs —µˆÇîÀc«þ"Í‘Rœç×>¶”av¯~Œóh”ûÀrà²eÀõË¬XS#Œî33ƒKâ.°4oÇêòŠžèª{Ñ}ò´`X×]hTAè}	åS×ò
!¢³Ó*hÎ4Ý©Ô˜ôDKÇ ŸhcOk>gƒŽÉÐOŽDË%~¬	åäƒ[+fÞàèdŸ€Î“ËZç®ì±ÈßyÀ­u>¡¡HGG“y4]C¶ ´Ž‚Ž"¢E;ïÕCM
›Çe¶bØ~ ãëL‘Î:T]Hr•GûŠ”œK@FÓ^2ôi€ÅXSfÆÀzÓ±XßC‘Ò›Áœ2™Ò;í±ï‡"›ó Þ±.pY£n$ð#_Ì.Ûóî£xë[/Ã—W/†F0`~UR4 é|¡PQßèÝ­èÑõö<ÞìâÒ¨éÇjæD×Jl€§ ÇÆî†3Àe7Ü«ÜÝ$ãdÒ½ìúC6öA­M1eiAÀØÖÌGb¥d(Œ‹Àýå,‹£|ž]ƒ›ãiöLvzÅâˆ,&Ñ¯äÑ×k5‚5ö(Þ×ÒZX¸N³ëê[cÏ©xãàEF3¸n4îº„Ý`Är+Øæ}àÍëÏ~ÖdñÙ›íè¯ø°–^ckK<•:ÚÚ—(ßª[/>O6/ëÏ}f*ö¹çƒÈ—8{‚R¤´ÛCÚDÖ”Ä§ª­)cBÍy÷îÞúÁŽìj l9`Û÷óne·ª‰hø¢òÿí?þoÓŸ‚¿9éä<0iåYGë{«‹LœŽG!Úµ…ŸºàSÐ|ŠŸØ;B.•zµE–ÎÁÃ8`@±ÃaçyèVU¦=•{–Í'PRÈàöø¼v´9†ðþÀ…@/1Æ› ÝtÉ+Ì}$¯Â’ÁEDõ+îQ³yú!K§Ñ'æwíxCú…°šðÃycÅU2±® ±ñ5“ÎÕ”[&c2X¦¾V+Â+–”ÝQ\JP0¯t+è”±geEC…ì°¶¿Ê©õEVF«Â›áÚ€AáQCü(<¨lñÊk‘ÜïŒMR££\’ÅÝH&…8D·>°4‹[ÍÅžPÌã`Gé ¹“sÂ¡òÝ­t©vórû„»lªÇC†?ÿ¦oÈñ¯TuéÑ$aƒÕc^MPø{ä³(›Õ^nzÉÚÝ»bMÈ‰âãóµÐgÂ¿Ú«"ê×î»"<X‘_Éì(‹ìG“ÏÊË)ùóÕ°»ÐÐHG(ð&?+âg"ßã(•§~DY,	v–’x’Ï3hnžAôiBu’;Iò¿ƒ\ƒ…S-ÕöêõuAÃ?R5\Å¡Ñj»+^2¡[šd¸ðwæñè¼8ì/³		Zùè”çž‘uòë¨báä?SuþÎò-Ï;`!ªy$ò]	ÛIžèÂéyÉb$éäd~6Nf²ë<ÕÆßg¿N¡±nº¾u”k”¾Z´(wÖl ›”Û½Âî×âMo0À ÂÄd{ Oi^ë€ïöWxÊ1¾ÀÂƒ€¥ùÉÌÇ‚%$–ƒ¶ˆ[Ì ÆÊ~0.=Z`Ôž^UF›ÍBþmždž·’E^Ýì§é(Ž&&<,]$èòÔy·©ËÄùÓ°¶?õÞÍ?ÇcH¡úÇù8òðoÝ.¬·›Ð”µRòã#óúó«ãWÜ Ú»ÂÐ…7—©µã”˜èþ«nsÊ¡‚ùÜŸq
?¹P}™}Î¿ñ>}ÞŸ¼¯%¶¶‡ñ+Z9›øõ¢ÌMöûÆ×ª×ƒçkê*c/˜÷[`5yxŸª? ß³V{ðµ)/î“‡^½Ð…]ú}ñS\¦ÀÁ”ñl©”.È¿+\ylüî—#„ññ
'/NªoÐœ¼]ÿžxº§$èw@’‡çù 15Â"D%%¾Ñá"t¸`5îïCýÅŽÔöú°wDxâÃ7ÍÁò
ß4‡ÚÂ¦[ã¿§™ V>¬²?–gèwÊŽ5&#ºüW©1´FŠ¢Ic
ü ®À
¤¿Ç? ©³F†ÿ^!rè!äW®!RËÞ¯J_øéäTép¹h„BÌý¦/<h}áznÊ=]ï­|•þr¾Z÷ßnô÷ö^¿Ùìom¿Ýüï(ƒoÜU\øv!»áÍÕ4ž“dã´ÑïÉÉü¬û[:\@	ý>¥µÏ!IaC–Ù0© TžU»îÐÏÔ’{Òøo4\„ÆýÓÁNú8y€LqÀw¡‘'ñYh!Mux_a‘@‹Ú¦ îC‚¬I%g¢ƒKÒ9¦Ü´ì´>¹AàY“rÚ¿Æ;¼¯´7§çgs8Cö`9Í¬lÛó&~ ÍÐf(t´¥™"!Ö/P¾ öÉGJîTRp¥¬S½Á]ü¸=üvƒt _N÷â‹Ü¶kÞÖM%I)Ç*oôÞY)¤%¨l£î–•­V¥¨†¡Îa´ÊÂ‘ÿð…_
vL&ô©º,v¿…Ïz­PZliËµdCNÔÞjÓ•¡Î5®b?–Òâˆí bh>ö gØ%{£Ëè:'ï§;hH>PFAwaâÂš6-ò–Ëi¥ÓØaV5Q®:"…7¦†±*¦«Íðt(ch™×°ÏF©¦† ÝnQ1·É(>¥Oò"¼aû 	Â ëd?¡Ì€‘Ç|JÞRËË?›åãG£JÜ8Ÿ¦d.©\¡#Ù¾aè› Ù‹/˜2.oZ½åýRý€„»]D9*Jpõõ
éšÇU·“ù˜÷µª-Içƒ¸ÓÉ“/Ü°#?/½ˆI´Çdã.ºR¸‚‰	yI˜J@@?wó¨«à!õ/ü<RÝ¶ZÉöHö-VìPBz¨j–ge¥Su"ÝäÿSŸb¼˜nÕÆ‰¡jNÃ¨½¯háÍÁ·TñxðQÖº­&{*4ê·fñK6Ö¤¬,²Á,º.¤ð5¼èUÕ®ðâ„è2‚W+¨W~è’c·_­@[”þ Š[”zi(gØ²ôõù`ÖG|±! v[ÏˆÃÈò2awl²RU²)b•CÄh%U—æèÙ
K³xK|ÄÄàÓ*ÊÌQÚÃ¦í¡lu1CŒEÊ/¼]Ê/avæêÌ¥¬ºwôLÚ“êP¡¡W~Õ}´¿u76Å:Ó“¹{?£ñi;Ô¥fÁTÕq•“üB0{di±ë0<¹ËÃðOiö™J¢|%¼÷Ð:ù0Ï¦i?Pb‡¹¶Fîj°¶þˆÊ“8»&téÒ9°˜“4‘KºÈ(QòøûwIüºsú!È€	*S€_ù,ž¾ZÙlxŠø›·vŽ´áÚ:ItÑ÷ÅµÐ‘)hc©‡ðãkèÀmâÞçøšù#Vß€¦ºŠiõ ¯¸7Í(ÿëø<šf!P3ª¹ØlüÏq{‚Ÿ¹+îÃ,I¼«íŽ-®Ò>„šZ7P'ë}¡áÌØÏDÅÍDè©UÊcK¢Ärvrðu_<A¸MJ09*ÒWuØÀt¼™FüÂp–¹bÁ»ÖŒ€;zoÂ°OFú(p@ÊbÌ€·ù0šÛõÙË@ K¬hÛo„Ùýè æÞ4ö¤ T²&˜d­9hj¸"¢aÎ§¢ìwQ˜"*ç¥=®bã'ØU™%³}Íƒû%|ªq5B "ˆ%	1è7H·;k¤’ž‹þ´¥ µÏ÷n½ì¢0ù­éœtx3íçµfÓe\Ô§Æãó/I2¼b*æÜ£ÑÕ¶zþÀÔIvs^ õ™:×Ë#jÂ27q<ÐCÙD«!Ž•Ç¶/Ä¡¹HëëÛå$ƒ~Ë	sÛáÜÀPAþ’¹zÙ¥šçn(‘Èf(ê‘*8È2º ª@šþ"Ý‹ëVÓyÓYÔ‹-<s¨yµ‚PVÍž þêO7Ês
eUH™Êƒ½ñ9Ç¤CÐgCÖi’ ¹bñ¢0»"Bn&é"»¨†„OYÙÇãå¥úÙ3²§<^¶GµÕäÃNïãÖêt àVVé2@ —è›èFÆØ=ûB’(ü•p¡BæWòE ÷#‰Xf®@ÂêèÚ–üæœb°Û`@¹Ò¬a©÷JŒÇñ˜îW‰CÁÔÌÂ[m„üL§~	¡]ük
â/¦Ñèå4‹ò‹;'´¼@Þè¯´³žªæ[D…lIÎ+¤rÏ•ºÖÜ°‹Šž)ÏòrŠl-s~' éÃ¨¯èí«Ó›=)¯»Mc/ƒ|í¶l\Œ×«Þ÷m«Ô÷M7dÊÄ1¤'Q)Ñæ\iêž%Hy;Æ)•µŒS&9·Ò¯ãð]$9ïÞ†ìÉfjWÔ!&,^ÙeñŠj)$]@"±ˆÀEgÀáƒ²òƒî\ài:e‡„ÄÅÁÌCKîÿÎHœ0ÏîYz»=#î
òµ J^eÎh2ëê±V›¢_‘›…it?Tqg…D.³Wf%×äoke!±1ÞÊ¸/Šª¾³‡§ÑUãzóbˆ"r*‚ýlmà]¶ãg­9°ÙjIEHî%ªY“ÊÞ_ÙåËF×éÂ‘ÿJ6ÿ¡…â²R!ÕÊîÞtJ™Qt6Š¢ÂpkãñÆÆF+àâ*‘àÛ‰s^¦ç-Cö¸R9ò·ó†=o 5púú„ž¸þ¸þã'K8sžœâ\Ó×¹ÊKõ§,þ’Ä—ä{¦¶p«i?ÉLc½±ÓÀv°U¶] ò >\Ø„¹ét™+¼¹í,‘WT±Ìòøí(fõ*Hpo¸SIdÏ]²ßn!0Á«Ù%ô€m ‡¢¿<| ê?AD'°7†±Äô~$š±Ãý‘tªïw~ˆ3Ø+öþ›kkT%¦˜!¨ç~Êgò±0sÔSéúCûÁÄ³ãÙ§œ>U{}Øéª…èŠ©ýw¼Ù™ØÈÛ De4mE‹—¢VI1Þ¢R$‰Ñï.Y<™»ÞÜ7Öi¾úS–æ95ƒfAö¾¯c˜ít{¬v¸jÑÒfRU;^¸§Åúp¯³ùk-,pù!}]XEAÝz4êbOçÙtëë-|§ßÊŠ‹.:g$¿Ó5×›½‰E?Šgäøô§“\¯-wõ	>ÿ€Vã¾qI°ÛµŽÛ½ƒrp{%gyÜ
û0ÉA{8Kãxœž¤ñá%Gƒ¢ŸìÏ‡Ôî8œœ§=@<€ òÆò^êw%ùéõ”nÅ›«A½¦ŠîH‰;p3ð­ŠG¨îç‰VÍ’@¡U™Œöz;6.i‘¼¨a6•ó×žšQ)Ö3Éé4$30—üÃ»@V“fË–¾û‰¾ §ÈYð¬# šðç$ú%#f¡E#H¨ƒ¼/Ê:UQ7¾M³ƒy–QÚ`XYNì Æód;3 <6`=Ý”Yê~ëõ-‚äÔRž¼CöÅ"vÞEW;í¬ãÚGvDÈƒ_ÓU¿"ß½zE&À)è©àùÅ¢ÙqzÉ‚ìÜ'Ï?£lºÔùDÈctJ€—ëÐ6Íü­Ë–g`³xJ¶v8Š¯œœ2ÚN(ÃÆßÂjÆW-Áf`½o€%VékÇƒnéîd>R02!Çíí¿Ø¶b™Ðg‘ÁYdA3ñh [=¹ú‡“Áh»‘L}‘;ªõ]+ÔöÖËõ‹m§Ã2¨”+nãÔhR‡z©Ä“ ÊBÂËÌ»Âñ;jo(ƒ¶ÑcrÆœ8ùtBÕå&ÇÌZó@®ÿÔ;ð(|‹l^×Ã$±TWrÊ¯ówrrý–næÉ £çµ3Ë¼•pÞÓÊ SeËíJrBM‡0Âç8•½Š–g&B»²^Ý3qw
BãXys5…h/#°J„/üÉÜ¨ð(2HÇÔ|„0“:²ñ,å öÉù,´§7'ÿ»‘>B€’¸¡¿Oç7î†.ú¡_6íKê:4®$ûªQ™ôWÜ”°„“tà.ÖæŒ‰Ýz%'G‹Ïg¥¤[ñ¼A:ES/êïËÙE½†æ,+IsOÄ© ÜÈ÷Ç8éâ—JÀ-ù¼œTÂÍVvÿðr}vÑhL9†È\ÿ—Í­‚;x°:©-<üŠjYÑÂP¬Å´ìë#0r_ îÿh?Kda@i[ë¼Í×ùd~F*p?m¾ŠË±Ô’Îáé?œ¶ð¦µª9Þ¯UKk·²`ùßCˆñð¨ñhú	ºìn>]ÙÝcëƒÞ“ùŒ^òrv–¯õ‰±ŽÐÔ âhFQgÜ¬‰uŠ*$%x²*‡pt€Òwuk ¿Ì[…¤Ã2gnƒQ©FJ©]0íà)žÐ\Èð\ÀñYÍ]G½¥öMÝ6†dt¨Ç)—;ÕpºË%‘dŽL]tl’?O
ùCOŽV,4óJ:ë3‹ñP#ÌLÃD«à¹½	½ë6 ®DÖàõ"ÑŒ–*ø–¤S:AùÇâÁØâì:âbâÅµ¾Î²~É(MBR0ÇIž8œ¢È‘ø1áãp¯Ì+®'ç½ó„žÂ)L9!â´N‹fS(¾ Ê/}ÃN1~Ôi/þxƒŠb¼òçìy@Úð§b[^hL
¾jÃ$ŸŽ¢R»!™Ð›éÒ—ÕVøÇÚ—«ÝUì“‘·p0~9N²Vä‡sR”LQÒ7¬bHÑöˆ¡w×ìÑºfKã‰åcÀ)*ŒG9¢E‡õÈä<hôò´…ä‹Q¯:•R½’®Q÷b-V¿×Š8mDFb«
Ä&&üj†UËkÓµÔÍoTÁ<¡>EU5Ÿ”ÍWæÊ‹ãËšÚÊˆ*ãàdžqk(hX[¤¬¡ ÇžÇÔªžpÞTj™$ÿåQé£Û¿†á¿z^QZ­Y¦£“Î£jQðš?l	Cø	\˜â«Fƒ‡€ÌØ\íŠ·%ô´'jfÓuÅ¬²dI¥Ðíp½`^w#Œ &Ç;d»}ÙÀ*qJ$–Ñ9¹rËv@ÎÖÐT¬÷Ä Ó‘=ËÞ’‘îž=<åñøÆú
2‚vÈ m©ü°ŸNÄÓè7¾õ
yVàq¾ŸÊz`%#æÚ@6ÿQ³¸ÛƒiCxQ²²[PPD‚ÕƒSQ×yÁ{·#Å?ÝiM¡=ZßTio”ë¨0jî
?kÕß-Éæwûkš›
¨Ìª ZÁß´†ßRÅa/å¼¼>œ@:Eœ–!¼®´Î¢ÉNé¡@u°î‡¼‰Žõpœ^â$ñZ‚tÀÄk¿¸_¢àY¼ÍÒqk„!“Ç«˜­‘…Ào`Ó_i´…ÞNKðøC*pyÝ;·ÞŒsñÅeÎÓtæ¹ ý¹*LÁk8^‘vrnâ#‡§žZó>*ÂŸZ>bá5i[¸†‘\‰¦\9M¯œj˜LDÆ6ª-$PeëÕÍö­C½-Òò‚•Y‘RÖ½Su¼TìU]°ÆŠ¸<©÷©p›<HÅfê.e%ÜQ«HÙÎ‚ÜûZ–ßt™vçBwûÂòt÷owï\âúB)ôA>yJo¼®†¹ÝnOD“ç©ét)æZFâ1¥ªq{w#®3…–§è?q”Š€­Àä™¥Žä ôøðm^NwßšN:§Gà:dã–Ž¨Û]1äï¯¨j-ø$š²ŽœRgsã1éoü@DõËš{6î­¼µ¦&®ÿ‘¥€vEÊ$Ÿ~—£‘£Šþ&±åÛºkfwrÅ2ÎóxÖÝ ÿÎ@œ>íiYß>ÊQ†Y:;3ƒ\W(§ÏÇ;@tå:}—ùM’1zãÃIPóq‘	Æ—ÝgW#nü//¶¿\ü¥@ûÔÜ/2/áŸå)ª$Î‹„Rì{ZÃ¨òðIØ8¾C?Çü>°ËâJ¥µ»èì^ÂòL»³81î­DÉúµmZ2±¥jÒüø¨

{u‡}$ÍÄÞe}Sç@8ø¸_s–hŒÃJÜ®lNØvùk…'Úh[5ß¨±ôèñÌàz«i¶®ho«;ùÛü'9Éç<s“~éÜH1±…|Åª½\¸þ´Š(§¡ó @viÓUÖ?]P^Ò{b,0	iË/kŽX@ä¢ ‹›ÑP’u0º*FD_ƒ¢CM“½«èƒ1XFÇñùy¤uKþ*C½ïèz^¬ÿ9Ž²þköüÍ¸(Ý#â,I‡¨ð-Ž¯¥"ê6.ÞÃêY©¹9F9®ýªôFgÙÊ½È?éWÅÓ?x›¾kgÍ;Ôi‰:8ï’‰©üW*+Æ`ÇùÍU2Ó‹xðøè§_/9’ŽN,5˜ìpßµì{2Js„³µåÿlP<zn»[ìJr2£ìšìG”S­¼NGL¨ýèr&\•<a1½ÔòHxÄöÒÑ”y¬#1ü>H\^`ÃÜn-Id“`Mæ¾+cÅë-œF_, &¼ìN¯ô*QóÈ>ª=sàÔÁ]6q%Lí!<Ý)Ô*ÞÀÇ7â©Þ×”x–‹œ¤—K¤[”êQ	î&.ÂfúêyeOË…ëëÏ74ëU5cèŽ« <aÈû§Ét‡pxq•ÝAU@M®E}¹3Ô9‡AI&Ê»ê]'4H©W}áH$t™¢dÒÀ¿&zV)oÇ~B(Øblh8].Ä¶5Ý-04·â†ÄM#
¹ùís³ÉÏ(É\SÚ˜2ïð&><[-N¥}IojsDW|[S«G~Êê~JÏhRîàe·¿]íP<3¬LÿzOê¯×¼ŒÜÿXc!}óç•·mR]Ž~é~¿þÒz½¹è´°ŸÅÑçùt	¯^Â(Ï£qYzÃµ¯	ðúf4©VGï…i‚«Ù­O+×´—_L²'¶"`?òrw4©ãÂÅ¨ºw>ÿÖjßáj«þ®eÕÀc·]_ŒÜ¬¾ø};EÂË(¾£á Õn·8õàÀ¢_+tÁGKUG×=/ÌèãU¤hŸˆZKÞ©pµ¥ XŸœ	ÄFÂÕõý›cÀUÁ=©ý2M¶f‰©–¡1#VBÀ3ÃÊ†s¬ 9lke~G)‘2ÍR&Ãðy.§@ÝCiub/jmê-è~=¾ïuk¯¸Œ£Öb	½é©¦2zNòÌ¾=9P5GÕ³§/¾úÎòúFq—¥ö--ÂBï»ä¢{6i[ÃúJ÷Éàbö%Þ«ùš‹ïUó9kšÚjÏðÊPErzd-¤J®ô4Tûôg¥öéšpmøThý­ Zê—^y¯à¢lí·ØþÛÏ66ZYŒ¥5ã®¼Wã5	–-òZ¨-}iåI«™/úÕsy¥Ô¨Æ$Ú°;}iÞMW®Á7zVðE£‡4P”1 åƒ×ƒÍðN¶£™ìn•Á&Ö€P²ÅËR€3äÜQxx’hP?xARÖíð”·gòÃj„ã—ž©kÀP‚È!À5š<²ç¸_>°]•¼Mq7ØÌîëü[«·žl,ÏŽi‘£¥…±cs+W­ãsÀŒÑa´¡‡í`/Ø
‹¢S)Êkp`æµÇ5‚ùæÎá¸p5Åv€«’>‹E@-.LX³¸¼¡ƒÞ†
6¹jýÁÚ8Hè’2–‰*å`=Ó;ÓÈ¹Â{œØ÷õ®_:2
+SäI3ìì†¬ã˜àR #CŒÅ³Öé€Mô­p¸ö«û>š-C¯ÀUËro Àbh/ÞŒ&›Á²ÀL>íÁ³°áÐÓ=Ë¸l ¸P|ÔÛ"é‡na7¢à[à
ƒä¾7—ÒÃïÊ¥üÖwçRzð]CºÀÕ ÌÞ"´ËvÜ	ï²å†wid%ý”AÎkà›çbºwéMÜÀ/
WáZL{¯î„ƒ)v2 Æõ44LkëÜP{£¶ß¿”àow9,¸‹ƒÆ4|ù<¦º¥<Å{Õ0xý q›l†ß„K¾çÈoéc#[\ë•ƒ[J&ÔhüM|q ÑCXÃ±Ñ’¹´}r‘^Bw¨Pöá ËX%pÂY2ÁOEß>ïbì1¬-f=-£X?tWÝmV@ªŽÞsÝH*'Õk}…ã¶–Œ¾÷öo>«ÖË»½{ˆ±g¼ŒçÖã7ÆïT7àíÞówO&¯ŸDÔ¬|Eç3²NNXøµ†ƒ‹ƒíƒ=G GúŠä¥s€=<,€¸D½Y#ê~©­ºæÆÀ-]Ûæn'öt§Fq”±­´-¡¢SÛãÿø–E9@™°]Õ!§[—(ïZQË^$ùjÇ?zjfGˆm“üg<¬õ®·ÂkØ_ßÉ&Øñ°à.@™ÿÌð]ŠžçQÎÏ;Ksfã«O¼ZD3ó»"ŠœñªY<!Œ§ah DõöÂÏ¾
âˆ'Ã¦ÄÁY=ùžPµý3ƒ„Ztï2Ëú•]fÄ:f–Ì‡hS{¦˜+È^;ºNùá'ª»Ð•ì<{ ·É®*·}Ú¡·wá»~	¸;Äµ£,€e®­èÔ"1ÏÈôŒªÌ‰Y’ØD‹c½¼Ø®)ÖÒ~4¹@¿+ü2ôûÁghTt™c¡»‘•^¾MFñ)Þ@ÑeÓõ©ù¸ÄPPLçæ<Ñ)ÄCö·ÔçíU5/×/¶Ëäª·Ty¬ì¾ýó¹ÉãQ< ëíŸ=|Ö©ýµ¬qæšQµâßæñd¿žOéTN°Õòõu/sWUÐ®Ù‡†rmq
Þ[Ë³¬šµÎ~–)Xoën‹U:Mû
Jì¡‹˜òOð+fqW."ÊUäH§æÜÍÉà"K'é(ýD¿®tã1ÙÜØ„ÿô{½ÞZ¸Føò8>§Ûwqp>ä‚àX›6wBç>êtÉ§b:‰ž|6[:ù–¥|èBÛ_€Ì™"]™
K${c*¢Â2?‘ºø÷{ù,šÍsV[¦i%.÷ç Í²˜Ybµê+ÄHƒhÂ‚Ö¯È<³ãtÄkW£á8™¬B¡üù0NÙ§ßi¯â¬ãC…`«<Ö3Ë³÷’¡Ï¹ WüM© ÐQŒù¹¯ÛGþX˜¦ÿ(Ë:a.?Ô–9
àé¤Æ!ûŸ$S.NÜÚXE<aG'« )ÎWpÝõçèg$ôPnBmñ§ª:Ê'ÈÇš¦Î°÷–¾Ao?ºÉÉç]S`ü&“óÔëµ¨[fõêlÆõ2>ÃÁK.‚Eë€JÃ§\™s+j`5Œ…#0ƒô”p´Q,°<>™åFgÑ%/=°¢¤ÈD¨DmÈ€dÖVÂòò˜mWÒ
ú,‘£Ï9HÃC¦1ôó4#ÇÑäSL Ùy‡ùÛÉj»‚çüç?QîK_¸‹ý‹îý÷ÞJ`é¿ý§Q{\’³'Áƒ“ä©$ÈÕ3ƒóëÛ£†Š[Ê¦µ*ðó|E¨¢7ÕôÏ®Ê·‰¿÷¯YìÏZ½Õ?Þcè¿ õèfB7à5ÕW;ú×ET¾Ò"{?í¯®Ý~/“‹Åtæƒ$Œªè¬eòÔßçç?…mHxR8}qÓßþõhN÷A†ÜÈv ö^±÷òEJ;Ï?´ì»ürY»þæ:¶lv9”m½þ:w²õw¾Ñ\iR;ÍÿDlõŠäœg›°@Twù·9U3(#Ý9 <‹É”>– è­œƒ5Ðï&}0²Ëý-ä&Öxü²÷}oèŽ)Æ¹Â:yýæý9JÉª|Cþz«d}úŽË7Å»˜X‹†ôWGt·þ¸þM$i"‰-èW*•¾É#·<b4ÿ»K¥·ú=Ê¥é<›Ž´-c$08å ãžOhy#„ÎeÄÜ/€2J®Óy&Ø÷„Ÿí×¦ôº‡“³ôÊ¹}búraŠçÊÒx8òkâ¼D¦ß#¼)ù C,åå)@TpÇÃœ¹°Øäv»eï-ÏXXdAo>®(0R®ìíƒ®ÇÎÒŽ^±Èâï"çX†OÄý–`Ín	c±µ)ñD£Î®”Â§kìy"ýà^÷oN“x%Cï˜[ª…ãð‚£U”[äÑÜ¬õ-"†™î²ãB+‰«"wCÎu¹—í‘áÍ‚ö~ÆœuµQîk·ðÍ“Ð q%üõ^J ‹@O¤K­ù‘ðôrº³Îjwu­G¥\œåqg­÷¯i2é¬®¯®q]…ü•°6JìçcÙP©Åh_$‘ãÚµ™~îêTúËáÍEíc^ŽtÄËdXLù›Ö
Ý¥3BEXb2\ÀBµtdFbÓz~2%©‘£ÜÉ1zG›4„É99¨ƒJÉ	•×Tûžœ¼Ó4æRA<Ó"IÐÔB.Z!a #e‚QC^NweÛQþN;E£QÐØg€œ™¹,»>Ê!ÇÑõÆ`kÄAÌ¿/=}Qy º™ï'oøMËÑQmÀÓf÷Éi¾£ÍÑQ*÷-	¼j"“Ì4
EVóQ¼^ êjg0õõ¦}&n°²Â”±¦}±Ù’ï³(Ëþ9MÙ³to>»H³dÆMæ×o©9õ¯Ÿ¢ìEIø£©<5`½¥LN¸B»M0ÝIiþ*UUÏ´8‹¢ƒ¨F æ­Þ2ä	UÛ#B¡¦Eù®G^]õlîh;Çá—2çÆÕ4ØZmžT'€Ž!†T¸šWo wÐ²´ê	ÄiàNs´–Z¦å•û²¼Ø0ÙW5eoÕÀ“ÙrT´Î	µ»0åhˆ½²7¼–0ÍÐ#Çñ8Ê>çä5Çð…¦3Q†‰ï–ÍHÊXøßð—´)·Ù"S™.Là³)Lc+¯‘fõluÛŒüŒÊ‘nZ÷³<L:xùëd˜Ú #J“|Ö„OIN¥SÕ,ÕK•vy ”O4Ï
L4ÑŸIùU[´êR…Ûlv×Ñ§Ùß¹W³Þ¤^B˜ñ5W	Ì·oaî…Ã¹Î+75ft‹hi¯÷1ŽknÞqÇš¤
Ç|Dòñµ68³Ì/¬ìÿØÅ—…Où~Ùr¨c»}®¼7Š³Ùi–D3ÕÅUü¾9s6ZçÆ@z˜‡Ô—Ì¨k3h›_‹·8»6Löï‡k¹08¦­"I<»¶Àœe˜…ËšmL»ç°ä`Su¶BöaÕÝ}œWlI	É!ðrá rÆBÓß’ø’šŽ‘gË^
Õ¤ÐŠàÉÃÉÇµÔðëHNü!ãÀ réñÏ`ã…xð swJKt&£4bUF^¿m˜X:J‰’~æ&#U.÷Z,NÒ%!Ãøð°‰¤$õ…ÉQDgºrÆ®žòúýÙg?eN8ª±ÊB®#¥“Ñ52þ¦V–äì5þ}ng©¹³÷Bÿl‚úé8VÿD¥ú\¹NÖ,ŽFå“Å>yð'«~¶ ÐÕÇ³CDIl0Ó9No˜ÖHhKä
“÷”²¬•§åm˜«¹C1m+§ž0QÄ«*í]§Fõôš/²vjaçá˜åÈÃzÊM©È»©Pªé7¦jocmSP0§ééêÓß¶)¸ªÛÎ ¢YŽã,%ƒ,fÐ*L8–ü^UWÍ@ÃrˆµòÎ¤TÃQg”wþAácF0½ú±«ÅABe¡\eB\œryVüúË/dïàÉþ¯§§ïvÈû£_þL¾$yE·¹¨žÈÉå5”ØþS:(²^A^½?%×ñŒð$áõ¨TÆ„áz%•Ä•ŠÆïª©áü³Ðø+[Ò»–Þœ½ŸÆ“ÂqÊÀ[âEG±¢ÐÈÊ‡'úÀògYìã¢Dz’ñxã9™ôÀÅT4Å¥J!ž~Ó©·-ûMQ ‚aFöY¢€dEA«Ì“H&C`)Zð{Žx‹'«ÙÑª5q(—‚´ÈÆÖJ‚ “póñ7Gm‰ž¨Ðn‘IafÑçvxc’Î*uv=¼Ð´*Qñ£"•Ædç‡Tºôª<o9Îwü™Ú{ýîðˆêF`›ìœ¾?:Ùá%7ßËô$ÂÏ2úˆµe7@Ã	0@^Ê¨‰½/K’œ ¯B—(†¿iQE`/Üâ£i{£º ËÖÊ^ +Àåq§yzÓÔ5öñƒÕØÚs]Ua©Hò**ØðcOÔ8–&ÿõ‚º1¿a/ïZÚ‡ÇµàÒ’'¥€­oú9ŸG#y ÏU1çraÇ¹a5ð=Ÿ\1ç W¿~Ÿ!µÈüˆJÒa§”ù’TþeNÕù‹X³:Aß y³-ðŒŒ%d{ö©v¤ÙB7:Ïòí1 ¤µBtP9í=ŸÎ@˜Ä<¤õÃó¤ÞÅªæªÓÀËM4LÞ»xÜþUˆHáÂ‘Ç©GØn€!q&.Ë“™:Çôez÷yØšxðÄøå<ï¿ýŸÿ}R7l»†Ff.à¥EªGM¯ã…p{15Ý–`ß)S9èV¸í`ysÛÝ|ÔyRÝ‹PÍ£svƒ°q$¤¼qåÅ‹¬‹<Ä
ÂüV¥¶Ý’Áˆ@8Ô¯í¡›õ‡ö7jˆB–´IìC}ðŠÅ%ü2%PÉŠ ž×Î¡ðK®Ð—Y9æØ¢Ò{ÉôœÊ±Õ}™+-;W\É‘Ó-ùXª«±ÊZEÒ+~ ï\ÜéÇâ/=j)>N6SÁ2l,&d[“oÃW7ß	ÜÜÐð'ç^fÞÕD½A3*1_4³àü©îº}VËÚ4û8¸Ñq#”Euy<{’Î`)ÒËxˆË|ë‘[
¯.¼´%m9Öû(…7<÷ïmÒµ×ÛãE¿XîBÌ»V=ð#Í§y÷‡:ð8/ùØê&™±ëpýÌÂ§íÙ­û<²†ÓjÛµîÆÁißä½ÅÑÐÞ£Ø1ªËWRÍÁ˜vŸ›ú|ªò!]q•³_à
U{êT‚¢jv,yfÏé~>'ã+Þ DÖî–¥"DìZ±²{”Š#<÷G”Mÿ¹Ôæ£ç,x9Ý¥³‘t>c½shÌFù2ü{ŸCaò)žPkŸ¥•|ƒ®‘]Ûjìc¾ßø±áÃÒ˜•ïu]“AAÏG£nÎúìñï"M‘e~ó®”º`º8…lbÆq½ó5ïÀ²
N¥WDûQ­Pšõæ{zY<Q²ë¬ÿÏã·ïÿGþÇõä1Ye0>âsúsÈ>2}óþü<¨¯J¤ç‰9
ä“X”ú$Ê"*óVKÇW5ŽF¬ïÞ:;Ô‘pô¨#ßXF<dµ‘ÂÍîá~‚6ÙÈÑè½îÁŽdëkä€:Pi¹¸¤ÜzŸÓ¿¨ñ!OYöpÞ©<}­ÔÃØûÂÐëBgÉeÉ$gÝòï±ñnŒŸ×Ÿ=aÎ©a–NAÍÎXŸ%k‰¡ø*ròñ4ÃP½N.’á0®iZ…N[:š$ç¤÷fQö	Z”¿zEâö}Ð)ûlúƒþV,w€j"æ¶ÌLj¬¢Þ4–d.9`ôo¹ìn_d'—ý/1µß‚VÞÙÖý§È¤t`Ñ&‹G¸‚„wÆ¤ËuMNÓ©`@§i::‹2Ò%?ó§%‚]uæê‹NÒî”Ý\	TM%X‡£ëlùål‹»XÿFÛæ¬÷«°Iö–cfÑª”]‘r5‰s5½ºØª	æB(Ã$Ë‘³®&]~c™-[Ö±ë…‰UVZÈžpŸï”d’ìaAþJ>ÄY’«_+`7«ÄöuÇ3/Uóno,`]”ÝVƒ×ôóò)°Ù‰­yÄõõ27Sr¹Å¿ÓŽ‘£–84jo,[{,¶¥&HcÏM1íºW~C5æuŽµÅ<R.(±e“µÔÕ Š/]ãO¨&Ë¦ öXÅó9CGoØÝ2÷È§ØWãÍ’S,ÚôÒSnJf†àoÛÔ¦¸}%¼ÁJ­½ÄD¥+$ºÐ‚Š¼¡XGÜõÞH	Gõ’‘ÍZ37YyjP1”ä—-mj4Tè(H€ àbúÕª7únæ[­‚ß\'êë¡[‚)UÎ“É(™P-[6ê&‹¥WqzÙ“¯ìÇ”pÄä(¿\;Âˆ7k]å‚%•Áî,[ç¢ôVQºÉÒ²ÓaL‰Ë9¦¤£•©»`Jÿ  ÿÿì}ývÛ8’ïÿûHOŸ¶¼cË‰¤“t>ŽÇ=>›N<¶»ïÎÍÉÙÐmqZ5¤Çãñ9ûû÷ÑöInUáƒ € Dçc&þ#‘D
…B¡êWU	%XKiäÌ¡Ð0â×W)(ÒÅòÂÈ>r\lÎ>BßîºÄä#kCg¼VLâ2BD02‚5ùwn‡'¼Õä¾ç¥GñNÓš´»¤’ý36ÖºÍƒ"Ipq.`Ç£4³e!c3)Ò¤Va—¬,w¨dúdøËme±¿g†¬Ê{r‹:Pßy4ÅÎ’5ƒÓ§‚N›ihyÄ ²³¦Wž“¡wK¬{6vœyh+ß†nßÑ8,2a˜é]ŒÍ’ÞA­w¿BLæ›zÊÔ¨t`MuKMK‹¯x¼'×Y‚þíÙÇw¡½Ç»¯)Šèå½ù>î¿9bû¿¾ÞcG/_½Ü=~êp8Ã“?¥U½œ'Žaš“Æ*7FÑC-\³gp³•–—ûX~ÆŒÑ¡Dða×T°}è¢³—M™Öü%X3U#ò¬é6(¼hC3ÎÂ³™ð»§Uô {ÀáHÂæób€Ä>Ì÷ŒûÉñ 3`.RbEýc· uÎ4RK/­UÒQm ³S¿%ÚôÆK¹[Š¡…Uëß#7Ê¹¶é<ä‚ûSÿ°Bàåñ	Û{y¸{tòËË×'’Ž¶}½¾Sëõ)öIöûãX¯€”uõ“«Õ]…ƒìç™_*6ØUåÿô±ÇèŽg,~¸*«`€`ªx»,º·—aöLÉEÒ&…µ0˜³ã„áäÎ‡[PyD”?WÅQ
3=FX¡ã¼c7ûd<ûv¥êˆE`,¾|[rÐóA‚âuÞoWnY5É§y­yŸº\áÃÁr~ž9ñM·q \ìêý5wìŸBÕwú²úÎRƒsÞqFâWpêâU“~`'ùü§Ü—pˆ£ÏE~,.‡=/:ð+^>Ý|T ÃóÍË}Pa‹ªø‡Gd9¨Ïh²¬¿;ê$âDT4’?ÉµnŸ€´•Ô»ÿB¦Õd—h-;ÞÀZ‹Ó¿¦÷‘ÍéÌ-Ç°¨r
a¼5pø];æÊ'ÖÆÀÅ(ôËéÂçæk]ôôÛ®÷ÚjÇÜ`á™"ÓÃ‹p7$ÿ*n¯+£†/ÈcÞ*ÒcÞXUâÁä½ï=W¯×E†óº(³ƒ¾Å¢8ì{×õëw2½„‡q|ÄsÆë`ÕMqh~ž}I}õ•äÔ#ˆa-&cHõÏ+-êˆ:.Éaÿßð»Äë*‰€vw¼ìr073ÍGÉœû'E9Ò¬¥%ú–&Éï”˜‹ôà9¹‰JZŠÏÒé`A!ÀkC¥7'¨æáò/.bô{É¶Ø8KN3Ø°\WM¥ ñ‚¤ðIJ6KùÃÃtždcèkrºÖð4ç>F¹9šþ%_°¤HèÃñ%+0¿YIÄÍY™LyˆfÀ‚†RÒvG)GfýfX+{¾(dÊ3XêÍ@pyë¶UêPÃ¯Ó9;:ùù˜qôÛ£z›°«¹Õuê~RÐT1ÅDèùˆ¿g‹0¸Ìxºy‚¦Ä3˜§I>L%ó ?ÌRŠS•ùÞŽð„WÌ[Í\x}Ê„r‡¢¯{‚YNÈà·Ô•if}K™;?s!¹8«b_ ½qóï@i=¬vÿnÄJ†ãñ¼°ö;„ý¡qlqê;ª/\q
Û`ýZFf‰º¯µy±ù \Ñg ·?Þšºjw’M7/Ðöu›Œ_¼zÛãÊ_ôÝÓàîÀÒ6ú<™þÎvZý‚O»$®çqáÀeÂMLÁ=> î°ßò¨ÅMh/.DøÞAÿ$ùÈþÈ~>>Y¿™÷…ëw¢Ðå´%™Ø¾ß==ÜÅTO0€ƒ×áÇájá7E%Åãù)ê;Ú«AðeCÐW˜øÐh-²,šF-Nºæ*Ëù”ÝnÈÕâma’Ìz=KÅåË†=P^ý#3d~šÚêg ÷ Å}mž—nH6oòAà-[„Í8'C?…ÃØ¶¶Ø.lN“džÁêÇ8ÏÑ²Æ¤¬âey>ñ~ËÎ`µ–%j¿Ù¬¤úóhíˆè ¶¾˜Ñö
½äûxÿŸÞg"Øc‚2ôØyœfýlÈ±Öú¢XS[³~Â…&§ºÖŠù;½eŠ›Bu~×ƒš®!¯$5"·|Å\ÿ†£QîYíâÚæZÃ<;ñöõ?Ô~O/Ÿ\ÁÚ¸6´×N7ª8Èãù°¨Ã†¾Jå< Ã,óNÃ ]ÉèéUÅ
×ñ…‡qå‚JÐF•9¾o{»Ì†DbªÞõÈ ¨Ö¤»—bUCE2Ç*ùbº\á+·9¨šTCþðåÏbn†zè(”+^—•MÄ¾©{
×ô¤Øõè®i.äh‡ÕËohÈ°Úû÷b™ILZ¹82I‘Úñª’ÖqÛÈ[ï^¶Î®«¿\+?l°ò@h`-ˆai‡Â‡-JÚ[æA›´³Ñ`mX(ø.Ú}úÂbÈ®Êý|@¯s´B¾ýÞøéú–G‹9á&n©ñó!64LËA‘ÍDz2š¿*³h<éb%þ­G'™Šl4>Oœ»&œ[,ƒõœÐÔIŽ¡çÙxÜ*mnóˆnN®Å­p¡Å·]º±Së`Tièð&l}Ð—‰¯Eª3ÏÖDÂÏŽfë+:ë ±P›¢îœ<¢=¯x».ÿñè
lÆ‹à ÍßWšh_L0z.?íâ#mÿu6f=ñíeÊ?ó¢]VT/¦ÒþyŒšrz»rŠf£Gr§†IÕ4Ï×[»-tÏ+eº¨ÏÎE_­sZBösžgB¦.þ<¬ð~MËêqÓ‹i&óñ1ÐüÉÕÃksVï§a÷&ºÙ<HDëY‘p!`*‹0eÂT	d7!ëŸß6x–çó¥|.@ª¬˜×Qpî¾w-úªLÿ\JTh2¿­=8™dÄ6ÃDOJme›}´%ÙÇrÙÎÆûod8fRô®vÍ œ®ë¬BÅž‘äz¥v‚®‡ZNj»ÞÊLáüíx„ zªÁŸá—Jp°í}ÄÂS“MEa¦úlh¢ €¨òê‹n&—äwîÉþ¯‡¡„5½ï@uÖªØÏËLG‹z×¯è½×Õ×):çy>ŸçvœO‚F‹<¡›¥r<O&³È‰šÍ7ïÜ7ÓŠ`ä%'#
ò-¼¡ÛÛA%Înk„îñpˆA«¢Â:ù„¹ …¡œXNŒ³‹$tžŒ‚&¶{Ãà6”{ð„Kç	óE±ÉœÏ£“ÝÔgS¥»þ»Ïá£
¦©§¦©
Œ[TK3QÆäžì6uöq“Xo›ù;ÚÔ¤¹ï^‘Éõì¾¨×erâ{97†ÊôhRWÃH»žÃ³XsØ®ºö8{)˜ž™àK˜¾eC|ÍiSàu5m*ìÂìá™³æ«Ëy!aÛšÞ+&f¬ý¤®×{¦¾Èür:à(kGÆr”_à5¼‰_¶"új®ÃZ¼m=ÞýÛKåÀÛ©åÀ»tAãrÏa^¹NÓÍñ<vîa»dšMP>K†éf6­íiHqolH=ù.³²AÙ™ä‚ÑÆLr‘ÑÄž:ºÒ3ÐVF/.Ãˆïzóêø]á¼oÀ22ºƒbï´dCO‰™_j cª!`Äø“ÂEö’ªrLé^aºYV–`ù_D!dËÍ1 ÅÙd–s	åC}cö
ã&A\ÂÇöòî%ã2uÊb›4<ÅÂ£XgÌ™j£ž–¡}.'áþÓÁ~NnóÉ[·ÎÌõåá%Û‚“YRF‘kUžv|Yäu“·E¬Ê@K¸YôÍ‘Ù*9-óñbÎ
c1‹|&KZ(›•gQfÓÙÂ³œ³6âÞ)î–iZ<ùNôíô’/Œ×ùl/Óùhã6DÄÍƒ5Iˆ±~ß“äëCšØ“+Dïñ&OÒbâV€½GxLSPÁ1žSÙOûÔ¬Gå¨#ÝgãÍ‡lVð42Uº×ö°ãV7¤À¢ãì`Q>B[¹EÔ—*ÝÄmÌAQpFœª”EµÞœ¼îO
ÂsW¿ÊJ÷1ïê,ÃêKI¯.áDŠ‚A"5]èsø6&?¯+Ñ&õéÂwügBwÂ$Lÿƒÿ™¿U>ð	†âz“¯øMTàeòATàF’6—ø0#Ú3 Oæ¤çÖÀ´ùˆ€M†´éOÏÙe¾(€å¹H(€ÖE–`Jý÷ÐÏI=ey6M¦”q$û^Ï_Ž•g=ÉüÛí`>÷“ÇãKàfKBéàÈ‚GÊæêp:Ê3ŒÊd¦øxFå¼ÜlMÌ´1÷Ô±¶'µcH`çéVÀÒìWíxà°­&6å‰Ao¬žF´Ó& Ý6iÒVó>¡Ñí·ê ×Jý”õn›¦»‚âÚ@û‡;…ë(‰
lIÔxy£Z$lTý¹f*)QiQÊ]‹H§QðÈêvŒû=†À‡˜J–ß{\Äßöœ©ô?žª†W¤;PŸ#ò&‚ýƒ;d;†=»^“€ži C8-R•üé~èáQöÁfÂ=ü£»¦ R¢úQx¢V€ÌÄpN¼žœÌ¶¬™âáüþ§š–Ól6ó$S·O\¡RµuàoƒØqØsÄJ²²Õ»vß;z™bL]×K}:àÒäÌñ&fPË dÎeÌ†ôT¿W§¬ âu$ÈM¤“Ñ8§ù©Þuô‹”²Ù›Å"§+¼ôÌÆKG¦å‰æ3èP×.xªÏy±˜PØr3áÛí»Ñ5ê;Ç_ÃgrƒÌ@‰gì}owëf‚°.Åc¥cAaÄ¸fHQ gÔ ºÔ"^!’9›-önh Tù2˜	
ÝèÏ:ÜÝNÌ&.¦¨^|1=Ëœ×T–Í}âŒ‡Ç'~˜ê]‘€>>bâ´Á¸æ,~à_ØõzCø¯.s‹ò=”é^°ÞØ1êÚËÚƒ½÷ºWäÅàþ÷†Š¡–’¼þ¾!¼‘±Ær¨"R…_ŸÀé{çÙµ“½ã5P®ðÿÿúùødmýí;ö„½}×ôFÞà()x¨®¦~Ð/´–h[„O½úf
¿¾}·Þ/óIÚKqÆÓ>ï|º—ja¾t£Š­W95òS¶sþÉ¦ñ3ÅkcÀ–‚xz¥¢Šù ª¯Mýß¾5 dS1ëöìög‹rÄ'8¶ªnåm†óGSSÚÊ¬¢ý{V£q°µ
³<ùœO’9¼†\'C±ÌXIbÿ'‘àE~g×‘ë‚9:x–/“ÁˆcØc¤[­ŸŽèíˆ&Ú…pÇ5h«*ÎÆmüÔZLP¸»¿Åå6™Š³~\E‹æh3¿H‹	ž+­rç=á8sþ×¸ð˜Ï^[âŸÁ”|ÝÅ–gUì+ÚÀ½/òQÉèBß©Åq¢˜rD_Ð9{-†\Œ6Üˆ»ï‰iG£» d;ù`H0ÞÂ>¬_¾£›-Gô¦iLwÀ¯Ù(‰á9ªß{¼Ò²DÛÔ%+IaÒœÉ–
r«Iß¸ÊÄ¸ eYÔ+X`¶BYò¾y©üO66&¤ž6œ-)¹é¼1Êñj¨Óž<4Õç]6¿µûøászÔ<Å^ó~žÏ# ;LEQ§Úy%ùøüÀ`î»Ê‡Ú9|k$Á‹ø¸-l øSCM\ýpCþX±ä «7@ªOX,í)`õa²Œ¾ñYŸ®FiS…Ò¨ú¤Ç>*8E¼!w—‹Ý¦SEJ?_ÅRòÝ»‹”¤mÏ<ØíÏQ¦t•"¥+—(uC™CõIƒ7—A¨aßë;ý¬I[‘”»ÄI	bb}K™†Ù?,'ÿ#£Hy7X´¼¾Á¢@JTH ·FN¤æáÞ~#xç­YÀÎÃ#Ë¯kQ%,Ñ@,GM¥w¡ÍæQŠÉ¯|2½…¿Úˆ¨•'‘Z]r;¨óÇ†
ÚÈYâ†
~5•ûòLO‹R_þ‰aüAUízKuy6ÿløä;-ófUÜë;ƒœUÍ/VµA¬UÍ¯*xÅ!®¼žš_í
Ë`pÙ]÷â(Ç²T­"N~&"ñÊÐ°&É`Ûça‘Órä+¦â*fCÐto¤¡¯ˆÀÿþ÷ÿ'.IðU¥iBÞ³+&…´P¨|ÊÏé4ªßQ•<j*näP(|›Î‹lÈð\©%?OT_ï
a¯R
è8€@hr0²”»¦TCÚÓÔºT–-òV)Û¾Uä8—/EW­ýÞ½i¯W.&LÀŸaû#SÙ`dV/î¤Ùûm»é=Ç)ZIØãìa!-ë
Žt0\ÆíÒ6H©³Nñs6!¬TÏ¢Y<¤72yL¾õn²­kÑøÞ¬îáÊ™ÖU)6¦u?‚[ô!*[ÒÏ­n&æüCCúè6mãöÅ÷i~Q$³ï¨vQg/øŽyGZÊTè6ÿë4›³-‘;`‹¿yÙaãšKÚ>„}2,ÆIÑñ¼ÍèH@á*Á^9Îo08„WµáÏÉÌ0þ8˜a¨&šyFï³˜|ÞÜ…9s¥ž®ÜgÊ	é÷:ˆ–ð>â;h±’¬ª	þÕp/o«†mò@›8²©©mØåžÀ?c½äóæOpà‰>ðê–ˆÑÃ&Ãü›N2~ÆA±-•ÐR™ËvJÚ^õnÁ5jÅÛLDzk•ÚÌ†ÍÙ­ë¶óÆ|q±	ê¶;JuÝæ}1L¤ hù1„Þªê¼uÙ³Ø‚£Ô.¥ÞV‹+.5Ý•ZZÑ	·+µ>Ÿ‚*¡ZˆÊL×)y¢HJ^[T‚I@×·àPbÿƒ¥îtÖM¨¯LhÇYRÛURêÎß—‘õª‡ÿÓJ0 NÑ[òh€rU;4ûèEKÃÒ,Á:n>¶­ór~âiîgkÕ"“ã©d5<Wª¦{8ž?Šž¬G¿gšÎe®ªÚKàšhÿÙ3ÖS”ß4;·ÞŒíâÎ1sHÑèŒÈÒ
ü/\›RÔé¬õoZZîè4¹Ñ)èEC•é-\ß«EÖá+œ~™<·IcZ¾ûöÎp²wüˆmRÏ†eC§b“í^‰…Û±Ù¢˜Ózß~>>aUÿx£]u‘ä CåB–ßW&¥Îóõ¸ÙôbNœ7ö‘\T1Ï1N†3e%º`ËøÊ««” ˆ²¥<Ôp7­B²²4²)7MèŽ¥¶X]‹$3ÚÆÿžÚ5UhW/éÀH:Y`§çÍÛzC–Zß ¯½ÕÜgËpîÚ¦Ìµº•ÆNSI2’ÔÞ[)IíŽfJÕÌÜÊÀÌ ÙW´¶vë	
ÌñÛ[ýCn"H0éhÐ¼ác%oºÑX+'†Çl‡ð#ZvJ¿a>!¡•`9,Ð:ÑóËGmR0èq¹÷ÑÔÂazÎvË2a0E£Ú^2OÈì‰Eu‹džûNÔ­_½ãü--²³ÆùEåEV®<hWJM·¯vÀ‘PÖÊR)€‘»²ð¡•¢r½%ÎI	B-“˜/¥†ÌBD8ÁOhÙ	UwÏÕTmÑUŽ|¿1ûýRÜãÁ³°c¥ãÿýïÿ'ƒ75oÄW'”ÈpsP$7 1ï&.\©áÛ&µ3Žpc|Y
Ö"®œË!ŒV¡ÎÑdŽùé ™h÷3Ó” ‚8wzÐO01¢‚¾ˆ¹oxöeíü©†Î=“{ªz*_£[²¾œ ôØ¿0Ýÿ7ìù7ì95àÎK÷ÉÁçöªÞÔ ‚¿uBWÁÀ2Aÿ÷T†.q	Gù³ÕrÑp,/»ãñÚ×Nç£þ†Iÿª1é|¿òüDgKÑÝûù7ºù£C1¨ö¾çœ4ÂÑ¿AÎ¿<È¹­/ìÉ|åAÎÛk ˆþ‰àçß¹µ>|	È\8Ûa£»Ã!,îð« ÃšÁ
õ6ù<™þ.¡ÂÕzìŽ°»¯ü§ÃFd°Z£U[$n¼«¿dÈçÍK&>4âÚ\§™÷0ƒ?ä[X® :SðT—ÏÙñ,( «m8}OàA†ÞlxÿÓX®ßßÂ›*±(…yÇ(ÓC.Ú.‘·±Û×YÁ'ˆŽæÒç&†UÕWÖÝàšÿ^?m«\—(®¡Âõ´õRuUMß­]Y2üFYô3€çá7àèQˆšh.½ ;+’œŸ«ê£Ü¢{ÎzÙŸêåU¥íN^îXWüMÅë˜µûéÐ:]À(D¢(¼e*~Ô l­¡ÂÄw”žÃþ†»]Œµ¯Ái—¥1U*£Õ"Ÿf þ/Ù‡,‘ç˜ç‹!ë_ÀsE>fÇ—ÐÙÉ'@h<õ$T$VS}Ò¥hâ9ÆÁÉ±Øï›ü©Á§t"¶R‰ÙŸœŠAçÞz9Ìd‚£ÞQŠY"	ÇÇü½—o¨\Öîp’MMtÇ•¾aÑÃüL€§¬¤z'ÈJhÓ!®ä‹ºÓ/êÆPuÕpâ¤¯­XlÖXõ©·VÏUj
¾•]Z“ár¾+Í÷`¹¯dÁ½lÊDÍ=ö÷<ŸàÿÃEA‘bNeÞ[~˜™7ç9ìrg´bÀ	óÍyšŒéó×æ±úè+°ñøù"#ÔuÛ´lh•üxi½Oä½Bž§u	‚p¼ó)šÊº©Ý÷3òµW‰ªÃ¼P-v°÷eUês%[¯yÇCÞuÎ¤•†^†îŒ–öomÕ|I%ÿŸÁþ
t?^œN²¹t'R\öùž2»ŒÅŽµYU¢J¨N,ŠU5D
¯5¥¬Ô¾šœ¨¡Ä÷ðÞíæ
¢>¡QU§¨LSþTÿÁºAUpä?Zµ”ç‹ÙGŽ¸â……‰üv¿iÅ­T¸ì-»°w‘:Wzª3„ü\>áØªØ·Nxz,LniÕ|JA¤ÎñQXwòôA¨£ö)ÜÏ.¾Ô1¼÷„[=Ø?¦ó†®Åì¼³ã7Üø‹Ö4ÓµyYf$@*e!òôÐq=V6J‘qršŽ±5Îe¥º4i¨½bîã[ö.ë‘˜õ»§ÿZ_ÔYç¥åË˜ŠŠ£s{—Œ/:Zß`£êŽõSÓþyŸýVÿÝÛX'õO‡'G "·ïnÞ¾}ÇÝõºoJÃ¦b6i¾ÓlÛ[Ž«h©Â ðÒ{X¦Ô¨JSÆ”ÿl½eŽÔÜÃ@Ü>>ê¸êÖB#«¯ÉˆU„º}ôn/yi¡QÊÓìo‹”ep™£½½`	©¢üÐ:e%Ï‰\@÷+vŸq>¿aŒ 7b NNò;ÍÆã²ïdé¼˜þíEªxšþ¸â5"|}N‘¥tÙvÚ¬7	»<eZÓœ¯¦_dÛZPÝˆÝyªWÅh—¤õ6
• ÞE¸Šh¢`²©P_° Øð½;
>FÞ””FæyçCôŒ±™µš¡+XÚx6þ“ä”4Úµc²Ár^C…T¿Á]zt­ýìîlVäXÊ˜[S×¼Ö5¬Aîí®­W/äÏ#
¶ÄW®%¢áµÀË÷24íÌ÷»V7¯…¨q”òé gdÒ4/ôBm¼JáÝœœ¶AÑiéB°’“{ÝÕ	-iN¼0>¢ÏE~Ÿ£Òîb¹kåæ4ÅYžî
p–ëˆkÚrþØšŽ¼>‡ðÈižÿÎ8‰Ýf»w .±bÒ¢èÞµDú:°‘%§e>bˆÙÙÜ­dØÖ6—1dØ¤c4+^òKºOÔg
[^å44?ÑñÓK¦e™Ù`|%Rq°·ÁF ‹à]—ý~?¨ÅŽ‰¾UÙðx%ö•õdœ«ï¥ãÍ‡lV¨h¾ì
%¤f.¥\r×5z§1(d¡s!EÝ8QN•}ò•½[¥ýSCŒ«÷ž-T„C„ªo‚ß_Õz€…áÍÚÔŽâãö1ëÄ¯)9fc<Ö\hçà}ˆGsC’è¹ÿ ç%ÖùPäÓ_gž÷Á˜ä=%÷Üå:JGýXÅxÃ{2À€Ý`N×¸ë\%\QÓŒ?4²éuNT›kAF…øÊOÓÌRðü°WMk=#/?‚6ÞÅœ¬Rä©Û¹¡!µ˜çI›GÑz†è=KÌ‘Óbd“@o„»Úÿ»:»¦à3ègÒ\e¿ÖÌ¾F¡Wýnrã7‡är˜m!ªÛ}—c³gôa“|“¨·¸¢—rkßAý±úºÃÆçÚ×ûBå£	UšmÉsóe’l»Ïa!C£ð¸Œ¸}U{ Ô…õZ‘mhBÝ!!–¶$&Š;|¹Yü*“T®(gF‹ëùˆ™ªÔ+#àVL\	Ÿ	‚Ï?Óy$›_Š+‹Ó]ó‡œžñÂ“œÁJÃ°!UbŽõ,ë#–›ó´ˆœ¤ß}÷´r <Þâ×¼þnoã7+T-%[•­\µ¼–V^?½B;¹h×Óâ³ß•»CÎ=¦™ïœu9CuË»’IkÌûÏÈ°|^b9VÜÝç¸8Î±~ø7Êvˆšu‰Ùž‹´Ø=È'³"¥°×9F0U¯­,œÔ,mrŽ×‘IDp
ßøÇ•HlÜýJ¤f;^‰bÅÙ+ñŸgõÑ £W¿{ÅÕ‡Ëá–KŠv½,C[Edöç,›ËžØ}â}µau‡²-Š”æ4:Ïk„X ÷h"ÂuSèM‚(Ú€€Êú‹¹¢	ä2Èßš¿’ñ¿‚x’K¹s%…F·"ªE5uáëK¢ßY´^P= …SÒ±pr] ÕôMj-/µÄ›ˆŽ¾7q‘ÅßT‰/×M«ÊG½$.åÚâE,¾	L>·8Ý¼1¡©	±ŽU;]:ÖDçW#)5âÇKËc5vM`–ŸLb:.È•öM˜~ÂTsæy—žô6C”ºoìX|›Ýãêoµæ¹üM¦ûd:6ëZ”ç]ÛÉòš‘ìkØo^FËé\5ûŒ¼¦Óm|þh7ÑK³yßÅJ§>g0GT4l‘¹XbIk¾ð˜G"¿$pA4¼ÈG/ö‘Õ\i¯nz‰ÕO*¡€pµyð-9Ò
êÙA8žŒÇØQ
£g‚âÆ
À!WÊ"ØäM“î³f¢0¤í´ò m~tçæ	9Ójžå`Ôáì£'¿aCPN  Š’É2ä`‹L§mñ.U™¡öže‚qhÎ9´(*7RöôV;y¸?kxyGÖÚé1W&¹\k5rï¥§Q[À>š(=Ä~FBw@LJÀmQëö<OÆˆ§î„š2Éw9Où;; hlø¾quMCSò³ö‰Ï\ oJ&OF³{Vî‰uñgM3c3w”lmªæ|´!UðÙ?I”&û!\“7Üù
ºd»tt©Œj„\Ü³Þ./ÄšUzcÛ$™Æz¢ŽDŠ…íkÈÆ›6Ì_ï×µÎõ"¶¨%Ów¿W©šh¨U+ÎÙ9åT,e1Ýþi¨Ý?£ÞöÉWÉöLÔ±å8¨uiƒyžçã4™Ê"«lâ Öy‘,¶ï(=©Èœ?ŽAö*3tÕW;|š¶ýóêã‚	Vëµ¬I8X²%•‰IV³iŸþ¬ŒHY«]³Jq™<SÆƒÆá|€}ULìe:àgÍêç`þLØ?š$£G© 	Ø oO5GoxF/à÷Ù#ññÝ_ýRŸ¶·p;g:èÈÔ’ø5áŸÞyX‰m>eÞh-–Žáä_ñ@Í 1ÓTkó_x¢–šžøå.+|Áe© Fe³yÕ ñîŸ¾µ‰’üð\5úna|N‘þ•@škë ·Šy¯—lœÒq]¥ŒM¨bôzÿ<Ÿd“ó›ÕÅSû¢g¤^&²òˆÍF‰\^â¨E[©¥0Sýý"9§œ%<e#½kóû+Ižë÷×¾ã2ÏÊt0…µ JÀ®"0;Ê/Yšµ7×T3ÒrïÝÞÚ¹­éi-“ÛŽKót³1g[½	³¹(Ýœü­¡¡À£‘±ýÐ™jÙµJˆw\ÉÍ»ù‰%†Ê“7Úù(ëL±DÓÌh¥¼ÕÀÛzþJ<Zé‹ÌuŠZáÝÝô]€²:nŠ‹¶]§.Ä.d+®ê2¸¬¯QLj!ãFÇ”)í6I‚…ÿ‹ª>]¥¢…FAªaÓ$Óù#Uý–Ih]³&¹Ã [¬üØ/gãlÞ[Û„-¥Há%%¦ç›áh½h•È_éya¬ ò5ÐðÈM)ù†6‚Šÿ]ÙÊ–?‡JÄ€ªÄAb7²–&	R®^’b“ûœ&sg«ë¯²Í_µÔ#Å$ÇŒ.ŽzU‘›wºÇ»*óÂ’ŒiçÅÌöÔ´Y…ù~ÆÞÿá{ëÇë÷¸ž6—]Píävc†×ƒìJ¾µßg"ßw3{½<¸_†û½•·L]Ôýw29g£MžVV«&ƒ2ÊQ?°R²Ý5
~[‚ïuÎ´Ææ #àß‚§rÙÏ¦@lÔ†þ’&EŸHLTt…ˆzLB‘ãyÊGã¼È·Ç"ÚXÈÛŸ¶þ±Ÿ‚õ½ë<[e]þùh÷õ;ys²û*X»»…Â·’ëg¥¬êYéåx|ï÷	oðÚBuæ²^%U	E•FVú%˜£dvçWù9OÈaf!I0SÉZ=;É0Í¿îí:~©J
=	½^Ù>åWhˆÆÙ+ÕÞ÷K2MÎSžfî~šRaBy*ý¢?bóô]?ÎSPëÕÕê2~ç§tÒK‹¹«¬g‚è÷~V¾³ASÞ3y²JšÌ+)¾þ}óí6VýœÕ’’x“«Ô«À'#·p9©%Öbh—M
¼JÊæŒ)Ú¶ÿ:ÇcòîLˆ"0+õÔ»jî8$©šËÆ'ó£’f¸?õæP])óÁ‹É×ûì©?Ï³ök<Ç‰†§Ž“ =Eú´¬×Šnc§—{4Î©ÒeiÁüî¸ÿZ4¿n³rôÀ½	¼ùçð©«ô"ŸžeÅÄ%…üÒ79d¿=Á¼©„Š¼Q)$§Ï¿íåÐmöøÒçóË–Eç­.BIÛÛÉ¤ÖDm’KV²>…°0Íjwn7dò^J­’À0Š`Ï`5½ðÿÚ¥–lKøe³N«;z—àÜïnÃý"6ŠÃ½}*T2Ú±I`ezÐ’ŸgãñáðlÉýáÇvûC­HÛÃªwá¬ô~‹»›‰ÿÚl%®¢Ž«´Ÿ»Å_z™RcßÕ+­Sƒ¯óG6Û÷1‰0ýó:×Ê¥únÚãžŽÐåï‡ð|ÕÙ÷ƒrÎÆÀËì	æƒÉÑ¶„x9æôµÄ¨Ç'û£"=ƒÇ-¢.Šqð)™Cž|´ÿ¯ïSK.Ù¿µëþlxöÞ×º'ÑÒzÈ‹Q6öð½á¡hþíÐlºH' —›nÞú¾ìjêá„uÞÃê±è”h˜q½ªºlvÉÂê­™³"¡Â¥Š%ÎÓ¹X1Ï/†½5dÜMàÓÍŒnÉ‘”ìO'¿¼:ØÇïâ^§!à‹¿öuú€j×žø?Ù–ŽüóV\}ØŽæú”^ÔÏø¾Ç¨Xjè±k6 ?Œõé"¶=Ï“ó¿"îîæ¿öLåé[._z¦MÓ¨ÕK­.¹t­­òbsöºsŸéP†f¿JÓÁòÿõèêùïé›SDÞÁ÷^}ë¨~3îåöÈÕLÖ²öÔÒ9ã-#Å}¤¾W©\õ¬äÖ6ïõ\U¯äØ5ëÙ”lSP»Rƒ'Wïk

Ìóõæy>>MŠ'·ß_»ÞPñÿ´9r§’”ç¤l:Å‹(W±Ì¥Ò]¿¥6GQ×\®ZŒÁÜ<é>æZÝÎ»VœdÊl9ø“Y_dÃ´tU6¥ºGª{+ƒÚ‰Ç;8ƒSã>âµ{D)VDþÀ~Kåªh8ŸÃóx<¯®K\Q)”ö7òrþZŒa;3ôé?ýT{i“-®Ë~6Œ@ï^ÚÏ,Ô´Ðj²¢œ¿DœÑþÛÛïê÷ÂxŸèH.H¶÷Ì@4;^La P"@ž9ã?TÔoÀØ¡:Ã<êM|Ð¦§Ž&ß_	ÍAv¿ŸÙy6½®_˜%ó¡žžQ;—8sO¾¿J§à_^ä“Y>Å#ÍÌÀ5b´N8A‚ñ†Ó6ƒÕGbª´¢uë½:2í¶µÎ›¦žÒ°"N…õß×Í?ÞÚZmªjRÛh.ÚFs‘Ú.Ýá{|1óUú§4Á·ºÀ¡óx9ÇüÎ®*ZøÕì4Ò²ÓÆ*ä²ìok*#Qò\·—Ô–+ñŠ?$â¢·ø rÏo;ZÍæ!ZpÜÀ’©Îµ’Wåtá³œñ?)íÙŽu:ü}ñš¶èª¥_½§¨i³t£GZ‹Ü­vÌï¯HjˆÐBV€ŒKÆÙßÓášUïÀÅ²ºlÔ2(x-9«G{y T"BÜ€ôÍà
X\Og°j)`UOÙpo}%Ë½ç‹†ƒ‡ûÖ¶|˜y„øó?ADîÔ²æ+-MKpð„ÕËŸÂ&é†ÓƒS¸+%­Hf£lÀ^$Å0JÈ[E¯OÍ¢×÷x9VQö¾‘:lIßÈ2Ý¡½XClFí ªÈ²Q÷³ø
ÂRÂÐ%±{Já±irîËC"¥®:>ø—Œç òäÄ Æ”¬eÂjï>7áŸi{x€Æ‡:çt ßD“T’yÖŽoï²¨©Ê!ØÖ›ÐK¥‘°ò‘V
é#ýàµp1þÐ”â(ôk{¹Ø÷‰
­p–¯ZŽ£¤”O’O@ÎFXU¨UŽ^¥âx ›ØE6“–3E4ÈÆÎžÌA»S%Z#U"OB
¥hð|Iç9GÍ{†NœúƒS[kA9‘}¸–Œž¾ìg)(?•‹gàtù”öK6¡9abvdK–Ì9š–>°P'œæÉt>¾ä§^~«ŒŽ°®ç ›_n°1Ê¾Éo@ÛCâ:T0øñr˜Áù$u–ìôî–a˜SåIçÉÄˆó,NŒ[9‰¶ÑîX}½[YdùBV¿ç·mËåí™ž€—4”›Å«Ø0
Ò‘ynv)ð¢Uf6!³*e Ø5cˆ¿&Î'Äõ®3üc-n¼^ó'5[&3ÎM’ø‰g?iKGGÍÓ•Ôç¥<Ì_:]EV˜niª*ƒË·øH,-FZ•z©Ïù¤;~ó’ÌŸzÜX¦Ñ®CŽÉö½¬ÀÜ”ÆùúÂø#çÃšAC¦Í;¾L›?zc†D—¹êSïù#'aü5Ù]ÖÁƒWÔô©ÈZQMß-ðÙ›)ìûžë"W¥eûŒÈ”(2bW•®3Ã®¬®OE5©ªJóŒ˜dñÓkb¶Óñ‡î˜cwíà=•hœ¾ÈgdÚÕ9¬º)NÉTóûJº˜E£MÛ§¿ò´éÓSÖ>8iyÍ$ƒ|–¥Ã?Ñê3ÍGæ‘FŽU¾¹[ó¼ÒU?SL÷_!aëë«]Í+kXXûV¿yèÑ+]S}L|±ÿo(ÉEûk/èû-"¶¼Hj×àòn/_9b=§6KvÂ“êÏ3-ØÒ}xêÔÖÅ—œý{—å-‹Âz×«LÞs™tÚÎý^ŠÕ¯e ò®eá.iclsƒ¯šK¼‹ó>z?|ýóJ«.çŸ`#|ò“á·ù¯þÚ wb§ŸÃ…ŽÇÙì` ßæ6‚‡P7dõ÷•viäÍª•k?ÓL†-x†¬|K1Ì+4qúþÓp<ø ÞÍu»NX%Vú|Ò‚;^^Ú¦ãHæ@²@%ê	¬Ñ¦êz½ÎÏšo¸9•÷ŠÓ™ŠôxP ŸË‰ð1EJãCÖÖ'Ìx÷¯F`=øl)|h oÀqO[°YÑ¶ç)ŠMä=ÃLMÿøüJ¶’êw™(ó¤hE³)cVgHŸ6ÒÇé"ª!}vPŸ—€Ÿí›üT‘]è)¥÷Fà~vÚá~îÆã~~ÆÛ¦hþ±üÂ…é&«›ð
ãçÍ	*ü+µlÓØûþœU¹÷¤#÷æS¡ÑÒŽZiS(ÇÂOlm±cX
lQŒÙßiq‰³"-¸žg0B@ìFY9Ïá°æ/ý"Á¢Hç˜Ä4 ÿçKš^pXŸºç<¼·wu½QÅ1cC¼ªó@å§Æº÷w%Ï¡J¸¹ëHx¾sÛ6ëCˆ/‡ñ¸°œNÃçgƒ&›‘r' f¶+°ræá¥èƒçj¨1ø^$tÌ[×;¹˜ªEŽÉ‚T“ÚÃÒ8WÂúÂŽ|f!Å9‚N×ò¼]÷†­¶D€yMvY©$é«œFãM‰f³5œ"¶—Ã©ÍÔ»„ŽÒ3˜òÑ‹“Ç‚ó ž›Vjå,›úÞ	ðóù´(ý "WÔ0åb’qÉÑ(gœ/qPÐ†pdž~¿ßŒ`×¡™¹åžš[Ô•ø9zà$¼hÅ×ÀÎñÿù…¹Ù”JV¦2¤z£8ñT7Qï§Ô?'E
¸]íäfÀÀ£»-P<Â—”@sd¯ó9ˆ&>¬Ó»«àJ1#j•¶¦‰F^Ì±8˜%‘Øa‚1ŽpZ;ØÛz½ÀŒì‡¿-òùOö	ñšÿÜg‡ º@Ê	$Î|”*¤N^0>U	:æà B/˜R«nàÍ2@UÏ+H’Hlîšô’IóMzï!-jO¾ãEÜ˜ˆõ‡eìNwZ^òç°3~l²«Žáoöã=³da é¬ÇÊ6Óæ!ÕšŒo! ×¾h`Mü¨½ ã3ÎõÙ¢úWÕkSÑŸÙ¤·Žûu>¼M…Â“™{BÕ ñÏ?ÁÞZž,B~HhÐç²ûæÄ/‘øˆ½âÁ¯j·WeÚl–Á½ÒÝ<®Wödi…{žÏSo:wK–]½§£QØçßÍp¯×<áÓ·®f¸“®R=Ä`osZGÕ}UÑ·‡>O| Òcµƒ—g*¶ÍlqÍÈ„0c 
*¡´vÉ¿Ðã#_y|<Bk²f ô%‚É}=OÌR¬äõÕ‹îP„êïTƒv5±8a…:ÖXÚÃƒî–ÇÂ*ÚK¡¡§MÑSoö÷^ì¾zõ¶ðz÷ÕÁÿ}¹‡A¾/ö^îgííîŸ°ç¯^±ÞÁkvxôæç£—ÇÇëà\ÔXÃÇãàEŠ†£°; Ç&ÜiHkæ³Ð@GKÜ-µ…)÷ðïóF¹ùC D¾7Ðr’â’ýGz‰ø÷"øK˜´CïðŠ¼Â¦ãßí0Ò%ð²7ñL×„•µxL‡qóãUx|ÐY‚Z'K)6Þ#(HšãAèý+ÄA~JòØ¬ýÚ_à|«ENË8gÓÔ­å«pª~Ñ4
Bp¾Îg"{Íª!pŠ~Ñô6ëX|:¢+ïrTÏiôESÂ;Ø–ÀItIÁØ(üsÅyxÂ¸¿^RÇ„ƒ8ßÖ‘[„ƒp‚Ýt”‚ô{‰%mô)?\´âæ@-¯/P¦ÖÏ¦¦òÃïˆQ~|óRÍ	o'ÅF!&Ñª±ËÑo¡i;-ƒÙ¯5b‚¾{j)ˆBNšËâC³€õ4e€èK-@”½Ê`YÆj®*ž§/¨_†¶%Ç™‹ˆõQT}êXïJGL‰BF×h±/²´\_rÆt hsl&þ‘÷èËÐœŒÐmÂæÜ}`ãrš¢‚ðìMÁróR~P@I4Ò Ð~â˜3„e±®ãË†íØc€—a5´:°ÌÝìfÉe*[¢â’3½%ºÚÐV\åHœ*ÉkF¢‡_×eŸZþ×<Ó÷h_dÒæˆlºy±ÉalÞQÐÛÅîoÌ UÏ0¸W»ärýÃ’ý‘Ýi Á7-«NVKêïãZ—›NE#YŒfÆz”2†²·(=¬èPjZ
èMŽ®EýO÷¶«lÎ2ÏƒÐÁ¢w5C{ÆÞ¢dxÄ¾¿šñBš„¸yÏK¯Ú+¹Â–3í§ªD5î&¯!²L¥µm³‰µÕÈ¿pËæý±"FÅºnºeÉ²”K«ÝÇé`Q å~?ÏçSÐ=Ù«ÁÃN0™E+r"‘ts{5k3ÞÁyVŠÙ&œ›CË@ª­¾R÷²s,"'•³´H§ƒ´½ÎgÅ<ìü§Ãý½Mô0l^‰ ¥~¹8-9«ÞÞ`ÖÅ#Ã‹¥ýòZýðiÝq«¸(ŽLO¯¦é™ÉzëÞ…zc§O3.Ä
¥]"S2¹)²2RL8-½ã& <Â	‹÷haµqˆhâ›ù—Qå„^t%ÃTO5 d Ó£ D"´&þ¡5n”Ìh·Õð/7)'ZÅ9®>åK¸Ñs‘Antï²@'
§ñé–Qq{8MÁ”Ov˜Ï3¥Ô¡¼î¦0$—•»cLGrù–eÁÒ/æ±ÖÓf?ý˜àœžÓÏ¢<Íðù%£±Øì³þ‚Î¼FcÏó|œ&Ó^\£êÖ¦=’Iq‰Çªø;½S ²¯¯û:G¥ry´[üKExœ¾h Û—ð‘‰ÇZSD>g’DvtÝÍ”’Mh‚s~÷k
ºÄ¶"™§·Æ“¹1´3)Q¨ðD?yqÙ]¬Þýå³rß»Ý}„Þ¶ŠÐ‰ÙUJîv5Ð*ã’ï í˜¸BfËáÖv¥xð¯?V ¯jòc(²ãeQhyD¡|!Ä\[m0'™~,m¾ÀG Œ­Y6a¶‹ü|ë?Ãz`É¨JðèM`Þq”žg%•y÷f[¢t\•YsîM êì¿ßÊú¸éì‚ù6Ÿ'ÓßåhV6=–ˆÑj´”]3ùdc,*„ºDB%âÑÞc8òåÓó§WJl¾rKWGF‘l£,Cü™õ>R²„I¤/ñ)X€øÃŸ½ø%ä%%çÅºë¾ƒ&óZo‚zI©“íô’)ÒØÛDEŠ °ñL¨r;àyLˆJ´Cj²ƒ=Ø¸»»áÅÃ:é±áIò{Š¯ÁH9©”/Ä¦†PŒãX<ùT#Hî[1]$bu&åå€1Bæ5òTØ1iËÙnÌ£êË–kc¶ÄËoP«•\Úì^:‡Å†>Õb5\t²O|³µTö¼‡x[ÐJe¤"‚O9·”¢©X4ËdÅüL¤2¥Ãäj0([æÝk¡’~É¿ÕÄMˆÿ|p&1ýLÓu°ü‚l+Í”Èj£âhš•2æ(XË$Œ|×\Óû¼°”Žç£ÝŒ¨XèÆY¹áyÑiÕ!Ñ3)Îýï32û*…¾T–áx'=}"¥¦þt#/Z(÷¬ðóròˆ>ùEƒÝÛCï®Ë¤ì¹91²`Š8 Ã¨.8ˆÅ0vÞH/Ž}T¢Ûcóáòéº^$ÓA:vNºß‰Q|q¢Ê¬o¥‰:ÉñÀg¸‹H¸ëœÈº÷¢k«ue›Ïjtò¶v™p}F‡pWyTûe÷aìÖœæ=ÀSþcaç<´ÜzÿŠ¥Ýü¿ùö+ô¡(çy>l?qŽüX7“íìˆYÑ
ñ"/Ðh¾WÊò!Q±ÜÑÛ{ù†Ì»ÃI6]7}hªzÜôIhýù·-L·‘Õ5ürcNµÚfŒóâÍ”vCEW	>Häõ}fá&Ò²9Q–)*q#…ÅÚ¢uf.°Ž-®±5ú¨'ð­Ö>þNÿ¼Î9î-|g¡éÖŽ¸1´;ËXígï¾Õ´gÕ÷«_ÜãòïW‹Šr~NzE§yóm0î3†ÚŽŠƒr›]„²¢_wªbO2Ð¥y]*´Žøð­Ýq@­
<Å±ª¸Ÿ`¸
‚!Iæ£>‘¡ÇMM½à£à‚×éüh~^¢™E?ÄGGÏÜöã¹œKÒMßšìlV~¢Ž¶Ó"ËÑX‹j‚Ø`ü¢ˆDÚuù‡|9/yðÍ3CÌvÍÄ¯«’Í/©þ¼ Y®”ÓCLe:Uáp§®Ëñ^j*õQš”XÓ˜÷·^¼È[º¨ç
Ä¸zëœÛ5ò÷I—'(&d-žd%¥S[Ûp?ô[¾Œ@hoñ<aÔÝø”Œ°Ä×Ãóô¹8Y¼Ï,f³¼à>;‘Ýu‹'"Äz1ðÆ.ù>€Öá%óä#ôôçãØÌ†> ‹1ÇÇ¦E‘®|ïxè¥È“¾°Ì1šB‚xÞâ;íOÚMiváØZ`–Ð'Ì·zEL¢æ®òõÊöD¶*raÄ÷Wæo×0ÁG‹±ü£ÿ%ìŠõû}Oï7˜èþ#½ƒ×þ\¿–YüŽÛ´h]”mQÓóªS[+Ô#œ”öAò—Hì·Aü‘™/hhi€pHh93'@„¤H“6Â¹­höDšúø†@QÌì X¾à2¢ôìÊ˜õ3>ýwÿ.æ“ßsAGwŠü¢|ru×5K"ÑbÃBv=éJ¸¸âr3³3zbaŒ”’”u||)b£NSÊµIB·Ä­` ^‡(“¿-21Á»ÝÑKûç*]'ÝzvÉ÷©,iÍ%¯’“ŒSZRÏb¯›ÔiŠk¡™rq:Éæ¸³¬{²\Ö’"tÝqbÙ1Êtd0fÍ·Ä‹"»“OÊC“•«RZ0cýÇÔQ[ÆutÕ@®«–ÎUr2J‰ê|\àÖ38ÌJ˜âK˜¨Y‘O sTÔxžW“v1Â ðQzÉ> .'F?§Hot¬ššeÀ/WCâ£Õ5‡Xc(Èg5< Ñ¨6|r••Çˆ›žWÏ7Y*nÜmÖ´G9E°ÛE(F³:ýeÉéYVLŒÏ“Áïûy&eöxÜ¹Õ ¥zœ&Ê'œËy/¢ËSzÃ­6:
¢C‚wLÛ/ž>EÏ¦9ÕÍ/Ò¡/yÜ#­¡ûuÝÌòÿXžW£®z˜ÎÓ4R¥]hŠˆ¥sèÁíÁx¬A1 áré-È–Ø—i°î²~pú„qÛ£9ü@ØIÙD7eY÷áÓøð¿¶('ëø77P¥	Š¨³”øãúu:Ì·¿BPµŽ¹â™rÔ÷çwÿÔ©Çùã¹ï+wýÔFõÙ?KiP_•H,ó(G.4ÿ=AžÕò…ùæŠssø„á7Ÿ8ŸËäÉ}Ò:­`eµ•û©mdi2©4í¸Y…´b£¤dpÐa—)š¢@!+R´p‹ªy[€ÚF®WõIqÀsB6'˜-¦8ù%~Ÿç*k¯HÎæ*Š•ZEÑWduÒ{ƒ%CÔ…1¾ƒŸ‡[ÃtœÎS™)m5ÓŒtÊé>®/Ô'W)e_’Kîu
ì s+£"xŠu[Üî)©|~/Ù€˜lñ<šîø6 Íœ$›ž<±…Æ?Ð^XÂj~éMÐÄÊ¥7
û£6ÑÂ3XY…ÙZvÛþytuÑúE¹çêZ²òÔ.}×œ}žtùæÄ=ßœsüïwÎ	y¼¥úˆ¶&¹’ÄþÎ¹çœw·qÍ-¹ÔÚ»å¨n.»]’ª g3A_Ò3æJgáþ·©µ+q—›DŽðâxr¯ ‹ÒçÆ÷!Än¬æjK¸{_›:˜~fgÛ¿‹oým—TY0ÙºÈY jRÇûæXkiÖÑ];BÂà“ßjË8Ô$Ý[Zù—¼âó«ñ-3ª ÏL“xº³Œ×L<zß¡7Ý ßÌa®vœéÔÄ´6Úî/f‚ëªƒÜ'ö”iÔGPtáúÙbÊ%é‹‰»È^‘Ï€Ó¦%lèp[IçÔÚß¾Û ±…Y¯ä7š`Ñêòât·ö[®|¾žãA[4 Á³V-à1 CÑÄ<›—ø’Š\íü+ô­ÐgÌÀ¿A¢`ƒFÙxmlÀÒ?†³!ë0øÓódŒÂGþ ¥ÛPi[°år”_Óh7è³×%ÿv¬*ÈSþáˆ·‹Å¬ÿÛõ#¬PÁ‹›òÓÐ[NU|'lÇâË;F9ÚHá‘ÅHåýÔy¿ø¸_ÌÌ¥xbW}½£“|‘þKèÉ\%£špÙ3í‡ÀsÀ¦È›Ð<.Ozrßø)ðl!§SmÜ­§¼ã)øª‡A Ì¦Ç
˜H<JÏóD»ìwì§|ß6TÜW=|à¼Upeu#WT¡uÜ!Î@Ó>í©ë”tn‹j ~l–Ð=?ƒ!ÊWˆÁÑ¤	-¾<;ƒŸ\‹X‡·¶@ª¼Æá_ÈÓyWr¥Áoqæ°gPÍŠÎìú'ãý¥Ù=¤þ-ƒ²}iž€‰SrïÍé_ájÿ÷ô²ìÙ#Xï—ù$íÁ52ÃÔ˜”´VßÂïØ­'OXÏý>ý>e2ÿ!iížë¥Ý->©Q»²DÔ&ª>.y3ßððßëSÎt<WùúN1æ[Oé0µÙôT°»`RÖÓ2A³ãËõ ß¢ ç• µçÖ´aÉw?“Ð‚kÛOÙK„¨†×%2k¤>èu“?ùYV
'FûÏTîMàMCÈ<ÓwPdÂI•J’HÌŽÊ|üÀËŠh,¥/—tgÔ%é•Õ°Ty»4‰Ù„È¥©µ¨¥ÃüuŠ‡k*hù/aÞwÏÏ‹ôœæ]Ql"tLeiÅ§/‰+ÐšŠ¯±zÀý€üét¸Ÿ|4æ\A¿5fKÓ‚#é©M;ãFEd"HEqí¦Ü%–Žd7ÝÞ«ÿl?X­4ëQ×…Ú[õ¥i¿Úy­Rü@w.)LübòH=4rÕóížñ|U@è>ÿÓS>üsæ~Ò]§5~ËSÉqgÎr82gjéa0>d/õçM{­hFÍþÈz|ÏD1îSÖÒlÐW7÷¦š<Z‚}áé¹b]©µ*¾MC‘¶“Niß)Ÿ´ö„o–©˜>ºš^ÑQƒoå¯±KPÝ±aUÈ›—ZÕóË¯K­­—gmäjb,þ®Ý(‹£>²V¤¥;×ê=«CÏúž=ò£åˆ-i·÷3ú7£y±H=9—¯=rE_í´òR±ä¬ÆùñT©ºNI¿Y[AUp_”ã4ÍYªQëÆŠrÜ ïMènéò-úà¥~ÅÒq™joÁN¹Þƒ¬4öˆn©›CÓÏ µ®­¢úyt<¡j²¨:…ËwíPmœÊQèçìJK¬NV¬$+?â‡,Ÿê-5Aƒå‘‰Q;Ò[-üPÍ‚œ4ãlWÍ…{u¬9()hÀõ¸Z	"Á=ö¶~£j<­Ž_9×.×Ö\Yã¹©ùiSð^WòÙ}½ÒkÇiC­u‹¥º*›´Ú‰Âž}Ð22ôB÷Z§[iàH®Õ«ï˜o2ä#ö~àxÔõ\Òô6kòUÚèz šFù€üÅý@5¯IíPã˜iCFšó]M‚9Ý>’7—ßÚ–´|ú[–¿&š¬üöX¢ÒÝ1$eÕ	ÇœæYižÍ×Ð8¡ŸÀ((Bhø¨×›õ¡ÒùQ‚Z?I°ÜSØÏ]G‡¬(çQgN«Ý··ß9ŽØ'ÙäºuÞ¨ÙýÔ¤³íT¦ÿë_[mJqqfˆ{ìÓ3\Á•¢Là‘UkÝ!†´Õ=˜·ZÛt»‡	qø·ìg`Ò­·ºfz@FræÙnÊ8•Ôzíù{ž;qsJ”¯]\x¸UøG¾	ÑöBTùÃ¤§¢§Nq©íùè¹h§ß]¿Õ}Ÿ$HÏ·>4gPÏ¢Š~\Ï¤„¸Cc›[šr‰+Ü¯°âa…R%S…”[¼Ó²yÃ¦§5½b„8ÀI¿`ÉÙËùˆÂ%K'³ù%êùgEZŽXµÒu:	MÈöm–jÏÓ]»š»Þc(•öEk“²/Ë	t»ž´Y¬†,œÏ1I RU)›PíÔ‹dJ0ëßÓtÆÐÁò]åˆT%àb„ÎØ4¦ææO´¾ÐÏ|h A#ªÝÙPl¤ú"8ƒGö¦¹°	);Éú†—2Ê™û›õ¡8Ù#È M,9¥¾ƒzÍ‡d†’JjWc*Nq¬ñÂ%/OÑ{iU&äô’‘xA@ÜŠtF%Ð8bDµzžÎf[¤Ï/…qô@¾™¿}§³åÎŸ¢÷Kn§óž”°Ò“ÒºÉÏµO£ö-Ñ¤¡ÉgÐÎd:SªQ<¡×õGIÙ«¬ä\J†CÇ%Ñô-Õ’ˆtÃõOMS²&dRš©}êxýèêó3=ƒYµÝ6Ö\)˜4‰8/a¸èIü´DVáÒ6<L. µN9‚«®)bž% ,~ROê–Õ>zK3­úõ:ìVBà·z:ŽCÜ¬æN!»'ªŒÐ†¸
[–kjðôm˜”]wiVvô¨ÓºÌ­œë.JX<z½¾!! u+ÃßáLÛ˜£cï\,bàd,.Ñp5£Ô\8Ý1ŠF%ó[Û¨9åÊoô¨:OÔå_*’"ê¨2ÿŸ.æô[eÀP’rþ¶	~gEˆÚjgýS_¾óŒÖÒOšXÃfp‰éË-ÄSîúÆ–Ý¸¾îÆºÂÎ»ôÆ+¬½±gñh|À’ðíVkµ„6œ€•È¥«oÎ5|lÁÝlao\¶E¾{)ksÔv1›–†ûÛ¨¾~*áïäBó°ISü•q¢þ€yv,²Šè-·yLÏM<ïÖßÍÆ¹“{sÓ·„Çc‹b:´Èã
¨rdŽdóQ‚’ÔU
6M¤OP¦ê’Z?×¹Ô›Ãf>HØ?¤C½§2›ÙäçŠO™—§t×lµ7>ó²”ô»Ïiœô,ÈYÎ§ÍeöÌ³òŒ_ÍÀS:„SKCƒH‡Rå¤çØØ&oÉ9pï’O“ùˆQfV´ êÂ<ÛÑŽÛ_A·Õ•”²ïùÖ¤²8$m?ð=NðšóÀ;ÔôÓ-{þ\ûÅKîÛ>b‡à)Ÿ®XrrÚvµõVWrë‹1¨å6Y…ð«zuuf¸e—¿v0¢iqÁæ
_S…¿ÁòÚ£V§TmPÚ×Œž_kßÙ&þ•öú
XBÕ¨±·x‹bo§æç`OÎ™à>Í8Ýµü¬Ž0´!G—ñ{¤0Ô¸¾aM8¢.|¨µÃÇáDãP¢ÆRHÂ[ˆíù/=]‹.(¼vjg×aIêš^…›áGS¡^vó|³êóól ¬uCeÄ¥äkÜLžð|³i	½ý7f‹B©‰ÄlyG2·°8Zf*ù«¡™%Eß—ysµ˜¬wû”Ûª+ÝtÀXÁöøMÛZìz¯ê}©O$WfèÕî2šKžNaZ½[¢nç+Sï-º^T3äáiíÿ  ÿÿ A«½xœìírÛ6òžÕt*ª‘d;­¯=Õ–&MkfÒ´÷zs“É\ 
²ØP¤JRþUî­îuîIn_@¢lÇ×^O?l»‹ýHrzzJh§!-¢4y1ëõÈæÑ¿œO§Á<Š–±YÙ˜¿9|;Œf½¯ìQ¯irÁêdâ±1j+¯¶„Å93ˆæ$01$Ÿ|B>ò€æé’šÜôÈé˜PÀ‹ìI`·[G‰õDcûHýÝöÉV}kò>)nV¬OØ,*¢äâEÁ–}²ÎY&'yS ¬ƒòt]¤ƒœÅ,Ä‘džf$O¡	Î®W,ñëŒ‘4‰oz0 @œÍçÐ78å‚BäÎÇyÐ5Æu‘‡ž
øxe Wr+j
1Š˜,Í‡ó(™µ`¼âž^PÅp¾NflvþýYîkÆ,¹(â‘¹|°:çˆtàŽ@ãäèIí…²×„÷TKCÍµ3!—&cÀÀ„ÜÉXB?™E—$Œiž¿¢KvÚ™Çìš\ÐÕàIg¬±=ËIñ2†\æë8&«ÁÑð˜LÓlÆ2’¥œ"R°ëb/;æÐK¯Ùé&l	èoÍgiòluº	˜ó©à›°aA³V9ÀÞWÕž€sšI¨<}
òy7uÏÏ×Ó]RVÝUM»JÒG†úô\}5»o·ÆMÆ~YG ©§® ¡‚(<¹v”Í_¯gÀžŸ"vÕ-!ŒX'éŠë¨X‘Ng¼)•ÎL&¤‹ˆ"‚ýy—ŒH÷\ˆ…hênO<sŠ²,rÜpIWA—§pxÏn@&@ý¶Z@ðfÿ³x¯¤åä@¦!©ÓuQ 4¤ã´#n:(Uq¾¡âSÎÒp½dI1Ïb†—_ó5.èt èëM†!
z[SØW×ƒ'dz1¸ÈèÍàèðÐøEzÉ²‘êð:pàwRw¨ETÄ ïél&ù×?>9kå< í?’w› _¤WB¨ùÏ)úXró.^‘­ùõWR®f)!V3Èø îÑ÷gIvšK¹vXFÚ™$ËàËºàÁÛrµÚ¯ù¹“*M4[û©Ç™Ö¤bw<ÆÄo?l“áµ–®-´n¥­i[[…œßsh4îù×T]ä^U×§¼÷¢¾’à‡T_>¥G}µ‹ ÞÒd%]ÖüÚüW[ªÕª•âZèSÝ{ÐS¿W.—ÂòÊæRXJ†[íÈˆSÙlË U]‹Ñ&¸½„h™í­Ãš ©ÙiR$f‘
hU¡¦U¡ÂªÐ‡·*%ºiXÔòíoZ™öE
ÆÓó°"¹©÷í" ËRxmÉ{÷µ%š¦“7<Ä^®^“¨U3¯ÓÍÜRÎ\jgþ_PO‡éû>·RSq¡”õàSòm4›±„DÉj]ä¤H	Kr,„)(¬ä<bñ,'š@b—QQ ªŸ(H'| äÝ‚Ãê„Ó«ìN5ã>·.ãÝN5Zn3¼tÒŸßn…©o½m(µØ¶æYm)É003Ÿç¦n§k¯LƒÊíf’²"©ñ“÷ˆšëRæs)«3³Æ;u+G£­²cªC7‡ð{6ØàW¿etfÛjÙ¸Û8§÷\?ñVPšâ
.ÝíHˆ6¦^MŒñ>¶DI[~\±õR¸È<MŠÁ’Í¢õÒ°ö_vì5 „#0"B»Tå8+×!S•c³ª7á¼G~tÿž®3¤»u@ï¡ÛpÃôxVŽ¯™Ã-Ÿ>ÞQ#‰ƒªóÐ5JÎã]_Ž¯ß rJlUþª¦ï*KCŸsWÁþW€Õ«‚â»ÏÓ,d$Z¢ŒÐ‚‘$-¢y$l6ßÛà„‘Áý7½LA d4½bzáà49Wû$‚=/¦È%§c ©ôæ}¢sŸ”¶O,?©·
p§!J.À”àz÷‰\—"lïø­—'®ðBKµgì@L†ÞtŒ.ªü‡VÅ3ŒðOVºS.úi,ÁÐÙ2Jº~~¡Ez–&ó([§8Ý¤kÂƒ*¼¸¢Ië$–Œ&×W³XD9v(NÈxFä‹h^ðÞzf7o„B¨qÃN_Éž‡y•Âò'¾QNËÖmhïÀ\v;'ÆµÆò`Ò¬ùÝQ6³B™ðe†<œ±ü[fú7~ãÎ¥f†—Óñšþ|A³ª2Ûð¨DÎ#. ª¼1¦*ßhYý„‰u‚	_~cmXø«?ÅÞí,ñÏ Lã¶¡—°Ô¶èÜ*:×:…U(]«¸ÕY¿ûî*ˆ7þò…y6¦6«}áÞs«SÜ×®€†¡(:‚%Ï³tY ~/#|Ñ¤?\`ìî(Oyì€ñœÅUmšÐ|¨Œ<·Žê¦§öryæ·¾–õGIX»šøfK›‰¨²ááÓ½€+Ï&±õßq´åÍ. ò)F©îÃª»«ØY>öéË—ßÿíì›€lœs&,o… ,i7“PŠƒ’ƒœoi°ðÂã°Œ…,ºä¡wú³ˆÔz	®^¨(ÜcÒÿZy‘ü˜Ýy¯OkÁË €Ã7JXðimL%€„Üùˆò×„T
˜¶nRõýEHÈ†pÑS!·=¡–@Ö7T )Ž\ Rÿ¸`à•^ãÆD+é¼TR“áüÈŽ\9‰Å\Üèa¼ÿæ­oê1Xƒ’ÇHõrÊ2ˆÖÊÉs1¤K°¿…·M«cëÛ\ Û³„±BèR4kRsEj½àå¸*a^Ò(E”´`ô¡)~c î#ºáJÓ7 V‚Ò­&qŠHN‡eß1 È5yêB¬d*¢å¡äíX·O/iÓiÉâ¿ÿù¯1Ñ°H_]1;/2¸zÛ^]dÑshÝö‚š¼Ýll—OïÈ¨0§þ Yõ½åÕžÌÚ—[×f×wÎ¯wdØ;sìvYö~yö>™v5×öÍþaÓí=îßTÊíKº½iwcâíd$ó¾{î]²£6ù†dLynz›$¼1oLÄ½µj“7Ÿ©fÞõI×”ú”³.›n™Oß2£¾÷œº)«®æÕþÌÚÊ­+fzw²í,®Ï¶óÙàÆÍ±Û:ÜSwmØ©ÝÂöí‡´;âR#Ù;+JRªŸZ ¸®-Õ¤ÑÜ1ü´}D¿¬OÊQªÀX¨P‹D	
½a=D‹ÑˆEÄ2š…‹›š@T@–³´„]‘sVøÌ€L”ÔÔ$s{óÆùÂþJƒÆÆ| ,à­ó¢Nz"»›ù‰ß¢Éø¿,>`ˆ—Fj‘27‡ñ»!ž’ósö6Ã†š"ƒíÕ{TgÍÔÿlÄfòpµÎ¾îá!}I’8_¥ËžKžÛÇÌÀKÌ;ìAh~ó–v·ÍN×.î.{$NEîa#‘‡ÝMð’k³Qö­àæf-uVüN€èÕ:üñMè+7t½‡?ÜÜøÛè—Ul)±^@&û^ÒbyÔk‘·Õº#Þ¹.Y¹‡t…øJxúµKLöMMöKNªïk4ä(ÂX`‡Zh·{á¨n²œ”Ã±^/UßH³ñ®Þø2J`¿%®÷5™‰†ÉÔdŸ´CñüöÓ{c» ™fÑE”ÐX©f}®TI×ˆaêZçX˜€9Ó4‰×]s,Vòûžw8¥ì¸Æ$‹CðKV3ÜP½‚¯±&ó¥pG¯ËH4àKZ[êÃßÆgìÝƒÑMá]†Ô]54ºÃTHPØç›(_ÅôæOøH6¹Eêd3`xŸ¤W‰,Õ5Î¯äK3bV»ØÞ¦Úp$XWÓXi{öµs+röb§Âhƒo&:ÍxƒrhePKso¢?½AaòlhU“ÄšGb¿¬SÖŒ)….-Ö¹8™˜±ŸùS‘0·±èÂÓo•‰ïñàn\¯WG§±ÃPnŒ˜Ùßn`Û%wË”‚±¶9ÂýXOš€?Iï|IÁÕ¢êÆ€Ædmüm+¥‘º°š?¹û»¥IzÐ—ô´5¡µ »]aÞSEs_"÷ž=qjŽç›£ÃÕõ[3°*£ñŒÁ·=¦iÌÏF¨Ç‡2vÖç&gä™	ÕyÆ£uTŠ|Â@§Ÿ×yÍoSV\1–x‚xŽd¾¢Éø)ÀqÐ_FË¨b»ÿëóJÁª‚^?³geùêVË‚Îò?²:{¯æ¥÷.úunkÁÕ‘?²vºŒ1M+ÎûyöÂê«_÷öùäÝ<ím½¬éaÅÌÙÉCyZŒÆõìþ0{GÜÖ;"÷;uÇ÷äŒ¥}‡Õyd‹—7zajMù•œãÂÈ|‘¼½^«mwÙ3pa+Oß«×¦Ê&©P‡ i¢Ž‹åŒ‘)i²n›’˜{l{÷yÁpÅ³Œ%¢§°Àéí# Q|vEíCB”¾„\± :õf®ö/íï„íˆè.²¤­{ê{c oŠ¸	y£€±UZœ;X)?trÌl™ø_}ê"2ïsuôk¢ÐSemCµJ¹B®ÏŠá;«¢¶úè¾¿”â¾Œe<R™¦ïë(ËL™ñ¾èâôo:û¸g}Olütˆõ-ï—@rã ùÿ?ýÑêÓâæ6è¾Ä7csºŽ‹Ÿƒ›”
óõîv_)i%î§'ür¢UÚ<©UÕô‰JÚÅñ-XÝ}±…Œ…–˜‰Û1¿¨Ý{é}h±û}²Â+vâ•¨„Ÿ×†«Aª|Ÿñœòäqÿd \9dÈ Xhª·—Åõ}÷$dZ`¸ÆÒ7þ@±ÒŠ—ßn(Ò‚Æg	8tªQÌûèÙÄéRÅøSž£ýÀ2Ý&x)lrŸß¿0ºˆöGÜ¹:“MÊÊšµl·§·º›h¨öà9Ÿn-È‰WF?.”xpª]À 79\¸zÜ#ËëÓ°XÓøGƒ2pð&¡ö‡,@¬‡0WI»¿«x4’ñ¤˜5Êñm”S‹ú¨ÿAåÉà¨”±,
_X}$Ð‰¤yÅ™˜`|8¢Ø–pŽ`Øw´XCÅê‰ðïà•`ON&°U0ðÈGîÕTðÊ·Òú´´8i­©$#+ÎøF€MM`$„ÐG=ò©wêÇÄ`6¤.,/‡9«–QbMâ…ß÷ èø¯äBˆê$þõOQMc¦–éX2:f’b)zõ‰Mï¾6Øh²zÖåZõp+#ÛŒ%VÆ ï ™®Ö¡Zr?–3òw‚$!Â”]Áú
nAŒ$¸{üXÍÇ™)`D%;Nn|?)_Žøu–^5Ö¸x8ó9Yðg…dÙ¤è4†áHÂŒˆV•rN¾´¤½œ/Ò+<8ÄK_ÖTe¹§1ƒVUÆÂÜoÇ ¥zH:ß5¤*æz0÷Ft²*	Ùø´Ý-wµãfõ8›ƒ±U#“ÇÃ©ÈÉ‘¯Æ'ƒT'7–Q™2ÂŒ,+îfáÕó5õ¦vG|,øPíVJ¼¤7ÀäêÉQ³œÐ+ŠÐ¬ ÈÃpïjtƒ¦†ë|”®‹¼á IÕ„…
^Ô7eÉùÐ>!Þ|Bost¸Öí88ŸoÇOŽ[v>ÈÇm!Cd‹x´ìÞ5èŒá¿»÷33vEõVû¦ô–†ß·»ãeÌŠÚÙI‘L,D<ÁÙ‚£žsdfå¸_4;Ý˜.‰[Ö73{ðÕ¢í”áøPC¥+BÐmµÁdyšð,c^a˜Ñ$ïhÝ!y±%NfÏ£4‘4;+õlÁ.3PÑ—l^ØQþçdvÜ~Ã{Æj_ÆºÌïœÍ?dì2J×yKNßžÑV’®#©ŸÀóÖÒ!|Ža8Jw°­+"ÌæWœë•t^ åžÜ¯üþ¡V°e?©$`ó"ºÁÕàÍgOp¿o1ø²a‹NÞ:à}#Ê]Sò±Çg9‚Çyá)¤ƒ3²7}Åc•ÿ«†/ !_ÐYzxu} !¯È­]Ps¨rC×u}ïš?9àY'Ÿ°Ùo ˜"¢‚ÜÓDß›ÞÉš<ÞÇš”Øÿ&ÍÊ+`\K“ò:ºX<”ñ.¹ö?Ãé—´½›¼5«}¢ÓmN7fžÑllŠkËý¼¢iÈˆ‘þº,6ŽøåÒIŒFÕˆ²b…#Gä5ƒtfÈÿ¾JgL‚*G)Õ¤¶—sM.¯Ÿ©5\Ç¥¥s6áu)Õ·}F„å?¬K:7ï '*ÅEÔ#1ùxÃÉÚ@ÃûÁá;ã­ªàAC0/kÁ2‡ÐÜd|tB“›qØ«g-×¼kœ²–×åKŽ7'6Ïµ¨Î8Ô²^­XÒœ¡P‡ï1ã(¸,ˆ-˜*0oø¿íXþ·rRkÚÍ;q¤a–óhÞ)¸À;~î¤ö‰™ój‘Dà¡ñÏ6‰Ÿ»(Q´Œ5$>‡Ë¶îü  ÿÿ ()õ~