/* ============================= CONFIG (Final—issue fixes) ============================= */
var CFG = {
  // Single live tab
  months: ["Month"],
  singleMonth: { sheet: "Month" },

  yearSheet: "Income",
  yearCell: "D20",

  // in CFG.formats
  formats: {
    money:    "$#,##0.00",
    moneyNeg: "$#,##0.00;-$#,##0.00",   // ← was [Red]… now plain “-”
    percent1: "0.0%",
    months1:  "0.0",
    date:     "m/d/yyyy"
  },

  /* ===== TRANSFERS & PAYMENTS ===== */
  transfersPayments: {
    sheet: "Transfers",
    rowStart: 4, rowEnd: 29,
    colName: "I",
    colDate: "N",
    colFreq: "O",          // ← NEW
    colFromAcct: "S",
    colAmount: "Y",
    colToAcct: "AC"
  },
  creditCards: {
    bankMarkCol: "O",
    paymentCategory: "Credit Card Payment"
  },
  /* ===== DATED ACCOUNT BALANCE ===== */
datedAccountBalance: {
  sheet: "Dated Account Balance",
  targetDate: "H2",
  accounts: {
    rowStart: 13, rowEnd: 21,
    colName: "I",
    colStartBalance: "L",
    colIncome: "O",
    colExpense: "S",
    colTransfers: "V",
    colPredBalance: "Z"
  },

  expenses: {
    rowStart: 25, rowEnd: 152,
    colName: "I",
    colDate: "M",
    colFreq: null,
    colCategory: "S",
    colPred: "W",           // ← was colAmount, rename as Pred
    colAct:  "Y",           // ← NEW: Actual
    colAccount: "Z",
    colInclude: "AD"
  },
  
  transfers: {
    rowStart: 156, rowEnd: 212,
    colName: "I",
    colDate: "N",
    colFreq: "O",
    colFromAcct: "S",
    colPred: "V",           // ← amount now lives in V as Pred
    colAct:  "Y",           // ← NEW: Actual
    colToAcct: "Z",
    colInclude: "AD"
  },
    income: {
    rowStart: 216, rowEnd: 248,
    colName: "I",
    colDate: "M",
    colFreq: "S",
    colPred: "V",
    colAct:  "Y",           // ← NEW: Actual
    colAcct: "Z",
    colInclude: "AD"
  }
},

lists: {
  accountsRange:      "_Lists!A2:A",
  expenseNamesRange:  "_Lists!C2:C",
  categoriesRange:    "_Lists!F2:F",
  transferNamesRange: "_Lists!H2:H",
  incomeNamesRange:   "_Lists!J2:J",
  savingsNamesRange:  "_Lists!L2:L"
},
  /* ===== EXPENSES ===== */
  expenses: {
    master:  { sheet:"Expenses", rowStart:12, rowEnd:148,
               colName:"I", colStart:"L", colFreq:"S", colCat:"V", colPred:"Y", colAcct:"AC" },

    monthly: { rowStart:144, rowEnd:400,  // ← CHANGED from 143
               colName:"I", colDate:"M", colCat:"S",
               colPred:"W", colAct:"Y", colAcct:"AC" },

    tiles:   {
      sheet:"Expenses",
      totalMonthly:"H2", predictedAnnual:"M2",
      upcomingExpense1:"R2", upcomingExpense2:"W2", largestDay:"AB2"
    },

    categoryTiles:{
      sheet:"Expenses",
      highestExpName:"R6", lowestExpName:"W6", totalSubs:"AB6",
      totalSpendingWithCC:"H6",  // = monthTotalNoCCPay + monthTotalCCPay
    totalCCPayments:"M6",      // = monthTotalCCPay
    }
  },

  /* ===== INCOME ===== */
  income: {
    master: { 
      sheet: "Income", 
      rowStart: 8, 
      rowEnd: 18,
      colName: "I", 
      colStart: "M", 
      colFreq: "S",
      colPerFreq: "V",
      colPredMonthly: "W",
      colAcct: "AC" 
    },

    monthly: { rowStart:16, rowEnd:26,  // ← CHANGED from 27
               colName:"I", colPredAnnual:"M", colPredMonthly:"R",
               colPredOnly:"U", colAcct:"Z" },

    tiles: {
      sheet:"Income",
      thisMonth:"H2",
      predictedAnnual:"M2",
      predictedMonthlyAvg:"N2",
      incomeVariance:"R2",
      daysLeft:"W2",
      unbudgeted:"AB2"
    },

    paychecks:{ 
    sheet:"Income", rowStart:22, rowEnd:126,
    colName:"I", colDate:"L", colFreq:"P", colPred:"U", colAct:"V" // ← ren U as Pred, add V as Actual
  }
},

  /* ===== SAVINGS ===== */
  savings: {
    sheet:"Savings",
    tiles: {
      totalSaved:"H2",
      amountLeft:"M2",
      monthsLeft:"R2",
      nearestToComplete:"W2",
      emergencyFundTarget:"AB2"
    },
 table: {
  rowStart:8, rowEnd:14,
  colName:"I", colStartAmount:"L", colGoalAmount:"O",
  colStartDate:"Q", colContribution:"T", colFreq:"X",
  colPct:"AC", colFinishDate:"AD"
},
    // In CFG.savings.monthly
    monthly: {
      rowStart: 45,
      rowEnd:   51,
      colName:"I", 
      colGoal:"L",
      colStart:"O",        // ← NEW: write start amount here
      colPred:"S", 
      colAct:"V", 
      colPct:"Z", 
      colFinish:"AD"
    },

    monthlyTiles: {
      totalSaved:"H6",
      amountLeft:"M6",
      monthsLeft:"R6",
      nearestGoal:"W6",
      savingsToIncome:"AB6"
    }
  },

  /* ===== BANK (Month) ===== */
  bank: {
    rowStart: 86, rowEnd: 94,
    colName: "I",
    colPredBegin: "L",   // Start (Predicted begin)
    // colActStart left unused
    colExpTotal: "P",    // Expense
    colIncomeTotal: "T", // Income
    colTransfer: "W",    // NEW: Transfers (net, +in / -out)
    colPredEnd: "Y",     // Predicted end
    colActEnd: "AC"
  },

  /* ===== MONTH sheet (dashboard) ===== */
  monthly: {
    tiles: { income:"H2", expense:"M2", avgDaily:"R2", ratio:"W2", flow:"AB2",
             hiCat:"H10", loCat:"M10", hiExp:"R10", loExp:"W10", subs:"AB10" },

    // Category summary table (unchanged)
    categoryTable: { rowStart:66, rowEnd:82, colLabel:"I", colTotal:"N", colPct:"O" },

    // Month expense table (moved down)
    expenseTable:  { rowStart:143, rowEnd:238, colName:"I", colCat:"S", colPred:"W", colAct:"Y", colAcct:"AC" },
    subscriptionsCategory: "Subscriptions",

    // Month income table (unchanged)
    incomeTable:  { rowStart:16, rowEnd:26, colName:"I", colPred:"R", colAct:"U", colAcct:"Z" },

    // Calendar (dynamic 6-week grid; Sunday..Saturday columns)
    calendar: {
      startRow: 99,
      perDayHeight: 7,          // 6 = date + 4 items + "+N more"
      visibleLines: 4,
      dayColGroups: [
        ["C","F"],["H","K"],["M","P"],["R","U"],["W","Z"],["AB","AE"],["AG","AJ"]
      ],
      spacerRows: [105,112,119,126,133],   // exact rows for the gap
      borderColor: "#b7e1cd",
      borderStyle: "MEDIUM",
      overflow1: { col:"AB", rowStart:134, rowEnd:139 },
      overflow2: { col:"AG", rowStart:134, rowEnd:139 }
    }
  }
};

var __ssCached = null;
function __ss() {
  return __ssCached || (__ssCached = SpreadsheetApp.getActive());
}

/* ============================= UTILITIES ============================= */
var _ = (function(){
  function c(l){ if(l==null)return 0; if(typeof l==='number')return l; l=String(l).trim(); let c=0; for(let i=0;i<l.length;i++){ const ch=l.charCodeAt(i); if(ch>=65&&ch<=90)c=c*26+(ch-64); else if(ch>=97&&ch<=122)c=c*26+(ch-96); } return c; }
  function n(v){ if(v==null||v==="")return 0; if(typeof v==="number"&&isFinite(v))return v; const n=parseFloat(String(v).replace(/[^0-9.\-]/g,"")); return isNaN(n)?0:n; }
  function toDate(v){ if(!v)return null; if(v instanceof Date)return v; if(typeof v==="number")return new Date(Math.round((v-25569)*86400000)); const d=new Date(v); return isNaN(d.getTime())?null:d; }
  function fmtMoney(r){ r.setNumberFormat(CFG.formats.money); return r; }
  function fmtPct(r){ r.setNumberFormat(CFG.formats.percent1); return r; }
  function fmtDate(r){ r.setNumberFormat(CFG.formats.date); return r; }
  function currentYM(){ const now=new Date(); return {y:now.getFullYear(), m: now.getMonth()+1}; }
  function year(){ return currentYM().y; }
  function monthIdx(){ return currentYM().m; }
  function safeDate(Y, M1_12, D){ const last=new Date(Y, M1_12, 0).getDate(); return new Date(Y, M1_12-1, Math.min(D, last)); }
  function normFreq(v){
    const f=String(v||"").trim().toUpperCase().replace(/\s|-/g,"");
    if(!f) return "MONTHLY";
    if(f.startsWith("WEEK")) return "WEEKLY";
    if(f.startsWith("BIWEEK")) return "BIWEEKLY";
    if(f.includes("ONETIME") || f==="ONE" || f==="ONCE") return "ONETIME";
    return "MONTHLY";
  }
  function periodsPerYear(freq){
    switch(normFreq(freq)){
      case"WEEKLY":return 52; case"BIWEEKLY":return 26; case"MONTHLY":return 12;
      case"ONETIME":return 1; default:return 12;
    }
  }
  function inWindowMonth(start, _end, y, m){ if(!start) return true; const s=(start.getFullYear()<y)?1:(start.getMonth()+1); return (m>=s); }
  // REPLACED sumAEP_block (kept your logic)
  function sumAEP_block(sh, r0, r1, colPredLtr, colActLtr){
    const colP=c(colPredLtr), colA=c(colActLtr);
    const predR = sh.getRange(r0, colP, r1-r0+1, 1).getValues();
    const actR  = sh.getRange(r0, colA, r1-r0+1, 1).getValues();
    let s=0;
    for (let i=0;i<predR.length;i++){
      const aRaw = actR[i][0];
      const pRaw = predR[i][0];
      const aHas = !(aRaw === "" || aRaw === null);
      s += aHas ? n(aRaw) : n(pRaw);
    }
    return s;
  }

  function expand(year, start, _end, freq){
    const Ystart = new Date(year,0,1), Yend = new Date(year,11,31), dates=[];
    if (!start || start>Yend) return dates;
    const F = normFreq(freq);
    switch(F){
      case"WEEKLY": case"BIWEEKLY":{
        const step = (F==="WEEKLY")?7:14; let d=new Date(start);
        while(d.getFullYear()<year || (d.getFullYear()===year && d<Ystart)) d.setDate(d.getDate()+step);
        while(d<=Yend){ dates.push(new Date(d)); d.setDate(d.getDate()+step); }
        break;
      }
      case"MONTHLY":{
        const s = start? new Date(start) : new Date(year,0,1);
        let m = s.getFullYear()<year?0:s.getMonth();
        const day = s.getDate();
        for(let mi=m; mi<12; mi++){
          const d = safeDate(year, mi+1, day);
          if (d>=Ystart && d>=s) dates.push(d);
        }
        break;
      }
      case"ONETIME": if (start && start>=Ystart && start<=Yend) dates.push(start); break;
    }
    return dates;
  }
  return { c,n,toDate,fmtMoney,fmtPct,fmtDate,year,monthIdx,safeDate,normFreq,periodsPerYear,inWindowMonth,sumAEP_block,expand };
})();

function _idxOrNeg(baseCol, targetCol){
  return targetCol ? (_.c(targetCol) - _.c(baseCol)) : -1;
}


(function ensureCCcfg(){
  if (!CFG.creditCards) {
    CFG.creditCards = { bankMarkCol: "O", paymentCategory: "Credit Card Payment" };
  }
})();

// Count occurrences of F in (year=y, month=m 1..12), anchored at `start`.
// F ∈ { "ONETIME","MONTHLY","WEEKLY","BIWEEKLY" } (case-insensitive via _.normFreq)
function _occurrencesInMonth(F, start, y, m){
  F = _.normFreq(F);
  if (!start) start = new Date(y,0,1);

  const monthStart = new Date(y, m-1, 1).getTime();
  const monthEnd   = new Date(y, m,   1).getTime(); // exclusive upper bound

  if (F === "ONETIME") {
    return (start.getFullYear()===y && (start.getMonth()+1)===m) ? 1 : 0;
  }
  if (F === "MONTHLY") {
    // If the anchor exists before month end, we expect exactly one hit in that month.
    return (start.getTime() < monthEnd) ? 1 : 0;
  }

  // WEEKLY / BIWEEKLY → arithmetic on the anchor; no arrays.
  const periodDays = (F === "WEEKLY") ? 7 : 14;
  const t0  = start.getTime();
  const pMs = periodDays * 86400000;

  if (monthEnd <= t0) return 0; // month entirely before anchor

  // Find first occurrence in/after monthStart
  const kFirst = Math.ceil((Math.max(monthStart, t0) - t0) / pMs);
  const first  = t0 + kFirst * pMs;
  if (first >= monthEnd) return 0;

  // Last occurrence index within [t0, monthEnd)
  const kLast = Math.floor((monthEnd - 1 - t0) / pMs);
  return Math.max(0, kLast - kFirst + 1);
}

function _buildIsCCMap_(mSh, B, markCol){
  const n = B.rowEnd - B.rowStart + 1;
  const cName = _.c(B.colName), cMark = _.c(markCol);
  const w = cMark - cName + 1;
  const vals = mSh.getRange(B.rowStart, cName, n, w).getValues();
  const idxMark = cMark - cName;
  const map = new Map();
  for (const r of vals){
    const name = String(r[0]||"").trim();
    if (!name) continue;
    const v = r[idxMark];
    const isCC = (v===true) || (v===1) || (String(v).toUpperCase()==="TRUE");
    map.set(name, isCC);
  }
  return map;
}
const __expandCache = new Map();
function expandMemo(y, anchor, end, F){
  const k = y + "|" + _dateKey(anchor) + "|" + (end? _dateKey(end):"") + "|" + F;
  const hit = __expandCache.get(k);
  if (hit) return hit;
  const out = _.expand(y, anchor, end, F);
  __expandCache.set(k, out);
  return out;
}
const MS_DAY = 86400000;
function monthBounds(y, mIdx){ // mIdx: 1..12
  const s = new Date(y, mIdx-1, 1);
  const e = new Date(y, mIdx, 0, 23, 59, 59, 999);
  return { s, e };
}
function daysInMonth(y, mIdx){ return new Date(y, mIdx, 0).getDate(); }

function _effectiveRows(sh, rowStart, colIdx, rowEnd){
  // scan the single column from bottom up to find last non-empty
  const n = rowEnd - rowStart + 1;
  const vals = sh.getRange(rowStart, colIdx, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--){
    const v = String(vals[i][0] || "").trim();
    if (v) return i + 1; // rows to read
  }
  return 0;
}


/* ============================= CATEGORY TABLE (MONTH) ============================= */
function Expenses_UpdateCategories_INO(){
  const ss = __ss(); sh=ss.getSheetByName(CFG.singleMonth.sheet); if(!sh) return;
  const E=CFG.monthly.categoryTable, t=CFG.monthly.expenseTable;

  const width = _.c(t.colAct) - _.c(t.colCat) + 1;
  const rows  = sh.getRange(t.rowStart, _.c(t.colCat), t.rowEnd-t.rowStart+1, width).getValues();
  const iPred = _.c(t.colPred)-_.c(t.colCat);
  const iAct  = _.c(t.colAct) -_.c(t.colCat);





const incomeAEP = _.sumAEP_block(sh, CFG.monthly.incomeTable.rowStart, CFG.monthly.incomeTable.rowEnd, CFG.monthly.incomeTable.colPred, CFG.monthly.incomeTable.colAct);
const map=new Map();

for (const r of rows){
  const cat=String(r[0]||"").trim(); if(!cat) continue;
  const aep = (_.n(r[iAct])>0?_.n(r[iAct]):_.n(r[iPred]));
  if (aep>0) map.set(cat,(map.get(cat)||0)+aep);
}
  const sorted=[...map.entries()].sort((a,b)=>b[1]-a[1]);
  const outH = E.rowEnd-E.rowStart+1;
  const outW = _.c(E.colPct)-_.c(E.colLabel)+1;
  const buf  = Array.from({length: outH}, ()=>Array(outW).fill(""));

  const offTotal=_.c(E.colTotal)-_.c(E.colLabel);
  const offPct  =_.c(E.colPct)  -_.c(E.colLabel);

  for (let i=0;i<Math.min(outH, sorted.length); i++){
    const [cat,total]=sorted[i];
    buf[i][0]=cat;
    buf[i][offTotal]=total;
    buf[i][offPct]= incomeAEP>0 ? total/incomeAEP : 0;
  }

  const outRange = sh.getRange(E.rowStart, _.c(E.colLabel), outH, outW);
  outRange.setValues(buf);
  _.fmtMoney(sh.getRange(E.rowStart, _.c(E.colTotal), Math.min(outH,sorted.length), 1));
  _.fmtPct(  sh.getRange(E.rowStart, _.c(E.colPct),   Math.min(outH,sorted.length), 1));
}

function _cacheGet_(k){
  try { return JSON.parse(CacheService.getDocumentCache().get(k) || "null"); } catch(_){ return null; }
}
function _cachePut_(k, obj, secs){
  try { CacheService.getDocumentCache().put(k, JSON.stringify(obj), secs||300); } catch(_){}
}
function _cacheDel_(k){
  try { CacheService.getDocumentCache().remove(k); } catch(_){}
}
function _toastThrottle_(key, ms){
  const props = PropertiesService.getDocumentProperties();
  const last = +props.getProperty(key) || 0;
  const now  = Date.now();
  if (now - last < ms) return false;
  props.setProperty(key, String(now));
  return true;
}


/* ============================= SMALL HELPERS ============================= */
/* ===== DUPLICATE PREVENTION (Accounts) ===== */
function Lists_ApplyValidations(){
  const ss = __ss();;
  const listsSh = ss.getSheetByName("_Lists");
  if (!listsSh) return;

  const mSh  = ss.getSheetByName(CFG.singleMonth.sheet);
  const dSh  = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (!mSh) return;

  // --- Source ranges from CFG.lists
  const accountsSrcRng   = ss.getRange(CFG.lists.accountsRange);
  const expenseNamesSrc  = ss.getRange(CFG.lists.expenseNamesRange);
  const categoriesSrc    = ss.getRange(CFG.lists.categoriesRange);
  const transferNamesSrc = ss.getRange(CFG.lists.transferNamesRange);
  const incomeNamesSrc   = ss.getRange(CFG.lists.incomeNamesRange);
  const savingsNamesSrc  = ss.getRange(CFG.lists.savingsNamesRange);

  // Build rules
  const ruleInAccounts = SpreadsheetApp.newDataValidation()
      .requireValueInRange(accountsSrcRng, true)
      .setAllowInvalid(true)
      .build();

  const ruleInExpenseNames = SpreadsheetApp.newDataValidation()
      .requireValueInRange(expenseNamesSrc, true)
      .setAllowInvalid(true)
      .build();

  const ruleInCategories = SpreadsheetApp.newDataValidation()
      .requireValueInRange(categoriesSrc, true)
      .setAllowInvalid(true)
      .build();

  const ruleInTransferNames = SpreadsheetApp.newDataValidation()
      .requireValueInRange(transferNamesSrc, true)
      .setAllowInvalid(true)
      .build();

  const ruleInIncomeNames = SpreadsheetApp.newDataValidation()
      .requireValueInRange(incomeNamesSrc, true)
      .setAllowInvalid(true)
      .build();

  const ruleInSavingsNames = SpreadsheetApp.newDataValidation()
      .requireValueInRange(savingsNamesSrc, true)
      .setAllowInvalid(true)
      .build();

  // Helper: compare + set-if-changed (keeps writes minimal)
  function _sameValidation_(a, b){
    if (!a && !b) return true;
    if (!a || !b) return false;
    if (a.getCriteriaType() !== b.getCriteriaType()) return false;
    if (a.getAllowInvalid() !== b.getAllowInvalid()) return false;
    // We won't deep-compare range objects (can be flaky); type+allowInvalid is sufficient to avoid thrash
    return true;
  }
  function _applyValidationIfChanged_(range, rule){
    const existing = range.getDataValidation();
    if (_sameValidation_(existing, rule)) return;
    range.setDataValidation(rule);
  }

  // ===== Targets =====
  // Month — Income table
  const IT = CFG.monthly.incomeTable;
  const iNameR = mSh.getRange(IT.rowStart, _.c(IT.colName), IT.rowEnd-IT.rowStart+1, 1);
  const iAcctR = mSh.getRange(IT.rowStart, _.c(IT.colAcct), IT.rowEnd-IT.rowStart+1, 1);

  // Month — Expense table
  const ET = CFG.monthly.expenseTable;
  const eNameR = mSh.getRange(ET.rowStart, _.c(ET.colName), ET.rowEnd-ET.rowStart+1, 1);
  const eCatR  = mSh.getRange(ET.rowStart, _.c(ET.colCat),  ET.rowEnd-ET.rowStart+1, 1);
  const eAcctR = mSh.getRange(ET.rowStart, _.c(ET.colAcct), ET.rowEnd-ET.rowStart+1, 1);

  // Month — Savings (optional names validation)
  const SM = CFG.savings.monthly;
  const sNameR = mSh.getRange(SM.rowStart, _.c(SM.colName), SM.rowEnd-SM.rowStart+1, 1);

  // Bank table — account names
  const B  = CFG.bank;
  const bNameR = mSh.getRange(B.rowStart, _.c(B.colName), B.rowEnd-B.rowStart+1, 1);

  // Dated AB (if sheet exists)
  if (dSh) {
    const DA = CFG.datedAccountBalance;

    // Accounts table names → Accounts list
    const daAccNameR = dSh.getRange(DA.accounts.rowStart, _.c(DA.accounts.colName), DA.accounts.rowEnd-DA.accounts.rowStart+1, 1);

    // Expenses: Category → categories, Account → accounts
    const deCatR  = dSh.getRange(DA.expenses.rowStart, _.c(DA.expenses.colCategory), DA.expenses.rowEnd-DA.expenses.rowStart+1, 1);
    const deAcctR = dSh.getRange(DA.expenses.rowStart, _.c(DA.expenses.colAccount),  DA.expenses.rowEnd-DA.expenses.rowStart+1, 1);

    // Transfers: Name → transfer names, From/To → accounts
    const dtNameR = dSh.getRange(DA.transfers.rowStart, _.c(DA.transfers.colName),     DA.transfers.rowEnd-DA.transfers.rowStart+1, 1);
    const dtFromR = dSh.getRange(DA.transfers.rowStart, _.c(DA.transfers.colFromAcct), DA.transfers.rowEnd-DA.transfers.rowStart+1, 1);
    const dtToR   = dSh.getRange(DA.transfers.rowStart, _.c(DA.transfers.colToAcct),   DA.transfers.rowEnd-DA.transfers.rowStart+1, 1);

    // Income: Name → income names, Account → accounts
    const diNameR = dSh.getRange(DA.income.rowStart, _.c(DA.income.colName), DA.income.rowEnd-DA.income.rowStart+1, 1);
    const diAcctR = dSh.getRange(DA.income.rowStart, _.c(DA.income.colAcct), DA.income.rowEnd-DA.income.rowStart+1, 1);

    // Apply (set only if changed)
    _applyValidationIfChanged_(daAccNameR, ruleInAccounts);
    _applyValidationIfChanged_(deCatR,     ruleInCategories);
    _applyValidationIfChanged_(deAcctR,    ruleInAccounts);
    _applyValidationIfChanged_(dtNameR,    ruleInTransferNames);
    _applyValidationIfChanged_(dtFromR,    ruleInAccounts);
    _applyValidationIfChanged_(dtToR,      ruleInAccounts);
    _applyValidationIfChanged_(diNameR,    ruleInIncomeNames);
    _applyValidationIfChanged_(diAcctR,    ruleInAccounts);
  }

  // Month sheet validations
  _applyValidationIfChanged_(iNameR, ruleInIncomeNames);
  _applyValidationIfChanged_(iAcctR, ruleInAccounts);

  _applyValidationIfChanged_(eNameR, ruleInExpenseNames);
  _applyValidationIfChanged_(eCatR,  ruleInCategories);
  _applyValidationIfChanged_(eAcctR, ruleInAccounts);

  _applyValidationIfChanged_(sNameR, ruleInSavingsNames);
  _applyValidationIfChanged_(bNameR, ruleInAccounts);
}




function _setRowWarning_(sh, row, on){
  const colH = 8; // column H
  const warn = sh.getRange(row, colH);
  warn.setValue(on ? "⚠️" : "");
  if (on) {
    warn.setNote("Duplicate account name");
  } else {
    if (String(warn.getNote()||"").indexOf("Duplicate account name") === 0){
      warn.setNote("");
    }
  }
}

// We no longer put a note on the NAME cell at all; only set/clear column H.
function _markDupeCell_(sh, row, col){
  _setRowWarning_(sh, row, true);
}

function _clearDupeCell_(sh, row, col){
  _setRowWarning_(sh, row, false);
}

/** Non-destructive dupe pass: never colors or styles cells */
function _preventDuplicateAccounts_(sh, rowStart, rowEnd, colName){
  const cIdx   = _.c(colName);
  const height = rowEnd - rowStart + 1;
  const names  = sh.getRange(rowStart, cIdx, height, 1).getValues().map(r=>String(r[0]||"").trim());
  const seen   = new Map();

  // Clear notes/⚠️ where unique or blank
  for (let i=0;i<height;i++){
    const nm = names[i];
    const isUnique = nm && !names.some((v, j)=> j!==i && v && v.toUpperCase()===nm.toUpperCase());
    if (!nm || isUnique) _clearDupeCell_(sh, rowStart+i, cIdx);
  }

  // Mark current dupes (note + ⚠️ only)
  for (let i=0;i<height;i++){
    const nm = names[i]; if (!nm) continue;
    const key = nm.toUpperCase();
    if (seen.has(key)){
      _markDupeCell_(sh, rowStart+i, cIdx);
      _markDupeCell_(sh, seen.get(key), cIdx);
    } else {
      seen.set(key, rowStart+i);
    }
  }
}

function CheckDuplicateAccounts_Auto(){
  const ss = __ss();;

  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    const B = CFG.bank;
    _preventDuplicateAccounts_(mSh, B.rowStart, B.rowEnd, B.colName);
  }

  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh){
    const A = CFG.datedAccountBalance.accounts;
    _preventDuplicateAccounts_(dSh, A.rowStart, A.rowEnd, A.colName);
  }
}

function ReapplyKeyFormats(){
  const ss = __ss();;

  // Month income table formats
  const mSh=ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    const I=CFG.monthly.incomeTable;
    _.fmtMoney(mSh.getRange(I.rowStart, _.c(I.colPred), I.rowEnd-I.rowStart+1, 1));
    _.fmtMoney(mSh.getRange(I.rowStart, _.c(I.colAct),  I.rowEnd-I.rowStart+1, 1));
  }

  // Dated accounts summary formats
  const dSh=ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh){
    const A=CFG.datedAccountBalance.accounts;
    _.fmtMoney(dSh.getRange(A.rowStart, _.c(A.colIncome),      A.rowEnd-A.rowStart+1, 1));
    _.fmtMoney(dSh.getRange(A.rowStart, _.c(A.colExpense),     A.rowEnd-A.rowStart+1, 1));
    _.fmtMoney(dSh.getRange(A.rowStart, _.c(A.colTransfers),   A.rowEnd-A.rowStart+1, 1));
    _.fmtMoney(dSh.getRange(A.rowStart, _.c(A.colPredBalance), A.rowEnd-A.rowStart+1, 1));
  }
}
function _groupWidth(i){ const g = CFG.monthly.calendar.dayColGroups[i]; return _.c(g[1]) - _.c(g[0]) + 1; }
function _ensureSheetSize(sh, minRows, minCols){
  if (sh.getMaxRows()    < minRows) sh.insertRowsAfter(sh.getMaxRows(), minRows - sh.getMaxRows());
  if (sh.getMaxColumns() < minCols) sh.insertColumnsAfter(sh.getMaxColumns(), minCols - sh.getMaxColumns());
}




function _overflowKeyForA1(a1) {
  return "CAL_OVERFLOW_" + a1;
}
function needsAccountPopulation(sh, tableConfig) {
  const accounts = sh.getRange(tableConfig.rowStart, _.c(tableConfig.colAcct), 
                               tableConfig.rowEnd - tableConfig.rowStart + 1, 1).getValues();
  for (const row of accounts) { if (String(row[0] || "").trim() !== "") return false; }
  return true;
}
function normalizeAccountName(name) { return String(name || "").trim().toUpperCase(); }

/* ============================= LISTS + VALIDATIONS ============================= */
function Lists_CollectAll(){
  const ss = __ss();;
  const accounts = new Set();
  const expenseNames = new Set();
  const categories = new Set();
  const transferNames = new Set();

  // === ACCOUNTS ===
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh) {
    const mtI = CFG.monthly.incomeTable;
    mSh.getRange(mtI.rowStart, _.c(mtI.colAcct), mtI.rowEnd-mtI.rowStart+1, 1)
      .getValues().forEach(([v])=>{ v=String(v||"").trim(); if(v) accounts.add(v); });

    const mtE = CFG.monthly.expenseTable;
    mSh.getRange(mtE.rowStart, _.c(mtE.colAcct), mtE.rowEnd-mtE.rowStart+1, 1)
      .getValues().forEach(([v])=>{ v=String(v||"").trim(); if(v) accounts.add(v); });

    const B = CFG.bank;
    mSh.getRange(B.rowStart, _.c(B.colName), B.rowEnd-B.rowStart+1, 1)
      .getValues().forEach(([v])=>{ v=String(v||"").trim(); if(v) accounts.add(v); });
  }

  try { INCOME.readMaster().forEach(r => { const v=String(r.acct||"").trim(); if(v) accounts.add(v); }); } catch(_){}
  try { EXPENSES.readMaster().forEach(r => { const v=String(r.acct||"").trim(); if(v) accounts.add(v); }); } catch(_){}

  const TP = CFG.transfersPayments, tpSh = ss.getSheetByName(TP.sheet);
  if (tpSh) {
    const rows = TP.rowEnd-TP.rowStart+1, width = _.c(TP.colToAcct)-_.c(TP.colName)+1;
    const data = tpSh.getRange(TP.rowStart, _.c(TP.colName), rows, width).getValues();
    const iFrom = _.c(TP.colFromAcct)-_.c(TP.colName);
    const iTo   = _.c(TP.colToAcct)  -_.c(TP.colName);
    data.forEach(r => {
      const a=String(r[iFrom]||"").trim(); if(a) accounts.add(a);
      const b=String(r[iTo]  ||"").trim(); if(b) accounts.add(b);
    });
  }

  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh) {
    const A = CFG.datedAccountBalance.accounts;
    dSh.getRange(A.rowStart, _.c(A.colName), A.rowEnd-A.rowStart+1, 1)
      .getValues().forEach(([v])=>{ v=String(v||"").trim(); if(v) accounts.add(v); });
  }

  // === EXPENSE NAMES ===
  try { EXPENSES.readMaster().forEach(r => { const v=String(r.name||"").trim(); if(v) expenseNames.add(v); }); } catch(_){}

  // === TRANSFER NAMES ===
  if (tpSh) {
    const rows = TP.rowEnd-TP.rowStart+1;
    const data = tpSh.getRange(TP.rowStart, _.c(TP.colName), rows, 1).getValues();
    data.forEach(([v])=>{ v=String(v||"").trim(); if(v) transferNames.add(v); });
  }

  // === CATEGORIES ===
  try { EXPENSES.readMaster().forEach(r => { const v=String(r.cat||"").trim(); if(v) categories.add(v); }); } catch(_){}
  if (mSh) {
    const mtE = CFG.monthly.expenseTable;
    mSh.getRange(mtE.rowStart, _.c(mtE.colCat), mtE.rowEnd-mtE.rowStart+1, 1)
      .getValues().forEach(([v])=>{ v=String(v||"").trim(); if(v) categories.add(v); });
  }
categories.add("Credit Card Payment");

  const savingsNames = new Set();
try { 
  SAVINGS.readMain().forEach(r => { 
    const v=String(r.name||"").trim(); 
    if(v) savingsNames.add(v); 
  }); 
} catch(_){}


  return {
  accounts: Array.from(accounts).sort((a,b)=>a.localeCompare(b)),
  expenseNames: Array.from(expenseNames).sort((a,b)=>a.localeCompare(b)),
  categories: Array.from(categories).sort((a,b)=>a.localeCompare(b)),
  transferNames: Array.from(transferNames).sort((a,b)=>a.localeCompare(b)),
  incomeNames: _listsCollectIncomeNames_(),
  savingsNames: Array.from(savingsNames).sort((a,b)=>a.localeCompare(b))  // ← NEW
};
}
// NEW: provide missing collector the rest of the code uses.
function Accounts_CollectNames(){ return Lists_CollectAll().accounts; }

function _ensureListSheet(){
 const ss = __ss(); name = "_Lists";
  let sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); sh.hideSheet(); }

  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());
  if (sh.getMaxRows()    < 2) sh.insertRowsAfter(sh.getMaxRows(), 2 - sh.getMaxRows());

  sh.getRange("A1").setValue("Accounts");
  sh.getRange("C1").setValue("Expense Names");
sh.getRange("J1").setValue("Income Names");
  sh.getRange("F1").setValue("Categories");
  sh.getRange("H1").setValue("Transfer Names");
  return sh;
}

function _writeAllLists(lists){
  const sh = _ensureListSheet();

  // Headers
  sh.getRange("A1").setValue("Accounts");
  sh.getRange("C1").setValue("Expense Names");
  sh.getRange("F1").setValue("Categories");
  sh.getRange("H1").setValue("Transfer Names");
  sh.getRange("J1").setValue("Income Names");
  sh.getRange("L1").setValue("Savings Names");  // ← NEW

  // Find max length
  const maxRows = Math.max(
    lists.accounts?.length||0,
    lists.expenseNames?.length||0,
    lists.categories?.length||0,
    lists.transferNames?.length||0,
    lists.incomeNames?.length||0,
    lists.savingsNames?.length||0  // ← NEW
  ) + 10;  // Add buffer

  // Clear old data (from row 2 down)
  if (maxRows > 0){
    sh.getRange(2, 1, maxRows, 1).clearContent();  // A
    sh.getRange(2, 3, maxRows, 1).clearContent();  // C
    sh.getRange(2, 6, maxRows, 1).clearContent();  // F
    sh.getRange(2, 8, maxRows, 1).clearContent();  // H
    sh.getRange(2,10, maxRows, 1).clearContent();  // J
    sh.getRange(2,12, maxRows, 1).clearContent();  // L ← NEW
  }

  // Write from row 2 (right below header)
  if (lists.accounts?.length){
    sh.getRange(2, 1, lists.accounts.length, 1).setValues(lists.accounts.map(v=>[v]));
  }
  if (lists.expenseNames?.length){
    sh.getRange(2, 3, lists.expenseNames.length, 1).setValues(lists.expenseNames.map(v=>[v]));
  }
  if (lists.categories?.length){
    sh.getRange(2, 6, lists.categories.length, 1).setValues(lists.categories.map(v=>[v]));
  }
  if (lists.transferNames?.length){
    sh.getRange(2, 8, lists.transferNames.length, 1).setValues(lists.transferNames.map(v=>[v]));
  }
  if (lists.incomeNames?.length){
    sh.getRange(2,10, lists.incomeNames.length, 1).setValues(lists.incomeNames.map(v=>[v]));
  }
  if (lists.savingsNames?.length){  // ← NEW
    sh.getRange(2,12, lists.savingsNames.length, 1).setValues(lists.savingsNames.map(v=>[v]));
  }

  return {
    accounts:      sh.getRange(2, 1, Math.max(1, (lists.accounts||[]).length), 1),
    expenseNames:  sh.getRange(2, 3, Math.max(1, (lists.expenseNames||[]).length), 1),
    categories:    sh.getRange(2, 6, Math.max(1, (lists.categories||[]).length), 1),
    transferNames: sh.getRange(2, 8, Math.max(1, (lists.transferNames||[]).length), 1),
    incomeNames:   sh.getRange(2,10, Math.max(1, (lists.incomeNames||[]).length), 1),
    savingsNames:  sh.getRange(2,12, Math.max(1, (lists.savingsNames||[]).length), 1)  // ← NEW
  };
}


function _applyValidationToRange(rng, listRange){
  const rule = SpreadsheetApp.newDataValidation()
   .requireValueInRange(listRange, true)
    .setAllowInvalid(true)
    .build();
  rng.setDataValidation(rule);
}
function _applyValidationToRange_NonDestructive(rng, listRange){
  if (!rng) return;                       // <- guard
  if (!listRange) return;                 // <- guard
  const rows = rng.getNumRows?.() || 0;   // <- safe calls
  const cols = rng.getNumColumns?.() || 0;
  if (!rows || !cols) return;

  const current = rng.getDataValidations?.() || null;
  const out = [];
  let needsWrite = false;

  const newRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange, true) // dropdown + allow search
    .setAllowInvalid(true)                // user can free-type (we capture via onEdit)
    .build();

  for (let r=0; r<rows; r++){
    const line = [];
    for (let c=0; c<cols; c++){
      const rule = current?.[r]?.[c] || null;
      if (!rule) { line.push(newRule); needsWrite = true; }
      else {
        const same = String(rule) === String(newRule);
        line.push(same ? rule : newRule);
        if (!same) needsWrite = true;
      }
    }
    out.push(line);
  }
  if (needsWrite) rng.setDataValidations(out);
}


function _writeAccountsList(names){
  const sh = _ensureListSheet();
  const last = Math.max(2, sh.getLastRow());
  sh.getRange(2,1,last-1,1).clearContent();
  if (names.length){ sh.getRange(2,1,names.length,1).setValues(names.map(n=>[n])); }
  return sh.getRange(2,1,Math.max(1,names.length),1);
}

function _mapMonthIncomeMeta_(){
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const map = new Map();
  if (!mSh) return map;

  const I = CFG.monthly.incomeTable;
  const width = _.c(I.colAcct) - _.c(I.colName) + 1;
  const data  = mSh.getRange(I.rowStart, _.c(I.colName), I.rowEnd - I.rowStart + 1, width).getValues();

  const idxName = 0;
  const idxFreq = _.c(I.colFreq) - _.c(I.colName);  // Month Income table freq col
  const idxAcct = _.c(I.colAcct) - _.c(I.colName);

  for (const r of data){
    const nm   = String(r[idxName]||"").trim();
    const freq = String(r[idxFreq]||"").trim();
    const acct = String(r[idxAcct]||"").trim();
    if (nm) map.set(nm, { acct, freq });
  }
  return map;
}

// One-time memo for _.c
// One-time memoization for _.c()
(function(){
  const _cache = new Map();
  const oldC = _.c;
  _.c = function(L){
    if (typeof L !== "string") return oldC(L);
    const key = L.toUpperCase();
    if (_cache.has(key)) return _cache.get(key);
    const v = oldC(key);
    _cache.set(key, v);
    return v;
  };
})();

function _dateKey(d){
  if (!(d instanceof Date)) d = new Date(d);
  const y = d.getFullYear();
  const m = d.getMonth()+1;
  const day = d.getDate();
  return y + "-" + (m<10?"0":"")+m + "-" + (day<10?"0":"")+day;
}


/* ============================= BOOL HELPERS ============================= */
function _boolCell_(v){
  // Normalizes checkbox / TRUE/FALSE / 1/0 / Y/N to a proper boolean
  const t = String(v).trim().toUpperCase();
  if (v === true || v === 1 || t === "TRUE" || t === "Y" || t === "YES") return true;
  if (v === false || v === 0 || t === "FALSE" || t === "N" || t === "NO") return false;
  return false;
}

// Replaces a bad Spreadsheet.getRange(A1) usage.
// Accepts either "A1" (current active sheet) or "Sheet Name!A1".
function _val_fromRange_(a1) {
  const ss = __ss();
  let sh = ss.getActiveSheet();
  let ref = String(a1 || "").trim();
  if (!ref) return null;

  // Support "Sheet Name!A1" syntax
  const bang = ref.indexOf("!");
  if (bang > 0) {
    const sheetName = ref.slice(0, bang);
    ref = ref.slice(bang + 1);
    const s = ss.getSheetByName(sheetName);
    if (s) sh = s;
  }

  try {
    return sh.getRange(ref).getValue();
  } catch (e) {
    // As a convenience, fall back to named range if provided
    try { return ss.getRangeByName(ref).getValue(); } catch(_) {}
    Logger.log("_val_fromRange_ failed for: " + a1 + " → " + e);
    return null;
  }
}


function _Bank_FormatMoneyColumns_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;
  const B = CFG.bank;

  const a1s = [
    sh.getRange(B.rowStart, _.c(B.colPredBegin),  B.rowEnd-B.rowStart+1, 1).getA1Notation(),
    sh.getRange(B.rowStart, _.c(B.colExpTotal),   B.rowEnd-B.rowStart+1, 1).getA1Notation(),
    sh.getRange(B.rowStart, _.c(B.colIncomeTotal),B.rowEnd-B.rowStart+1, 1).getA1Notation(),
    sh.getRange(B.rowStart, _.c(B.colTransfer),   B.rowEnd-B.rowStart+1, 1).getA1Notation(),
    sh.getRange(B.rowStart, _.c(B.colPredEnd),    B.rowEnd-B.rowStart+1, 1).getA1Notation(),
    sh.getRange(B.rowStart, _.c(B.colActEnd),     B.rowEnd-B.rowStart+1, 1).getA1Notation()
  ];

  sh.getRangeList(a1s).setNumberFormat(CFG.formats.money);
}

// Return first occurrence (ms epoch) of freq within (year=y, month=m 1..12)
// anchored at `start`. Also returns the period length in ms for iteration.
// If there is no hit in that month, returns { firstTime: -1, periodMs: 0 }.
function _firstOccurrenceTimeInMonth(freq, start, y, m){
  const F = _.normFreq(freq);
  const t0 = (start ? new Date(start) : new Date(y,0,1)).getTime();
  const monthStart = new Date(y, m-1, 1).getTime();
  const monthEnd   = new Date(y, m,   1).getTime(); // exclusive
  if (F === "ONETIME"){
    const s = start && start.getFullYear()===y && (start.getMonth()+1)===m ? start.getTime() : -1;
    return { firstTime: s, periodMs: 0 };
  }
  if (F === "MONTHLY"){
    if (t0 >= monthEnd) return { firstTime: -1, periodMs: 0 };
    // same day-of-month as anchor, clamped to month length
    const d = _.safeDate(y, m, (start||new Date()).getDate());
    return { firstTime: d.getTime(), periodMs: 30*86400000 /*unused*/ };
  }
  // WEEKLY / BIWEEKLY
  const pMs = (F==="WEEKLY" ? 7 : 14) * 86400000;
  if (monthEnd <= t0) return { firstTime: -1, periodMs: 0 };
  const kFirst = Math.ceil((Math.max(monthStart, t0) - t0) / pMs);
  const first  = t0 + kFirst * pMs;
  return (first >= monthEnd) ? { firstTime: -1, periodMs: 0 }
                             : { firstTime: first, periodMs: pMs };
}

/* ===== Public list helpers ===== */

function _mapMonthIncomeAcct_(){
 const ss = __ss(); mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const map = new Map(); if (!mSh) return map;
  const t = CFG.monthly.incomeTable;
  const width = _.c(t.colAcct) - _.c(t.colName) + 1;
  const data = mSh.getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, width).getValues();
  const iName = 0, iAcct = _.c(t.colAcct) - _.c(t.colName);
  for (const r of data){
    const nm   = String(r[iName]||"").trim();
    const acct = String(r[iAcct]||"").trim();
    if (nm && acct && !map.has(nm)) map.set(nm, acct);
  }
  return map;
}

function _inferIncomeAcct_(name){
  try {
    // Prefer Month Income mapping
    if (typeof _mapMonthIncomeAcct_ === "function") {
      const m = _mapMonthIncomeAcct_();
      if (m && m.has(name)) return m.get(name);
    }

    // Fall back to Income Master
    if (typeof INCOME !== "undefined" && INCOME.readMaster) {
      const rows = INCOME.readMaster();
      const hit = rows.find(r => String(r.name||"").trim() === String(name||"").trim() && r.acct);
      if (hit && hit.acct) return hit.acct;
    }
  } catch(_) {}
  return ""; // don’t guess
}



function DatedIncome_UpsertFromPaychecks(){
  const ss  = SpreadsheetApp.getActive();
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  const pSh = ss.getSheetByName(CFG.income.paychecks.sheet);
  if (!dSh || !pSh) return;

  // H2 target month
  const targetDate = _.toDate(dSh.getRange(CFG.datedAccountBalance.targetDate).getValue());
  if (!targetDate){ SpreadsheetApp.getActive().toast("❗ Set a valid target date (H2) on Dated sheet", "Dated Income", 4); return; }
  targetDate.setHours(0,0,0,0);
  const y = targetDate.getFullYear();
  const m = targetDate.getMonth() + 1;

  // name → account from Month table

  // read Paychecks table
  const P = CFG.income.paychecks;
  const pRows = P.rowEnd - P.rowStart + 1;
  const pW    = _.c(P.colAct) - _.c(P.colName) + 1;
  const pVals = pSh.getRange(P.rowStart, _.c(P.colName), pRows, pW).getValues();
const iNm = 0, iDt = _.c(P.colDate) - _.c(P.colName), iPr = _.c(P.colPred) - _.c(P.colName);
const iFq = _.c(P.colFreq) - _.c(P.colName); // paycheck freq index
const metaByName = _mapMonthIncomeMeta_();   // month-level name→{acct,freq}
const acctMap    = _mapMonthIncomeAcct_ ? _mapMonthIncomeAcct_() : new Map(); // fallback

  // current dated income (to preserve Include)
  const DI   = CFG.datedAccountBalance.income;
  const dN   = DI.rowEnd - DI.rowStart + 1;
  const dW   = _.c(DI.colInclude) - _.c(DI.colName) + 1;
  const dRng = dSh.getRange(DI.rowStart, _.c(DI.colName), dN, dW);
  const dOld = dRng.getValues();

const oNm=0, oDt=_.c(DI.colDate)-_.c(DI.colName), oPr=_.c(DI.colPred)-_.c(DI.colName),
      oFrq=_.c(DI.colFreq)-_.c(DI.colName),            // DI freq index
      oAcnt=_.c(DI.colAcct)-_.c(DI.colName),
      oInc=_.c(DI.colInclude)-_.c(DI.colName);

  const existing = new Map();
  const keyOf = (nm, dt, acct) => `${String(nm||"").trim()}|${_.toDate(dt)?.getTime()||""}|${String(acct||"").trim()}`;
  for (const row of dOld){
    const nm   = String(row[oNm]||"").trim();
    const dt   = _.toDate(row[oDt]);
    const acct = String(row[oAcnt]||"").trim();
    if (!nm || !dt) continue;
    existing.set(keyOf(nm, dt, acct), _boolCell_(row[oInc]));
  }

  // build rows for only H2 month
  const rows = [];
  const seen = new Set();
  for (const r of pVals){
    const nm = String(r[iNm]||"").trim(); if (!nm) continue;
    const dt = _.toDate(r[iDt]);          if (!dt || dt.getFullYear()!==y || (dt.getMonth()+1)!==m) continue;
const pred = _.n(r[iPr]);
const freq = String(r[iFq]||"").trim();                // <-- NEW: get paycheck freq label
const acct = (acctMap.get(nm) || _inferIncomeAcct_(nm) || "");  // <-- fallback auto acct
    const k = keyOf(nm, dt, acct);
    if (seen.has(k)) continue;
    seen.add(k);

    const include = existing.has(k) ? existing.get(k) : (dt <= targetDate);
rows.push({ nm, dt, pred, acct, include, freq }); // ← add freq
  }

  // write (clear remainder)
  const out = Array.from({length:dN}, ()=>Array(dW).fill(""));
  for (let i=0;i<Math.min(rows.length, dN); i++){
    const r = rows[i];
out[i][oNm]   = r.nm;
out[i][oDt]   = r.dt;
out[i][oFrq] = r.freq || "";
out[i][oPr]   = r.pred>0 ? r.pred : "";
out[i][oAcnt] = r.acct;
out[i][oInc]  = !!r.include;

  }
  dRng.setValues(out);

  // restore controls & formats
  _.fmtDate( dSh.getRange(DI.rowStart, _.c(DI.colDate), dN, 1) );
  _.fmtMoney(dSh.getRange(DI.rowStart, _.c(DI.colPred), dN, 1) );
}
function DatedIncome_RecheckIncludeForH2(){
  const ss  = SpreadsheetApp.getActive();
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet); if(!dSh) return;

  const DI = CFG.datedAccountBalance.income;
  const rows = DI.rowEnd - DI.rowStart + 1;
  const w    = _.c(DI.colInclude) - _.c(DI.colName) + 1;
  const rng  = dSh.getRange(DI.rowStart, _.c(DI.colName), rows, w);
  const data = rng.getValues();

  const target = _.toDate(dSh.getRange(CFG.datedAccountBalance.targetDate).getValue());
  if (!target) return; target.setHours(0,0,0,0);

  const oNm = 0, oDt = _.c(DI.colDate)-_.c(DI.colName), oInc = _.c(DI.colInclude)-_.c(DI.colName);
  let touched = false;

  for (let i=0;i<data.length;i++){
    const nm = String(data[i][oNm]||"").trim();
    const dt = _.toDate(data[i][oDt]);
    if (!nm || !dt) continue;
    const should = dt.getTime() <= target.getTime();
    if (!!data[i][oInc] !== should){ data[i][oInc] = should; touched = true; }
  }
  if (touched){
    rng.setValues(data);
  }
}

function Lists_AddAccount(name){
  const nm = String(name||"").trim(); if (!nm) return false;
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),1).getValues().map(([v])=>String(v||"").trim());
  if (vals.some(v=>v.toUpperCase()===nm.toUpperCase())) return false;
  sh.getRange(Math.max(2, sh.getLastRow()+1),1).setValue(nm);
  Lists_ApplyEverywhere();
  return true;
}
function Lists_AddExpenseName(name){
  const nm = String(name||"").trim(); if (!nm) return false;
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,3,Math.max(0,sh.getLastRow()-1),1).getValues().map(([v])=>String(v||"").trim());
  if (vals.some(v=>v.toUpperCase()===nm.toUpperCase())) return false;
  sh.getRange(Math.max(2, sh.getLastRow()+1),3).setValue(nm);
  Lists_ApplyEverywhere();
  return true;
}
function _listsCollectIncomeNames_(){
  const ss = __ss();;
  const set = new Set();
  try { // Income master
    const M = CFG.income.master, sh = ss.getSheetByName(M.sheet);
    if (sh){
      const n = M.rowEnd - M.rowStart + 1;
      sh.getRange(M.rowStart, _.c(M.colName), n, 1).getValues()
        .forEach(([v])=>{ v=String(v||"").trim(); if(v) set.add(v); });
    }
  } catch(_){}
  try { // Month income table
    const I = CFG.monthly.incomeTable, sh = ss.getSheetByName(CFG.singleMonth.sheet);
    if (sh){
      const n = I.rowEnd - I.rowStart + 1;
      sh.getRange(I.rowStart, _.c(I.colName), n, 1).getValues()
        .forEach(([v])=>{ v=String(v||"").trim(); if(v) set.add(v); });
    }
  } catch(_){}
  try { // Paychecks block
    const P = CFG.income.paychecks, sh = ss.getSheetByName(P.sheet);
    if (sh){
      const n = P.rowEnd - P.rowStart + 1;
      sh.getRange(P.rowStart, _.c(P.colName), n, 1).getValues()
        .forEach(([v])=>{ v=String(v||"").trim(); if(v) set.add(v); });
    }
  } catch(_){}
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

function Lists_AddTransferName(name){
  const nm = String(name||"").trim(); if (!nm) return false;
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,8,Math.max(0,sh.getLastRow()-1),1).getValues().map(([v])=>String(v||"").trim());
  if (vals.some(v=>v.toUpperCase()===nm.toUpperCase())) return false;
  sh.getRange(Math.max(2, sh.getLastRow()+1),8).setValue(nm);
  Lists_ApplyEverywhere();
  return true;
}
function Lists_AddCategory(name){
  const nm = String(name||"").trim(); if (!nm) return false;
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,6,Math.max(0,sh.getLastRow()-1),1).getValues().map(([v])=>String(v||"").trim());
  if (vals.some(v=>v.toUpperCase()===nm.toUpperCase())) return false;
  sh.getRange(Math.max(2, sh.getLastRow()+1),6).setValue(nm);
  Lists_ApplyEverywhere();
  return true;
}

function Lists_AddIncomeName(name){
  const nm = String(name||"").trim(); if (!nm) return false;
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,10,Math.max(0,sh.getLastRow()-1),1).getValues().map(([v])=>String(v||"").trim());
  if (vals.some(v=>v.toUpperCase()===nm.toUpperCase())) return false;
  sh.getRange(Math.max(2, sh.getLastRow()+1),10).setValue(nm);
  Lists_ApplyEverywhere();
  return true;
}

function _fillBlanksWithList(sh, rowStart, rowEnd, colLtr, names){
  const n = rowEnd - rowStart + 1;
  const vals = sh.getRange(rowStart, _.c(colLtr), n, 1).getValues();
  let i = 0;
  for (let r=0; r<n; r++){
    const cur = String(vals[r][0]||"").trim();
    if (!cur && i < names.length){ vals[r][0] = names[i++]; }
  }
  sh.getRange(rowStart, _.c(colLtr), n, 1).setValues(vals);
}

function _listsReadSet(){
  const sh = _ensureListSheet();
  const vals = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1),1).getValues();
  const set = new Set();
  for (const [v] of vals){ const s = String(v||"").trim(); if (s) set.add(s); }
  return set;
}

/* ===== ACCOUNTS: sync & validations ===== */
function _syncNamesIntoBlock(sh, rowStart, rowEnd, colLtr, allNames){
  const nRows = rowEnd - rowStart + 1;
  if (nRows <= 0) return;

  const col = _.c(colLtr);
  const vals = sh.getRange(rowStart, col, nRows, 1).getValues().map(r => String(r[0] || "").trim());
  const have = new Set(vals.filter(v => v).map(v => v.toUpperCase()));

  const toInsert = [];
  for (const nm of allNames){ if (!have.has(nm.toUpperCase())) toInsert.push(nm); }
  if (!toInsert.length) return;

  let insIdx = 0;
  for (let r = 0; r < nRows && insIdx < toInsert.length; r++){
    if (!vals[r]) { sh.getRange(rowStart + r, col).setValue(toInsert[insIdx++]); }
  }
}

function Accounts_SyncBankBlocks(){
  const ss = __ss();;
  const names = Accounts_CollectNames();
  if (!names.length) return;

  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    _syncNamesIntoBlock(mSh, CFG.bank.rowStart, CFG.bank.rowEnd, CFG.bank.colName, names);
  }

  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh){
    const A = CFG.datedAccountBalance.accounts;
    _syncNamesIntoBlock(dSh, A.rowStart, A.rowEnd, A.colName, names);
  }
}

function Accounts_ListRangeForValidation(){
  const sh = _ensureListSheet();
  if (sh.getMaxRows() < 2000) sh.insertRowsAfter(sh.getMaxRows(), 2000 - sh.getMaxRows());
  return sh.getRange(2, 1, 1999, 1); // _Lists!A2:A2000
}

function Accounts_ReapplyValidationsOnly(){
  const ss = __ss();;
  const listRange = Accounts_ListRangeForValidation(); // _Lists!A2:A

  function apply(rng){ _applyValidationToRange_NonDestructive(rng, listRange); }

  // MONTH
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    const I = CFG.monthly.incomeTable, E = CFG.monthly.expenseTable, B = CFG.bank;
    apply(mSh.getRange(I.rowStart, _.c(I.colAcct), I.rowEnd - I.rowStart + 1, 1));
    apply(mSh.getRange(E.rowStart, _.c(E.colAcct), E.rowEnd - E.rowStart + 1, 1));
    apply(mSh.getRange(B.rowStart, _.c(B.colName), B.rowEnd - B.rowStart + 1, 1));
  }

  // DATED AB
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh){
    const D = CFG.datedAccountBalance;
    apply(dSh.getRange(D.accounts.rowStart, _.c(D.accounts.colName), D.accounts.rowEnd - D.accounts.rowStart + 1, 1));
    apply(dSh.getRange(D.expenses.rowStart, _.c(D.expenses.colAccount), D.expenses.rowEnd - D.expenses.rowStart + 1, 1));
    apply(dSh.getRange(D.transfers.rowStart, _.c(D.transfers.colFromAcct), D.transfers.rowEnd - D.transfers.rowStart + 1, 1));
    apply(dSh.getRange(D.transfers.rowStart, _.c(D.transfers.colToAcct), D.transfers.rowEnd - D.transfers.rowStart + 1, 1));
  }

  // Income/Expense masters (acct)
  const incSh = ss.getSheetByName(CFG.income.master.sheet);
  if (incSh){
    const M = CFG.income.master;
    apply(incSh.getRange(M.rowStart, _.c(M.colAcct), M.rowEnd - M.rowStart + 1, 1));
  }
  const expSh = ss.getSheetByName(CFG.expenses.master.sheet);
  if (expSh){
    const M = CFG.expenses.master;
    apply(expSh.getRange(M.rowStart, _.c(M.colAcct), M.rowEnd - M.rowStart + 1, 1));
  }
}

  
function Accounts_ApplyEverywhere(){ return Lists_ApplyEverywhere(); }


function Lists_ApplyEverywhere(){
  const ss = __ss();;
  const lists = Lists_CollectAll();
  const ranges = _writeAllLists(lists);

  function apply(sheet, rowStart, rowEnd, colLtr, listType){
    if (!sheet) return;
    const rng = sheet.getRange(rowStart, _.c(colLtr), rowEnd - rowStart + 1, 1);
    _applyValidationToRange_NonDestructive(rng, ranges[listType]);
  }

  // === MONTH SHEET ===
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    const mtI = CFG.monthly.incomeTable;
    const mtE = CFG.monthly.expenseTable;
    const B   = CFG.bank;
    
    apply(mSh, mtI.rowStart, mtI.rowEnd, mtI.colName, 'incomeNames');
    apply(mSh, mtI.rowStart, mtI.rowEnd, mtI.colAcct, 'accounts');
    
    apply(mSh, mtE.rowStart, mtE.rowEnd, mtE.colName, 'expenseNames');
    apply(mSh, mtE.rowStart, mtE.rowEnd, mtE.colCat, 'categories');
    apply(mSh, mtE.rowStart, mtE.rowEnd, mtE.colAcct, 'accounts');
    
    apply(mSh, B.rowStart, B.rowEnd, B.colName, 'accounts');
  }

  // === DATED AB ===
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh){
    const D = CFG.datedAccountBalance;
    apply(dSh, D.accounts.rowStart,  D.accounts.rowEnd,  D.accounts.colName, 'accounts');
    apply(dSh, D.expenses.rowStart,  D.expenses.rowEnd,  D.expenses.colAccount, 'accounts');
    apply(dSh, D.transfers.rowStart, D.transfers.rowEnd, D.transfers.colFromAcct, 'accounts');
    apply(dSh, D.transfers.rowStart, D.transfers.rowEnd, D.transfers.colToAcct, 'accounts');
    apply(dSh, D.income.rowStart,    D.income.rowEnd,    D.income.colAcct, 'accounts');
  }

  // === TRANSFERS ===
  const TP = CFG.transfersPayments, tpSh = ss.getSheetByName(TP.sheet);
  if (tpSh){
    apply(tpSh, TP.rowStart, TP.rowEnd, TP.colName, 'transferNames');
    apply(tpSh, TP.rowStart, TP.rowEnd, TP.colFromAcct, 'accounts');
    apply(tpSh, TP.rowStart, TP.rowEnd, TP.colToAcct, 'accounts');
  }

  // === INCOME MASTER ===
  const incSh = ss.getSheetByName(CFG.income.master.sheet);
  if (incSh){
    const M = CFG.income.master;
    apply(incSh, M.rowStart, M.rowEnd, M.colName, 'incomeNames');
    apply(incSh, M.rowStart, M.rowEnd, M.colAcct, 'accounts');
  }

  // === EXPENSE MASTER ===
  const expSh = ss.getSheetByName(CFG.expenses.master.sheet);
  if (expSh){
    const M = CFG.expenses.master;
    apply(expSh, M.rowStart, M.rowEnd, M.colName, 'expenseNames');
    apply(expSh, M.rowStart, M.rowEnd, M.colCat, 'categories');
    apply(expSh, M.rowStart, M.rowEnd, M.colAcct, 'accounts');
  }

  // === PAYCHECKS ===
  const paySh = ss.getSheetByName(CFG.income.paychecks.sheet);
  if (paySh){
    const P = CFG.income.paychecks;
    apply(paySh, P.rowStart, P.rowEnd, P.colName, 'incomeNames');
  }
}
/* ============================= CALENDAR & OVERFLOW ============================= */
function clearOverflowPanels(){
  const props = PropertiesService.getDocumentProperties().getProperties();
  for (const key in props) {
    if (key.startsWith("CAL_OVERFLOW_")) {
      PropertiesService.getDocumentProperties().deleteProperty(key);
    }
  }
}

/* ============================= STATE ============================= */
var STATE = (function(){
  const SHEET="_State", KEY_LAST_YM="last_ym";
  function _ss(){ return SpreadsheetApp.getActive(); }
  function _ensure(){ const ss=_ss(); let sh=ss.getSheetByName(SHEET); if(!sh){ sh=ss.insertSheet(SHEET); sh.hideSheet(); sh.getRange(1,1).setValue("key"); sh.getRange(1,2).setValue("value"); } return sh; }
  function get(key){ const sh=_ensure(); const last=Math.max(2, sh.getLastRow()); const vals=sh.getRange(2,1,last-1,2).getValues(); for (const r of vals){ if (String(r[0]||"").trim()===key) return String(r[1]||""); } return ""; }
  function set(key,val){ const sh=_ensure(); const last=Math.max(2, sh.getLastRow()); for (let r=2;r<=last;r++){ if (String(sh.getRange(r,1).getValue()||"").trim()===key){ sh.getRange(r,2).setValue(String(val)); return; } } sh.getRange(last+1,1).setValue(key); sh.getRange(last+1,2).setValue(String(val)); }
  function ymStr(y,m){ return y+"-"+("0"+m).slice(-2); }
  return { get,set,ymStr, KEY_LAST_YM };
})();

function _expKey(nm, d, amt, acct){
  return `E|${String(nm||'').trim()}|${_dateKey(d)}|${Number(amt||0).toFixed(2)}|${String(acct||'').trim()}`;
}
function _xferKey(nm, d, amt, from, to){
  return `X|${String(nm||'').trim()}|${_dateKey(d)}|${Number(amt||0).toFixed(2)}|${String(from||'').trim()}|${String(to||'').trim()}`;
}
function _incKey(nm, d, amt, acct){
  const ds = d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  return `INC|${nm||''}|${ds}|${Number(amt||0).toFixed(2)}|${acct||''}`;
}



/* ============================= EXPENSES (master → month + occurrences) ============================= */
var EXPENSES = (function () {
  const M  = CFG.expenses.master;
  const Mm = CFG.expenses.monthly;
  const ss = __ss();;

  function readMaster() {
    const sh = ss.getSheetByName(M.sheet); if (!sh) return [];
    _ensureSheetSize(sh, M.rowEnd, _.c(M.colAcct));

    const width = _.c(M.colAcct) - _.c(M.colName) + 1;
    const n     = M.rowEnd - M.rowStart + 1;
    const data  = sh.getRange(M.rowStart, _.c(M.colName), n, width).getValues();
    const idx   = L => _.c(L) - _.c(M.colName);

    const out = [];
    for (let i = 0; i < n; i++) {
      const r  = data[i];
      const nm = String(r[idx(M.colName)] || "").trim();
      if (!nm) continue;

      out.push({
        name: nm,
        start: _.toDate(r[idx(M.colStart)]),
        freq:  _.normFreq(r[idx(M.colFreq)]),
        cat:   String(r[idx(M.colCat)]  || "").trim(),
        amt:   _.n(r[idx(M.colPred)]),
        acct:  String(r[idx(M.colAcct)] || "").trim()
      });
    }
    return out;
  }

  function computeOccurrences(rows, y, m) {
    const out = [];
    for (const r of rows) {
      const F   = _.normFreq(r.freq);
      const amt = r.amt || 0;
      if (amt <= 0) continue;

      if (F === "ONETIME") {
        if (r.start && r.start.getFullYear() === y && (r.start.getMonth() + 1) === m) {
          out.push({ name: r.name, date: r.start, cat: r.cat || "", amt, acct: r.acct || "" });
        }
      } else {
        const anchor = r.start || new Date(y, 0, 1);
        const dates  = _.expand(y, anchor, null, F).filter(d => (d.getMonth() + 1) === m);
        for (const d of dates) {
          out.push({ name: r.name, date: d, cat: r.cat || "", amt, acct: r.acct || "" });
        }
      }
    }
    return out;
  }

// FAST populate for Expenses Month table (no _.expand, one write)
// FAST populate for Expenses Month table (now writes Date too)
function populateMonth(sh, year, mIdx, allowPast, masterRows){
  const T = CFG.expenses.monthly; if (!sh) return;
  const cap = T.rowEnd - T.rowStart + 1;

  // cache columns once
  const cName = _.c(T.colName);
  const cDate = _.c(T.colDate);
  const cCat  = _.c(T.colCat);
  const cPred = _.c(T.colPred);
  const cAct  = _.c(T.colAct);
  const cAcct = _.c(T.colAcct);
  const width = cAcct - cName + 1;

  const oName = 0;
  const oDate = cDate - cName;
  const oCat  = cCat  - cName;
  const oPred = cPred - cName;
  const oAct  = cAct  - cName;
  const oAcct = cAcct - cName;

  const out = Array.from({length: cap}, () => Array(width).fill(""));
  let w = 0;

  // prefilter once
  const valid = masterRows.filter(r => {
    const nm = String(r.name||"").trim();
    return nm && _.n(r.amt) > 0;
  });

  const monthEnd = new Date(year, mIdx, 1).getTime(); // exclusive

  for (const r of valid){
    if (w >= cap) break;

    const F     = _.normFreq(r.freq);
    const amt   = _.n(r.amt);
    const nm    = String(r.name||"").trim();
    const acct  = String(r.acct||"").trim();
    const cat   = String(r.cat||"").trim();
    const start = r.start || new Date(year,0,1);

    const { firstTime, periodMs } = _firstOccurrenceTimeInMonth(F, start, year, mIdx);
    if (firstTime < 0) continue;

    if (F === "ONETIME" || F === "MONTHLY"){
      if (w >= cap) break;
      out[w][oName] = nm;
      out[w][oDate] = new Date(firstTime);
      out[w][oCat]  = cat;
      out[w][oPred] = amt;
      out[w][oAct]  = "";
      out[w][oAcct] = acct;
      w++;
    } else {
      const occ = _occurrencesInMonth(F, start, year, mIdx);
      for (let k=0; k<occ && w<cap; k++){
        const t = firstTime + k*periodMs;
        if (t >= monthEnd) break;
        out[w][oName] = nm;
        out[w][oDate] = new Date(t);
        out[w][oCat]  = cat;
        out[w][oPred] = amt;
        out[w][oAct]  = "";
        out[w][oAcct] = acct;
        w++;
      }
    }
  }

  const rng = sh.getRange(T.rowStart, cName, cap, width);
  rng.clearContent();
  rng.setValues(out);

  if (w > 0) {
    sh.getRange(T.rowStart, cDate, w, 1).setNumberFormat(CFG.formats.date);
    sh.getRange(T.rowStart, cPred, w, 1).setNumberFormat(CFG.formats.money);
    sh.getRange(T.rowStart, cAct,  w, 1).setNumberFormat(CFG.formats.money);
  }
}



return { readMaster, populateMonth };
})();



/* ============================= SAVINGS ============================= */
function getMonthlyEquivalent(freq, contrib) {
  switch (_.normFreq(freq)) {
    case "WEEKLY":   return (contrib||0)*52/12;
    case "BIWEEKLY": return (contrib||0)*26/12;
    case "MONTHLY":  return (contrib||0);
    default:         return contrib||0;
  }
}
function calculateMonthlyPredicted(freq, contrib, year, month, startDate){
  const F = _.normFreq(freq);
  if (!contrib) return 0;
  
  if (F === "ONETIME") {
    return (startDate && startDate.getFullYear() === year && (startDate.getMonth() + 1) === month) ? contrib : 0;
  }
  
  // Use _.expand to get ACTUAL occurrences in this month
  const anchor = startDate || new Date(year, 0, 1);
  const allDates = _.expand(year, anchor, null, F);
  const monthDates = allDates.filter(d => (d.getMonth() + 1) === month);
  
  return contrib * monthDates.length;
}
function Savings_GetContribForMonth(goalName, mIdx){
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet); if(!sh) return 0;
  const S=CFG.savings.monthly, rows=S.rowEnd-S.rowStart+1;
  const names = sh.getRange(S.rowStart, _.c(S.colName), rows, 1).getValues().map(r=>String(r[0]||"").trim());
  const preds = sh.getRange(S.rowStart, _.c(S.colPred), rows, 1).getValues().map(r=>_.n(r[0]));
  const acts  = sh.getRange(S.rowStart, _.c(S.colAct),  rows, 1).getValues().map(r=>_.n(r[0]));
  const i = names.indexOf(String(goalName||"").trim());
  return i<0 ? 0 : (acts[i]>0 ? acts[i] : preds[i]||0);
}
function Savings_SumContribThroughMonth(goalName, throughM){ return Savings_GetContribForMonth(goalName, throughM); }

function Savings_UpdateEmergencyFundTile() {
  const ss = __ss();;
  const sSh = ss.getSheetByName(CFG.savings.sheet);
  if (!sSh) return;
  
  // Get CURRENT MONTH expenses (not previous)
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!mSh) return;
  
  const expAEP = _.sumAEP_block(
    mSh, 
    CFG.monthly.expenseTable.rowStart, 
    CFG.monthly.expenseTable.rowEnd, 
    CFG.monthly.expenseTable.colPred, 
    CFG.monthly.expenseTable.colAct
  );
  
  // 6 months of expenses
  const emergencyTarget = expAEP * 6;
  
  sSh.getRange(CFG.savings.tiles.emergencyFundTarget)
    .setValue(emergencyTarget)
    .setNumberFormat(CFG.formats.money);
}

function Savings_RecomputeMonthFromMonthOnly() {
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const savingsSh = ss.getSheetByName(CFG.savings.sheet);
  if (!mSh || !savingsSh) return;

  const M = CFG.savings.monthly;
  const MASTER = CFG.savings.table;
  const rows = M.rowEnd - M.rowStart + 1;

  // Read Month block once (Name..Finish)
  const left  = _.c(M.colName);
  const right = _.c(M.colFinish);
  const width = right - left + 1;
  const data  = mSh.getRange(M.rowStart, left, rows, width).getValues();

  const iName = _.c(M.colName)  - left;
  const iGoal = _.c(M.colGoal)  - left;
  const iStart= _.c(M.colStart) - left;
  const iPred = _.c(M.colPred)  - left;
  const iAct  = _.c(M.colAct)   - left;

  // Read master once
  const mLeft = _.c(MASTER.colName);
  const mRight= _.c(MASTER.colFinishDate);
  const mWidth= mRight - mLeft + 1;
  const master = savingsSh.getRange(MASTER.rowStart, mLeft, MASTER.rowEnd - MASTER.rowStart + 1, mWidth).getValues();
  const mIdxStart = _.c(MASTER.colStartDate) - mLeft;
  const mIdxFreq  = _.c(MASTER.colFreq)      - mLeft;

  const masterLookup = new Map();
  for (const r of master){
    const name = String(r[0]||"").trim();
    if (!name) continue;
    masterLookup.set(name, { startDate: _.toDate(r[mIdxStart]), freq: _.normFreq(r[mIdxFreq]) });
  }

  const outPct = Array(rows).fill(null).map(()=>[""]);
  const outFinish = Array(rows).fill(null).map(()=>[""]);
  const now = new Date();

  for (let i=0;i<rows;i++){
    const name = String(data[i][iName]||"").trim();
    if (!name || !masterLookup.has(name)) continue;

    const goal     = _.n(data[i][iGoal]);
    const startAmt = _.n(data[i][iStart]);
    const monthly  = (_.n(data[i][iAct])>0) ? _.n(data[i][iAct]) : _.n(data[i][iPred]);
    const { startDate, freq } = masterLookup.get(name);

    if (!(goal>0) || !startDate) {
      outPct[i][0] = (goal>0) ? Math.min(1, startAmt/goal) : 0;
      continue;
    }
    // elapsed
    const daysElapsed = Math.max(0, Math.floor((now - startDate)/86400000));
    let totalContrib = 0;
    if (freq === "MONTHLY")   totalContrib = monthly * Math.floor(daysElapsed/30);
    else if (freq === "WEEKLY")   totalContrib = monthly * Math.floor(daysElapsed/7);
    else if (freq === "BIWEEKLY") totalContrib = monthly * Math.floor(daysElapsed/14);

    const currentBalance = startAmt + totalContrib;
    outPct[i][0] = Math.min(1, goal>0 ? currentBalance/goal : 0);

    if (goal > currentBalance && monthly > 0) {
      const remaining = goal - currentBalance;
      const periods = Math.ceil(remaining / monthly);
      const finish = new Date(now);
      if (freq === "MONTHLY")      finish.setMonth(finish.getMonth() + periods);
      else if (freq === "WEEKLY")  finish.setDate(finish.getDate() + 7*periods);
      else if (freq === "BIWEEKLY")finish.setDate(finish.getDate() + 14*periods);
      outFinish[i][0] = finish;
    }
  }

  if (rows > 0){
    mSh.getRange(M.rowStart, _.c(M.colPct), rows, 1).setValues(outPct);
    _.fmtPct(mSh.getRange(M.rowStart, _.c(M.colPct), rows, 1));
    mSh.getRange(M.rowStart, _.c(M.colFinish), rows, 1).setValues(outFinish);
    _.fmtDate(mSh.getRange(M.rowStart, _.c(M.colFinish), rows, 1));
  }
}

var SAVINGS = (function(){
  const S = CFG.savings, ss = SpreadsheetApp.getActive();

  function readMain(){
    const sh=ss.getSheetByName(S.sheet); if(!sh) return [];
    const T=S.table, n=T.rowEnd-T.rowStart+1;
    const width = _.c(T.colFreq)-_.c(T.colName)+1;
    const data=sh.getRange(T.rowStart, _.c(T.colName), n, width).getValues();
    const idx=L=>_.c(L)-_.c(T.colName);
    const rows=[];
    for (let i=0;i<n;i++){
      const r=data[i], nm=String(r[idx(T.colName)]||"").trim(); if(!nm) continue;
      rows.push({
        name:nm,
        startAmt: _.n(r[idx(T.colStartAmount)]),
        start:    _.toDate(r[idx(T.colStartDate)]),
        contrib:  _.n(r[idx(T.colContribution)]),
        freq:     _.normFreq(r[idx(T.colFreq)]),
        goalAmt:  _.n(r[idx(T.colGoalAmount)])
      });
    }
    return rows;
  }

  function monthlyEquivalent(freq, contrib){
    switch (_.normFreq(freq)){
      case "WEEKLY":   return (contrib||0)*52/12;
      case "BIWEEKLY": return (contrib||0)*26/12;
      case "MONTHLY":  return (contrib||0);
      default:         return contrib||0;
    }
  }

function populateMonth(sh, year, mIdx, _allowPast, mainRows){
  const S = CFG.savings.monthly;
  const cap = S.rowEnd - S.rowStart + 1;
  const width = _.c(S.colPred) - _.c(S.colName) + 1; // write through Pred in one go

  const oName = 0;
  const oGoal = _.c(S.colGoal) - _.c(S.colName);
  const oStart= _.c(S.colStart)- _.c(S.colName);
  const oPred = _.c(S.colPred) - _.c(S.colName);

  const out = Array.from({length: cap}, () => Array(width).fill(""));
  let w = 0;

  for (const g of mainRows){
    if (w >= cap) break;
    const nm = String(g.name||"").trim();
    if (!nm) continue;

    const goalAmt  = _.n(g.goalAmt)  || 0;
    const startAmt = _.n(g.startAmt) || 0;
    const monthlyEq= getMonthlyEquivalent(g.freq, g.contrib) || 0;

    out[w][oName] = nm;
    out[w][oGoal] = goalAmt;
    out[w][oStart]= startAmt;
    out[w][oPred] = monthlyEq;
    w++;
  }

  const rng = sh.getRange(S.rowStart, _.c(S.colName), cap, width);
  rng.clearContent();
  rng.setValues(out);

  if (w > 0) {
    sh.getRange(S.rowStart, _.c(S.colGoal),  w, 1).setNumberFormat(CFG.formats.money);
    sh.getRange(S.rowStart, _.c(S.colStart), w, 1).setNumberFormat(CFG.formats.money);
    sh.getRange(S.rowStart, _.c(S.colPred),  w, 1).setNumberFormat(CFG.formats.money);
  }
};


function recomputeTiles(_sh, mainRows){
  console.time("SAVINGS.recomputeTiles");
  Logger.log("=== SAVINGS.recomputeTiles CALLED FROM: " + new Error().stack);
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet); 
  if(!mSh) return;

  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;

  const M = CFG.savings.monthly;
  const rows = M.rowEnd - M.rowStart + 1;
  
  // Read Month table
  const names = mSh.getRange(M.rowStart, _.c(M.colName), rows, 1).getValues().map(([v])=>String(v||"").trim());
  const preds = mSh.getRange(M.rowStart, _.c(M.colPred), rows, 1).getValues().map(([v])=>_.n(v));
  const acts  = mSh.getRange(M.rowStart, _.c(M.colAct),  rows, 1).getValues().map(([v])=>_.n(v));
  const goals = mSh.getRange(M.rowStart, _.c(M.colGoal), rows, 1).getValues().map(([v])=>_.n(v));
  const starts = mSh.getRange(M.rowStart, _.c(M.colStart), rows, 1).getValues().map(([v])=>_.n(v));
  const finishes = mSh.getRange(M.rowStart, _.c(M.colFinish), rows, 1).getValues().map(([v])=>_.toDate(v));

  let monthSave = 0;
  let amountLeftFromCurrentBalance = 0;  // ← Changed variable name
  let latestFinishDate = null;
  let earliestFinishDate = null;
  let nearestName = "";
  let totalStartAmounts = 0;

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name) continue;

    totalStartAmounts += starts[i];

    // Month contribution (from Month table, actual first then pred)
    const contrib = acts[i] > 0 ? acts[i] : preds[i];
    monthSave += contrib;

    // Current balance = start + this month's contribution
    const currentBalance = starts[i] + contrib;
    
    // Amount left = Goal - CurrentBalance
    const remainingFromCurrentBalance = Math.max(0, goals[i] - currentBalance);
    amountLeftFromCurrentBalance += remainingFromCurrentBalance;

    // Track finish dates from Month table
    const finishDate = finishes[i];
    if (finishDate) {
      if (!latestFinishDate || finishDate > latestFinishDate) {
        latestFinishDate = finishDate;
      }
      if (!earliestFinishDate || finishDate < earliestFinishDate) {
        earliestFinishDate = finishDate;
        nearestName = name;
      }
    }
  }

  // Calculate months left from latest finish date
  let monthsLeft = 0;
  if (latestFinishDate) {
    const diffTime = latestFinishDate - now;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    monthsLeft = Math.max(0, Math.ceil(diffDays / 30));
  }

  const MT = CFG.savings.monthlyTiles;
  const sSh = ss.getSheetByName(CFG.savings.sheet);

  // Month tiles (H6, M6, R6, W6, AB6)
  _.fmtMoney(mSh.getRange(MT.totalSaved)).setValue(monthSave);                          // H6 = this month's contributions
  _.fmtMoney(mSh.getRange(MT.amountLeft)).setValue(amountLeftFromCurrentBalance);       // M6 = Goal - (Start + ThisMonthContrib)
  mSh.getRange(MT.monthsLeft).setValue(monthsLeft).setNumberFormat(CFG.formats.months1); // R6 = months to latest finish
  mSh.getRange(MT.nearestGoal).setValue(nearestName);                                   // W6 = nearest goal name

  const incomeAEP = _.sumAEP_block(
    mSh, CFG.monthly.incomeTable.rowStart, CFG.monthly.incomeTable.rowEnd,
    CFG.monthly.incomeTable.colPred, CFG.monthly.incomeTable.colAct
  );
  mSh.getRange(MT.savingsToIncome)
     .setValue(incomeAEP > 0 ? monthSave / incomeAEP : 0)
     .setNumberFormat(CFG.formats.percent1);  // AB6 = savings to income ratio

  // Savings sheet tiles (H2, M2, R2, W2) - these still look at Savings sheet calculations
  if (sSh) {
    let savingsSheetTotal = 0;
    let savingsSheetLeft = 0;
    let savingsSheetLongest = 0;
    let savingsSheetNearest = 1e9;
    let savingsSheetNearestName = "";

    for (const goal of mainRows) {
      savingsSheetTotal += (goal.startAmt || 0);

      const remainingFromStart = Math.max(0, (goal.goalAmt || 0) - (goal.startAmt || 0));
      savingsSheetLeft += remainingFromStart;

      const monthlyEq = getMonthlyEquivalent(goal.freq, goal.contrib);
      if (monthlyEq > 0 && remainingFromStart > 0) {
        const monthsLeft = Math.ceil(remainingFromStart / monthlyEq);
        savingsSheetLongest = Math.max(savingsSheetLongest, monthsLeft);
        if (monthsLeft < savingsSheetNearest) {
          savingsSheetNearest = monthsLeft;
          savingsSheetNearestName = goal.name;
        }
      }
    }

    _.fmtMoney(sSh.getRange(CFG.savings.tiles.totalSaved)).setValue(savingsSheetTotal);
    _.fmtMoney(sSh.getRange(CFG.savings.tiles.amountLeft)).setValue(savingsSheetLeft);
    sSh.getRange(CFG.savings.tiles.monthsLeft).setValue(savingsSheetLongest).setNumberFormat(CFG.formats.months1);
    sSh.getRange(CFG.savings.tiles.nearestToComplete).setValue(savingsSheetNearestName);
  }
    console.timeEnd("SAVINGS.recomputeTiles");

}
function recomputeMain(mainSh, mainRows){
  const T = CFG.savings.table;
  const n = mainRows.length;
  if (n === 0) return;

  const outPct = [];
  const outFinish = [];

  const now = new Date();

  for (const r of mainRows) {
    const goal = _.n(r.goalAmt);
    const start = _.n(r.startAmt);
    const contribPerPeriod = _.n(r.contrib);
    const freq = _.normFreq(r.freq);
    const startDate = r.start || new Date();
    
    // Calculate days elapsed from start date
    const daysElapsed = Math.max(0, Math.floor((now - startDate) / (1000 * 60 * 60 * 24)));
    
    // Calculate total contributions based on frequency
    let totalContributions = 0;
    if (freq === "MONTHLY") {
      const monthsElapsed = Math.floor(daysElapsed / 30);
      totalContributions = contribPerPeriod * monthsElapsed;
    } else if (freq === "WEEKLY") {
      const weeksElapsed = Math.floor(daysElapsed / 7);
      totalContributions = contribPerPeriod * weeksElapsed;
    } else if (freq === "BIWEEKLY") {
      const biweeksElapsed = Math.floor(daysElapsed / 14);
      totalContributions = contribPerPeriod * biweeksElapsed;
    }
    
    const currentBalance = start + totalContributions;
    const pct = (goal > 0) ? Math.max(0, Math.min(1, currentBalance / goal)) : 0;
    
    outPct.push([pct]);
    
    // Calculate finish date
    if (goal > currentBalance && contribPerPeriod > 0) {
      const remaining = goal - currentBalance;
      let periodsNeeded = 0;
      
      if (freq === "MONTHLY") {
        periodsNeeded = Math.ceil(remaining / contribPerPeriod);
        const finish = new Date(now);
        finish.setMonth(finish.getMonth() + periodsNeeded);
        outFinish.push([finish]);
      } else if (freq === "WEEKLY") {
        periodsNeeded = Math.ceil(remaining / contribPerPeriod);
        const finish = new Date(now);
        finish.setDate(finish.getDate() + (periodsNeeded * 7));
        outFinish.push([finish]);
      } else if (freq === "BIWEEKLY") {
        periodsNeeded = Math.ceil(remaining / contribPerPeriod);
        const finish = new Date(now);
        finish.setDate(finish.getDate() + (periodsNeeded * 14));
        outFinish.push([finish]);
      } else {
        outFinish.push([""]);
      }
    } else {
      outFinish.push([""]);
    }
  }

  if (n > 0) {
    const pctCol = _.c(T.colPct);
    const finishCol = _.c(T.colFinishDate);
    
    if (pctCol > 0) {
      mainSh.getRange(T.rowStart, pctCol, n, 1).setValues(outPct);
      _.fmtPct(mainSh.getRange(T.rowStart, pctCol, n, 1));
    }
    
    if (finishCol > 0) {
      mainSh.getRange(T.rowStart, finishCol, n, 1).setValues(outFinish);
      _.fmtDate(mainSh.getRange(T.rowStart, finishCol, n, 1));
    }
  }
}
return { readMain, populateMonth, recomputeMain, recomputeTiles };

})();

/* ============================= INCOME ============================= */
var INCOME = (function(){
  const M  = CFG.income.master,
        Mm = CFG.income.monthly,
        ss = SpreadsheetApp.getActive();

  function readMaster(){
    const sh = ss.getSheetByName(M.sheet); if (!sh) return [];
    _ensureSheetSize(sh, M.rowEnd, _.c(M.colAcct));

    const width = _.c(M.colAcct) - _.c(M.colName) + 1;
    const n = M.rowEnd - M.rowStart + 1;

    const data = sh.getRange(M.rowStart, _.c(M.colName), n, width).getValues();
    const idx = L => _.c(L) - _.c(M.colName);

    const rows = [];
    for (let i = 0; i < n; i++){
      const r  = data[i];
      const nm = String(r[idx(M.colName)] || "").trim();
      if (!nm) continue;
      
      const perFreq = _.n(r[idx(M.colPerFreq)]);
      const freq = _.normFreq(r[idx(M.colFreq)]);
      const annual = perFreq * _.periodsPerYear(freq);
      
      rows.push({
        name: nm,
        start: _.toDate(r[idx(M.colStart)]),
        freq: freq,
        perFreq: perFreq,
        annual: annual,
        acct: String(r[idx(M.colAcct)] || "").trim()
      });
    }
    return rows;
  }

function computePlan(rows, year, m){
  const out = [];
  for (const r of rows){
    const F = _.normFreq(r.freq);
    if (!_.inWindowMonth(r.start, null, year, m) && F!=="ONETIME") {
      // keep your original gating if you rely on inWindowMonth;
      // MONTHLY/WEEKLY/BIWEEKLY still handled by the counter below
    }
    const occ = _occurrencesInMonth(F, r.start || new Date(year,0,1), year, m);
    if (F === "ONETIME") {
      if (occ === 1){
        out.push({ name:r.name, annual:r.annual, monthly:r.perFreq, acct:r.acct || "" });
      }
    } else if (occ > 0){
      out.push({ name:r.name, annual:r.annual, monthly: r.perFreq * occ, acct:r.acct || "" });
    }
  }
  return out;
}


function populateMonth(sh, year, mIdx, allowPast, masterRows) {
  const M = CFG.monthly.incomeTable;
  const cap = M.rowEnd - M.rowStart + 1;

  // cache columns once
  const cName = _.c(M.colName);
  const cPred = _.c(M.colPred);
  const cAcct = _.c(M.colAcct);
  const width = cAcct - cName + 1;

  const oName = 0;
  const oPred = cPred - cName;
  const oAcct = cAcct - cName;

  const out = Array.from({length: cap}, () => Array(width).fill(""));
  let w = 0;

  // (optional) prefilter by name existence
  const valid = masterRows.filter(inc => String(inc.name||"").trim());

  for (const inc of valid){
    if (w >= cap) break;
    const nm = String(inc.name||"").trim();

    const occ = _occurrencesInMonth(inc.freq, inc.start || new Date(year,0,1), year, mIdx);
    const monthlyAmt = (inc.perFreq || 0) * occ;

    out[w][oName] = nm;
    out[w][oPred] = monthlyAmt;
    out[w][oAcct] = String(inc.acct||"").trim();
    w++;
  }

  const rng = sh.getRange(M.rowStart, cName, cap, width);
  rng.clearContent();
  rng.setValues(out);

  if (w > 0) {
    sh.getRange(M.rowStart, cPred, w, 1).setNumberFormat(CFG.formats.money);
  }
}


return { readMaster, computePlan, populateMonth };
})();  // ← closes var INCOME = (function(){ ... })();
/* ============================= TRANSFERS (expand by frequency) ============================= */
var TRANSFERS = (function(){
  const TP = CFG.transfersPayments, ss = SpreadsheetApp.getActive();

  function readAndExpandForMonth(y, m){
    const CK = `xfer:${y}:${m}`;
const hit = _cacheGet_(CK);
if (hit) return hit;

    const sh = ss.getSheetByName(TP.sheet); if (!sh) return [];
    const rows  = TP.rowEnd - TP.rowStart + 1;

    // cache columns once
    const cName = _.c(TP.colName);
    const cDate = _.c(TP.colDate);
    const cFreq = _.c(TP.colFreq);
    const cFrom = _.c(TP.colFromAcct);
    const cTo   = _.c(TP.colToAcct);
    const cAmt  = _.c(TP.colAmount);
    const width = cTo - cName + 1;

    const data  = sh.getRange(TP.rowStart, cName, rows, width).getValues();
    const idx = L => _.c(L) - cName;

    const out = [];
    const monthEnd = new Date(y, m, 1).getTime(); // m = 1..12 → JS month next month (exclusive)

    // prefilter once
    const valid = data.filter(r => String(r[0]||"").trim() !== "");

    for (const r of valid){
      const name   = String(r[0]||"").trim();
      const start  = _.toDate(r[idx(TP.colDate)]); if (!start) continue;
      const freq   = _.normFreq(r[idx(TP.colFreq)]);
      const from   = String(r[idx(TP.colFromAcct)]||"").trim();
      const to     = String(r[idx(TP.colToAcct)]  ||"").trim();
      const amount = _.n(r[idx(TP.colAmount)]);
      if (!from || !to || amount<=0) continue;

      const { firstTime, periodMs } = _firstOccurrenceTimeInMonth(freq, start, y, m);
      if (firstTime < 0) continue;

      if (freq === "ONETIME" || freq === "MONTHLY"){
        out.push({ name, date: new Date(firstTime), from, to, amount, freq });
      } else {
        const occ = _occurrencesInMonth(freq, start, y, m);
        for (let k=0; k<occ; k++){
          const t = firstTime + k*periodMs;
          if (t >= monthEnd) break;
          out.push({ name, date:new Date(t), from, to, amount, freq });
        }
      }
    }
_cachePut_(CK, out, 180); // 3 minutes
return out;
  }
  return { readAndExpandForMonth };
})();


/* ============================= PAYCHECKS ============================= */
/* ============================= PAYCHECKS (full-year, Pred only) ============================= */
var PAYCHECKS = (function(){
  const P = CFG.income.paychecks;
  const ss = __ss();

  function rebuild(){
    const ck = "paychecks:" + (new Date()).getFullYear();
    try {
      const hit = CacheService.getDocumentCache().get(ck);
      if (hit) {
        const rows = JSON.parse(hit);
        _writeRows(rows); // helper: does the writing/formatting
        return;
      }
    } catch(_) {}

    const sh = ss.getSheetByName(P.sheet); if (!sh) return;
    const yr = (new Date()).getFullYear();
    const master = (typeof INCOME!=='undefined' && INCOME.readMaster) ? INCOME.readMaster() : [];
    const rows = [];

    for (const r of master){
      const F = _.normFreq(r.freq);
      if (F === "ONETIME"){
        if (r.start && r.start.getFullYear()===yr){
          rows.push({ name:r.name, date:r.start, freq:"One Time", amount:r.perFreq||0 });
        }
        continue;
      }
      const anchor = r.start || new Date(yr,0,1);
      const dates  = expandMemo(yr, anchor, null, F);
      const label  = (F==="WEEKLY")?"Weekly":(F==="BIWEEKLY"?"Bi-weekly":"Monthly");
      for (const d of dates){ rows.push({ name:r.name, date:d, freq:label, amount:r.perFreq||0 }); }
    }

    rows.sort((a,b)=>a.date-b.date);
    _writeRows(rows);

    try { CacheService.getDocumentCache().put(ck, JSON.stringify(rows), 180); } catch(_){}
  }

  function _writeRows(rows){
    const sh = ss.getSheetByName(P.sheet); if (!sh) return;
    const cap = P.rowEnd - P.rowStart + 1;
    const n = Math.min(rows.length, cap);
    const cName = _.c(P.colName), cDate = _.c(P.colDate), cFreq = _.c(P.colFreq), cPred = _.c(P.colPred);

    sh.getRangeList([
      `${colLetter(P.colName)}${P.rowStart}:${colLetter(P.colName)}${P.rowEnd}`,
      `${colLetter(P.colDate)}${P.rowStart}:${colLetter(P.colDate)}${P.rowEnd}`,
      `${colLetter(P.colFreq)}${P.rowStart}:${colLetter(P.colFreq)}${P.rowEnd}`,
      `${colLetter(P.colPred)}${P.rowStart}:${colLetter(P.colPred)}${P.rowEnd}`
    ]).clearContent();

    if (n===0) return;
    sh.getRange(P.rowStart, cName, n, 1).setValues(rows.slice(0,n).map(r=>[r.name]));
    sh.getRange(P.rowStart, cDate, n, 1).setValues(rows.slice(0,n).map(r=>[r.date]));
    sh.getRange(P.rowStart, cFreq, n, 1).setValues(rows.slice(0,n).map(r=>[r.freq]));
    sh.getRange(P.rowStart, cPred, n, 1).setValues(rows.slice(0,n).map(r=>[r.amount]));
    _.fmtDate( sh.getRange(P.rowStart, cDate, n, 1));
    _.fmtMoney(sh.getRange(P.rowStart, cPred, n, 1));
  }

  return { rebuild };
})();



/* ============================= DATED ACCOUNT BALANCE (FIXED) ============================= */
var DATED_ACCOUNT_BALANCE = (function(){
  const D  = CFG.datedAccountBalance;
  const TP = CFG.transfersPayments;
  const ss = __ss();;

function updateDatedAccountBalance(){
  const startTime = new Date();
  Logger.log("=== updateDatedAccountBalance START ===");

  const ss = __ss();
  const sh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const D = CFG.datedAccountBalance;

  // Build isCC map once
  const isCC = _buildIsCCMap_(mSh, CFG.bank, (CFG.creditCards && CFG.creditCards.bankMarkCol) || "O");

  // Target date
  const targetDate = _.toDate(sh.getRange(D.targetDate).getValue());
  if (!targetDate) return;

  // ---------- Read account names + starts (gap-safe) ----------
  console.time("Read accounts");
  const aNameCol = _.c(D.accounts.colName);
  const aStartCol = _.c(D.accounts.colStartBalance);
  const aRowsCfg = D.accounts.rowEnd - D.accounts.rowStart + 1;
  const aRowsEff = _effectiveRows(sh, D.accounts.rowStart, aNameCol, D.accounts.rowEnd); // count non-blank names
  const aRowsUse = Math.max(aRowsEff, 0);

  const names  = (aRowsUse > 0)
    ? sh.getRange(D.accounts.rowStart, aNameCol, aRowsUse, 1).getValues().map(r => String(r[0] || "").trim())
    : [];
  const starts = (aRowsUse > 0)
    ? sh.getRange(D.accounts.rowStart, aStartCol, aRowsUse, 1).getValues().map(r => _.n(r[0]))
    : [];
  console.timeEnd("Read accounts");

  // Working map by account
  const map = new Map();
  for (let i = 0; i < names.length; i++){
    const nm = names[i];
    if (nm) map.set(nm, { start: starts[i] || 0, inc: 0, exp: 0, xfer: 0 });
  }

  // ---------- Rollups ----------
  console.time("_sumExpenses");
  _sumExpenses(sh, targetDate, map);
  console.timeEnd("_sumExpenses");

  console.time("_sumTransfers");
  _sumTransfers(sh, targetDate, map, isCC);
  console.timeEnd("_sumTransfers");

  console.time("_sumIncomeToDate");
  _sumIncomeToDate(sh, targetDate, map);
  console.timeEnd("_sumIncomeToDate");

  // ---------- Build outputs ----------
  console.time("Write results");
  const outInc = [];
  const outExp = [];
  const outXf  = [];
  const outEnd = [];

  for (const nm of names){
    if (!nm){ outInc.push([""]); outExp.push([""]); outXf.push([""]); outEnd.push([""]); continue; }
    const v = map.get(nm) || { start: 0, inc: 0, exp: 0, xfer: 0 };
    const endBal = v.start + v.inc - v.exp + v.xfer;
    outInc.push([v.inc]);
    outExp.push([v.exp]);
    outXf.push([v.xfer]);
    outEnd.push([endBal]);
  }

  const cInc  = _.c(D.accounts.colIncome);
  const cExp  = _.c(D.accounts.colExpense);
  const cXfer = _.c(D.accounts.colTransfers);
  const cEnd  = _.c(D.accounts.colPredBalance);

  // Write only used rows
  if (aRowsUse > 0) {
    sh.getRange(D.accounts.rowStart, cInc,  aRowsUse, 1).setValues(outInc);
    sh.getRange(D.accounts.rowStart, cExp,  aRowsUse, 1).setValues(outExp);
    sh.getRange(D.accounts.rowStart, cXfer, aRowsUse, 1).setValues(outXf);
    sh.getRange(D.accounts.rowStart, cEnd,  aRowsUse, 1).setValues(outEnd);

    // Format only used rows
    sh.getRangeList([
      `${D.accounts.colIncome}${D.accounts.rowStart}:${D.accounts.colIncome}${D.accounts.rowStart + aRowsUse - 1}`,
      `${D.accounts.colExpense}${D.accounts.rowStart}:${D.accounts.colExpense}${D.accounts.rowStart + aRowsUse - 1}`,
      `${D.accounts.colTransfers}${D.accounts.rowStart}:${D.accounts.colTransfers}${D.accounts.rowStart + aRowsUse - 1}`,
      `${D.accounts.colPredBalance}${D.accounts.rowStart}:${D.accounts.colPredBalance}${D.accounts.rowStart + aRowsUse - 1}`,
    ]).setNumberFormat(CFG.formats.money);
  }

  // Clear any remaining configured rows *below* the used block (keeps the sheet tidy)
  // NOTE: This does NOT clear your filled rows — only the trailing unused rows.
  if (aRowsUse < aRowsCfg) {
    const clearStart = D.accounts.rowStart + aRowsUse;
    const clearH = aRowsCfg - aRowsUse;
    sh.getRangeList([
      `${D.accounts.colIncome}${clearStart}:${D.accounts.colIncome}${clearStart + clearH - 1}`,
      `${D.accounts.colExpense}${clearStart}:${D.accounts.colExpense}${clearStart + clearH - 1}`,
      `${D.accounts.colTransfers}${clearStart}:${D.accounts.colTransfers}${clearStart + clearH - 1}`,
      `${D.accounts.colPredBalance}${clearStart}:${D.accounts.colPredBalance}${clearStart + clearH - 1}`,
    ]).clearContent();
  }
  console.timeEnd("Write results");

  const duration = (new Date() - startTime) / 1000;
  Logger.log("=== updateDatedAccountBalance END === " + duration + "s");
}


function refreshSourceData(isH2Change){
  const startTime = new Date();
  Logger.log("=== refreshSourceData START ===");
  
  const ss = __ss();;
  const sh = ss.getSheetByName(D.sheet);
  if (!sh) return;
  
  const targetDate = _.toDate(sh.getRange(D.targetDate).getValue());
  if (!targetDate) return;

  const expenseMaster = EXPENSES.readMaster();

  // ✅ NEW: Read ALL existing data in ONE batch
  console.time("Read all existing data");
  const existingData = _readAllExistingDatedData(sh, isH2Change);
  console.timeEnd("Read all existing data");

  // ✅ Pass existing data to each function
  _populateExpensesFromMonth(sh, targetDate, !!isH2Change, expenseMaster, existingData.expenses);
  _populateTransfersFromTP(sh, targetDate, !!isH2Change, existingData.transfers);
// FIX: use sh (the dated sheet you already have), and pass existingData.income
_populateIncomeFromPaychecks(sh, targetDate, !!isH2Change, existingData.income, _buildAcctByNameFromIncomeMaster_());

  SpreadsheetApp.flush();
  updateDatedAccountBalance();
  
  const duration = (new Date() - startTime) / 1000;
  Logger.log("=== refreshSourceData END === " + duration + "s");
}

// ✅ NEW: Read all existing data in one go
function _readAllExistingDatedData(sh, isH2Change){
  const D = CFG.datedAccountBalance;
  
 // Read expenses (gap-safe)
const eName = _.c(D.expenses.colName);
const eRowsEff = _effectiveRows(sh, D.expenses.rowStart, eName, D.expenses.rowEnd);
const eWidth   = _.c(D.expenses.colInclude) - eName + 1;
const expensesRaw = eRowsEff > 0
  ? sh.getRange(D.expenses.rowStart, eName, eRowsEff, eWidth).getValues()
  : [];

// Read transfers (gap-safe)
const tName = _.c(D.transfers.colName);
const tRowsEff = _effectiveRows(sh, D.transfers.rowStart, tName, D.transfers.rowEnd);
const tWidth   = _.c(D.transfers.colInclude) - tName + 1;
const transfersRaw = tRowsEff > 0
  ? sh.getRange(D.transfers.rowStart, tName, tRowsEff, tWidth).getValues()
  : [];

// Read income (gap-safe)
const iName = _.c(D.income.colName);
const iRowsEff = _effectiveRows(sh, D.income.rowStart, iName, D.income.rowEnd);
const iWidth   = _.c(D.income.colInclude) - iName + 1;
const incomeRaw = iRowsEff > 0
  ? sh.getRange(D.income.rowStart, iName, iRowsEff, iWidth).getValues()
  : [];
  
  // Process expenses
  const expenses = _processExistingExpenses(expensesRaw, isH2Change);
  
  // Process transfers
  const transfers = _processExistingTransfers(transfersRaw, isH2Change);
  
  // Process income
  const income = _processExistingIncome(incomeRaw, isH2Change);
  
  return { expenses, transfers, income };
}

// ✅ NEW: Process existing expenses
function _processExistingExpenses(data, isH2Change){
  const D = CFG.datedAccountBalance.expenses;
  const iName = 0;
  const iDate = _.c(D.colDate) - _.c(D.colName);
  const iPred = _.c(D.colPred) - _.c(D.colName);
  const iAct = _.c(D.colAct) - _.c(D.colName);
  const iAcct = _.c(D.colAccount) - _.c(D.colName);
  const iInc = _.c(D.colInclude) - _.c(D.colName);

  const actuals = new Map();
  const includes = new Map();

  for (const r of data) {
    const nm = String(r[iName] || "").trim();
    if (!nm) continue;
    
    const dt = _.toDate(r[iDate]);
    const pred = _.n(r[iPred]);
    const act = _.n(r[iAct]);
    const acct = String(r[iAcct] || "").trim();
    
    if (dt) {
      const key = _expKey(nm, dt, pred, acct);
      if (act > 0) actuals.set(key, act);
      if (!isH2Change) includes.set(key, r[iInc] === true);
    }
  }
  
  return { actuals, includes };
}

// ✅ NEW: Process existing transfers
function _processExistingTransfers(data, isH2Change){
  const D = CFG.datedAccountBalance.transfers;
  const iName = 0;
  const iDate = _.c(D.colDate) - _.c(D.colName);
  const iPred = _.c(D.colPred) - _.c(D.colName);
  const iAct = _.c(D.colAct) - _.c(D.colName);
  const iFrom = _.c(D.colFromAcct) - _.c(D.colName);
  const iTo = _.c(D.colToAcct) - _.c(D.colName);
  const iInc = _.c(D.colInclude) - _.c(D.colName);

  const actuals = new Map();
  const includes = new Map();

  for (const row of data) {
    const nm = String(row[iName] || "").trim();
    if (!nm) continue;
    
    const dt = _.toDate(row[iDate]);
    const pred = _.n(row[iPred]);
    const from = String(row[iFrom] || "").trim();
    const to = String(row[iTo] || "").trim();
    
    if (dt) {
      const key = _xferKey(nm, dt, pred, from, to);
      const act = _.n(row[iAct]);
      if (act > 0) actuals.set(key, act);
      if (!isH2Change) includes.set(key, row[iInc] === true);
    }
  }
  
  return { actuals, includes };
}

// ✅ NEW: Process existing income
function _processExistingIncome(data, isH2Change){
  const DI = CFG.datedAccountBalance.income;
  const iName = 0;
  const iDate = _.c(DI.colDate) - _.c(DI.colName);
  const iPred = _.c(DI.colPred) - _.c(DI.colName);
  const iAct = _.c(DI.colAct) - _.c(DI.colName);
  const iAcct = _.c(DI.colAcct) - _.c(DI.colName);
  const iInc = _.c(DI.colInclude) - _.c(DI.colName);

  const actuals = new Map();
  const includes = new Map();

  for (const row of data) {
    const nm = String(row[iName] || "").trim();
    if (!nm) continue;
    
    const dt = _.toDate(row[iDate]);
    const pred = _.n(row[iPred]);
    const acct = String(row[iAcct] || "").trim();
    
    if (dt) {
const key = _incKey(nm, dt, pred, acct);
      const act = _.n(row[iAct]);
      if (act > 0) actuals.set(key, act);
      if (!isH2Change) includes.set(key, row[iInc] === true);
    }
  }
  
  return { actuals, includes };
}

function _populateExpensesFromMonth(datedSh, targetDate, isH2Change, masterRows, existingData){  // ← ADD existingData parameter

const DE = CFG.datedAccountBalance.expenses;
const deName = _.c(DE.colName);
const deDate = _.c(DE.colDate);
const deFreq = _idxOrNeg(DE.colName, DE.colFreq);
const deCat  = _.c(DE.colCategory);
const dePred = _.c(DE.colPred);
const deAct  = _.c(DE.colAct);
const deAcct = _.c(DE.colAccount);
const deInc  = _.c(DE.colInclude);

  const t = new Date();
  Logger.log("_populateExpensesFromMonth START");
  
  const existingActuals = existingData.actuals;  // ← Now it works!
  const existingIncludes = existingData.includes;
  
  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;
  const D = CFG.datedAccountBalance;

  const dRows = D.expenses.rowEnd - D.expenses.rowStart + 1;
  const dWidth = _.c(D.expenses.colInclude) - _.c(D.expenses.colName) + 1;
  

  console.time("Expand from master");
  const expanded = [];

  function formatFreqForDisplay(freq) {
    switch(_.normFreq(freq)) {
      case "WEEKLY": return "Weekly";
      case "BIWEEKLY": return "Bi-weekly";
      case "MONTHLY": return "Monthly";
      case "ONETIME": return "One Time";
      default: return "Monthly";
    }
  }

  for (const r of masterRows) {
    const F = _.normFreq(r.freq);
    const amt = r.amt || 0;
    if (amt <= 0 || !r.acct) continue;

    const freqDisplay = formatFreqForDisplay(r.freq);

    if (F === "ONETIME") {
      if (r.start && r.start.getFullYear() === y && (r.start.getMonth() + 1) === targetMonth) {
        expanded.push({ name: r.name, date: r.start, freq: freqDisplay, cat: r.cat || "", amt, acct: r.acct });
      }
    } else {
      const anchor = r.start || new Date(y, 0, 1);
     const dates = expandMemo(y, anchor, null, F).filter(d => (d.getMonth() + 1) === targetMonth);
      for (const d of dates) {
        expanded.push({ name: r.name, date: d, freq: freqDisplay, cat: r.cat || "", amt, acct: r.acct });
      }
    }
  }
  console.timeEnd("Expand from master");

  console.time("Build output array");
  const outData = Array.from({length: dRows}, () => Array(dWidth).fill(""));
  
  const oName = 0;
  const oDate = _.c(D.expenses.colDate) - _.c(D.expenses.colName);
  const oFreq = _idxOrNeg(D.expenses.colName, D.expenses.colFreq);
  const oCat = _.c(D.expenses.colCategory) - _.c(D.expenses.colName);
  const oPred = _.c(D.expenses.colPred) - _.c(D.expenses.colName);
  const oAct = _.c(D.expenses.colAct) - _.c(D.expenses.colName);
  const oAcct = _.c(D.expenses.colAccount) - _.c(D.expenses.colName);
  const oInc = _.c(D.expenses.colInclude) - _.c(D.expenses.colName);

  let wrow = 0;
  for (const e of expanded) {
    if (wrow >= dRows) break;
    
    const key = _expKey(e.name, e.date, e.amt, e.acct);
    
    outData[wrow][oName] = e.name;
    outData[wrow][oDate] = e.date;
    if (oFreq >= 0) outData[wrow][oFreq] = e.freq;
    outData[wrow][oCat] = e.cat;
    outData[wrow][oPred] = e.amt;
    outData[wrow][oAct] = existingActuals.get(key) || "";
    outData[wrow][oAcct] = e.acct;
    
    const defaultInc = (e.date <= targetDate);
    outData[wrow][oInc] = isH2Change ? defaultInc : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);
    
    wrow++;
  }
  console.timeEnd("Build output array");


  datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colName), dRows, dWidth).setValues(outData);

// === Batch format (dates + money) for just the rows we filled
if (wrow > 0) {
  datedSh.getRangeList([
    `${D.expenses.colDate}${D.expenses.rowStart}:${D.expenses.colDate}${D.expenses.rowStart + wrow - 1}`
  ]).setNumberFormat("M/d/yyyy");

  datedSh.getRangeList([
    `${D.expenses.colPred}${D.expenses.rowStart}:${D.expenses.colPred}${D.expenses.rowStart + wrow - 1}`,
    `${D.expenses.colAct}${D.expenses.rowStart}:${D.expenses.colAct}${D.expenses.rowStart + wrow - 1}`
  ]).setNumberFormat("$#,##0.00");
}
  
  Logger.log("_populateExpensesFromMonth END: " + ((new Date() - t) / 1000) + "s");
}
function _populateTransfersFromTP(datedSh, targetDate, isH2Change, existingData){
const DT = CFG.datedAccountBalance.transfers;
const dtName = _.c(DT.colName);
const dtDate = _.c(DT.colDate);
const dtFreq = _.c(DT.colFreq);
const dtFrom = _.c(DT.colFromAcct);
const dtPred = _.c(DT.colPred);
const dtAct  = _.c(DT.colAct);
const dtTo   = _.c(DT.colToAcct);
const dtInc  = _.c(DT.colInclude);
const offDate = dtDate - dtName;

// Hoisted columns for Transfers & Payments source sheet
const TP = CFG.transfersPayments;
const tpCName = _.c(TP.colName);
const tpCDate = _.c(TP.colDate);
const tpCFreq = _.c(TP.colFreq);
const tpCFrom = _.c(TP.colFromAcct);
const tpCTo   = _.c(TP.colToAcct);
const tpCAmt  = _.c(TP.colAmount);

  const t = new Date();
  Logger.log("_populateTransfersFromTP START");

  // === pull existing actuals/includes once ===
  const existingActuals  = existingData.actuals;
  const existingIncludes = existingData.includes;

  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;

  // === shorthand ===
  const D  = CFG.datedAccountBalance;

  // === cache dated-transfers columns once ===
  

  // === row/width for output grid ===
  const n      = D.transfers.rowEnd - D.transfers.rowStart + 1;
  const dWidth = dtInc - dtName + 1;

  // === read Transfers sheet (cached ss via __ss) ===
  console.time("Read Transfers sheet");
  const ss   = __ss();
  const tpSh = ss.getSheetByName(TP.sheet);
  if (!tpSh) {
    Logger.log("_populateTransfersFromTP END: " + ((new Date() - t) / 1000) + "s");
    return;
  }

  // === cache TP columns once ===

const tpWidth  = tpCTo - tpCName + 1;
const tpRowsEff = _effectiveRows(tpSh, TP.rowStart, tpCName, TP.rowEnd);
if (tpRowsEff <= 0) { Logger.log("_populateTransfersFromTP: no rows"); return; }
const tpData = tpSh.getRange(TP.rowStart, tpCName, tpRowsEff, tpWidth).getValues();

  console.timeEnd("Read Transfers sheet");

  // indexers (0-based relative to tpCName)
  const iTpName   = 0;
  const iTpDate   = tpCDate - tpCName;
  const iTpFreq   = tpCFreq - tpCName;
  const iTpFrom   = tpCFrom - tpCName;
  const iTpTo     = tpCTo   - tpCName;
  const iTpAmount = tpCAmt  - tpCName;

  function displayFreq(freq) {
    switch (_.normFreq(freq)) {
      case "WEEKLY":   return "Weekly";
      case "BIWEEKLY": return "Bi-weekly";
      case "MONTHLY":  return "Monthly";
      case "ONETIME":  return "One Time";
      default:         return "Monthly";
    }
  }

  // === expand transfers for this month ===
console.time("Process transfers");
const expanded = [];
const validTp  = tpData.filter(r => String(r[iTpName] || "").trim() !== "");

for (const r of validTp) {
  const name      = String(r[iTpName] || "").trim();
  const startDate = _.toDate(r[iTpDate]); if (!startDate) continue;
  const freq      = _.normFreq(r[iTpFreq]);
  const freqDisp  = displayFreq(r[iTpFreq]);
  const from      = String(r[iTpFrom] || "").trim();
  const to        = String(r[iTpTo]   || "").trim();
  const amt       = _.n(r[iTpAmount]);
  if (!from || !to || amt <= 0) continue;

  const { firstTime, periodMs } = _firstOccurrenceTimeInMonth(freq, startDate, y, targetMonth);
  if (firstTime >= 0) {
    if (freq === "ONETIME" || freq === "MONTHLY") {
      expanded.push({ name, date:new Date(firstTime), freq:freqDisp, from, to, amount:amt });
    } else {
      const monthEnd = new Date(y, targetMonth, 1).getTime(); // exclusive
      for (let t = firstTime; t < monthEnd; t += periodMs) {
        expanded.push({ name, date:new Date(t), freq:freqDisp, from, to, amount:amt });
      }
    }
  }
}
console.timeEnd("Process transfers");


  // === pull CC payments from Dated Expenses (single read) ===
  console.time("Read CC payments");
  const DE         = D.expenses;
  const expNameCol = _.c(DE.colName);
  const expDateCol = _.c(DE.colDate);
  const expCatCol  = _.c(DE.colCategory);
  const expPredCol = _.c(DE.colPred);
  const expActCol  = _.c(DE.colAct);
  const expAcctCol = _.c(DE.colAccount);
  const expWidth   = expAcctCol - expNameCol + 1;

  const expRows = DE.rowEnd - DE.rowStart + 1;
  const expData = datedSh.getRange(DE.rowStart, expNameCol, expRows, expWidth).getValues();

  const iExpName = 0;
  const iExpDate = expDateCol - expNameCol;
  const iExpCat  = expCatCol  - expNameCol;
  const iExpPred = expPredCol - expNameCol;
  const iExpAct  = expActCol  - expNameCol;
  const iExpAcct = expAcctCol - expNameCol;

  const ccCat = ((CFG.creditCards && CFG.creditCards.paymentCategory) || "Credit Card Payment").toUpperCase();

  const monthStart = new Date(y, targetMonth - 1, 1).getTime();
  const monthEnd   = new Date(y, targetMonth, 0, 23, 59, 59).getTime();

  for (const r of expData) {
    const cat  = String(r[iExpCat]  || "").trim().toUpperCase();
    if (cat !== ccCat) continue;

    const name = String(r[iExpName] || "").trim(); if (!name) continue;

    const date = _.toDate(r[iExpDate]); if (!date) continue;
    const tms  = date.getTime();
    if (tms < monthStart || tms > monthEnd) continue;

    const fromAcct = String(r[iExpAcct] || "").trim(); if (!fromAcct) continue;
    const amt = (_.n(r[iExpAct]) > 0) ? _.n(r[iExpAct]) : _.n(r[iExpPred]);
    if (amt <= 0) continue;

    const toAcct = name.replace(/\s*payment\s*$/i, "").trim();
    expanded.push({ name, date, freq:"One Time", from:fromAcct, to:toAcct, amount:amt });
  }
  console.timeEnd("Read CC payments");

  // === build output grid ===
  console.time("Build output");
  const outData = Array.from({length: n}, () => Array(dWidth).fill(""));

  // 0-based offsets relative to dtName
  const oName = 0;
  const oDate = dtDate - dtName;
  const oFreq = dtFreq - dtName;
  const oFrom = dtFrom - dtName;
  const oPred = dtPred - dtName;
  const oAct  = dtAct  - dtName;
  const oTo   = dtTo   - dtName;
  const oInc  = dtInc  - dtName;

  const targetTime = targetDate.getTime();

  let wrow = 0;
  for (const xfer of expanded) {
    if (wrow >= n) break;

    const key = _xferKey(xfer.name, xfer.date, xfer.amount, xfer.from, xfer.to);

    outData[wrow][oName] = xfer.name;
    outData[wrow][oDate] = xfer.date;
    outData[wrow][oFreq] = xfer.freq;
    outData[wrow][oFrom] = xfer.from;
    outData[wrow][oPred] = xfer.amount;
    outData[wrow][oAct]  = existingActuals.get(key) || "";
    outData[wrow][oTo]   = xfer.to;

    const defaultInc = (xfer.date.getTime() <= targetTime);
    outData[wrow][oInc] = isH2Change ? defaultInc
                                     : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);

    wrow++;
  }
  console.timeEnd("Build output");

  // === single write + format ===
// === single write + format ===
console.time("Write and format");
datedSh.getRange(D.transfers.rowStart, dtName, n, dWidth).setValues(outData);

if (wrow > 0) {
  datedSh.getRangeList([
    `${D.transfers.colDate}${D.transfers.rowStart}:${D.transfers.colDate}${D.transfers.rowStart + wrow - 1}`
  ]).setNumberFormat("M/d/yyyy");

  datedSh.getRangeList([
    `${D.transfers.colPred}${D.transfers.rowStart}:${D.transfers.colPred}${D.transfers.rowStart + wrow - 1}`,
    `${D.transfers.colAct}${D.transfers.rowStart}:${D.transfers.colAct}${D.transfers.rowStart + wrow - 1}`
  ]).setNumberFormat("$#,##0.00");
}
console.timeEnd("Write and format");


  Logger.log("_populateTransfersFromTP END: " + ((new Date() - t) / 1000) + "s");
}


function _populateIncomeFromPaychecks(datedSh, targetDate, isH2Change, existingData, acctByName){
  // Hoisted columns for Dated Income
const DI = CFG.datedAccountBalance.income;
const diName = _.c(DI.colName);
const diDate = _.c(DI.colDate);
const diFreq = _.c(DI.colFreq);
const diPred = _.c(DI.colPred);
const diAct  = _.c(DI.colAct);
const diAcct = _.c(DI.colAcct);
const diInc  = _.c(DI.colInclude);
const offDate = diDate - diName;

// Hoisted columns for Paychecks source sheet
const P = CFG.income.paychecks;
const pNameCol = _.c(P.colName);
const pDateCol = _.c(P.colDate);
const pFreqCol = _.c(P.colFreq);
const pPredCol = _.c(P.colPred);

  const t = new Date();
  Logger.log("_populateIncomeFromPaychecks START");

  const existingActuals = existingData.actuals;
  const existingIncludes = existingData.includes;
  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;
  const D = CFG.datedAccountBalance;
  const ss = __ss();

  const n = DI.rowEnd - DI.rowStart + 1;
  const dWidth = _.c(DI.colInclude) - _.c(DI.colName) + 1;

  // Read Paychecks (Name, Date, Freq, Pred)
  console.time("Read paychecks");
  const paySh = ss.getSheetByName(CFG.income.paychecks.sheet);
  if (!paySh) {
    Logger.log("_populateIncomeFromPaychecks END: " + ((new Date() - t) / 1000) + "s");
    return;
  }
const pRowsEff = _effectiveRows(paySh, P.rowStart, pNameCol, P.rowEnd);
if (pRowsEff <= 0) { Logger.log("_populateIncomeFromPaychecks: no rows"); return; }

  const pWidth = pPredCol - pNameCol + 1;
const pData = paySh.getRange(P.rowStart, pNameCol, pRowsEff, pWidth).getValues();
  console.timeEnd("Read paychecks");

  const pNm = 0;
  const pDt = pDateCol - pNameCol;
  const pFq = pFreqCol - pNameCol;
  const pPr = pPredCol - pNameCol;

  // Optionally (re)build acctByName from Month income if the map was not provided
  if (!acctByName) acctByName = new Map();
  if (acctByName.size === 0) {
    console.time("Build account map from Month income");
    const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
    const it = CFG.monthly.incomeTable;
    const iNameCol = _.c(it.colName);
    const iAcctCol = _.c(it.colAcct);
    const iWidth = iAcctCol - iNameCol + 1;
    const incData = mSh ? mSh.getRange(it.rowStart, iNameCol, it.rowEnd - it.rowStart + 1, iWidth).getValues() : [];
    const idxAcct = iAcctCol - iNameCol;

    for (const r of incData) {
      const nm = String(r[0] || "").trim();
      if (!nm) continue;
      const ac = String(r[idxAcct] || "").trim();
      if (ac) acctByName.set(nm, ac);
    }
    console.timeEnd("Build account map from Month income");
  }

  // Pre-compute month window
  console.time("Process paychecks");
  const monthStart = new Date(y, targetMonth - 1, 1).getTime();
  const monthEnd   = new Date(y, targetMonth, 0, 23, 59, 59).getTime();

  const expanded = [];
  const validPayData = pData.filter(r => String(r[pNm] || "").trim() !== "");

  for (const r of validPayData) {
    const nm = String(r[pNm] || "").trim();
    const d  = _.toDate(r[pDt]); if (!d) continue;
    const ts = d.getTime();
    if (ts < monthStart || ts > monthEnd) continue;

    const freqLabel = String(r[pFq] || "").trim();
    const pred = _.n(r[pPr]);
    if (pred <= 0) continue;

    const acct = acctByName.get(nm) || "";
    expanded.push({ name: nm, date: d, freq: freqLabel, pred, acct });
  }
  console.timeEnd("Process paychecks");

  // Build output
  console.time("Build output");
  const outData = Array.from({length: n}, () => Array(dWidth).fill(""));

  const oNm   = 0;
  const oDt   = _.c(DI.colDate) - _.c(DI.colName);
  const oFq   = _.c(DI.colFreq) - _.c(DI.colName);
  const oPr   = _.c(DI.colPred) - _.c(DI.colName);
  const oAc   = _.c(DI.colAct)  - _.c(DI.colName);
  const oAcct = _.c(DI.colAcct) - _.c(DI.colName);
  const oIn   = _.c(DI.colInclude) - _.c(DI.colName);

  const targetTime = targetDate.getTime();

  let w = 0;
  for (const z of expanded) {
    if (w >= n) break;

    const key = _incKey(z.name, z.date, z.pred, z.acct);

    outData[w][oNm]   = z.name;
    outData[w][oDt]   = z.date;
    outData[w][oFq]   = z.freq;
    outData[w][oPr]   = z.pred;
    outData[w][oAc]   = existingActuals.get(key) || "";
    outData[w][oAcct] = z.acct;

    const defaultInc = (z.date.getTime() <= targetTime);
    outData[w][oIn]  = isH2Change ? defaultInc
                                  : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);
    w++;
  }
  console.timeEnd("Build output");

  // Write + format
datedSh.getRange(DI.rowStart, _.c(DI.colName), n, dWidth).setValues(outData);

if (w > 0) {
  datedSh.getRangeList([
    `${DI.colDate}${DI.rowStart}:${DI.colDate}${DI.rowStart + w - 1}`
  ]).setNumberFormat("M/d/yyyy");

  datedSh.getRangeList([
    `${DI.colPred}${DI.rowStart}:${DI.colPred}${DI.rowStart + w - 1}`,
    `${DI.colAct}${DI.rowStart}:${DI.colAct}${DI.rowStart + w - 1}`
  ]).setNumberFormat("$#,##0.00");
}

Logger.log("_populateIncomeFromPaychecks END: " + ((new Date() - t) / 1000) + "s");
}


function _sumExpenses(sh, targetDate, map){
  const D = CFG.datedAccountBalance;
  const cName = _.c(D.expenses.colName);
  const cInc  = _.c(D.expenses.colInclude);

  const nEff = _effectiveRows(sh, D.expenses.rowStart, cName, D.expenses.rowEnd);
  if (nEff <= 0) return;

  const width = cInc - cName + 1;
  const data = sh.getRange(D.expenses.rowStart, cName, nEff, width).getValues();

  const iName = 0;
  const iDate = _.c(D.expenses.colDate)    - cName;
  const iPred = _.c(D.expenses.colPred)    - cName;
  const iAct  = _.c(D.expenses.colAct)     - cName;
  const iAcc  = _.c(D.expenses.colAccount) - cName;
  const iInc  = _.c(D.expenses.colInclude) - cName;

  for (const r of data){
    const nm = String(r[iName]||"").trim(); if (!nm) continue;
    if (!_boolCell_(r[iInc])) continue;

    const d = _.toDate(r[iDate]); if (!d || d > targetDate) continue;

    const act  = +r[iAct]  || 0;
    const pred = +r[iPred] || 0;
    const amt  = act > 0 ? act : pred;
    if (amt <= 0) continue;

    const acc = String(r[iAcc]||"").trim(); if (!acc) continue;
    const v = map.get(acc); if (v) v.exp += amt;
  }
}


function _sumTransfers(sh, targetDate, map, isCC){
  const D = CFG.datedAccountBalance;
  const cName = _.c(D.transfers.colName);
  const cInc  = _.c(D.transfers.colInclude);

  const nEff = _effectiveRows(sh, D.transfers.rowStart, cName, D.transfers.rowEnd);
  if (nEff <= 0) return;

  const width = cInc - cName + 1;
  const data = sh.getRange(D.transfers.rowStart, cName, nEff, width).getValues();

  const iName = 0;
  const iDate = _.c(D.transfers.colDate)     - cName;
  const iFrm  = _.c(D.transfers.colFromAcct) - cName;
  const iPred = _.c(D.transfers.colPred)     - cName;
  const iAct  = _.c(D.transfers.colAct)      - cName;
  const iTo   = _.c(D.transfers.colToAcct)   - cName;
  const iInc  = _.c(D.transfers.colInclude)  - cName;

  for (const r of data){
    const nm = String(r[iName]||"").trim(); if (!nm) continue;
    if (!_boolCell_(r[iInc])) continue;

    const d = _.toDate(r[iDate]); if (!d || d > targetDate) continue;

    const from = String(r[iFrm]||"").trim();
    const to   = String(r[iTo] ||"").trim();
    if (!from || !to) continue;

    const act  = +r[iAct]  || 0;
    const pred = +r[iPred] || 0;
    const amt  = act > 0 ? act : pred;
    if (amt <= 0) continue;

    const fromIsCC = !!isCC.get(from);
    const toIsCC   = !!isCC.get(to);

    if (fromIsCC && !toIsCC) {
      if (map.has(from)) map.get(from).xfer += amt;
      if (map.has(to))   map.get(to).xfer   += amt;
    } else if (!fromIsCC && toIsCC) {
      if (map.has(from)) map.get(from).xfer -= amt;
      if (map.has(to))   map.get(to).xfer   -= amt;
    } else {
      if (map.has(from)) map.get(from).xfer -= amt;
      if (map.has(to))   map.get(to).xfer   += amt;
    }
  }
}



function _sumIncomeToDate(sh, targetDate, map){
  const DI = CFG.datedAccountBalance.income;
  const cName = _.c(DI.colName);
  const cInc  = _.c(DI.colInclude);

  const nEff = _effectiveRows(sh, DI.rowStart, cName, DI.rowEnd);
  if (nEff <= 0) return;

  const width = cInc - cName + 1;
  const data = sh.getRange(DI.rowStart, cName, nEff, width).getValues();

  const iDat = _.c(DI.colDate) - cName;
  const iPrd = _.c(DI.colPred) - cName;
  const iAcc = _.c(DI.colAcct) - cName;
  const iInc = _.c(DI.colInclude) - cName;

  for (const r of data){
    if (!_boolCell_(r[iInc])) continue;

    const d = _.toDate(r[iDat]); if (!d || d > targetDate) continue;

    const amt = +r[iPrd] || 0; if (amt <= 0) continue;

    const acc = String(r[iAcc]||"").trim(); if (!acc) continue;

    const v = map.get(acc); if (v) v.inc += amt;
  }
}





  return {
    updateDatedAccountBalance: updateDatedAccountBalance,
    refreshSourceData: refreshSourceData
  };
})(); // ← end of DATED_ACCOUNT_BALANCE IIFE
/* ============================= MONTH TAB: Tiles + Calendar + Bank ============================= */

var _calendarUpdatePending = false;
var _calendarLock = false;

function Monthly_RenderCalendar_Safe() {
  try {
    const ss = __ss();
if (_toastThrottle_("__calendar_toast__", 1200)) {
  ss.toast("Rendering calendar...", "Budget Tracker", -1);
}

    const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
    if (!mSh) {
      ss.toast("", "", 1);  // Clear message
      return;
    }
    
    const y = _.year();
    const m = _.monthIdx();
    
    const T = CFG.expenses.monthly;
    
    const colNameNum = _.c(T.colName);
    const colDateNum = _.c(T.colDate);
    const colPredNum = _.c(T.colPred);
    const colActNum = _.c(T.colAct);
    const colAcctNum = _.c(T.colAcct);
    
    const startCol = Math.min(colNameNum, colDateNum, colPredNum, colActNum, colAcctNum);
    const endCol = Math.max(colNameNum, colDateNum, colPredNum, colActNum, colAcctNum);
    
    const dataStartRow = T.rowStart;
    const numRows = T.rowEnd - dataStartRow + 1;
    const numCols = endCol - startCol + 1;
    
    const monthData = mSh.getRange(dataStartRow, startCol, numRows, numCols).getValues();
    
    const idxName = colNameNum - startCol;
    const idxDate = colDateNum - startCol;
    const idxPred = colPredNum - startCol;
    const idxAct = colActNum - startCol;
    
    const occFromMonth = [];
    for (const r of monthData) {
      const name = String(r[idxName] || "").trim();
      const date = _.toDate(r[idxDate]);
      if (!name || !date) continue;
      
      if (date.getFullYear() === y && (date.getMonth() + 1) === m) {
        const amt = (_.n(r[idxAct]) > 0) ? _.n(r[idxAct]) : _.n(r[idxPred]);
        if (amt > 0) {
          occFromMonth.push({ name, date, amt });
        }
      }
    }
    
    const transfers = TRANSFERS.readAndExpandForMonth(y, m);
    Monthly_RenderCalendar(mSh, y, m, occFromMonth, transfers);
    
    ss.toast("", "", 1);  // ← Clear message
    
  } catch(e) {
    SpreadsheetApp.getActive().toast("", "", 1);  // Clear on error
  }
}

function Monthly_RenderCalendar(mSh, y, m, occExp, occXfer){
  const stack = new Error().stack;
  Logger.log("=== Monthly_RenderCalendar CALLED ===");
  const C = CFG.monthly.calendar;
  const sh = mSh || SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const per = C.perDayHeight;
  const contentH = 6;
  const firstDow = new Date(y, m-1, 1).getDay();
  const daysInMonth = new Date(y, m, 0).getDate();

  const leftCol = _.c("C");
  const rightCol = _.c("AJ");
  const totalRows = C.perDayHeight * 6;
  const totalCols = rightCol - leftCol + 1;

  // Map transfers to same format
  const occXferMapped = (occXfer || []).map(t => ({ name: `Transfer: ${t.name}`, date: t.date, amt: t.amount }));
  const occ = (occExp || []).concat(occXferMapped);

  // Group by day AND average duplicates with same name
  const byDay = new Map();
  for (const o of occ){
    const d = o.date.getDate();
    if (!byDay.has(d)) byDay.set(d, new Map());
    
    const dayMap = byDay.get(d);
    if (!dayMap.has(o.name)) {
      dayMap.set(o.name, { name: o.name, amounts: [] });
    }
    dayMap.get(o.name).amounts.push(o.amt);
  }
  
  // Average and sort
  const byDayFinal = new Map();
  for (const [day, nameMap] of byDay.entries()) {
    const items = [];
    for (const [name, data] of nameMap.entries()) {
      const avg = data.amounts.reduce((sum, amt) => sum + amt, 0) / data.amounts.length;
      items.push({ name, amount: avg });
    }
    items.sort((a, b) => b.amount - a.amount);
    byDayFinal.set(day, items);
  }

  // Clear and prep calendar area
  const all = sh.getRange(C.startRow, leftCol, totalRows, totalCols);
  all.clearContent();
  all.breakApart();
  all.setBorder(false,false,false,false,false,false);
  
  clearOverflowPanels();

  // Batch operations
  const mergesToDo = [];
  const valuesToSet = [];
  const borderRanges = [];
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const color = C.borderColor || "#b7e1cd";

  for (let w=0; w<6; w++){
    const top = C.startRow + w*per;

    for (let dgi=0; dgi<7; dgi++){
      const idx = w*7 + dgi;
      const dayNum = 1 + idx - firstDow;
      if (dayNum < 1 || dayNum > daysInMonth) continue;

      const [cStart, cEnd] = C.dayColGroups[dgi];
      const left = _.c(cStart);
      const right = _.c(cEnd);
      const width = right - left + 1;

      // Queue merges
      for (let r=0; r<contentH; r++){
        mergesToDo.push(sh.getRange(top + r, left, 1, width));
      }

      // Queue border
      borderRanges.push(sh.getRange(top, left, contentH, width));

      // Queue day number
      valuesToSet.push({
        range: sh.getRange(top, left),
        value: ordinal(dayNum)
      });

      const items = byDayFinal.get(dayNum) || [];
      const visN = C.visibleLines || 4;

      // Queue visible items
      for (let i=0; i<Math.min(items.length, visN); i++){
        valuesToSet.push({
          range: sh.getRange(top+1+i, left, 1, width),
          value: `${items[i].name} ($${Number(items[i].amount||0).toFixed(0)})`
        });
      }

      // Queue +N more
      const hidden = Math.max(0, items.length - visN);
      if (hidden > 0){
        const plusCell = sh.getRange(top + (contentH - 1), left);
        valuesToSet.push({
          range: plusCell,
          value: `+${hidden} more`
        });
        
  _ensureCalendarProtection_(sh, C.startRow, _.c("C"), C.perDayHeight * 6, _.c("AJ") - _.c("C") + 1);

      }
    }
  }

  // Execute all merges
  for (const r of mergesToDo) r.merge();

  // Execute all borders
  for (const r of borderRanges) {
    r.setBorder(true, true, true, true, false, false, color, borderStyle);
  }

  // Execute all values
  for (const {range, value} of valuesToSet) {
    range.setValue(value);
  }

  SpreadsheetApp.flush();

  // Protect
  try {
    const calendarRange = sh.getRange(C.startRow, leftCol, totalRows, totalCols);
    const p = calendarRange.protect();
    p.setDescription("Calendar (auto-generated)");
    p.setWarningOnly(true);
  } catch(e){ }
}
function _shouldRun_(key, ms){
  const p = PropertiesService.getDocumentProperties();
  const last = Number(p.getProperty(key) || "0");
  const now  = Date.now();
  if (now - last < (ms||400)) return false;
  p.setProperty(key, String(now));
  return true;
}

function ordinal(n) {
  const s = ["th", "st", "nd", "rd"];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}
function Lists_SyncCreditCardPayment() {
  const ss = __ss();;
  const lists = ss.getSheetByName("_Lists");
  if (!lists) return;
  
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!mSh) return;
  
  const B = CFG.bank;
  const n = B.rowEnd - B.rowStart + 1;
  const markCol = (CFG.creditCards && CFG.creditCards.bankMarkCol) || "O";
  
  const accounts = mSh.getRange(B.rowStart, _.c(B.colName), n, 1).getValues().map(r => String(r[0]||"").trim());
  const ccMarks = mSh.getRange(B.rowStart, _.c(markCol), n, 1).getValues();
  
  // Get CC account names
  const ccAccounts = [];
  for (let i = 0; i < accounts.length; i++) {
    const v = ccMarks[i][0];
    const isCC = v === true || String(v).toUpperCase()==="TRUE" || v === 1;
    if (isCC && accounts[i]) {
      ccAccounts.push(accounts[i] + " Payment");
    }
  }
  
  // Get current expense names from column C
  const maxRows = 200;
  const currentNames = lists.getRange(2, 3, maxRows, 1).getValues().map(r => String(r[0]||"").trim()).filter(n => n !== "");
  
  // Remove old CC payment names
  const filtered = currentNames.filter(n => !n.endsWith(" Payment"));
  
  // Add new CC payment names
  const updated = [...filtered, ...ccAccounts].sort();
  
  lists.getRange(2, 3, maxRows, 1).clearContent();
  if (updated.length > 0) {
    lists.getRange(2, 3, updated.length, 1).setValues(updated.map(n => [n]));
  }
  
  Lists_ApplyValidations();
}

/* ============================= BANK TOTALS + CATEGORY TABLE + TOP TILES ============================= */
function Monthly_UpdateBankAndCategoryFormatting(){
  console.time("Monthly_UpdateBankAndCategoryFormatting");
  const ss = __ss();
  const sh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const B = CFG.bank;
  const n = B.rowEnd - B.rowStart + 1;

  // ✅ OPTIMIZATION 1: Read bank data in ONE batch
  const bankCols = Math.max(_.c(B.colName), _.c(B.colPredBegin), _.c((CFG.creditCards && CFG.creditCards.bankMarkCol) || "O")) - _.c(B.colName) + 1;
  const bankData = sh.getRange(B.rowStart, _.c(B.colName), n, bankCols).getValues();
  
  const accounts    = bankData.map(r => String(r[0]||"").trim());
  const predBegins  = bankData.map(r => _.n(r[_.c(B.colPredBegin) - _.c(B.colName)]));
  const markCol     = (CFG.creditCards && CFG.creditCards.bankMarkCol) || "O";
  const ccMarks     = bankData.map(r => {
    const v = r[_.c(markCol) - _.c(B.colName)];
    return v === true || String(v).toUpperCase()==="TRUE" || v === 1;
  });
  
  const isCC = new Map();
  for (let i=0; i<accounts.length; i++){
    if (accounts[i]) isCC.set(accounts[i], !!ccMarks[i]);
  }

  // ✅ OPTIMIZATION 2: Read income & expense tables ONCE
  const t = CFG.monthly.incomeTable;
  const iWidth  = _.c(t.colAcct) - _.c(t.colName) + 1;
  const incData = sh.getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, iWidth).getValues();
  
  const et = CFG.monthly.expenseTable;
  const eWidth  = _.c(et.colAcct) - _.c(et.colName) + 1;
  const expData = sh.getRange(et.rowStart, _.c(et.colName), et.rowEnd - et.rowStart + 1, eWidth).getValues();

  // ✅ OPTIMIZATION 3: Process income & expenses in single pass
  const incByAcct = new Map();
  const iPred = _.c(t.colPred) - _.c(t.colName);
  const iAct  = _.c(t.colAct)  - _.c(t.colName);
  const iAcc  = _.c(t.colAcct) - _.c(t.colName);

  for (const r of incData){
    const acct = String(r[iAcc]||"").trim();
    if (!acct) continue;
    const act  = +r[iAct]  || 0;
    const pred = +r[iPred] || 0;
    const aep  = act > 0 ? act : pred;
    if (aep <= 0) continue;
    incByAcct.set(acct, (incByAcct.get(acct) || 0) + aep);
  }

  // Expense processing
  const expByAcct = new Map();
  const ccPaymentsByFromAcct = new Map();
  
  let monthTotalNoCCPay = 0;
  let monthTotalCCPay   = 0;

  const eCat  = _.c(et.colCat)  - _.c(et.colName);
  const ePred = _.c(et.colPred) - _.c(et.colName);
  const eAct  = _.c(et.colAct)  - _.c(et.colName);
  const eAcc  = _.c(et.colAcct) - _.c(et.colName);

  const ccPaymentCatUpper = ((CFG.creditCards && CFG.creditCards.paymentCategory) || "Credit Card Payment").toUpperCase();

  // ✅ OPTIMIZATION 4: Filter empty rows BEFORE processing
  const validExpRows = expData.filter(r => String(r[0]||"").trim() !== "");
  
  for (const r of validExpRows) {
    const cat  = String(r[eCat] || "").trim().toUpperCase();
    const acct = String(r[eAcc] || "").trim();
    const act  = +r[eAct]  || 0;
    const pred = +r[ePred] || 0;
    const aep  = act > 0 ? act : pred;
    if (aep <= 0) continue;

    const isOnCCAccount = isCC.get(acct);

    if (cat === ccPaymentCatUpper) {
      monthTotalCCPay += aep;
      if (acct) {
        expByAcct.set(acct, (expByAcct.get(acct)||0) + aep);
        ccPaymentsByFromAcct.set(acct, (ccPaymentsByFromAcct.get(acct)||0) + aep);
      }
    } else if (!isOnCCAccount) {
      monthTotalNoCCPay += aep;
      if (acct) expByAcct.set(acct, (expByAcct.get(acct)||0) + aep);
    }
  }

  // Write tiles
  try {
    sh.getRange("M2:M10").setValues([
      [monthTotalNoCCPay],
      [""], [""], [""], [""], [""], [""], [""],
      [monthTotalCCPay]
    ]).setNumberFormat(CFG.formats.money);
    
    sh.getRange("H10").setValue(monthTotalNoCCPay + monthTotalCCPay).setNumberFormat(CFG.formats.money);
  } catch(_) {}

  // ✅ OPTIMIZATION 5: Read transfers ONCE
  const y = _.year(), m = _.monthIdx();
  const xferByAcct = new Map();
  
  const allTransfers = TRANSFERS.readAndExpandForMonth(y, m);
  for (const tfr of allTransfers){
    const fromIsCC = isCC.get(tfr.from);
    const toIsCC   = isCC.get(tfr.to);
    
    if (fromIsCC && !toIsCC) {
      xferByAcct.set(tfr.from, (xferByAcct.get(tfr.from)||0) + tfr.amount);
      xferByAcct.set(tfr.to,   (xferByAcct.get(tfr.to)||0)   + tfr.amount);
    } else if (!fromIsCC && toIsCC) {
      xferByAcct.set(tfr.from, (xferByAcct.get(tfr.from)||0) - tfr.amount);
      xferByAcct.set(tfr.to,   (xferByAcct.get(tfr.to)||0)   - tfr.amount);
    } else {
      xferByAcct.set(tfr.from, (xferByAcct.get(tfr.from)||0) - tfr.amount);
      xferByAcct.set(tfr.to,   (xferByAcct.get(tfr.to)||0)   + tfr.amount);
    }
  }

  // adjust CC acct balances for payments recorded as expenses-from-cash
  for (const [fromAcct, paymentAmt] of ccPaymentsByFromAcct.entries()) {
    for (const ccAcct of accounts) {
      if (isCC.get(ccAcct)) {
        xferByAcct.set(ccAcct, (xferByAcct.get(ccAcct)||0) - paymentAmt);
      }
    }
  }

// Build CC charges map from rows we've already read (no second sheet read)
const chargesByAcct = new Map();
for (const r of validExpRows) {
  const acct = String(r[eAcc] || "").trim(); if (!acct) continue;
  const cat  = String(r[eCat] || "").trim().toUpperCase();
  const act  = +r[eAct]  || 0;
  const pred = +r[ePred] || 0;
  const aep  = act > 0 ? act : pred;
  if (aep <= 0) continue;
  // “charges” = all non-CC-Payment spend on card accounts; exclude CC Payment category
  if (cat !== ccPaymentCatUpper && isCC.get(acct)) {
    chargesByAcct.set(acct, (chargesByAcct.get(acct) || 0) + aep);
  }
}

// Compute predicted end by account
const outExp=[], outInc=[], outXf=[], outEnd=[];
for (let i=0; i<accounts.length; i++){
  const acc = accounts[i];
  if (!acc){ outExp.push([""]); outInc.push([""]); outXf.push([""]); outEnd.push([""]); continue; }

  const inc   = incByAcct.get(acc)  || 0;
  const xfers = xferByAcct.get(acc) || 0;
  const begin = predBegins[i]       || 0;

  let exp, endBalance;
  if (isCC.get(acc)) {
    const charges = chargesByAcct.get(acc) || 0;
    exp = charges;
    endBalance = begin + charges + xfers;  // CC “balance” grows by charges; xfers adjust it
  } else {
    exp = expByAcct.get(acc) || 0;
    endBalance = begin + inc - exp + xfers;
  }

  outExp.push([exp]);
  outInc.push([inc]);
  outXf.push([xfers]);
  outEnd.push([endBalance]);
}


  // ===== BATCH WRITE =====
  sh.getRange(B.rowStart, _.c(B.colExpTotal),   n, 1).setValues(outExp);
  sh.getRange(B.rowStart, _.c(B.colIncomeTotal),n, 1).setValues(outInc);
  sh.getRange(B.rowStart, _.c(B.colTransfer),   n, 1).setValues(outXf);
  sh.getRange(B.rowStart, _.c(B.colPredEnd),    n, 1).setValues(outEnd);

  // ===== 4D: Batch format the four money columns (optional if your _Bank_FormatMoneyColumns_ already does this) =====
  sh.getRangeList([
    `${B.colExpTotal}${B.rowStart}:${B.colExpTotal}${B.rowEnd}`,
    `${B.colIncomeTotal}${B.rowStart}:${B.colIncomeTotal}${B.rowEnd}`,
    `${B.colTransfer}${B.rowStart}:${B.colTransfer}${B.rowEnd}`,
    `${B.colPredEnd}${B.rowStart}:${B.colPredEnd}${B.rowEnd}`
  ]).setNumberFormat(CFG.formats.money);

  // If your helper formats these, you can keep this or remove one of them
  try { _Bank_FormatMoneyColumns_(); } catch(_) {}

  // Category rollup (reuse validExpRows!)
  _Expenses_UpdateCategories_Fast(sh, validExpRows, eCat, ePred, eAct);

  // Top tiles
  const TILES = CFG.monthly.tiles;
  const incomeTotal  = _.sumAEP_block(sh, t.rowStart,  t.rowEnd,  t.colPred,  t.colAct);
  const expenseTotal = _.sumAEP_block(sh, et.rowStart, et.rowEnd, et.colPred, et.colAct);

  const { y:yy, m:mm } = _.currentYM ? _.currentYM() : { y:_.year(), m:_.monthIdx() };
  const dim      = new Date(yy, mm, 0).getDate();
  const avgDaily = (expenseTotal>0 && dim>0) ? (expenseTotal/dim) : 0;
  const ratio    = (incomeTotal>0) ? (expenseTotal/incomeTotal) : 0;
  const flow     = incomeTotal - expenseTotal;

  sh.getRange(TILES.income).setValue(incomeTotal).setNumberFormat(CFG.formats.money);
  sh.getRange(TILES.expense).setValue(expenseTotal).setNumberFormat(CFG.formats.money);
  sh.getRange(TILES.avgDaily).setValue(avgDaily).setNumberFormat(CFG.formats.money);
  sh.getRange(TILES.ratio).setValue(ratio).setNumberFormat(CFG.formats.percent1);
  sh.getRange(TILES.flow).setValue(flow).setNumberFormat(CFG.formats.money);

  // Hi/Lo tiles (reuse validExpRows!)
  _updateHiLoTiles_Fast(sh, validExpRows, eCat, ePred, eAct, TILES);
  
  console.timeEnd("Monthly_UpdateBankAndCategoryFormatting");
}


// ✅ NEW: Faster category update (reuses already-read data)
function _Expenses_UpdateCategories_Fast(sh, expData, idxCat, idxPred, idxAct){
  const E = CFG.monthly.categoryTable;
  const IT = CFG.monthly.incomeTable;
  
  const incomeAEP = _.sumAEP_block(sh, IT.rowStart, IT.rowEnd, IT.colPred, IT.colAct);
  const map = new Map();

  for (const r of expData){
    const cat = String(r[idxCat]||"").trim();
    if(!cat) continue;
    const aep = (_.n(r[idxAct])>0 ? _.n(r[idxAct]) : _.n(r[idxPred]));
    if (aep>0) map.set(cat,(map.get(cat)||0)+aep);
  }
  
  const sorted = [...map.entries()].sort((a,b)=>b[1]-a[1]);
  const outH = E.rowEnd-E.rowStart+1;
  const outW = _.c(E.colPct)-_.c(E.colLabel)+1;
  const buf = Array.from({length: outH}, ()=>Array(outW).fill(""));

  const offTotal = _.c(E.colTotal)-_.c(E.colLabel);
  const offPct = _.c(E.colPct)-_.c(E.colLabel);

  for (let i=0;i<Math.min(outH, sorted.length); i++){
    const [cat,total] = sorted[i];
    buf[i][0] = cat;
    buf[i][offTotal] = total;
    buf[i][offPct] = incomeAEP>0 ? total/incomeAEP : 0;
  }

  sh.getRange(E.rowStart, _.c(E.colLabel), outH, outW).setValues(buf);
  
  const filled = Math.min(outH,sorted.length);
  if (filled > 0) {
    sh.getRange(E.rowStart, _.c(E.colTotal), filled, 1).setNumberFormat(CFG.formats.money);
    sh.getRange(E.rowStart, _.c(E.colPct), filled, 1).setNumberFormat(CFG.formats.percent1);
  }
}

// ✅ NEW: Faster hi/lo tiles (reuses already-read data)
function _updateHiLoTiles_Fast(sh, expData, idxCat, idxPred, idxAct, T){
  let hiExpVal = 0, hiExpName = "";
  let loExpVal = Number.POSITIVE_INFINITY, loExpName = "";
  const catTotals = new Map();
  let subsTotal = 0;

  for (const row of expData){
    const name = String(row[0] || "").trim();
    if (!name) continue;

    const cat = String(row[idxCat] || "").trim();
    const pred = _.n(row[idxPred]);
    const act = _.n(row[idxAct]);
    const aep = (act > 0) ? act : pred;
    if (aep <= 0) continue;

    if (aep > hiExpVal){ hiExpVal = aep; hiExpName = name; }
    if (aep < loExpVal){ loExpVal = aep; loExpName = name; }

    if (cat){
      catTotals.set(cat, (catTotals.get(cat) || 0) + aep);
      if (cat === CFG.monthly.subscriptionsCategory) subsTotal += aep;
    }
  }

  sh.getRange(T.hiExp).setValue(hiExpName || "")
    .setNote(hiExpVal>0 ? ("Amount: $" + hiExpVal.toFixed(2)) : "");
  sh.getRange(T.loExp).setValue((isFinite(loExpVal)&&loExpVal>0)?loExpName:"")
    .setNote((isFinite(loExpVal)&&loExpVal>0)?("Amount: $" + loExpVal.toFixed(2)):"");
  sh.getRange(T.subs).setValue(subsTotal).setNumberFormat(CFG.formats.money);
}
/* ============================= OPTIONAL: PUBLIC COMPOSER ============================= */


/* ============================= INCOME SHEET DASH (tiles) ============================= */
function updateIncomeTiles(){
  console.time("updateIncomeTiles");
  const ss = __ss();;
  const sh = ss.getSheetByName(CFG.income.tiles.sheet);
  if (!sh) return;

  const y = _.year();
  const master = INCOME.readMaster();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);

  // Predicted annual (unchanged logic)
  let predAnnual = 0;
  for (const r of master){
    const F = _.normFreq(r.freq);
    const occs = _.expand(y, r.start || new Date(y,0,1), null, F);
    const perOcc = r.perFreq || 0;
    predAnnual += perOcc * occs.length;
  }

  const predictedMonthlyAvg = predAnnual / 12;

  // This-month income AEP
  const t = CFG.monthly.incomeTable;
  const thisMonthIncome = _.sumAEP_block(mSh, t.rowStart, t.rowEnd, t.colPred, t.colAct);

  // Variance
  const variance = thisMonthIncome - predictedMonthlyAvg;

  // Days left in month
  const now = new Date();
const { e:last } = monthBounds(now.getFullYear(), now.getMonth()+1);
const daysLeft = Math.max(0, Math.ceil((last - now) / MS_DAY));


  // Unbudgeted = income - (expenses + savings)
  const expAEP = _.sumAEP_block(mSh, CFG.monthly.expenseTable.rowStart, CFG.monthly.expenseTable.rowEnd, CFG.monthly.expenseTable.colPred, CFG.monthly.expenseTable.colAct);
  const saveAEP = _.sumAEP_block(mSh, CFG.savings.monthly.rowStart, CFG.savings.monthly.rowEnd, CFG.savings.monthly.colPred, CFG.savings.monthly.colAct);
  const unbudgeted = thisMonthIncome - expAEP - saveAEP;

  // ===== BATCH TILE WRITES =====
  const tiles = [
    [CFG.income.tiles.predictedAnnual,     predAnnual,          CFG.formats.money],
    [CFG.income.tiles.predictedMonthlyAvg, predictedMonthlyAvg, CFG.formats.money],
    [CFG.income.tiles.thisMonth,           thisMonthIncome,     CFG.formats.money],
    [CFG.income.tiles.daysLeft,            daysLeft,            "0"],
    [CFG.income.tiles.unbudgeted,          unbudgeted,          CFG.formats.money]
  ];

  // set all values first
  tiles.forEach(([a1, val]) => sh.getRange(a1).setValue(val));
  // then format in one pass
  tiles.forEach(([a1, , fmt]) => sh.getRange(a1).setNumberFormat(fmt));

  // Variance keeps your conditional behavior
  const varCell = sh.getRange(CFG.income.tiles.incomeVariance);
  if (Math.abs(variance) < 0.01) {
    varCell.clearContent();
  } else {
    varCell.setValue(variance).setNumberFormat(CFG.formats.moneyNeg);
  }

  console.timeEnd("updateIncomeTiles");
}

function WarnDupes_WriteColumnH_Block(sh, rowStart, rowEnd, colName, warnColLtr){
  const warnCol = _.c(warnColLtr||"H");
  const n = rowEnd - rowStart + 1;
  const vals = sh.getRange(rowStart, _.c(colName), n, 1).getValues().map(r=>String(r[0]||"").trim());
  // clear
  sh.getRange(rowStart, warnCol, n, 1).clearContent().clearNote();
  const seen = new Map();
  for (let i=0;i<n;i++){
    const nm = vals[i]; if (!nm) continue;
    const k = nm.toUpperCase();
    const row = rowStart + i;
    if (seen.has(k)){
      const first = seen.get(k);
      sh.getRange(first, warnCol).setValue("⚠️").setNote("Duplicate account name");
      sh.getRange(row,   warnCol).setValue("⚠️").setNote("Duplicate account name");
    } else {
      seen.set(k, row);
    }
  }
}
function WarnDupes_UpdateAll(){
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh) WarnDupes_WriteColumnH_Block(mSh, CFG.bank.rowStart, CFG.bank.rowEnd, CFG.bank.colName, "H");
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (dSh) WarnDupes_WriteColumnH_Block(dSh, CFG.datedAccountBalance.accounts.rowStart, CFG.datedAccountBalance.accounts.rowEnd, CFG.datedAccountBalance.accounts.colName, "H");
}

/* ============================= EXPENSES SHEET DASH HELPERS ============================= */
function Expenses_ComputePredictedAnnual(y) {
  const rows = EXPENSES.readMaster();
  let total = 0;
  for (const r of rows) {
    const F = _.normFreq(r.freq), amt = _.n(r.amt); if (amt<=0) continue;
    let count = 0;
    if (F==="ONETIME") count = (r.start && r.start.getFullYear()===y) ? 1 : 0;
    else if (!r.start) count = _.periodsPerYear(F);
    else count = _.expand(y, r.start || new Date(y,0,1), null, F).length;
    total += amt * count;
  }
  return total;
}

function Expenses_UpdateSheetTiles() {
   console.time("Expenses_UpdateSheetTiles");
  const ss = __ss();;
  const sheet = ss.getSheetByName(CFG.expenses.tiles.sheet); 
  if (!sheet) return;
  
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!mSh) return;

  // READ values directly from Month sheet (batched read)
  const monthValues = mSh.getRange("M2:M10").getValues();
  const monthM2 = _.n(monthValues[0][0]);
  const monthM10 = _.n(monthValues[8][0]);
  
  const monthH10 = _.n(mSh.getRange("H10").getValue());


  // Predicted annual
  const y = _.year();
  const predictedAnnual = Expenses_ComputePredictedAnnual(y);

  sheet.getRangeList(["H2","M2","H6","M6"])
  .getRanges()
  .forEach((rng, i) => {
    const v = [monthM2, monthM2 * 12, monthH10, monthM10][i];
    rng.setValue(v);
  });
sheet.getRangeList(["H2","M2","H6","M6"])
  .getRanges()
  .forEach(rng => rng.setNumberFormat(CFG.formats.money));

  
  sheet.getRange(CFG.expenses.tiles.predictedAnnual).setValue(predictedAnnual).setNumberFormat(CFG.formats.money);

  // Upcoming expenses, largest day, etc.
  updateExpenseExtras();
  
    console.timeEnd("Expenses_UpdateSheetTiles");

}
function updateExpenseExtras(){
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const sheet = ss.getSheetByName(CFG.expenses.tiles.sheet);
  if (!mSh || !sheet) return;

  const CT = CFG.expenses.categoryTiles;
  const now = new Date();
  const currentY = now.getFullYear();
  const currentM = now.getMonth() + 1;

  const t = CFG.expenses.monthly;
  
  // OPTIMIZATION: Read all columns at once (instead of calculating width)
  const dataRange = mSh.getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, _.c(t.colAcct) - _.c(t.colName) + 1);
  const data = dataRange.getValues();

  const idxDate = _.c(t.colDate) - _.c(t.colName);
  const idxCat = _.c(t.colCat) - _.c(t.colName);
  const idxPred = _.c(t.colPred) - _.c(t.colName);
  const idxAct = _.c(t.colAct) - _.c(t.colName);

  // Pre-filter: Skip empty rows (HUGE speedup!)
  const validRows = data.filter(r => String(r[0] || "").trim() !== "");

  // Upcoming expenses
  const upcoming = [];
  for (const r of validRows){
    const nm = String(r[0] || "").trim();
    const d = _.toDate(r[idxDate]);
    if (!d || d < now) continue;
    
    const days = Math.ceil((d.setHours(0,0,0,0) - (new Date()).setHours(0,0,0,0)) / 86400000);
    upcoming.push({ nm, d: new Date(d), days });
  }
  
  upcoming.sort((a,b) => a.d - b.d);

  // Batch write upcoming expenses
  const up1 = sheet.getRange(CFG.expenses.tiles.upcomingExpense1);
  const up2 = sheet.getRange(CFG.expenses.tiles.upcomingExpense2);
  
  if (upcoming[0]) {
    up1.setValue(upcoming[0].nm).setNote(upcoming[0].days + " days");
  } else {
    up1.clearContent().setNote("");
  }
  
  if (upcoming[1]) {
    up2.setValue(upcoming[1].nm).setNote(upcoming[1].days + " days");
  } else {
    up2.clearContent().setNote("");
  }

  // Largest day - CURRENT MONTH ONLY
  const perDay = new Map();
  for (const r of validRows){
    const d = _.toDate(r[idxDate]);
    if (!d || d.getFullYear() !== currentY || (d.getMonth() + 1) !== currentM) continue;
    
 const act  = +r[idxAct]  || 0;
const pred = +r[idxPred] || 0;
const aep = act > 0 ? act : pred;
if (aep <= 0) continue;
    
    const key = d.toISOString().slice(0,10);
    perDay.set(key, (perDay.get(key) || 0) + aep);
  }
  
  let bestKey = "", bestAmt = 0;
  for (const [k,v] of perDay.entries()){ 
    if (v > bestAmt){ 
      bestAmt = v; 
      bestKey = k; 
    } 
  }
  
  const AB2 = sheet.getRange(CFG.expenses.tiles.largestDay);
  if (bestKey){
    const d = new Date(bestKey + "T12:00:00");
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const displayText = MON[d.getMonth()] + " " + d.getDate();
    AB2.setValue(displayText).setNote("$" + bestAmt.toFixed(2));
  } else {
    AB2.clearContent().clearNote();
  }

  // Category & item rollups
  const catTotals = new Map();
  const nameTotals = new Map();
  let subsTotal = 0;

  for (const r of validRows){
    const nm = String(r[0] || "").trim();
    const cat = String(r[idxCat] || "").trim();
const act  = +r[idxAct]  || 0;
const pred = +r[idxPred] || 0;
const aep = act > 0 ? act : pred;
if (aep <= 0) continue;
    
nameTotals.set(nm, (nameTotals.get(nm) || 0) + aep);   
if (cat){
     catTotals.set(cat, (catTotals.get(cat) || 0) + aep);
      if (cat === CFG.monthly.subscriptionsCategory) subsTotal += aep;
    }
  }

  // Write category tiles
  if (nameTotals.size){
    const items = [...nameTotals.entries()].sort((a,b) => b[1] - a[1]);
    const [hiName, hiAmt] = items[0];
    const [loName, loAmt] = items[items.length - 1];
    
    sheet.getRange(CT.highestExpName).setValue(hiName).setNote(`$${hiAmt.toFixed(2)}`);
    sheet.getRange(CT.lowestExpName).setValue(loName).setNote(`$${loAmt.toFixed(2)}`);
  } else {
    sheet.getRange(CT.highestExpName).setValue("—").setNote("");
    sheet.getRange(CT.lowestExpName).setValue("—").setNote("");
  }

  sheet.getRange(CT.totalSubs).setValue("$" + subsTotal.toFixed(2));
}

/* ============================= CATEGORY TABLE (MONTH) (alt composer) ============================= */
function Expenses_UpdateCategories(){
 const ss = __ss(); E=CFG.monthly.categoryTable, sh=ss.getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const incomeTotal = _.sumAEP_block(sh, CFG.monthly.incomeTable.rowStart, CFG.monthly.incomeTable.rowEnd, CFG.monthly.incomeTable.colPred, CFG.monthly.incomeTable.colAct);
  const t=CFG.expenses.monthly, width = _.c(t.colAct) - _.c(t.colCat) + 1;
  const rows = sh.getRange(t.rowStart, _.c(t.colCat), t.rowEnd - t.rowStart + 1, width).getValues();
  const idxPred = _.c(t.colPred) - _.c(t.colCat), idxAct = _.c(t.colAct) - _.c(t.colCat);

  const catTotals=new Map();
  for (const row of rows){
    const cat=String(row[0]||"").trim(); if (!cat) continue;
    const aep = Math.max(_.n(row[idxAct]), _.n(row[idxPred]));
    if (aep>0) catTotals.set(cat, (catTotals.get(cat)||0) + aep);
  }

  const outH = E.rowEnd - E.rowStart + 1, outW = _.c(E.colPct) - _.c(E.colLabel) + 1;
  const buf = Array.from({length: outH}, () => Array(outW).fill(""));
  const offTotal = _.c(E.colTotal) - _.c(E.colLabel), offPct = _.c(E.colPct) - _.c(E.colLabel);

  const sorted = [...catTotals.entries()].sort((a,b)=>b[1]-a[1]);
  for (let i=0;i<Math.min(outH, sorted.length);i++){
    const [cat,total]=sorted[i];
    buf[i][0]=cat;
    buf[i][offTotal]=total;
    buf[i][offPct]= incomeTotal>0? total/incomeTotal : 0;
  }

  const outRange = sh.getRange(E.rowStart, _.c(E.colLabel), outH, outW);
  outRange.setValues(buf);

  const filled = Math.min(outH, sorted.length);
  if (filled>0){
    _.fmtMoney(sh.getRange(E.rowStart, _.c(E.colTotal), filled, 1));
    _.fmtPct(  sh.getRange(E.rowStart, _.c(E.colPct),   filled, 1));
  }
}
function _ensureCalendarProtection_(sh, startRow, leftCol, height, width){
  const range = sh.getRange(startRow, leftCol, height, width);
  const a1 = range.getA1Notation();
  const prot = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (prot.some(p => p.getRange().getA1Notation() === a1)) return;
  const p = range.protect();
  p.setDescription("Calendar (auto-generated)");
  p.setWarningOnly(true);
}

function _buildAcctByNameFromIncomeMaster_(){
  const master = (typeof INCOME!=='undefined' && INCOME.readMaster) ? INCOME.readMaster() : [];
  const map = new Map();
  for (const r of master){
    const nm = String(r.name||"").trim();
    const acct = String(r.acct||"").trim();
    if (nm && acct) map.set(nm, acct);
  }
  return map;
}


/* ============================= SINGLE-MONTH ROLLOVER HELPERS ============================= */
function SingleMonth_EnsureSheet(){
 const ss = __ss(); name=CFG.singleMonth.sheet;
  let sh=ss.getSheetByName(name); if (!sh){ sh=ss.insertSheet(name, 0); }
  return sh;
}
function SingleMonth_ClearMonthlyRows(sh){
  // Income - clear Name only
  const I = CFG.monthly.incomeTable;
  sh.getRange(I.rowStart, _.c(I.colName), I.rowEnd-I.rowStart+1, 1).clearContent();
  
  // Expenses - clear Name, Cat, Pred (skip Date since it doesn't exist in monthly table)
  const E = CFG.monthly.expenseTable;
  sh.getRange(E.rowStart, _.c(E.colName), E.rowEnd-E.rowStart+1, 1).clearContent();
  sh.getRange(E.rowStart, _.c(E.colCat), E.rowEnd-E.rowStart+1, 1).clearContent();
  sh.getRange(E.rowStart, _.c(E.colPred), E.rowEnd-E.rowStart+1, 1).clearContent();
  // NOT clearing colDate - doesn't exist in expenseTable
  // NOT clearing colAct - preserve actuals

  // Savings - clear Name, Goal, Start, Pred, Pct, Finish
  const S = CFG.savings.monthly;
  sh.getRange(S.rowStart, _.c(S.colName), S.rowEnd-S.rowStart+1, 1).clearContent();
  sh.getRange(S.rowStart, _.c(S.colGoal), S.rowEnd-S.rowStart+1, 1).clearContent();
  sh.getRange(S.rowStart, _.c(S.colStart), S.rowEnd-S.rowStart+1, 1).clearContent();
  sh.getRange(S.rowStart, _.c(S.colPred), S.rowEnd-S.rowStart+1, 1).clearContent();
  sh.getRange(S.rowStart, _.c(S.colPct), S.rowEnd-S.rowStart+1, 1).clearContent();
  sh.getRange(S.rowStart, _.c(S.colFinish), S.rowEnd-S.rowStart+1, 1).clearContent();
  // NOT clearing colAct - preserve actuals
}
function SingleMonth_CarryBankBegins_FromPrevEnd(sh){
  const B = CFG.bank;
  const nameCol = _.c(B.colName);
  const lastRow = B.rowEnd;

  // find last row with an account name
  const names = sh.getRange(B.rowStart, nameCol, lastRow - B.rowStart + 1, 1).getValues();
  let lastUsed = B.rowStart - 1;
  for (let i = 0; i < names.length; i++){
    if (String(names[i][0]||"").trim()) lastUsed = B.rowStart + i;
  }
  if (lastUsed < B.rowStart) return; // nothing to do

  const firstTail = lastUsed + 1;
  if (firstTail > lastRow) return;

  // only clear totals in *tail* rows (do NOT touch Pred Begin)
  const cols = [B.colExpTotal, B.colIncomeTotal, B.colTransfer, B.colPredEnd, B.colActEnd];
  const a1s  = cols.map(L => `${L}${firstTail}:${L}${lastRow}`);
  sh.getRangeList(a1s).clearContent();
}


/* ============================= MISC HEADER + BOUNDS ============================= */
function Month_WriteHeader() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;
  const now = new Date();
  const label = now.toLocaleDateString(undefined, { month: "long" });
  sh.getRange("B7").setValue(label);
}

function getCurrentMonthBounds() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-11
  return { 
    start: new Date(y, m, 1), 
    end: new Date(y, m + 1, 0), 
    now: now 
  };
}


/* ============================= MAIN ORCHESTRATOR ============================= */
function updateAll(){
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(0)) {
    Logger.log("updateAll skipped: lock busy");
    return;
  }
  try {
    const totalStart = new Date();
    Logger.log("=== updateAll START ===");
    console.time("updateAll");

const ss = __ss();
    const sh = ss.getSheetByName(CFG.singleMonth.sheet) || ss.insertSheet(CFG.singleMonth.sheet);
    const y = _.year(), mIdx = _.monthIdx();

    try { PAYCHECKS.rebuild(); } catch(_){}
    const incomeMaster = INCOME.readMaster();
    const expenseMaster = EXPENSES.readMaster();
    const savingsMaster = SAVINGS.readMain();

    INCOME.populateMonth(sh, y, mIdx, true, incomeMaster);
    EXPENSES.populateMonth(sh, y, mIdx, true, expenseMaster);
    SAVINGS.populateMonth(sh, y, mIdx, true, savingsMaster);

    console.time("Savings_RecomputeMonthFromMonthOnly");
    try { Savings_RecomputeMonthFromMonthOnly(); } catch(e){ Logger.log(e); }
    console.timeEnd("Savings_RecomputeMonthFromMonthOnly");

    console.time("Savings_UpdateEmergencyFundTile");
    try { Savings_UpdateEmergencyFundTile(); } catch(e){ Logger.log(e); }
    console.timeEnd("Savings_UpdateEmergencyFundTile");

    const savingsSh = ss.getSheetByName(CFG.savings.sheet);
    if (savingsSh){
      console.time("SAVINGS.recomputeMain");
      try { SAVINGS.recomputeMain(savingsSh, savingsMaster); } catch(e){ Logger.log(e); }
      console.timeEnd("SAVINGS.recomputeMain");
    }

    try { Monthly_UpdateBankAndCategoryFormatting(); } catch(e){ Logger.log("Bank error: " + e); }

    try { updateIncomeTiles && updateIncomeTiles(); } catch(e){ Logger.log(e); }
    try { Expenses_UpdateSheetTiles && Expenses_UpdateSheetTiles(); } catch(e){ Logger.log(e); }

    try {
      const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
      if (dSh) {
        const targetDate = _.toDate(dSh.getRange(CFG.datedAccountBalance.targetDate).getValue());
        if (targetDate) {
          DATED_ACCOUNT_BALANCE.refreshSourceData(false); // includes updateDatedAccountBalance
        }
      }
    } catch(e){ Logger.log("Dated AB error: " + e); }

    if (savingsSh){
      try { SAVINGS.recomputeTiles(savingsSh, savingsMaster); } catch(e){ Logger.log(e); }
    }

    console.timeEnd("updateAll");
    Logger.log("=== updateAll END === " + ((new Date()-totalStart)/1000) + "s");
  } finally {
    lock.releaseLock();
  }
}

function CreditCards_EnsureBankCheckboxes(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;
  const B = CFG.bank, markCol = CFG.creditCards.bankMarkCol || "O";
  const rows = B.rowEnd - B.rowStart + 1;
  const rng = sh.getRange(B.rowStart, _.c(markCol), rows, 1);
  rng.setHorizontalAlignment("center").setVerticalAlignment("middle");
}

function EnsureCreditCardCategory() {
  const ccCat = (CFG.creditCards && CFG.creditCards.paymentCategory) || "Credit Card Payment";
  try { Lists_AddCategory(ccCat); } catch(e) { Logger.log("CC Category error: " + e); }
}



/* ============================= ACTUALS → MONTH ============================= */
/* Income: sum dated-income Actuals (Include=TRUE) for the current month by Name,
   write totals to the Month Income table's Actual column (match by Name). */
function Actuals_WriteToMonthIncome(){
const ss = __ss();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (!mSh || !dSh) return;

  const y = _.year(), m = _.monthIdx(); // current month (1..12)
  const DI = CFG.datedAccountBalance.income;
  const rows = DI.rowEnd - DI.rowStart + 1;
  const width = _.c(DI.colInclude) - _.c(DI.colName) + 1;

  const data = dSh.getRange(DI.rowStart, _.c(DI.colName), rows, width).getValues();

  const idxName = 0;
  const idxDate = _.c(DI.colDate)   - _.c(DI.colName);
  const idxAct  = _.c(DI.colAct)    - _.c(DI.colName);
const idxInc  = _.c(DI.colInclude)- _.c(DI.colName);

  const byName = new Map();
  for (const r of data){
const inc = _boolCell_(r[idxInc]);
    if (!inc) continue;
    const nm = String(r[idxName]||"").trim(); if (!nm) continue;
    const d  = _.toDate(r[idxDate]);          if (!d || d.getFullYear()!==y || (d.getMonth()+1)!==m) continue;
    const a  = _.n(r[idxAct]); if (!(a>0)) continue;
    byName.set(nm, (byName.get(nm)||0) + a);
  }

  const IT = CFG.monthly.incomeTable;
  const n = IT.rowEnd - IT.rowStart + 1;
  const names = mSh.getRange(IT.rowStart, _.c(IT.colName), n, 1).getValues().map(v=>String(v[0]||"").trim());
  const outRng = mSh.getRange(IT.rowStart, _.c(IT.colAct), n, 1);
  const vals = outRng.getValues();

  for (let i=0;i<n;i++){
    const nm = names[i];
    const sum = nm ? (byName.get(nm)||0) : 0;
    vals[i][0] = sum>0 ? sum : "";
  }
  outRng.setValues(vals);
  _.fmtMoney(outRng);
}

/* Expenses: sum dated-expenses (Include=TRUE) for the current month by Name,
   write totals to the Month Expense table's Actual column (match by Name).
   Interprets “total of all frequencies” as “sum of all dated occurrences this month.” */
function Actuals_WriteToMonthExpenses(){
  const ss = __ss();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (!mSh || !dSh) return;

  const y = _.year(), m = _.monthIdx();
  const E  = CFG.datedAccountBalance.expenses;

  const rows  = E.rowEnd - E.rowStart + 1;
  const width = _.c(E.colInclude) - _.c(E.colName) + 1;
  const data  = dSh.getRange(E.rowStart, _.c(E.colName), rows, width).getValues();

  const iName = 0;
  const iDate = _.c(E.colDate)    - _.c(E.colName);
  const iPred = _.c(E.colPred)    - _.c(E.colName);
  const iAct  = _.c(E.colAct)     - _.c(E.colName);
  const iInc  = _.c(E.colInclude) - _.c(E.colName);

  const byName = new Map();
  for (const r of data){
const inc = _boolCell_(r[idxInc]);
    if (!inc) continue;
    const nm = String(r[iName]||"").trim(); if (!nm) continue;
    const d  = _.toDate(r[iDate]);          if (!d || d.getFullYear()!==y || (d.getMonth()+1)!==m) continue;

  const act  = +r[iAct]  || 0;
const pred = +r[iPred] || 0;
const aep = act > 0 ? act : pred;
if (aep <= 0) continue;


    byName.set(nm, (byName.get(nm)||0) + aep);
  }

  const MT = CFG.monthly.expenseTable;
  const n = MT.rowEnd - MT.rowStart + 1;
  const names = mSh.getRange(MT.rowStart, _.c(MT.colName), n, 1).getValues().map(v=>String(v[0]||"").trim());
  const outRng = mSh.getRange(MT.rowStart, _.c(MT.colAct), n, 1);
  const vals = outRng.getValues();

  for (let i=0;i<n;i++){
    const nm = names[i];
    const sum = nm ? (byName.get(nm)||0) : 0;
    vals[i][0] = sum>0 ? sum : "";
  }
  outRng.setValues(vals);
  _.fmtMoney(outRng);
}


/* One-click writer + refresh tiles/bank/category */
function WriteActualsToMonth_All(){
  try { Actuals_WriteToMonthIncome(); } catch(e){ Logger.log(e); }
  try { Actuals_WriteToMonthExpenses(); } catch(e){ Logger.log(e); }

  // Fast refresh path you already have:
  try { Monthly_UpdateBankAndCategoryFormatting(); } catch(e){ Logger.log(e); }
  try { updateIncomeTiles && updateIncomeTiles(); } catch(e){ Logger.log(e); }
  try { Expenses_UpdateSheetTiles && Expenses_UpdateSheetTiles(); } catch(e){ Logger.log(e); }

  SpreadsheetApp.getActive().toast("✅ Actuals written to Month (Income & Expenses)", "Done", 3);
}

/* Minimal Reset: clears Month Actuals only (Income/Expenses/Savings).
   If you already have Reset_All(), keep yours; else this is a safe default. */
function Reset_All(){
  const ss = __ss();;
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (mSh){
    const I = CFG.monthly.incomeTable;
    const E = CFG.monthly.expenseTable;
    const S = CFG.savings.monthly;
    mSh.getRange(I.rowStart, _.c(I.colAct), I.rowEnd - I.rowStart + 1, 1).clearContent();
    mSh.getRange(E.rowStart, _.c(E.colAct), E.rowEnd - E.rowStart + 1, 1).clearContent();
    mSh.getRange(S.rowStart, _.c(S.colAct), S.rowEnd - S.rowStart + 1, 1).clearContent();
  }
  try { Monthly_UpdateBankAndCategoryFormatting(); } catch(e){ Logger.log(e); }
  try { updateIncomeTiles && updateIncomeTiles(); } catch(e){ Logger.log(e); }
  try { Expenses_UpdateSheetTiles && Expenses_UpdateSheetTiles(); } catch(e){ Logger.log(e); }
  SpreadsheetApp.getActive().toast("🧹 Cleared Month Actuals", "Reset", 3);
}


/* ============================= TRIGGERS ============================= */



function onOpen(){
  const totalStart = Date.now();
  Logger.log("=== onOpen START ===");

  const ss = __ss();
  const sh = SingleMonth_EnsureSheet();

  // Month rollover once per calendar month
  const now = new Date();
  const ym = STATE.ymStr(now.getFullYear(), now.getMonth()+1);
  const last = STATE.get(STATE.KEY_LAST_YM);
  if (last !== ym) { 
    console.time("Rollover");
    SingleMonth_ClearMonthlyRows(sh); 
    SingleMonth_CarryBankBegins_FromPrevEnd(sh); 
    console.timeEnd("Rollover");
  }
  STATE.set(STATE.KEY_LAST_YM, ym);

  // UI bits (menus must be rebuilt each open)
  console.time("UI setup");
  try { Month_WriteHeader(); } catch(_){}
  try { Lists_AddCategory("Credit Card Payment"); } catch(_){}
  try { createBudgetToolsMenu(); } catch(_){}
  console.timeEnd("UI setup");

  // Toast while we compute
  ss.toast("Updating budget data...", "Budget Tracker", -1);

  // Heavy work
  console.time("updateAll");
  updateAll();
  console.timeEnd("updateAll");

  // Validations (keep if cheap; else move to onEdit of lists)
  console.time("Validations");
  try { Lists_ApplyValidations(); } catch(_) {}
  console.timeEnd("Validations");

  ss.toast("Done!", "Budget Tracker", 1);

  Logger.log("=== onOpen END === " + ((Date.now() - totalStart)/1000).toFixed(3) + "s");
}

function onSelectionChange(e) {
if (!e || !e.range) return;

  const props = PropertiesService.getDocumentProperties();
const last = Number(props.getProperty("__debounce_onSelection_ms") || "0");
  const now = Date.now();
  if (now - last < 250) return; // 250ms debounce window
props.setProperty("__debounce_onSelection_ms", String(now));

  try {
    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    const topLeftCell = sh.getRange(row, col);
    const cellA1 = topLeftCell.getA1Notation();
    const value = range.getValue();
    
    // === AUTOMATIC "+N more" EXPANSION ===
    if (sheetName === CFG.singleMonth.sheet && value) {
      const str = String(value).trim();
      
      if (str.match(/^\+\d+\s+more$/i)) {
        const key = "CAL_OVERFLOW_" + cellA1;
const json = PropertiesService.getDocumentProperties().getProperty(key);
        
        if (json) {
          const data = JSON.parse(json);
          const { y, m, day, items } = data;
          const msg = `📅 ${items.length} expenses on ${m}/${day}/${y}:\n\n` +
            items.map((it, i) => `${i+1}. ${it.name} - $${Number(it.amount||0).toFixed(2)}`).join("\n");
          SpreadsheetApp.getUi().alert(msg);
          return;
        }
      }
    }
    
    // === DATED H2 - Only trigger when SELECTING, not editing ===
    // (Editing is handled by onEdit)
    if (sheetName === "Dated Account Balance" && cellA1 === "H2") {
      const d = _.toDate(sh.getRange("H2").getValue());
      if (!d) return;

      const stamp = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const last = STATE.get("H2_LAST");
      if (stamp === last) return;  // Date hasn't changed, don't refresh

      STATE.set("H2_LAST", stamp);
      sh.getRange("K2").setValue("Updating...");
      SpreadsheetApp.flush();
      
      // Use FALSE because user is just SELECTING, not editing
      // If they edit, onEdit will handle it with TRUE
      DATED_ACCOUNT_BALANCE.refreshSourceData(false);
      
      sh.getRange("K2").clearContent();
      return;
    }
    
  } catch(err) {
  }
}

function onEdit(e){
  if (!e || !e.range) return;

  // --- ultra-light debounce so rapid keystrokes don’t re-run logic
  const props = PropertiesService.getDocumentProperties();
  const last = Number(props.getProperty("__debounce_onEdit_ms") || "0");
  const now = Date.now();
  if (now - last < 250) return;
  props.setProperty("__debounce_onEdit_ms", String(now));

  const sh   = e.range.getSheet();
  const name = sh.getName();
  const row  = e.range.getRow();
  const col  = e.range.getColumn();
  const val  = e.value;

  const D = CFG.datedAccountBalance;

  // ========== 1) Live add to _Lists + bust transfer cache if TP changed ==========
  (function maybeAdd(){
    const v = String(val||"").trim(); if (!v) return;
    function inCol(L){ return col === _.c(L); }
    function inRows(rs,re){ return row>=rs && row<=re; }

    if (name === CFG.singleMonth.sheet){
      const I=CFG.monthly.incomeTable, E=CFG.monthly.expenseTable, B=CFG.bank;
      if (inRows(I.rowStart,I.rowEnd) && inCol(I.colAcct)) return void Lists_AddAccount(v);
      if (inRows(E.rowStart,E.rowEnd) && inCol(E.colAcct)) return void Lists_AddAccount(v);
      if (inRows(B.rowStart,B.rowEnd) && inCol(B.colName)) return void Lists_AddAccount(v);
      if (inRows(E.rowStart,E.rowEnd) && inCol(E.colCat))  return void Lists_AddCategory(v);
    }

    if (name === D.sheet){
      if (inRows(D.accounts.rowStart,D.accounts.rowEnd) && inCol(D.accounts.colName)) return void Lists_AddAccount(v);
      if (inRows(D.expenses.rowStart,D.expenses.rowEnd) && inCol(D.expenses.colAccount)) return void Lists_AddAccount(v);
      if (inRows(D.transfers.rowStart,D.transfers.rowEnd) && (inCol(D.transfers.colFromAcct)||inCol(D.transfers.colToAcct))) return void Lists_AddAccount(v);
      if (inRows(D.income.rowStart,D.income.rowEnd) && inCol(D.income.colAcct)) return void Lists_AddAccount(v);
    }

    if (name === CFG.transfersPayments.sheet){
      const TP=CFG.transfersPayments;
      if (inRows(TP.rowStart,TP.rowEnd) && (inCol(TP.colFromAcct)||inCol(TP.colToAcct))) return void Lists_AddAccount(v);
      if (inRows(TP.rowStart,TP.rowEnd) && inCol(TP.colName)) return void Lists_AddTransferName(v);
    }

    // invalidate current-month expanded transfers (both our helper and CacheService)
    try { _cacheDel_(`xfer:${_.year()}:${_.monthIdx()}`); } catch(_) {}
    try { CacheService.getDocumentCache().remove(`xfer:${_.year()}:${_.monthIdx()}`); } catch(_) {}

    if (name === CFG.income.master.sheet){
      const M=CFG.income.master;
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colAcct)) return void Lists_AddAccount(v);
      // paycheck cache bust (year list)
      try { CacheService.getDocumentCache().remove("paychecks:" + _.year()); } catch(_){}
    }

    if (name === CFG.expenses.master.sheet){
      const M=CFG.expenses.master;
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colAcct)) return void Lists_AddAccount(v);
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colName)) return void Lists_AddExpenseName(v);
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colCat))  return void Lists_AddCategory(v);
    }

    // cheap live refresh for Month
    try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
  })();

  // ========== 2) _Lists sheet → just re-apply validations ==========
  if (name === "_Lists") {
    try { Lists_ApplyValidations(); } catch(err) { Logger.log("Validation update failed: " + err); }
    return;
  }

  // ========== 3) Savings MASTER sheet ==========
  if (name === CFG.savings.sheet){
    const T = CFG.savings.table;
    if (row >= T.rowStart && row <= T.rowEnd) {
      const editCols = [ _.c(T.colStartAmount), _.c(T.colGoalAmount), _.c(T.colStartDate), _.c(T.colContribution), _.c(T.colFreq) ];
      if (editCols.includes(col)) {
        SpreadsheetApp.flush();
        const ssLoc = __ss();
        const savingsSh = ssLoc.getSheetByName(CFG.savings.sheet);
        const mainRows = SAVINGS.readMain();
        SAVINGS.recomputeMain(savingsSh, mainRows);
        const mSh = ssLoc.getSheetByName(CFG.singleMonth.sheet);
        if (mSh) SAVINGS.populateMonth(mSh, _.year(), _.monthIdx(), true, mainRows);
        SAVINGS.recomputeTiles(savingsSh, mainRows);
        return;
      }
    }
  }

  // ========== 4) Dated Account Balance sheet ==========
  if (name === D.sheet){
    const tCell = sh.getRange(D.targetDate);
    const tRow  = tCell.getRow();
    const tCol  = tCell.getColumn();
    const A = D.accounts;

    if (row >= A.rowStart && row <= A.rowEnd && col === _.c(A.colStartBalance)) {
      try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_) {}
      return;
    }

    if (row === tRow && col === tCol) {
      try { DATED_ACCOUNT_BALANCE.refreshSourceData(true); } catch(_) {}
      return;
    }

    if ((row >= D.expenses.rowStart && row <= D.expenses.rowEnd) ||
        (row >= D.transfers.rowStart && row <= D.transfers.rowEnd) ||
        (row >= D.income.rowStart && row <= D.income.rowEnd)) {
      try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_) {}
      return;
    }

    if (row >= A.rowStart && row <= A.rowEnd && col === _.c(A.colName)) {
      try { WarnDupes_UpdateAll(); } catch(_) {}
    }
  }

  // ========== 5) Transfers sheet → refresh dated, bank/category, calendar (COOLDOWN) ==========
  if (name === CFG.transfersPayments.sheet){
    const TP = CFG.transfersPayments;
    if (row >= TP.rowStart && row <= TP.rowEnd) {
      if (_shouldRun_("__cooldown_transfers__", 400)) {
        try {
          DATED_ACCOUNT_BALANCE.refreshSourceData(false);
          Monthly_UpdateBankAndCategoryFormatting();
          Monthly_RenderCalendar_Safe();
        } catch(_) {}
      }
    }
    return;
  }

  // ========== 6) Master expenses → repopulate Month + calendar + tiles ==========
  if (name === CFG.expenses.master.sheet){
    const ssLoc = __ss();
    const mSh = ssLoc.getSheetByName(CFG.singleMonth.sheet);
    if (mSh){
      EXPENSES.populateMonth(mSh, _.year(), _.monthIdx(), true, EXPENSES.readMaster());
      Monthly_RenderCalendar_Safe();
      Expenses_UpdateSheetTiles();
    }
    return;
  }

  // ========== 7) SINGLE MONTH sheet (COOLDOWNS) ==========
  if (name === CFG.singleMonth.sheet){
    const M = CFG.savings.monthly;
    const ET = CFG.expenses.monthly;
    const B = CFG.bank;
    const markCol = (CFG.creditCards && CFG.creditCards.bankMarkCol) || "O";

    if (row >= M.rowStart && row <= M.rowEnd) {
      try { Savings_RecomputeMonthFromMonthOnly(); } catch(err) { Logger.log("Savings error: " + err); }
    }

    if (row >= B.rowStart && row <= B.rowEnd && col === _.c(markCol)) {
      try { Lists_SyncCreditCardPayment(); } catch(_) {}
    }

    if (row >= ET.rowStart && row <= ET.rowEnd) {
      if (_shouldRun_("__cooldown_calendar__", 400)) {
        try { Monthly_RenderCalendar_Safe(); } catch(e) { Logger.log("Calendar error: " + e); }
      }
    }

    if (row >= B.rowStart && row <= B.rowEnd && col === _.c(B.colName)) {
      try { WarnDupes_UpdateAll(); } catch(_) {}
    }

    try {
      const ssLoc2 = __ss();
      const savingsSh = ssLoc2.getSheetByName(CFG.savings.sheet);
      if (savingsSh) {
        const main = SAVINGS.readMain();
        SAVINGS.recomputeMain(savingsSh, main);
        SAVINGS.recomputeTiles(savingsSh, main);
      }
    } catch(e){ Logger.log(e); }

    SpreadsheetApp.flush();

    if (_shouldRun_("__cooldown_bank_tiles__", 400)) {
      try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
      try { updateIncomeTiles && updateIncomeTiles(); } catch(_){}
      try { Expenses_UpdateSheetTiles && Expenses_UpdateSheetTiles(); } catch(_){}
    }
    return;
  }

  // ========== 8) Income tiles sheet → rebuild paychecks, repopulate month, refresh ==========
  if (name === CFG.income.tiles.sheet){
    try { PAYCHECKS.rebuild(); } catch(_){}
    const ssLoc = __ss();
    const mSh = ssLoc.getSheetByName(CFG.singleMonth.sheet);
    if (mSh) INCOME.populateMonth(mSh, _.year(), _.monthIdx(), true, INCOME.readMaster());
    try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
    try { updateIncomeTiles(); } catch(_){}
    try { Expenses_UpdateSheetTiles(); } catch(_){}
    return;
  }

  // ========== 9) Expenses tiles sheet → full orchestration ==========
  if (name === CFG.expenses.tiles.sheet){
    updateAll();
    return;
  }

  // ========== 10) Paychecks sheet → sync dated income & refresh (CACHE-BUST + COOLDOWN ok) ==========
  if (name === CFG.income.paychecks.sheet){
    // paycheck cache bust (year list)
    try { CacheService.getDocumentCache().remove("paychecks:" + _.year()); } catch(_){}
    const P = CFG.income.paychecks;
    if (row>=P.rowStart && row<=P.rowEnd && [_.c(P.colName), _.c(P.colDate), _.c(P.colPred)].includes(col)){
      DatedIncome_UpsertFromPaychecks();
      try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_){}
      try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
      try { updateIncomeTiles(); } catch(_){}
      try { Expenses_UpdateSheetTiles(); } catch(_){}
    }
  }
}

/* ============================= MENU ============================= */
function createBudgetToolsMenu() {
  let ui=null; try { ui = SpreadsheetApp.getUi(); } catch(e){ return; }
  ui.createMenu('🛠️ Budget Tools')
    .addItem('📝 Write Actuals → Month', 'WriteActualsToMonth_All')
    .addItem('🗑️ Reset All Actuals', 'Reset_All')
    .addToUi();
}
