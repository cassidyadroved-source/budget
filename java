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
  /* ===== EXPENSES ===== */
  expenses: {
    master:  { sheet:"Expenses", rowStart:12, rowEnd:148,
               colName:"I", colStart:"L", colFreq:"S", colCat:"V", colPred:"Y", colAcct:"AC" },

    monthly: { rowStart:144, rowEnd:238,  // ← CHANGED from 143
               colName:"I", colDate:"M", colCat:"S",
               colPred:"W", colAct:"Y", colAcct:"AC" },

    tiles:   {
      sheet:"Expenses",
      totalMonthly:"H2", predictedAnnual:"M2",
      upcomingExpense1:"R2", upcomingExpense2:"W2", largestDay:"AB2"
    },

    categoryTiles:{
      sheet:"Expenses",
      highestCatName:"H6", lowestCatName:"M6",
      highestExpName:"R6", lowestExpName:"W6", totalSubs:"AB6"
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




/* ============================= CATEGORY TABLE (MONTH) ============================= */
function Expenses_UpdateCategories_INO(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(CFG.singleMonth.sheet); if(!sh) return;
  const E=CFG.monthly.categoryTable, t=CFG.monthly.expenseTable;

  const width = _.c(t.colAct) - _.c(t.colCat) + 1;
  const rows  = sh.getRange(t.rowStart, _.c(t.colCat), t.rowEnd-t.rowStart+1, width).getValues();
  const iPred = _.c(t.colPred)-_.c(t.colCat);
  const iAct  = _.c(t.colAct) -_.c(t.colCat);

  const ccPaymentCat = (CFG.creditCards && CFG.creditCards.paymentCategory || "Credit Card Payment").toUpperCase();
// ...
for (const row of rows){
  const catRaw = String(row[0]||"").trim();
  if (!catRaw) continue;
  const cat = catRaw.toUpperCase();
  if (cat === ccPaymentCat) continue;              // ← exclude CC payment from category rollup

  const aep = Math.max(_.n(row[iAct]), _.n(row[iPred]));
  if (aep>0) catTotals.set(catRaw, (catTotals.get(catRaw)||0) + aep);
}


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

/* ============================= SMALL HELPERS ============================= */
/* ===== DUPLICATE PREVENTION (Accounts) ===== */
function Lists_ApplyValidations(){
  const ss = SpreadsheetApp.getActive();
  const lists = ss.getSheetByName("_Lists"); if (!lists) return;

  function valFromRange(a1){ 
    const r = lists.getRange(a1);
    return SpreadsheetApp.newDataValidation().requireValueInRange(r, true).build();
  }

  const accountsVal  = valFromRange(CFG.lists.accountsRange);   // e.g. "_Lists!A2:A"
  const categoriesVal= valFromRange(CFG.lists.categoriesRange); // e.g. "_Lists!B2:B"
  const expNamesVal  = valFromRange(CFG.lists.expenseNamesRange);
  const xferNamesVal = valFromRange(CFG.lists.transferNamesRange);

  // Month sheet
  (function(){
    const sh = ss.getSheetByName(CFG.singleMonth.sheet); if (!sh) return;
    const I = CFG.monthly.incomeTable, E = CFG.monthly.expenseTable, B = CFG.bank;

    sh.getRange(I.rowStart, _.c(I.colAcct), I.rowEnd-I.rowStart+1, 1).setDataValidation(accountsVal);
    sh.getRange(E.rowStart, _.c(E.colAcct), E.rowEnd-E.rowStart+1, 1).setDataValidation(accountsVal);
    sh.getRange(E.rowStart, _.c(E.colCat),  E.rowEnd-E.rowStart+1, 1).setDataValidation(categoriesVal);
    sh.getRange(B.rowStart, _.c(B.colName), B.rowEnd-B.rowStart+1, 1).setDataValidation(accountsVal);
    sh.getRange(E.rowStart, _.c(E.colName), E.rowEnd-E.rowStart+1, 1).setDataValidation(expNamesVal);

  })();

  // Expenses master
  (function(){
    const sh = ss.getSheetByName(CFG.expenses.master.sheet); if (!sh) return;
    const M = CFG.expenses.master;
    sh.getRange(M.rowStart, _.c(M.colName), M.rowEnd-M.rowStart+1, 1).setDataValidation(expNamesVal);
    sh.getRange(M.rowStart, _.c(M.colAcct), M.rowEnd-M.rowStart+1, 1).setDataValidation(accountsVal);
    sh.getRange(M.rowStart, _.c(M.colCat),  M.rowEnd-M.rowStart+1, 1).setDataValidation(categoriesVal);
    
  })();

  // Transfers sheet
  (function(){
    const sh = ss.getSheetByName(CFG.transfersPayments.sheet); if (!sh) return;
    const T = CFG.transfersPayments;
    sh.getRange(T.rowStart, _.c(T.colName),     T.rowEnd-T.rowStart+1, 1).setDataValidation(xferNamesVal);
    sh.getRange(T.rowStart, _.c(T.colFromAcct), T.rowEnd-T.rowStart+1, 1).setDataValidation(accountsVal);
    sh.getRange(T.rowStart, _.c(T.colToAcct),   T.rowEnd-T.rowStart+1, 1).setDataValidation(accountsVal);
  })();

  (function(){
  const sh = ss.getSheetByName(CFG.datedAccountBalance.sheet); if (!sh) return;
  const D = CFG.datedAccountBalance;

  // Expenses section
  sh.getRange(D.expenses.rowStart, _.c(D.expenses.colName),     D.expenses.rowEnd - D.expenses.rowStart + 1, 1).setDataValidation(expNamesVal);
  sh.getRange(D.expenses.rowStart, _.c(D.expenses.colCategory), D.expenses.rowEnd - D.expenses.rowStart + 1, 1).setDataValidation(categoriesVal);
  sh.getRange(D.expenses.rowStart, _.c(D.expenses.colAccount),  D.expenses.rowEnd - D.expenses.rowStart + 1, 1).setDataValidation(accountsVal);

  // Transfers section
  sh.getRange(D.transfers.rowStart, _.c(D.transfers.colName),     D.transfers.rowEnd - D.transfers.rowStart + 1, 1).setDataValidation(xferNamesVal);
  sh.getRange(D.transfers.rowStart, _.c(D.transfers.colFromAcct), D.transfers.rowEnd - D.transfers.rowStart + 1, 1).setDataValidation(accountsVal);
  sh.getRange(D.transfers.rowStart, _.c(D.transfers.colToAcct),   D.transfers.rowEnd - D.transfers.rowStart + 1, 1).setDataValidation(accountsVal);

  // Income section
  sh.getRange(D.income.rowStart, _.c(D.income.colAcct), D.income.rowEnd - D.income.rowStart + 1, 1).setDataValidation(accountsVal);
})();
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
  const ss = SpreadsheetApp.getActive();

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
  const ss=SpreadsheetApp.getActive();

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

function ordinal(n){ const s=["th","st","nd","rd"], v=n%100; return n + (s[(v-20)%10] || s[v] || s[0]); }
function formatLongDate(d){
  const wk = d.toLocaleDateString('en-US',{weekday:'long'});
  const mo = d.toLocaleDateString('en-US',{month:'long'});
  return `${wk}, ${mo} ${ordinal(d.getDate())}, ${d.getFullYear()}`;
}
function _groupWidth(i){ const g = CFG.monthly.calendar.dayColGroups[i]; return _.c(g[1]) - _.c(g[0]) + 1; }
function _groupIndexForCol(colLtr){
  const G = CFG.monthly.calendar.dayColGroups;
  const x = _.c(colLtr);
  for (let i=0;i<G.length;i++){ if (x >= _.c(G[i][0]) && x <= _.c(G[i][1])) return i; }
  return -1;
}
function _ensureSheetSize(sh, minRows, minCols){
  if (sh.getMaxRows()    < minRows) sh.insertRowsAfter(sh.getMaxRows(), minRows - sh.getMaxRows());
  if (sh.getMaxColumns() < minCols) sh.insertColumnsAfter(sh.getMaxColumns(), minCols - sh.getMaxColumns());
}
function _panelRect(panelKey){
  const C = CFG.monthly.calendar;
  const p  = (panelKey==="AB") ? C.overflow1 : C.overflow2;
  const gi = _groupIndexForCol(p.col);
  const [cStart, cEnd] = C.dayColGroups[gi];
  const left  = _.c(cStart);
  const width = _.c(cEnd) - left + 1;
return { rowStart:p.rowStart, rowEnd:p.rowEnd, left, width, headerA1: `${p.col}${p.rowStart}` };
}
function _centerOverflowHeader(panelKey){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  const R  = _panelRect(panelKey); // gives rowStart + full merged width
  sh.getRange(R.rowStart, R.left, 1, R.width)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}
function centerOverflowPanels(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  ["AB134:AE134", "AG134:AJ134"].forEach(rng => {
    sh.getRange(rng).setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
}
function _outlineOverflowPanel(panelKey, on){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  const C  = CFG.monthly.calendar;
  const R  = _panelRect(panelKey);
  const color = C.borderColor || "#b7e1cd";
  if (on) {
    sh.getRange(R.rowStart, R.left, R.rowEnd - R.rowStart + 1, R.width)
      .setBorder(true, true, true, true, false, false, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } else {
    sh.getRange(R.rowStart, R.left, R.rowEnd - R.rowStart + 1, R.width).setBorder(false,false,false,false,false,false);
  }
  STATE.set("OV_PANEL_ACTIVE", on ? panelKey : "");
}
function _overflowKeyForA1(a1){ return `OV_${a1}`; }

function needsAccountPopulation(sh, tableConfig) {
  const accounts = sh.getRange(tableConfig.rowStart, _.c(tableConfig.colAcct), 
                               tableConfig.rowEnd - tableConfig.rowStart + 1, 1).getValues();
  for (const row of accounts) { if (String(row[0] || "").trim() !== "") return false; }
  return true;
}
function normalizeAccountName(name) { return String(name || "").trim().toUpperCase(); }

/* ============================= LISTS + VALIDATIONS ============================= */
function Lists_CollectAll(){
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive(), name = "_Lists";
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
  const ss = SpreadsheetApp.getActive();
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

/* ============================= BOOL HELPERS ============================= */
function _boolCell_(v){
  // Normalizes checkbox / TRUE/FALSE / 1/0 / Y/N to a proper boolean
  const t = String(v).trim().toUpperCase();
  if (v === true || v === 1 || t === "TRUE" || t === "Y" || t === "YES") return true;
  if (v === false || v === 0 || t === "FALSE" || t === "N" || t === "NO") return false;
  return false;
}

function _val_fromRange_(rangeA1){
  const ss = SpreadsheetApp.getActive();
  const rng = ss.getRange(rangeA1);
  return SpreadsheetApp.newDataValidation()
          .requireValueInRange(rng, true)   // show dropdown
          .setAllowInvalid(true)
          .build();
}

/* ===== Public list helpers ===== */

function _mapMonthIncomeAcct_(){
  const ss = SpreadsheetApp.getActive(), mSh = ss.getSheetByName(CFG.singleMonth.sheet);
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
    rows.push({ nm, dt, pred, acct, include });
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
  dSh.getRange(DI.rowStart, _.c(DI.colInclude), dN, 1).insertCheckboxes();
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
    dSh.getRange(DI.rowStart, _.c(DI.colInclude), rows, 1).insertCheckboxes();
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
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
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
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;
  for (const key of ["AB","AG"]){
    const R = _panelRect(key);
    sh.getRange(R.rowStart, R.left, R.rowEnd - R.rowStart + 1, R.width)
      .clearContent()
      .setBorder(false,false,false,false,false,false);
  }
  STATE.set("OV_PANEL_ACTIVE", "");
}

function _renderPanel(panelKey, y, m, day, items){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const C = CFG.monthly.calendar;
  const panel = (panelKey === "AB") ? C.overflow1 : C.overflow2;
  const colIdx = _.c(panel.col);

  const fullDate = new Date(y, m - 1, day);
  const header = formatLongDate(fullDate);

  const R = _panelRect(panelKey);
  const headerRange = sh.getRange(R.rowStart, R.left, 1, R.width);
  headerRange.breakApart().merge();
  headerRange.setValue(header);  // Remove setHorizontalAlignment

  const visN = C.visibleLines || 4;
  const hiddenItems = items.slice(visN);

  const firstItemRow = panel.rowStart + 1;
  const lastItemRow  = panel.rowEnd - 1;
  sh.getRange(firstItemRow, colIdx, lastItemRow - firstItemRow + 1, 1).clearContent();

  let r = firstItemRow, i = 0;
  while (r <= lastItemRow && i < hiddenItems.length){
    const it = hiddenItems[i++];
    sh.getRange(r++, colIdx).setValue(`${it.name} ($${Number(it.amount||0).toFixed(0)})`);
  }

  const remaining = hiddenItems.length - i;
  sh.getRange(panel.rowEnd, colIdx).setValue(remaining > 0 ? `+${remaining} more` : "");

  STATE.set(panelKey === "AB" ? "OV_PANEL_AB" : "OV_PANEL_AG", JSON.stringify({ y, m, day, items }));
  _outlineOverflowPanel(panelKey, true);
  STATE.set(_overflowKeyForA1(R.headerA1), JSON.stringify({ y, m, day, items }));
}

function renderOverflowPanel(sh, headerA1){
  const raw = STATE.get(_overflowKeyForA1(headerA1));
  if (!raw) return;
  let payload; try { payload = JSON.parse(raw); } catch(_) { return; }
  if (!payload || !payload.items) return;
  const C = CFG.monthly.calendar;
const panelKey = (headerA1 === `${C.overflow1.col}${C.overflow1.rowStart}`) ? "AB" : "AG";
  _renderPanel(panelKey, payload.y, payload.m, payload.day, payload.items);
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
  const ds = d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  return `E|${nm}|${ds}|${Number(amt||0).toFixed(2)}|${acct||''}`;
}
function _xferKey(nm, d, amt, from, to){
  const ds = d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  return `X|${nm}|${ds}|${Number(amt||0).toFixed(2)}|${from||''}|${to||''}`;
}
/* ============================= EXPENSES (master → month + occurrences) ============================= */
var EXPENSES = (function () {
  const M  = CFG.expenses.master;
  const Mm = CFG.expenses.monthly;
  const ss = SpreadsheetApp.getActive();

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

function populateMonth(sh, y, mIdx, _allowPast, masterRows) {
  const cap = Mm.rowEnd - Mm.rowStart + 1;

  // Read existing categories and accounts BEFORE clearing
  const existingCats = sh.getRange(Mm.rowStart, _.c(Mm.colCat), cap, 1).getValues().map(([v])=>String(v||"").trim());
  const existingAccts = sh.getRange(Mm.rowStart, _.c(Mm.colAcct), cap, 1).getValues().map(([v])=>String(v||"").trim());

  // Clear only Name, Date, Pred (NOT Cat or Acct)
  sh.getRange(Mm.rowStart, _.c(Mm.colName), cap, 1).clearContent();
  sh.getRange(Mm.rowStart, _.c(Mm.colDate), cap, 1).clearContent();
  sh.getRange(Mm.rowStart, _.c(Mm.colPred), cap, 1).clearContent();

  const occ = computeOccurrences(masterRows, y, mIdx);
  const map = new Map();
  for (const o of occ){
    const cur = map.get(o.name);
    if (cur){
      cur.amt += o.amt;
      if (o.date < cur.date) cur.date = o.date;
    } else {
      map.set(o.name, { date: o.date, cat: o.cat, amt: o.amt, acct: o.acct });
    }
  }
  const rows = Array.from(map.entries()).map(([name,v])=>({name, ...v}));

  const n = Math.min(rows.length, cap);
  for (let i=0; i<n; i++){
    const r = Mm.rowStart + i, o = rows[i];
    
    // Always write name, date, pred
    sh.getRange(r, _.c(Mm.colName)).setValue(o.name);
    const d = o.date || new Date(y, mIdx-1, 1);
    sh.getRange(r, _.c(Mm.colDate)).setValue(_.safeDate(y, mIdx, d.getDate()));
    sh.getRange(r, _.c(Mm.colPred)).setValue(o.amt || 0);
    
    // Only write category if cell is EMPTY
    if (!existingCats[i]) {
      sh.getRange(r, _.c(Mm.colCat)).setValue(o.cat || "");
    }
    
    // Only write account if cell is EMPTY
    if (!existingAccts[i]) {
      sh.getRange(r, _.c(Mm.colAcct)).setValue(o.acct || "");
    }
  }
  
  if (n>0){
    _.fmtDate( sh.getRange(Mm.rowStart, _.c(Mm.colDate), n, 1));
    _.fmtMoney(sh.getRange(Mm.rowStart, _.c(Mm.colPred), n, 1));
  }
}

  return { readMaster, computeOccurrences, populateMonth };
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
  const ss = SpreadsheetApp.getActive();
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

function Savings_RecomputeMonthFromMonthOnly(){
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!mSh) return;

  const M = CFG.savings.monthly;
  const mRows = M.rowEnd - M.rowStart + 1;
  
  const nameCol = mSh.getRange(M.rowStart, _.c(M.colName), mRows, 1).getValues();
  const goalCol = mSh.getRange(M.rowStart, _.c(M.colGoal), mRows, 1).getValues();
  const startCol = mSh.getRange(M.rowStart, _.c(M.colStart), mRows, 1).getValues();
  const predCol = mSh.getRange(M.rowStart, _.c(M.colPred), mRows, 1).getValues();
  const actCol = mSh.getRange(M.rowStart, _.c(M.colAct), mRows, 1).getValues();

  Logger.log("DEBUG Savings: First row data:");
  Logger.log("Name: " + nameCol[0][0]);
  Logger.log("Goal: " + goalCol[0][0]);
  Logger.log("Start: " + startCol[0][0]);
  Logger.log("Pred: " + predCol[0][0]);
  Logger.log("Act: " + actCol[0][0]);

  const savingsMaster = SAVINGS.readMain();
  const masterByName = new Map();
  for (const s of savingsMaster) {
    masterByName.set(s.name, s);
  }

  const pctOut = [];
  const finishOut = [];

  for (let r = 0; r < mRows; r++){
    const nm = String(nameCol[r][0] || "").trim();
    
    if (!nm) {
      pctOut.push([""]);
      finishOut.push([""]);
      continue;
    }

    const goal = _.n(goalCol[r][0]);
    const start = _.n(startCol[r][0]);
    const pred = _.n(predCol[r][0]);
    const act = _.n(actCol[r][0]);
    
    const contrib = (act > 0) ? act : pred;
    const bal = start + contrib;

    if (r === 0) {
      Logger.log("DEBUG: contrib=" + contrib + ", bal=" + bal + ", goal=" + goal);
    }

    const pct = (goal > 0) ? Math.max(0, Math.min(1, bal / goal)) : 0;
    pctOut.push([pct]);

    const remaining = Math.max(0, goal - bal);
    
    const master = masterByName.get(nm);
    if (master && remaining > 0) {
      const monthlyEq = getMonthlyEquivalent(master.freq, master.contrib);
      if (monthlyEq > 0) {
        const monthsToFinish = Math.ceil(remaining / monthlyEq);
        const finish = new Date();
        finish.setMonth(finish.getMonth() + monthsToFinish);
        finishOut.push([finish]);
      } else {
        finishOut.push([""]);
      }
    } else {
      finishOut.push([""]);
    }
  }

  Logger.log("DEBUG: Writing to column Z, first value: " + pctOut[0][0]);

  mSh.getRange(M.rowStart, _.c(M.colPct), mRows, 1).setValues(pctOut);
  mSh.getRange(M.rowStart, _.c(M.colFinish), mRows, 1).setValues(finishOut);
  
  _.fmtPct(mSh.getRange(M.rowStart, _.c(M.colPct), mRows, 1));
  _.fmtDate(mSh.getRange(M.rowStart, _.c(M.colFinish), mRows, 1));
  
  Logger.log("DEBUG: Savings percent update complete");
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
  const M = S.monthly;
  const cap = M.rowEnd - M.rowStart + 1;
  
  sh.getRange(M.rowStart, _.c(M.colName),   cap, 1).clearContent();
  sh.getRange(M.rowStart, _.c(M.colGoal),   cap, 1).clearContent();
  sh.getRange(M.rowStart, _.c(M.colStart),  cap, 1).clearContent();
  sh.getRange(M.rowStart, _.c(M.colPred),   cap, 1).clearContent();
  sh.getRange(M.rowStart, _.c(M.colPct),    cap, 1).clearContent();
  sh.getRange(M.rowStart, _.c(M.colFinish), cap, 1).clearContent();

  let r = 0;
  for (const m of mainRows){
    if (r >= cap) break;

    // Calculate THIS month's predicted using calendar occurrences
    const pred = calculateMonthlyPredicted(m.freq, m.contrib, year, mIdx, m.start);
    const bal = (m.startAmt || 0) + pred;
    const pct = (m.goalAmt > 0) ? Math.min(1, Math.max(0, bal / m.goalAmt)) : 0;

    const row = M.rowStart + r;
    sh.getRange(row, _.c(M.colName)).setValue(m.name);
    sh.getRange(row, _.c(M.colGoal)).setValue(m.goalAmt || 0);
    sh.getRange(row, _.c(M.colStart)).setValue(m.startAmt || 0);
    sh.getRange(row, _.c(M.colPred)).setValue(pred);
    sh.getRange(row, _.c(M.colPct)).setValue(pct);

    const remaining = Math.max(0, (m.goalAmt || 0) - bal);
    const monthlyEq = getMonthlyEquivalent(m.freq, m.contrib);
    if (remaining > 0 && monthlyEq > 0){
      const monthsToFinish = Math.ceil(remaining / monthlyEq);
      const finish = new Date(year, mIdx - 1, 1);
      finish.setMonth(finish.getMonth() + monthsToFinish);
      sh.getRange(row, _.c(M.colFinish)).setValue(finish);
    }
    r++;
  }

  if (r > 0){
    _.fmtMoney(sh.getRange(M.rowStart, _.c(M.colGoal),   r, 1));
    _.fmtMoney(sh.getRange(M.rowStart, _.c(M.colStart),  r, 1));
    _.fmtMoney(sh.getRange(M.rowStart, _.c(M.colPred),   r, 1));
    _.fmtPct(  sh.getRange(M.rowStart, _.c(M.colPct),    r, 1));
    _.fmtDate( sh.getRange(M.rowStart, _.c(M.colFinish), r, 1));
  }
}

  function recomputeTiles(_sh, mainRows){
    const mSh = ss.getSheetByName(CFG.singleMonth.sheet); if(!mSh) return;

    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth() + 1;

    const M = CFG.savings.monthly;
    const rows = M.rowEnd - M.rowStart + 1;
    const names = mSh.getRange(M.rowStart, _.c(M.colName), rows, 1).getValues().map(([v])=>String(v||"").trim());
    const preds = mSh.getRange(M.rowStart, _.c(M.colPred), rows, 1).getValues().map(([v])=>_.n(v));
    const acts  = mSh.getRange(M.rowStart, _.c(M.colAct),  rows, 1).getValues().map(([v])=>_.n(v));

    let monthSave=0, amountLeft=0, longest=0, nearest=1e9, nearestName="", totalStartAmounts=0;

    for (const goal of mainRows){
      totalStartAmounts += (goal.startAmt || 0);

      let contrib = 0;
      const i = names.indexOf(goal.name);
      if (i>=0) contrib = (acts[i]>0 ? acts[i] : preds[i]);
      else      contrib = calculateMonthlyPredicted(goal.freq, goal.contrib, y, m, goal.start);

      monthSave += contrib;

      const currentBalance = (goal.startAmt || 0) + contrib;
      const rem = Math.max(0, (goal.goalAmt || 0) - currentBalance);
      amountLeft += rem;

      const monthlyEq = getMonthlyEquivalent(goal.freq, goal.contrib);
      if (monthlyEq > 0 && rem > 0){
        const monthsLeft = Math.ceil(rem / monthlyEq);
        longest = Math.max(longest, monthsLeft);
        if (monthsLeft < nearest){ nearest = monthsLeft; nearestName = goal.name; }
      }
    }

    const MT = CFG.savings.monthlyTiles;
    const sSh = ss.getSheetByName(CFG.savings.sheet);

    _.fmtMoney(mSh.getRange(MT.totalSaved)).setValue(monthSave);
    _.fmtMoney(mSh.getRange(MT.amountLeft)).setValue(amountLeft);
    mSh.getRange(MT.monthsLeft).setValue(longest).setNumberFormat(CFG.formats.months1);
    mSh.getRange(MT.nearestGoal).setValue(nearestName);

    const incomeAEP = _.sumAEP_block(
      mSh, CFG.monthly.incomeTable.rowStart, CFG.monthly.incomeTable.rowEnd,
      CFG.monthly.incomeTable.colPred, CFG.monthly.incomeTable.colAct
    );
    mSh.getRange(MT.savingsToIncome)
       .setValue(incomeAEP>0 ? monthSave/incomeAEP : 0)
       .setNumberFormat(CFG.formats.percent1);

    _.fmtMoney(sSh.getRange(CFG.savings.tiles.totalSaved)).setValue(totalStartAmounts);
    _.fmtMoney(sSh.getRange(CFG.savings.tiles.amountLeft)).setValue(amountLeft);
    sSh.getRange(CFG.savings.tiles.monthsLeft).setValue(longest).setNumberFormat(CFG.formats.months1);
    sSh.getRange(CFG.savings.tiles.nearestToComplete).setValue(nearestName);
  }
  function recomputeMain(mainSh, mainRows){
  const T = S.table;
  const names = mainSh.getRange(T.rowStart, _.c(T.colName), T.rowEnd - T.rowStart + 1, 1).getValues().map(r => String(r[0] || "").trim());
  
  // Get Month savings actuals to calculate real progress
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const M = CFG.savings.monthly;
  const monthNames = mSh ? mSh.getRange(M.rowStart, _.c(M.colName), M.rowEnd - M.rowStart + 1, 1).getValues().map(r => String(r[0] || "").trim()) : [];
  const monthPreds = mSh ? mSh.getRange(M.rowStart, _.c(M.colPred), M.rowEnd - M.rowStart + 1, 1).getValues().map(r => _.n(r[0])) : [];
  const monthActs  = mSh ? mSh.getRange(M.rowStart, _.c(M.colAct),  M.rowEnd - M.rowStart + 1, 1).getValues().map(r => _.n(r[0])) : [];

  for (const m of mainRows){
    const i = names.indexOf(m.name);
    if (i < 0) continue;

    // Get this month's contribution (Actual if exists, else Pred)
    let monthContrib = 0;
    const mi = monthNames.indexOf(m.name);
    if (mi >= 0) {
      monthContrib = (monthActs[mi] > 0) ? monthActs[mi] : monthPreds[mi];
    } else {
      // Calculate from master if not in month table
      const y = _.year();
      const mIdx = _.monthIdx();
      monthContrib = calculateMonthlyPredicted(m.freq, m.contrib, y, mIdx, m.start);
    }

    const bal = (m.startAmt || 0) + monthContrib;
    const pct = (m.goalAmt > 0) ? Math.min(1, Math.max(0, bal / m.goalAmt)) : 0;
    mainSh.getRange(T.rowStart + i, _.c(T.colPct)).setValue(pct);

    const remaining = Math.max(0, (m.goalAmt || 0) - bal);
    const monthlyEq = getMonthlyEquivalent(m.freq, m.contrib);
    if (remaining > 0 && monthlyEq > 0){
      const monthsToFinish = Math.ceil(remaining / monthlyEq);
      const finish = new Date();
      finish.setMonth(finish.getMonth() + monthsToFinish);
      mainSh.getRange(T.rowStart + i, _.c(T.colFinishDate)).setValue(finish);
    } else {
      mainSh.getRange(T.rowStart + i, _.c(T.colFinishDate)).clearContent();
    }
  }

  _.fmtPct(  mainSh.getRange(T.rowStart, _.c(T.colPct),        T.rowEnd - T.rowStart + 1, 1));
  _.fmtDate( mainSh.getRange(T.rowStart, _.c(T.colFinishDate), T.rowEnd - T.rowStart + 1, 1));
  
  recomputeTiles(mainSh, mainRows);
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
      if (!_.inWindowMonth(r.start, null, year, m)) continue;
      const F = _.normFreq(r.freq);

      if (F === "ONETIME"){
        if (r.start && r.start.getFullYear()===year && (r.start.getMonth()+1)===m){
          out.push({ name:r.name, annual:r.annual, monthly:r.perFreq, acct:r.acct || "" });
        }
      } else {
        const occs = _.expand(year, r.start || new Date(year,0,1), null, F);
        const monthOccs = occs.filter(d => (d.getMonth()+1) === m);
        if (monthOccs.length > 0){
          out.push({ name:r.name, annual:r.annual, monthly: r.perFreq * monthOccs.length, acct:r.acct || "" });
        }
      }
    }
    return out;
  }

  function populateMonth(sh, y, mIdx, _allowPast, masterRows){
    const plan = computePlan(masterRows, y, mIdx);
    const rows = Mm.rowEnd - Mm.rowStart + 1;

    const populateAccounts = needsAccountPopulation(sh, Mm);

    sh.getRange(Mm.rowStart, _.c(Mm.colName),        rows, 1).clearContent();
    sh.getRange(Mm.rowStart, _.c(Mm.colPredAnnual),  rows, 1).clearContent();
    sh.getRange(Mm.rowStart, _.c(Mm.colPredMonthly), rows, 1).clearContent();

    const n = Math.min(plan.length, rows);
    for (let i=0; i<n; i++){
      const p = plan[i], r = Mm.rowStart + i;
      sh.getRange(r, _.c(Mm.colName)).setValue(p.name);
      sh.getRange(r, _.c(Mm.colPredAnnual)).setValue(p.annual || 0);
      sh.getRange(r, _.c(Mm.colPredMonthly)).setValue(p.monthly || 0);
      if (populateAccounts){ sh.getRange(r, _.c(Mm.colAcct)).setValue(p.acct || ""); }
    }
    if (n > 0){
      _.fmtMoney(sh.getRange(Mm.rowStart, _.c(Mm.colPredAnnual),  n, 1));
      _.fmtMoney(sh.getRange(Mm.rowStart, _.c(Mm.colPredMonthly), n, 1));
    }
  }

  return { readMaster, computePlan, populateMonth };
})();

/* ============================= TRANSFERS (expand by frequency) ============================= */
var TRANSFERS = (function(){
  const TP = CFG.transfersPayments, ss = SpreadsheetApp.getActive();
  function readAndExpandForMonth(y, m){
    const sh = ss.getSheetByName(TP.sheet); if (!sh) return [];
    const rows  = TP.rowEnd - TP.rowStart + 1;
    const width = _.c(TP.colToAcct) - _.c(TP.colName) + 1;
    const data  = sh.getRange(TP.rowStart, _.c(TP.colName), rows, width).getValues();
    const idx   = L => _.c(L) - _.c(TP.colName);
    const out = [];
    
    for (const r of data){
  
      const name   = String(r[0]||"").trim(); if (!name) continue;
      const start  = _.toDate(r[idx(TP.colDate)]); if (!start) continue;
      const freq   = _.normFreq(r[idx(TP.colFreq)]);
      const from   = String(r[idx(TP.colFromAcct)]||"").trim();
      const to     = String(r[idx(TP.colToAcct)]  ||"").trim();
      const amount = _.n(r[idx(TP.colAmount)]);
      if (!from || !to || amount<=0) continue;

      const dates = (freq==="ONETIME")
        ? ((start.getFullYear()===y && (start.getMonth()+1)===m) ? [start] : [])
        : _.expand(y, start, null, freq).filter(d => (d.getMonth()+1)===m);

      for (const d of dates){ out.push({ name, date:d, from, to, amount, freq }); }
    }
    return out;
  }
  return { readAndExpandForMonth };
})();

/* ============================= PAYCHECKS ============================= */
/* ============================= PAYCHECKS (full-year, Pred only) ============================= */
var PAYCHECKS = (function(){
  const P = CFG.income.paychecks, ss = SpreadsheetApp.getActive();

  function rebuild(){
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
      const dates  = _.expand(yr, anchor, null, F);
      const label  = (F==="WEEKLY")?"Weekly":(F==="BIWEEKLY"?"Bi-weekly":"Monthly");
      for (const d of dates){ rows.push({ name:r.name, date:d, freq:label, amount:r.perFreq||0 }); }
    }

    rows.sort((a,b)=>a.date-b.date);

    const cap = P.rowEnd - P.rowStart + 1;
    // Clear only what we own (Name/Date/Freq/Pred). We do NOT touch Act.
    for (const col of [P.colName, P.colDate, P.colFreq, P.colPred]){
      sh.getRange(P.rowStart, _.c(col), cap, 1).clearContent();
    }

    const n = Math.min(rows.length, cap);
    if (n===0) return;

    sh.getRange(P.rowStart, _.c(P.colName), n, 1).setValues(rows.slice(0,n).map(r=>[r.name]));
    sh.getRange(P.rowStart, _.c(P.colDate), n, 1).setValues(rows.slice(0,n).map(r=>[r.date]));
    sh.getRange(P.rowStart, _.c(P.colFreq), n, 1).setValues(rows.slice(0,n).map(r=>[r.freq]));
    sh.getRange(P.rowStart, _.c(P.colPred), n, 1).setValues(rows.slice(0,n).map(r=>[r.amount]));

    _.fmtDate( sh.getRange(P.rowStart, _.c(P.colDate), n, 1));
    _.fmtMoney(sh.getRange(P.rowStart, _.c(P.colPred), n, 1));
  }
  return { rebuild };
})();


/* ============================= DATED ACCOUNT BALANCE (FIXED) ============================= */
var DATED_ACCOUNT_BALANCE = (function(){
  const D  = CFG.datedAccountBalance;
  const TP = CFG.transfersPayments;
  const ss = SpreadsheetApp.getActive();

  function updateDatedAccountBalance(){
    const sh = ss.getSheetByName(D.sheet); if (!sh) return;
    const targetDate = _.toDate(sh.getRange(D.targetDate).getValue());
    if (!targetDate) return;

    // Accounts & start balances
    const aRows = D.accounts.rowEnd - D.accounts.rowStart + 1;
    const names = sh.getRange(D.accounts.rowStart, _.c(D.accounts.colName), aRows, 1).getValues().map(r=>String(r[0]||"").trim());
    const starts= sh.getRange(D.accounts.rowStart, _.c(D.accounts.colStartBalance), aRows, 1).getValues().map(r=>_.n(r[0]));
    
    const map = new Map();
    for (let i=0;i<names.length;i++){ if (names[i]) map.set(names[i], { start:starts[i]||0, inc:0, exp:0, xfer:0 }); }

    _sumExpenses(sh, targetDate, map);
_sumTransfers(sh, targetDate, map);
_sumIncomeToDate(sh, targetDate, map);

    const outInc=[], outExp=[], outXf=[], outEnd=[];
    for (const nm of names){
      if (!nm){ outInc.push([""]); outExp.push([""]); outXf.push([""]); outEnd.push([""]); continue; }
      const v = map.get(nm) || {start:0,inc:0,exp:0,xfer:0};
      const endBal = v.start + v.inc - v.exp + v.xfer;
      outInc.push([v.inc]); 
      outExp.push([v.exp]); 
      outXf.push([v.xfer]);
      outEnd.push([endBal]);
    }
    
    sh.getRange(D.accounts.rowStart, _.c(D.accounts.colIncome),     aRows, 1).setValues(outInc);
    sh.getRange(D.accounts.rowStart, _.c(D.accounts.colExpense),    aRows, 1).setValues(outExp);
    sh.getRange(D.accounts.rowStart, _.c(D.accounts.colTransfers),  aRows, 1).setValues(outXf);
    sh.getRange(D.accounts.rowStart, _.c(D.accounts.colPredBalance),aRows, 1).setValues(outEnd);

    _.fmtMoney(sh.getRange(D.accounts.rowStart, _.c(D.accounts.colIncome),     aRows, 1));
    _.fmtMoney(sh.getRange(D.accounts.rowStart, _.c(D.accounts.colExpense),    aRows, 1));
    _.fmtMoney(sh.getRange(D.accounts.rowStart, _.c(D.accounts.colTransfers),  aRows, 1));
    _.fmtMoney(sh.getRange(D.accounts.rowStart, _.c(D.accounts.colPredBalance),aRows, 1));
  }

function refreshSourceData(isH2Change){
  const sh = ss.getSheetByName(D.sheet); if (!sh) return;
  const targetDate = _.toDate(sh.getRange(D.targetDate).getValue()); if (!targetDate) return;

  try { PAYCHECKS.rebuild(); } catch(_) {}

  // isH2Change should be TRUE only when H2 changes, FALSE for everything else
  _populateExpensesFromMonth(sh, targetDate, !!isH2Change);
  _populateTransfersFromTP(sh, targetDate, !!isH2Change);
  _populateIncomeFromPaychecks(sh, targetDate, !!isH2Change);

  SpreadsheetApp.flush();
  updateDatedAccountBalance();
}
 function _populateExpensesFromMonth(datedSh, targetDate, isH2Change){
  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;
  const D = CFG.datedAccountBalance;

  const dRows = D.expenses.rowEnd - D.expenses.rowStart + 1;
  const dWidth = _.c(D.expenses.colInclude) - _.c(D.expenses.colName) + 1;
  
  // Read existing data to preserve Actuals and Include preferences
  const existing = datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colName), dRows, dWidth).getValues();
  
  const iName = 0;
  const iDate = _.c(D.expenses.colDate)    - _.c(D.expenses.colName);
  const iPred = _.c(D.expenses.colPred)    - _.c(D.expenses.colName);
  const iAct  = _.c(D.expenses.colAct)     - _.c(D.expenses.colName);
  const iAcct = _.c(D.expenses.colAccount) - _.c(D.expenses.colName);
  const iInc  = _.c(D.expenses.colInclude) - _.c(D.expenses.colName);

  // Store existing Actuals and Include preferences by key
  const existingActuals = new Map();
  const existingIncludes = new Map();
  

    if (!isH2Change) {
  for (const r of existing) {
    const nm   = String(r[iName] || "").trim();
    const dt   = _.toDate(r[iDate]);
    const pred = _.n(r[iPred]);
    const act  = _.n(r[iAct]);
    const acct = String(r[iAcct] || "").trim();
    if (nm && dt) {
      const key = _expKey(nm, dt, pred, acct);
      if (act > 0) existingActuals.set(key, act);
      existingIncludes.set(key, r[iInc] === true);
    }
  }
}
  

  // Expand from master
  const masterRows = EXPENSES.readMaster();
  const expanded = [];

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
      const dates = _.expand(y, anchor, null, F).filter(d => (d.getMonth() + 1) === targetMonth);
      for (const d of dates) {
        expanded.push({ name: r.name, date: d, freq: freqDisplay, cat: r.cat || "", amt, acct: r.acct });
      }
    }
  }

  function formatFreqForDisplay(freq) {
    switch(_.normFreq(freq)) {
      case "WEEKLY": return "Weekly";
      case "BIWEEKLY": return "Bi-weekly";
      case "MONTHLY": return "Monthly";
      case "ONETIME": return "One Time";
      default: return "Monthly";
    }
  }

  // Build output preserving Actuals
  const outData = Array.from({length: dRows}, () => Array(dWidth).fill(""));
  
  const oName = 0;
  const oDate = _.c(D.expenses.colDate) - _.c(D.expenses.colName);
  const oFreq = _idxOrNeg(D.expenses.colName, D.expenses.colFreq);
  const oCat  = _.c(D.expenses.colCategory) - _.c(D.expenses.colName);
  const oPred = _.c(D.expenses.colPred) - _.c(D.expenses.colName);
  const oAct  = _.c(D.expenses.colAct) - _.c(D.expenses.colName);
  const oAcct = _.c(D.expenses.colAccount) - _.c(D.expenses.colName);
  const oInc  = _.c(D.expenses.colInclude) - _.c(D.expenses.colName);

  let wrow = 0;
  for (const e of expanded) {
    if (wrow >= dRows) break;
    
    const key = _expKey(e.name, e.date, e.amt, e.acct);
    
    outData[wrow][oName] = e.name;
    outData[wrow][oDate] = e.date;
    if (oFreq >= 0) outData[wrow][oFreq] = e.freq;
    outData[wrow][oCat]  = e.cat;
    outData[wrow][oPred] = e.amt;
    outData[wrow][oAct]  = existingActuals.get(key) || "";  // RESTORE ACTUAL
    outData[wrow][oAcct] = e.acct;
    
    const defaultInc = (e.date <= targetDate);
    outData[wrow][oInc] = isH2Change ? defaultInc : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);
    
    wrow++;
  }

  datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colName), dRows, dWidth).setValues(outData);
  
  if (wrow > 0) {
    datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colInclude), dRows, 1).insertCheckboxes();
    _.fmtDate(datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colDate), wrow, 1));
    _.fmtMoney(datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colPred), wrow, 1));
    _.fmtMoney(datedSh.getRange(D.expenses.rowStart, _.c(D.expenses.colAct), wrow, 1));
  }

} // ← add this line (end _populateExpensesFromMonth)

 function _populateTransfersFromTP(datedSh, targetDate, isH2Change){
  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;
  const D = CFG.datedAccountBalance;
  const TP = CFG.transfersPayments;
  const ss = SpreadsheetApp.getActive();

  const n = D.transfers.rowEnd - D.transfers.rowStart + 1;
  const dWidth = _.c(D.transfers.colInclude) - _.c(D.transfers.colName) + 1;
  
  const existing = datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colName), n, dWidth).getValues();

  const iName = 0;
  const iDate = _.c(D.transfers.colDate) - _.c(D.transfers.colName);
  const iPred = _.c(D.transfers.colPred) - _.c(D.transfers.colName);
  const iAct  = _.c(D.transfers.colAct) - _.c(D.transfers.colName);
  const iFrom = _.c(D.transfers.colFromAcct) - _.c(D.transfers.colName);
  const iTo   = _.c(D.transfers.colToAcct) - _.c(D.transfers.colName);
  const iInc  = _.c(D.transfers.colInclude) - _.c(D.transfers.colName);

  const existingActuals = new Map();
  const existingIncludes = new Map();

  if (!isH2Change) {
    for (const row of existing) {
      const nm = String(row[iName] || "").trim();
      const dt = _.toDate(row[iDate]);
      const pred = _.n(row[iPred]);
      const from = String(row[iFrom] || "").trim();
      const to = String(row[iTo] || "").trim();
      
      if (nm && dt) {
        const key = _xferKey(nm, dt, pred, from, to);
        const act = _.n(row[iAct]);
        if (act > 0) existingActuals.set(key, act);
        existingIncludes.set(key, row[iInc] === true);
      }
    }
  }

  try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
try { Monthly_RenderCalendar(); } catch(_){}


  const tpSh = ss.getSheetByName(TP.sheet);
  if (!tpSh) return;

  const tpRows = TP.rowEnd - TP.rowStart + 1;
  const tpWidth = _.c(TP.colToAcct) - _.c(TP.colName) + 1;
  const tpData = tpSh.getRange(TP.rowStart, _.c(TP.colName), tpRows, tpWidth).getValues();
  
  const iTpName = 0;
  const iTpDate = _.c(TP.colDate) - _.c(TP.colName);
  const iTpFreq = _.c(TP.colFreq) - _.c(TP.colName);
  const iTpFrom = _.c(TP.colFromAcct) - _.c(TP.colName);
  const iTpTo = _.c(TP.colToAcct) - _.c(TP.colName);
  const iTpAmount = _.c(TP.colAmount) - _.c(TP.colName);

  function formatFreqForDisplay(freq) {
    switch(_.normFreq(freq)) {
      case "WEEKLY": return "Weekly";
      case "BIWEEKLY": return "Bi-weekly";
      case "MONTHLY": return "Monthly";
      case "ONETIME": return "One Time";
      default: return "Monthly";
    }
    } // ← add this line to close the helper


  const expanded = [];
  for (const r of tpData) {
    const name = String(r[iTpName] || "").trim();
    if (!name) continue;
    
    const startDate = _.toDate(r[iTpDate]);
    if (!startDate) continue;
    
    const freq = _.normFreq(r[iTpFreq]);
    const freqDisplay = formatFreqForDisplay(r[iTpFreq]);
    const from = String(r[iTpFrom] || "").trim();
    const to = String(r[iTpTo] || "").trim();
    const amt = _.n(r[iTpAmount]);
    
    if (!from || !to || amt <= 0) continue;
    
    const dates = _.expand(y, startDate, null, freq).filter(d => (d.getMonth() + 1) === targetMonth);
    for (const d of dates) {
      expanded.push({ name, date: d, freq: freqDisplay, from, to, amount: amt });
    }
  }

  const outData = Array.from({length: n}, () => Array(dWidth).fill(""));
  
  const oName = 0;
  const oDate = _.c(D.transfers.colDate) - _.c(D.transfers.colName);
  const oFreq = _.c(D.transfers.colFreq) - _.c(D.transfers.colName);
  const oFrom = _.c(D.transfers.colFromAcct) - _.c(D.transfers.colName);
  const oPred = _.c(D.transfers.colPred) - _.c(D.transfers.colName);
  const oAct  = _.c(D.transfers.colAct) - _.c(D.transfers.colName);
  const oTo   = _.c(D.transfers.colToAcct) - _.c(D.transfers.colName);
  const oInc  = _.c(D.transfers.colInclude) - _.c(D.transfers.colName);

  let wrow = 0;
  for (const t of expanded) {
    if (wrow >= n) break;
    
    const key = _xferKey(t.name, t.date, t.amount, t.from, t.to);
    
    outData[wrow][oName] = t.name;
    outData[wrow][oDate] = t.date;
    outData[wrow][oFreq] = t.freq;
    outData[wrow][oFrom] = t.from;
    outData[wrow][oPred] = t.amount;
    outData[wrow][oAct]  = existingActuals.get(key) || "";  // RESTORE ACTUAL
    outData[wrow][oTo]   = t.to;
    
    const defaultInc = (t.date <= targetDate);
    outData[wrow][oInc] = isH2Change ? defaultInc : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);
    
    wrow++;
  }

  datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colName), n, dWidth).setValues(outData);
  
  if (wrow > 0) {
    datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colInclude), n, 1).insertCheckboxes();
    _.fmtDate(datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colDate), wrow, 1));
    _.fmtMoney(datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colPred), wrow, 1));
    _.fmtMoney(datedSh.getRange(D.transfers.rowStart, _.c(D.transfers.colAct), wrow, 1));
  }

} // ← add this line (end _populateTransfersFromTP)

function _populateIncomeFromPaychecks(datedSh, targetDate, isH2Change){
  const y = targetDate.getFullYear();
  const targetMonth = targetDate.getMonth() + 1;
  const D = CFG.datedAccountBalance;
  const ss = SpreadsheetApp.getActive();

  const DI = D.income;
  const n = DI.rowEnd - DI.rowStart + 1;
  const dWidth = _.c(DI.colInclude) - _.c(DI.colName) + 1;
  
  const existing = datedSh.getRange(DI.rowStart, _.c(DI.colName), n, dWidth).getValues();

  const iName = 0;
  const iDate = _.c(DI.colDate) - _.c(DI.colName);
  const iPred = _.c(DI.colPred) - _.c(DI.colName);
  const iAct  = _.c(DI.colAct) - _.c(DI.colName);
  const iAcct = _.c(DI.colAcct) - _.c(DI.colName);
  const iInc  = _.c(DI.colInclude) - _.c(DI.colName);

  const existingActuals = new Map();
  const existingIncludes = new Map();

  if (!isH2Change) {
    for (const row of existing) {
      const nm = String(row[iName] || "").trim();
      const dt = _.toDate(row[iDate]);
      const pred = _.n(row[iPred]);
      const acct = String(row[iAcct] || "").trim();
      
      if (nm && dt) {
        const key = `INC|${nm}|${dt.toDateString()}|${pred}|${acct}`;
        const act = _.n(row[iAct]);
        if (act > 0) existingActuals.set(key, act);
        existingIncludes.set(key, row[iInc] === true);
      }
    }
  }

  const paySh = ss.getSheetByName(CFG.income.paychecks.sheet);
  if (!paySh) return;

  const P = CFG.income.paychecks;
  const pRows = P.rowEnd - P.rowStart + 1;
  const pWidth = _.c(P.colPred) - _.c(P.colName) + 1;
  const pData = paySh.getRange(P.rowStart, _.c(P.colName), pRows, pWidth).getValues();
  
  const pNm = 0;
  const pDt = _.c(P.colDate) - _.c(P.colName);
  const pFq = _.c(P.colFreq) - _.c(P.colName);
  const pPr = _.c(P.colPred) - _.c(P.colName);

  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const it = CFG.monthly.incomeTable;
  const iWidth = _.c(it.colAcct) - _.c(it.colName) + 1;
  const incData = mSh ? mSh.getRange(it.rowStart, _.c(it.colName), it.rowEnd - it.rowStart + 1, iWidth).getValues() : [];
  const idxAcct = _.c(it.colAcct) - _.c(it.colName);
  
  const acctByName = new Map();
  for (const r of incData) {
    const nm = String(r[0] || "").trim();
    const ac = String(r[idxAcct] || "").trim();
    if (nm && ac) acctByName.set(nm, ac);
  }

  const expanded = [];
  for (const r of pData) {
    const nm = String(r[pNm] || "").trim();
    if (!nm) continue;
    
    const d = _.toDate(r[pDt]);
    if (!d || d.getFullYear() !== y || (d.getMonth() + 1) !== targetMonth) continue;
    
    const freqLabel = String(r[pFq] || "").trim();
    const pred = _.n(r[pPr]);
    const acct = acctByName.get(nm) || "";
    
    if (pred > 0) expanded.push({ name: nm, date: d, freq: freqLabel, pred, acct });
  }

  const outData = Array.from({length: n}, () => Array(dWidth).fill(""));
  
  const oNm = 0;
  const oDt = _.c(DI.colDate) - _.c(DI.colName);
  const oFq = _.c(DI.colFreq) - _.c(DI.colName);
  const oPr = _.c(DI.colPred) - _.c(DI.colName);
  const oAc = _.c(DI.colAct) - _.c(DI.colName);
  const oAcct = _.c(DI.colAcct) - _.c(DI.colName);
  const oIn = _.c(DI.colInclude) - _.c(DI.colName);

  let w = 0;
  for (const z of expanded) {
    if (w >= n) break;
    
    const key = `INC|${z.name}|${z.date.toDateString()}|${z.pred}|${z.acct}`;
    
    outData[w][oNm] = z.name;
    outData[w][oDt] = z.date;
    outData[w][oFq] = z.freq;
    outData[w][oPr] = z.pred;
    outData[w][oAc] = existingActuals.get(key) || "";  // RESTORE ACTUAL
    outData[w][oAcct] = z.acct;
    
    const defaultInc = (z.date <= targetDate);
    outData[w][oIn] = isH2Change ? defaultInc : (existingIncludes.has(key) ? existingIncludes.get(key) : defaultInc);
    
    w++;
  }

  datedSh.getRange(DI.rowStart, _.c(DI.colName), n, dWidth).setValues(outData);
  
  if (w > 0) {
    datedSh.getRange(DI.rowStart, _.c(DI.colInclude), n, 1).insertCheckboxes();
    _.fmtDate(datedSh.getRange(DI.rowStart, _.c(DI.colDate), w, 1));
    _.fmtMoney(datedSh.getRange(DI.rowStart, _.c(DI.colPred), w, 1));
    _.fmtMoney(datedSh.getRange(DI.rowStart, _.c(DI.colAct), w, 1));
  }

} // ← add this line (end _populateIncomeFromPaychecks)



function _sumExpenses(sh, targetDate, map){
  const D = CFG.datedAccountBalance;
  const n = D.expenses.rowEnd - D.expenses.rowStart + 1;
  const width = _.c(D.expenses.colInclude) - _.c(D.expenses.colName) + 1;
  const data = sh.getRange(D.expenses.rowStart, _.c(D.expenses.colName), n, width).getValues();
  
  const iDate = _.c(D.expenses.colDate)    - _.c(D.expenses.colName);
  const iPred = _.c(D.expenses.colPred)    - _.c(D.expenses.colName);
  const iAct  = _.c(D.expenses.colAct)     - _.c(D.expenses.colName);
  const iAcc  = _.c(D.expenses.colAccount) - _.c(D.expenses.colName);
  const iInc  = _.c(D.expenses.colInclude) - _.c(D.expenses.colName);

  for (const r of data){
    const inc  = (r[iInc]===true) || (String(r[iInc]).toUpperCase()==="TRUE") || (r[iInc]===1);
    if (!inc) continue;
    
    const d    = _.toDate(r[iDate]);
    if (!d || d > targetDate) continue;
    
    const amtA = _.n(r[iAct]);
    const amtP = _.n(r[iPred]);
    const amt  = (amtA>0 ? amtA : amtP);
    if (amt<=0) continue;
    
    const acc  = String(r[iAcc]||"").trim();
    if (!acc) continue;

    if (map.has(acc)) map.get(acc).exp += amt;
  }
}

function _sumTransfers(sh, targetDate, map){
  const D = CFG.datedAccountBalance;
  const n = D.transfers.rowEnd - D.transfers.rowStart + 1;
  const width = _.c(D.transfers.colInclude) - _.c(D.transfers.colName) + 1;
  const data = sh.getRange(D.transfers.rowStart, _.c(D.transfers.colName), n, width).getValues();
  
  const iDate = _.c(D.transfers.colDate)     - _.c(D.transfers.colName);
  const iFrm  = _.c(D.transfers.colFromAcct) - _.c(D.transfers.colName);
  const iPred = _.c(D.transfers.colPred)     - _.c(D.transfers.colName);
  const iAct  = _.c(D.transfers.colAct)      - _.c(D.transfers.colName);
  const iTo   = _.c(D.transfers.colToAcct)   - _.c(D.transfers.colName);
  const iInc  = _.c(D.transfers.colInclude)  - _.c(D.transfers.colName);

  for (const r of data){
    const inc = (r[iInc]===true) || (String(r[iInc]).toUpperCase()==="TRUE") || (r[iInc]===1);
    if (!inc) continue;
    
    const d   = _.toDate(r[iDate]);
    if (!d || d > targetDate) continue;
    
    const from= String(r[iFrm]||"").trim();
    const to  = String(r[iTo] ||"").trim();
    const amtA = _.n(r[iAct]);
    const amtP = _.n(r[iPred]);
    const amt  = (amtA>0 ? amtA : amtP);
    
    if (!from || !to || amt<=0) continue;
    
    if (map.has(from)) map.get(from).xfer -= amt;
    if (map.has(to))   map.get(to).xfer   += amt;
  }
 }

 function _sumIncomeToDate(sh, targetDate, map){
  const DI = CFG.datedAccountBalance.income;
  const n = DI.rowEnd - DI.rowStart + 1;
  const width = _.c(DI.colInclude) - _.c(DI.colName) + 1;
  const data = sh.getRange(DI.rowStart, _.c(DI.colName), n, width).getValues();
  const iDat = _.c(DI.colDate) - _.c(DI.colName);
  const iPrd = _.c(DI.colPred) - _.c(DI.colName);
  const iAcc = _.c(DI.colAcct) - _.c(DI.colName);
  const iInc = _.c(DI.colInclude) - _.c(DI.colName);

  for (const r of data){
    const inc = (r[iInc]===true) || (String(r[iInc]).toUpperCase()==="TRUE") || (r[iInc]===1);
    if (!inc) continue;
    const d   = _.toDate(r[iDat]);
    if (!d || d > targetDate) continue;                // ← (unchanged) date gate
    const amt = _.n(r[iPrd]); if (amt<=0) continue;
    const acc = String(r[iAcc]||"").trim(); if (!acc) continue;
    if (map.has(acc)) map.get(acc).inc += amt;
  }
 }


  return {
    updateDatedAccountBalance: updateDatedAccountBalance,
    refreshSourceData: refreshSourceData
  };
})(); // ← end of DATED_ACCOUNT_BALANCE IIFE
/* ============================= MONTH TAB: Tiles + Calendar + Bank ============================= */
function Monthly_RenderCalendar(){
  const C = CFG.monthly.calendar;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;

  const per = C.perDayHeight;
  const contentH = 6;
  const G = C.dayColGroups;
  const y = _.year();
  const m = _.monthIdx();

   const firstDow    = new Date(y, m-1, 1).getDay();
  const daysInMonth = new Date(y, m, 0).getDate();

  const leftCol  = _.c("C");
  const rightCol = _.c("AJ");

  // Clear calendar area and borders, then rebuild
  const all = sh.getRange(C.startRow, leftCol, C.perDayHeight*6, rightCol-leftCol+1);
  all.clearContent();
  all.breakApart();
  all.setBorder(false,false,false,false,false,false);

  // build expense & transfer occurrences for the current month
  const occExp  = EXPENSES.computeOccurrences(EXPENSES.readMaster(), y, m);
const occXfer = TRANSFERS.readAndExpandForMonth(y, m)
  .map(t => ({ name: `Transfer: ${t.name}`, date: t.date, amt: t.amount }));

  const occ = occExp.concat(occXfer);

  // Group by day, sort items per day (desc amount)
  const byDay = new Map();
  for (const o of occ){
    const d = o.date.getDate();
    if (!byDay.has(d)) byDay.set(d, []);
    byDay.get(d).push({ name:o.name, amount:o.amt });
  }
  for (const items of byDay.values()){
    items.sort((a,b)=> b.amount - a.amount);
  }

  // Draw 6-week grid
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const color = C.borderColor || "#b7e1cd";

  clearOverflowPanels();

  for (let w=0; w<6; w++){
  const top = C.startRow + w*C.perDayHeight;

  for (let dgi=0; dgi<7; dgi++){
    const idx    = w*7 + dgi;
    const dayNum = 1 + idx - firstDow;
    if (dayNum < 1 || dayNum > daysInMonth) continue;

    const [cStart, cEnd] = C.dayColGroups[dgi];
    const left  = _.c(cStart);
    const right = _.c(cEnd);
    const width = right - left + 1;

    // Merge each line in the day cell vertically
    for (let r=0; r<contentH; r++){
      sh.getRange(top + r, left, 1, width).merge();
    }

    // Outline the day box
    sh.getRange(top, left, contentH, width)
      .setBorder(true, true, true, true, false, false, color, borderStyle);

    // Header = day number
    sh.getRange(top, left).setValue(ordinal(dayNum));

    // Fill visible lines
    const items = byDay.get(dayNum) || [];
    const visN  = C.visibleLines || 4;

    for (let i=0; i<Math.min(items.length, visN); i++){
      sh.getRange(top+1+i, left, 1, width)
        .setValue(`${items[i].name} ($${Number(items[i].amount||0).toFixed(0)})`);
    }

    // Handle overflow ("+N more") and store payload for click
    const hidden = Math.max(0, items.length - visN);
    if (hidden > 0){
      const plusCell = sh.getRange(top + (contentH - 1), left);
      plusCell.setValue(`+${hidden} more`);

      const cellA1 = plusCell.getA1Notation();
      STATE.set(_overflowKeyForA1(cellA1), JSON.stringify({ y, m, day: dayNum, items }));
    }
  }
}

  // Protect with warning-only (best effort)
  try {
    const calendarRange = sh.getRange(C.startRow, leftCol, C.perDayHeight*6, rightCol-leftCol+1);
    const p = calendarRange.protect();
    p.setDescription("Calendar (auto-generated)");
    p.setWarningOnly(true);
  } catch(e){ /* simple triggers / perms can fail here; ignore */ }
}

/* ============================= BANK TOTALS + CATEGORY TABLE + TOP TILES ============================= */
function Monthly_UpdateBankAndCategoryFormatting(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.singleMonth.sheet); if (!sh) return;

  const B = CFG.bank;
  const n = B.rowEnd - B.rowStart + 1;

  // Account names + Pred Begin (user-managed)
  const accounts   = sh.getRange(B.rowStart, _.c(B.colName), n, 1).getValues().map(r => String(r[0]||"").trim());
  const predBegins = sh.getRange(B.rowStart, _.c(B.colPredBegin), n, 1).getValues().map(r => _.n(r[0]));

  const markCol = (CFG.creditCards && CFG.creditCards.bankMarkCol) || "O";
const ccMarks = sh.getRange(B.rowStart, _.c(markCol), n, 1).getValues().map(r => {
  const v = r[0];
  return v === true || String(v).toUpperCase()==="TRUE" || v === 1;
});
const isCC = new Map();
for (let i=0;i<accounts.length;i++){ if (accounts[i]) isCC.set(accounts[i], !!ccMarks[i]); }



  // Income AEP by account
  const incByAcct = new Map();
  {
    const t = CFG.monthly.incomeTable;
    const width = _.c(t.colAcct) - _.c(t.colName) + 1;
    const data  = sh.getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, width).getValues();
    const iPred = _.c(t.colPred) - _.c(t.colName);
    const iAct  = _.c(t.colAct)  - _.c(t.colName);
    const iAcc  = _.c(t.colAcct) - _.c(t.colName);

    for (const r of data){
      const acct = String(r[iAcc]||"").trim(); if (!acct) continue;
      const aep  = (_.n(r[iAct]) > 0) ? _.n(r[iAct]) : _.n(r[iPred]);
      if (aep > 0) incByAcct.set(acct, (incByAcct.get(acct)||0) + aep);
    }
  }

// Expense AEP by account + CC-payment split + month totals for tiles
const expByAcct = new Map();
let monthTotalNoCCPay = 0;   // -> Month!M2
let monthTotalCCPay   = 0;   // -> Month!M10

{
  const t = CFG.monthly.expenseTable;
  const width = _.c(t.colAcct) - _.c(t.colName) + 1;
  const data  = sh.getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, width).getValues();

  const iName = 0;
  const iCat  = _.c(t.colCat)  - _.c(t.colName);
  const iPred = _.c(t.colPred) - _.c(t.colName);
  const iAct  = _.c(t.colAct)  - _.c(t.colName);
  const iAcc  = _.c(t.colAcct) - _.c(t.colName);

  const ccPaymentCatUpper = ((CFG.creditCards && CFG.creditCards.paymentCategory) || "Credit Card Payment").toUpperCase();

  for (const r of data){
    const name = String(r[iName]||"").trim();
    const cat  = String(r[iCat] ||"").trim().toUpperCase();
    const acct = String(r[iAcc] ||"").trim();
    const aep  = (_.n(r[iAct]) > 0) ? _.n(r[iAct]) : _.n(r[iPred]);
    if (aep <= 0) continue;

    if (cat === ccPaymentCatUpper) {
      monthTotalCCPay += aep;                      // for M10
      if (acct) expByAcct.set(acct, (expByAcct.get(acct)||0) + aep);
    } else {
      monthTotalNoCCPay += aep;                    // for M2
      if (acct) expByAcct.set(acct, (expByAcct.get(acct)||0) + aep);
    }
  }
}

// Write Month tiles
try {
  sh.getRange("M2").setValue(monthTotalNoCCPay).setNumberFormat(CFG.formats.money);
  sh.getRange("M10").setValue(monthTotalCCPay).setNumberFormat(CFG.formats.money);
  sh.getRange("H10").setValue(monthTotalNoCCPay + monthTotalCCPay).setNumberFormat(CFG.formats.money);
} catch(_) {}


// Write Month tiles you asked for
try {
  sh.getRange("M2").setValue(monthTotalNoCCPay).setNumberFormat(CFG.formats.money); // total spending (no CC payments)
  sh.getRange("M10").setValue(monthTotalCCPay).setNumberFormat(CFG.formats.money);  // total CC payments
  sh.getRange("H10").setValue(monthTotalNoCCPay + monthTotalCCPay).setNumberFormat(CFG.formats.money);
} catch(_) {}


  // Net transfers for current month only (expanded by frequency)
  const y = _.year(), m = _.monthIdx();
  const xferByAcct = new Map();
  for (const t of TRANSFERS.readAndExpandForMonth(y, m)){
    xferByAcct.set(t.from, (xferByAcct.get(t.from)||0) - t.amount);
    xferByAcct.set(t.to,   (xferByAcct.get(t.to)||0)   + t.amount);
  }

  // Write P/T/W/Y and money formats
  const outExp=[], outInc=[], outXf=[], outEnd=[];
  for (let i=0;i<accounts.length;i++){
    const acc   = accounts[i];
    if (!acc){ outExp.push([""]); outInc.push([""]); outXf.push([""]); outEnd.push([""]); continue; }
    const inc   = incByAcct.get(acc)  || 0;
    const exp   = expByAcct.get(acc)  || 0;
    const xfers = xferByAcct.get(acc) || 0;
    const begin = predBegins[i]       || 0;

    outExp.push([exp]);
  outInc.push([inc]);
  outXf.push([xfers]);
  outEnd.push([begin + inc - exp + xfers]);
}

sh.getRange(B.rowStart, _.c(B.colExpTotal),    n, 1).setValues(outExp);
sh.getRange(B.rowStart, _.c(B.colIncomeTotal), n, 1).setValues(outInc);
sh.getRange(B.rowStart, _.c(B.colTransfer),    n, 1).setValues(outXf);
sh.getRange(B.rowStart, _.c(B.colPredEnd),     n, 1).setValues(outEnd);

  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colPredBegin),   n, 1));
  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colExpTotal),    n, 1));
  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colIncomeTotal), n, 1));
  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colTransfer),    n, 1));
  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colPredEnd),     n, 1));
  _.fmtMoney(sh.getRange(B.rowStart, _.c(B.colActEnd),      n, 1));

  // Refresh the Month category rollup table
  Expenses_UpdateCategories_INO();

  // Refresh top tiles (income, expense, avgDaily, ratio, flow, hi/lo/subs)
  const T  = CFG.monthly.tiles;
  const IT = CFG.monthly.incomeTable;
  const ET = CFG.monthly.expenseTable;

  const incomeTotal  = _.sumAEP_block(sh, IT.rowStart, IT.rowEnd, IT.colPred, IT.colAct);
  const expenseTotal = _.sumAEP_block(sh, ET.rowStart, ET.rowEnd, ET.colPred, ET.colAct);

  const { y:yy, m:mm } = _.currentYM ? _.currentYM() : { y:_.year(), m:_.monthIdx() };
  const daysInMonth  = new Date(yy, mm, 0).getDate();
  const avgDaily     = (expenseTotal>0 && daysInMonth>0) ? (expenseTotal/daysInMonth) : 0;
  const ratio        = (incomeTotal>0) ? (expenseTotal/incomeTotal) : 0;
  const flow         = incomeTotal - expenseTotal;

  sh.getRange(T.income).setValue(incomeTotal).setNumberFormat(CFG.formats.money);
  sh.getRange(T.expense).setValue(expenseTotal).setNumberFormat(CFG.formats.money);
  sh.getRange(T.avgDaily).setValue(avgDaily).setNumberFormat(CFG.formats.money);
  sh.getRange(T.ratio).setValue(ratio).setNumberFormat(CFG.formats.percent1);
  sh.getRange(T.flow).setValue(flow).setNumberFormat(CFG.formats.money);

  // Hi/Lo tiles & subs from the Month expense table (AEP)
  (function(){
    const startCol = _.c(ET.colName);
    const width    = _.c(ET.colAct) - startCol + 1;
    const rowsN    = ET.rowEnd - ET.rowStart + 1;
    const data     = sh.getRange(ET.rowStart, startCol, rowsN, width).getValues();

    const idxName = 0;
    const idxCat  = _.c(ET.colCat)  - startCol;
    const idxPred = _.c(ET.colPred) - startCol;
    const idxAct  = _.c(ET.colAct)  - startCol;

    let hiExpVal = 0, hiExpName = "";
    let loExpVal = Number.POSITIVE_INFINITY, loExpName = "";
    const catTotals = new Map();
    let subsTotal = 0;

    for (const row of data){
      const name = String(row[idxName] || "").trim();
      if (!name) continue;

      const cat = String(row[idxCat] || "").trim();
      const pred = _.n(row[idxPred]);
      const act  = _.n(row[idxAct]);
      const aep  = (act > 0) ? act : pred;
      if (aep <= 0) continue;

      if (aep > hiExpVal){ hiExpVal = aep; hiExpName = name; }
      if (aep < loExpVal){ loExpVal = aep; loExpName = name; }

      if (cat){
        catTotals.set(cat, (catTotals.get(cat) || 0) + aep);
        if (cat === CFG.monthly.subscriptionsCategory) subsTotal += aep;
      }
    }

    let hiCatName = "", loCatName = "";
    if (catTotals.size){
      const sorted = [...catTotals.entries()].filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
      if (sorted.length){
        hiCatName = sorted[0][0];
        loCatName = sorted[sorted.length-1][0];
      }
    }

    sh.getRange(T.hiCat).setValue(hiCatName || "")
      .setNote(hiCatName ? ("Total: $" + (catTotals.get(hiCatName)||0).toFixed(2)) : "");
    sh.getRange(T.loCat).setValue(loCatName || "")
      .setNote(loCatName ? ("Total: $" + (catTotals.get(loCatName)||0).toFixed(2)) : "");
    sh.getRange(T.hiExp).setValue(hiExpName || "")
      .setNote(hiExpVal>0 ? ("Amount: $" + hiExpVal.toFixed(2)) : "");
    sh.getRange(T.loExp).setValue((isFinite(loExpVal)&&loExpVal>0)?loExpName:"")
      .setNote((isFinite(loExpVal)&&loExpVal>0)?("Amount: $" + loExpVal.toFixed(2)):"");

    sh.getRange(T.subs).setValue(subsTotal).setNumberFormat(CFG.formats.money);
  })();
}

/* ============================= “TILES ONLY” REFRESH (fast path for onEdit) ============================= */
function _updateTilesOnly(){
  Monthly_UpdateBankAndCategoryFormatting(); // includes top tiles + categories
}

/* ============================= OPTIONAL: PUBLIC COMPOSER ============================= */
function Monthly_UpdateAllTiles(){
  Monthly_RenderCalendar();                  // calendar grid
  Monthly_UpdateBankAndCategoryFormatting(); // bank + category table + top tiles
}



/* ============================= INCOME SHEET DASH (tiles) ============================= */
function updateIncomeTiles(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(CFG.income.tiles.sheet);
  if(!sh) return;

  const y=_.year();
  const master=INCOME.readMaster();
  const mSh=ss.getSheetByName(CFG.singleMonth.sheet);

  // Predicted annual (frequency-aware, start-date aware)
  let predAnnual = 0;
  for (const r of master){
    const F = _.normFreq(r.freq);
    const occs = _.expand(y, r.start || new Date(y,0,1), null, F);
    const perOcc = r.perFreq || 0;
    predAnnual += perOcc * occs.length;
  }
  _.fmtMoney(sh.getRange(CFG.income.tiles.predictedAnnual)).setValue(predAnnual);

  const predictedMonthlyAvg = predAnnual / 12;
  sh.getRange(CFG.income.tiles.predictedMonthlyAvg)
    .setValue(predictedMonthlyAvg)
    .setNumberFormat(CFG.formats.money);

  // This-month income AEP
  const t=CFG.monthly.incomeTable;
  const thisMonthIncome = _.sumAEP_block(mSh, t.rowStart, t.rowEnd, t.colPred, t.colAct);
  _.fmtMoney(sh.getRange(CFG.income.tiles.thisMonth)).setValue(thisMonthIncome);

  // Variance (actual total - monthly avg)
  const actuals = mSh.getRange(t.rowStart, _.c(t.colAct), t.rowEnd-t.rowStart+1, 1).getValues();
  let actualMonthlyTotal = 0;
  for (const row of actuals) { actualMonthlyTotal += _.n(row[0]); }
 const varCell = sh.getRange(CFG.income.tiles.incomeVariance);
const variance = thisMonthIncome - predictedMonthlyAvg; // Use thisMonthIncome (already calculated above)
if (Math.abs(variance) < 0.01) { varCell.clearContent(); }
else { varCell.setValue(variance).setNumberFormat(CFG.formats.moneyNeg); }



  // Days left this month
  const now=new Date(), last=new Date(now.getFullYear(), now.getMonth()+1, 0);
  sh.getRange(CFG.income.tiles.daysLeft).setValue(Math.max(0, Math.ceil((last-now)/86400000))).setNumberFormat("0");

  // Unbudgeted = thisMonthIncome - expenses AEP - savings AEP
  const expAEP = _.sumAEP_block(mSh, CFG.monthly.expenseTable.rowStart, CFG.monthly.expenseTable.rowEnd, CFG.monthly.expenseTable.colPred, CFG.monthly.expenseTable.colAct);
  const saveAEP = _.sumAEP_block(mSh, CFG.savings.monthly.rowStart, CFG.savings.monthly.rowEnd, CFG.savings.monthly.colPred, CFG.savings.monthly.colAct);
  _.fmtMoney(sh.getRange(CFG.income.tiles.unbudgeted)).setValue(thisMonthIncome - expAEP - saveAEP);
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
  const ss = SpreadsheetApp.getActive();
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
  const ss=SpreadsheetApp.getActive(), sheet=ss.getSheetByName(CFG.expenses.tiles.sheet); if (!sheet) return;
  const mSh=ss.getSheetByName(CFG.singleMonth.sheet);

  const monthTotal = mSh ? _.sumAEP_block(mSh, CFG.expenses.monthly.rowStart, CFG.expenses.monthly.rowEnd, CFG.expenses.monthly.colPred, CFG.expenses.monthly.colAct) : 0;
  _.fmtMoney(sheet.getRange(CFG.expenses.tiles.totalMonthly)).setValue(monthTotal);

  const y=_.year(), predictedAnnual = Expenses_ComputePredictedAnnual(y);
  _.fmtMoney(sheet.getRange(CFG.expenses.tiles.predictedAnnual)).setValue(predictedAnnual);
try {
  const mSh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  const monthM2  = _.n(mSh.getRange("M2").getValue());
  const monthH10 = _.n(mSh.getRange("H10").getValue());
  const monthM10 = _.n(mSh.getRange("M10").getValue());

  // Expenses sheet targets
  const eSh = SpreadsheetApp.getActive().getSheetByName(CFG.expenses.tiles.sheet);
  if (eSh) {
    eSh.getRange("H2").setValue(monthM2).setNumberFormat(CFG.formats.money);        // total month spending
    eSh.getRange("M2").setValue(monthM2 * 12).setNumberFormat(CFG.formats.money);   // predicted annual from M2

    eSh.getRange("H6").setValue(monthH10).setNumberFormat(CFG.formats.money);       // all-out expenses w/ CC
    eSh.getRange("M6").setValue(monthM10).setNumberFormat(CFG.formats.money);       // cc payments total
  }
} catch(_) {}

  updateExpenseExtras();
}

function updateExpenseExtras(){
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const sheet = ss.getSheetByName(CFG.expenses.tiles.sheet);
  if (!mSh || !sheet) return;

  const now = new Date();
  const currentY = now.getFullYear();
  const currentM = now.getMonth() + 1;

  const t = CFG.expenses.monthly;
  const data = mSh
    .getRange(t.rowStart, _.c(t.colName), t.rowEnd - t.rowStart + 1, _.c(t.colAcct) - _.c(t.colName) + 1)
    .getValues();

  const idxDate = _.c(t.colDate) - _.c(t.colName);
  const idxCat  = _.c(t.colCat)  - _.c(t.colName);
  const idxPred = _.c(t.colPred) - _.c(t.colName);
  const idxAct  = _.c(t.colAct)  - _.c(t.colName);

  // Upcoming (>= today) – top 2 by date
  const up1 = sheet.getRange(CFG.expenses.tiles.upcomingExpense1);
  const up2 = sheet.getRange(CFG.expenses.tiles.upcomingExpense2);
  up1.clearContent().setNote(""); 
  up2.clearContent().setNote("");

  const upcoming = [];
  for (const r of data){
    const nm = String(r[0] || "").trim(); if (!nm) continue;
    const d  = _.toDate(r[idxDate]);        if (!d)  continue;
    if (d >= now){
      const days = Math.ceil((d.setHours(0,0,0,0) - (new Date()).setHours(0,0,0,0)) / 86400000);
      upcoming.push({ nm, d: new Date(d), days });
    }
  }
  upcoming.sort((a,b)=>a.d - b.d);
  if (upcoming[0]) up1.setValue(upcoming[0].nm).setNote(upcoming[0].days + " days");
  if (upcoming[1]) up2.setValue(upcoming[1].nm).setNote(upcoming[1].days + " days");

  // Largest day - CURRENT MONTH ONLY
  const perDay = new Map();
  for (const r of data){
    const d = _.toDate(r[idxDate]); 
    if (!d) continue;
    
    // CRITICAL: Filter to current month
    if (d.getFullYear() !== currentY || (d.getMonth() + 1) !== currentM) continue;
    
    const amt = Math.max(_.n(r[idxAct]), _.n(r[idxPred])); 
    if (amt <= 0) continue;
    
    const key = d.toISOString().slice(0,10);
    perDay.set(key, (perDay.get(key) || 0) + amt);
  }
  
  let bestKey = "", bestAmt = 0;
  for (const [k,v] of perDay.entries()){ 
    if (v > bestAmt){ 
      bestAmt = v; 
      bestKey = k; 
    } 
  }
  
  const AB2 = sheet.getRange(CFG.expenses.tiles.largestDay);
  AB2.clearContent().clearNote();  // CLEAR FIRST
  if (bestKey){
    const d = new Date(bestKey + "T12:00:00");  // Force noon to avoid timezone issues
    const displayText = d.toLocaleString('en-US', {month:'short'}) + " " + d.getDate();
    AB2.setValue(displayText).setNote("$" + bestAmt.toFixed(2));
  }

  // Category & item rollups (+ subscriptions)
  const catTotals = new Map();
  const nameTotals = new Map();
  let subsTotal = 0;

  for (const r of data){
    const nm  = String(r[0] || "").trim();
    const cat = String(r[idxCat] || "").trim();
    const amt = Math.max(_.n(r[idxAct]), _.n(r[idxPred]));
    if (!nm || amt <= 0) continue;
    nameTotals.set(nm, (nameTotals.get(nm) || 0) + amt);
    if (cat){
      catTotals.set(cat, (catTotals.get(cat) || 0) + amt);
      if (cat === CFG.monthly.subscriptionsCategory) subsTotal += amt;
    }
  }

  const incomeTotal = _.sumAEP_block(
    mSh, CFG.monthly.incomeTable.rowStart, CFG.monthly.incomeTable.rowEnd,
    CFG.monthly.incomeTable.colPred, CFG.monthly.incomeTable.colAct
  );

  const CT = CFG.expenses.categoryTiles;
  if (catTotals.size){
    const sorted = [...catTotals.entries()].sort((a,b)=>b[1]-a[1]);
    const [hiName, hiAmt] = sorted[0];
    const low = sorted.filter(([,v])=>v>0).slice(-1)[0] || ["—",0];

    sheet.getRange(CT.highestCatName)
      .setValue((hiName||"").split(/\s+/)[0])
      .setNote("$" + hiAmt.toFixed(2) + (incomeTotal>0 ? (" (" + (hiAmt/incomeTotal*100).toFixed(1) + "%)") : ""));

    sheet.getRange(CT.lowestCatName)
      .setValue((low[0]||"—").split(/\s+/)[0])
      .setNote(low[1]>0?`$${low[1].toFixed(2)}${incomeTotal>0?` (${(low[1]/incomeTotal*100).toFixed(1)}%)`:""}`:"No non-zero category");
  } else {
    sheet.getRange(CT.highestCatName).setValue("—").setNote("");
    sheet.getRange(CT.lowestCatName).setValue("—").setNote("");
  }

  if (nameTotals.size){
    const items = [...nameTotals.entries()].sort((a,b)=>b[1]-a[1]);
    const [hiName, hiAmt] = items[0];
    const [loName, loAmt] = items[items.length-1];
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
  const ss=SpreadsheetApp.getActive(), E=CFG.monthly.categoryTable, sh=ss.getSheetByName(CFG.singleMonth.sheet);
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

/* ============================= SINGLE-MONTH ROLLOVER HELPERS ============================= */
function SingleMonth_EnsureSheet(){
  const ss=SpreadsheetApp.getActive(), name=CFG.singleMonth.sheet;
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
  const n = B.rowEnd - B.rowStart + 1;

  // Only clear totals; leave Pred Begin (L) alone
  const colsToClear = [B.colExpTotal, B.colIncomeTotal, B.colTransfer, B.colPredEnd, B.colActEnd];
  for (const col of colsToClear){
    sh.getRange(B.rowStart, _.c(col), n, 1).clearContent();
  }
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
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(CFG.singleMonth.sheet) || ss.insertSheet(CFG.singleMonth.sheet);
  const y=_.year(), mIdx=_.monthIdx();

  // Rebuild paychecks ONCE
  try { PAYCHECKS.rebuild(); } catch(_) {}

  const incomeMaster=INCOME.readMaster();
  const expenseMaster=EXPENSES.readMaster();
  const savingsMaster=SAVINGS.readMain();

  // Populate Month tables
  INCOME.populateMonth(sh, y, mIdx, true, incomeMaster);
  EXPENSES.populateMonth(sh, y, mIdx, true, expenseMaster);
  SAVINGS.populateMonth(sh, y, mIdx, true, savingsMaster);

  // === FIX: Recalculate savings percent AFTER populating ===
  try { Savings_RecomputeMonthFromMonthOnly(); } catch(_) {}
  
  // === FIX: Update emergency fund tile ===
  try { Savings_UpdateEmergencyFundTile(); } catch(_) {}

  // Update Month visuals (calendar + tiles)
  Monthly_UpdateAllTiles();

  // Update Savings (main table + tiles)
  const savingsSh = ss.getSheetByName(CFG.savings.sheet);
  if (savingsSh) {
    SAVINGS.recomputeMain(savingsSh, savingsMaster);
  }

  // Update other dashboards
  updateIncomeTiles();
  Expenses_UpdateSheetTiles();

  // Dated AB - only if date is set
  try {
    const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
    if (dSh) {
      const targetDate = _.toDate(dSh.getRange(CFG.datedAccountBalance.targetDate).getValue());
      if (targetDate) {
        DATED_ACCOUNT_BALANCE.refreshSourceData(false);
        DATED_ACCOUNT_BALANCE.updateDatedAccountBalance();
      }
    }
  } catch(_){}
}
/* ===== CREDIT CARDS ===== */
(function ensureCCcfg(){
  if (!CFG.creditCards) {
    CFG.creditCards = { bankMarkCol: "O", paymentCategory: "Credit Card Payment" };
  }
})();

function CreditCards_EnsureBankCheckboxes(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
  if (!sh) return;
  const B = CFG.bank, markCol = CFG.creditCards.bankMarkCol || "O";
  const rows = B.rowEnd - B.rowStart + 1;
  const rng = sh.getRange(B.rowStart, _.c(markCol), rows, 1);
  rng.insertCheckboxes();
  rng.setHorizontalAlignment("center").setVerticalAlignment("middle");
}

function _CC_computeExpenseSplits_(){
  const ss = SpreadsheetApp.getActive(), mSh = ss.getSheetByName(CFG.singleMonth.sheet); 
  if(!mSh) return {chargesByAcct:new Map(), paymentsByAcct:new Map()};
  const t = CFG.monthly.expenseTable;
  const width = _.c(t.colAcct)-_.c(t.colName)+1;
  const data  = mSh.getRange(t.rowStart, _.c(t.colName), t.rowEnd-t.rowStart+1, width).getValues();
  const idxCat=_.c(t.colCat)-_.c(t.colName), idxPred=_.c(t.colPred)-_.c(t.colName), idxAct=_.c(t.colAct)-_.c(t.colName), idxAcc=_.c(t.colAcct)-_.c(t.colName);
  const payCat = (CFG.creditCards.paymentCategory||"Credit Card Payment").toUpperCase();
  const chargesByAcct = new Map(), paymentsByAcct = new Map();
  for (const row of data){
    const acct = String(row[idxAcc]||"").trim(); if(!acct) continue;
    const cat  = String(row[idxCat]||"").trim().toUpperCase();
    const amt  = (_.n(row[idxAct])>0)?_.n(row[idxAct]):_.n(row[idxPred]);
    if (amt<=0) continue;
    if (cat===payCat) paymentsByAcct.set(acct,(paymentsByAcct.get(acct)||0)+amt);
    else              chargesByAcct.set(acct,(chargesByAcct.get(acct)||0)+amt);
  }
  return { chargesByAcct, paymentsByAcct };
}

function EnsureCreditCardCategory() {
  const ccCat = (CFG.creditCards && CFG.creditCards.paymentCategory) || "Credit Card Payment";
  try { Lists_AddCategory(ccCat); } catch(e) { Logger.log("CC Category error: " + e); }
}

function CreditCards_Enable(){ 
  CreditCards_EnsureBankCheckboxes(); 
  try{ Monthly_UpdateBankAndCategoryFormatting(); }catch(_){} 
}

/* Optional debug helpers */
function DebugCreditCardMath(){
  try {
    const res = _CC_computeExpenseSplits_();
    Logger.log(res);
    SpreadsheetApp.getActive().toast("Logged: CC splits → View > Logs", "Debug", 4);
  } catch(e){ Logger.log(e); }
}

/* ============================= ACTUALS → MONTH ============================= */
/* Income: sum dated-income Actuals (Include=TRUE) for the current month by Name,
   write totals to the Month Income table's Actual column (match by Name). */
function Actuals_WriteToMonthIncome(){
  const ss  = SpreadsheetApp.getActive();
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
    const inc = (r[idxInc]===true) || (String(r[idxInc]).toUpperCase()==="TRUE") || (r[idxInc]===1);
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
  const ss  = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
  const dSh = ss.getSheetByName(CFG.datedAccountBalance.sheet);
  if (!mSh || !dSh) return;

  const y = _.year(), m = _.monthIdx();
  const D  = CFG.datedAccountBalance;
  const E  = D.expenses;

  const rows  = E.rowEnd - E.rowStart + 1;
  const width = _.c(E.colInclude) - _.c(E.colName) + 1;
  const data  = dSh.getRange(E.rowStart, _.c(E.colName), rows, width).getValues();

  const idxName = 0;
  const idxDate = _.c(E.colDate)    - _.c(E.colName);
  const idxAmt  = _.c(E.colAmount)  - _.c(E.colName);
  const idxInc  = _.c(E.colInclude) - _.c(E.colName);

  const byName = new Map();
  for (const r of data){
    const inc = (r[idxInc]===true) || (String(r[idxInc]).toUpperCase()==="TRUE") || (r[idxInc]===1);
    if (!inc) continue;
    const nm = String(r[idxName]||"").trim(); if (!nm) continue;
    const d  = _.toDate(r[idxDate]);          if (!d || d.getFullYear()!==y || (d.getMonth()+1)!==m) continue;
    const a  = _.n(r[idxAmt]); if (!(a>0)) continue;
    byName.set(nm, (byName.get(nm)||0) + a);
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
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();

  const sh = SingleMonth_EnsureSheet();
  const now = new Date();
  const ym = STATE.ymStr(now.getFullYear(), now.getMonth()+1);
  const last = STATE.get(STATE.KEY_LAST_YM);
  
  // === MONTH ROLLOVER CHECK ===
  if (last !== ym) { 
    SingleMonth_ClearMonthlyRows(sh); 
    SingleMonth_CarryBankBegins_FromPrevEnd(sh); 
  }
  STATE.set(STATE.KEY_LAST_YM, ym);

  // === WRITE MONTH HEADER ===
  try { Month_WriteHeader(); } catch(_){}

  try { 
    Lists_AddCategory("Credit Card Payment");  // Force add first
  } catch(_){}



  // === CREATE MENU ===
  try { createBudgetToolsMenu(); } catch(_){}

  // === UPDATE ALL SHEETS ===
  updateAll();
  try { Lists_ApplyValidations(); } catch(_) {}

}
function onSelectionChange(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sh = range.getSheet();
  const sheetName = sh.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const topLeftA1 = sh.getRange(row, col).getA1Notation();
  const a1 = sh.getRange(row, col).getA1Notation(); // <-- use this everywhere

  // === NAVIGATION - run first and short-circuit ===
  try {
    const navMap = {
      "C2": "Month",
      "C3": "Income",
      "C4": "Expenses",
      "C5": "Savings",
      "C28": "Transfers",
      "C29": "Dated Account Balance"
    };
    if (navMap[a1] && sh.getName() !== "_Lists") {
      const target = navMap[a1];
      const ss = SpreadsheetApp.getActive();
      const t = ss.getSheetByName(target);
      if (t) { ss.setActiveSheet(t); t.getRange("A1").activate(); SpreadsheetApp.flush(); return; }
    }
  } catch(_) {}

  // === DATED H2 ===
  if (sheetName === CFG.datedAccountBalance.sheet && a1 === "H2") {
    const d = _.toDate(sh.getRange("H2").getValue());
    if (!d) return;

    const stamp = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const last = STATE.get("H2_LAST");
    if (stamp === last) return;

    STATE.set("H2_LAST", stamp);
    sh.getRange("K2").setValue("Updating...");
    SpreadsheetApp.flush();
    DATED_ACCOUNT_BALANCE.refreshSourceData(false);  // preserve actuals
    sh.getRange("K2").clearContent();
    return;
  }

  // === CALENDAR OVERFLOW ===
  if (sheetName !== CFG.singleMonth.sheet) return;

  const C = CFG.monthly.calendar;
  const calLeft = _.c("C");
  const calRight = _.c("AJ");
  const calTop = C.startRow;
  const calBottom = calTop + (C.perDayHeight * 6) - 1;

  if (row < calTop || row > calBottom || col < calLeft || col > calRight) {
    clearOverflowPanels();
    return;
  }

  const RAB = _panelRect("AB");
  const RAG = _panelRect("AG");

  if (a1 === RAB.headerA1) {
    const raw = STATE.get("OV_PANEL_AB");
    if (raw) {
      const p = JSON.parse(raw);
      _renderPanel("AB", p.y, p.m, p.day, p.items);
    }
    return;
  }

  if (a1 === RAG.headerA1) {
    const raw = STATE.get("OV_PANEL_AG");
    if (raw) {
      const p = JSON.parse(raw);
      _renderPanel("AG", p.y, p.m, p.day, p.items);
    }
    return;
  }

  const inAB = row >= RAB.rowStart && row <= RAB.rowEnd && col >= RAB.left && col < (RAB.left + RAB.width);
  const inAG = row >= RAG.rowStart && row <= RAG.rowEnd && col >= RAG.left && col < (RAG.left + RAG.width);
  if (inAB || inAG) return;

  const contentH = 6;
  const weekIdx = Math.floor((row - C.startRow) / C.perDayHeight);
  if (weekIdx < 0 || weekIdx > 5) { clearOverflowPanels(); return; }

  const dayTop = C.startRow + weekIdx * C.perDayHeight;
  const plusRow = dayTop + (contentH - 1);
  if (row !== plusRow) { clearOverflowPanels(); return; }

  const cellA1 = sh.getRange(row, col).getA1Notation();
  const colLtr = cellA1.replace(/\d+/g, "");
  const dgi = _groupIndexForCol(colLtr);
  if (dgi < 0) { clearOverflowPanels(); return; }

  const [cStart] = C.dayColGroups[dgi];
  const firstCol = _.c(cStart);
  const cellValue = String(sh.getRange(plusRow, firstCol).getValue() || "").trim();
  if (!/^\+\d+\s+more$/.test(cellValue)) { clearOverflowPanels(); return; }

  const y = _.year();
  const m = _.monthIdx();
  const firstDow = new Date(y, m - 1, 1).getDay();
  const idx = weekIdx * 7 + dgi;
  const dayNum = 1 + idx - firstDow;
  const dim = new Date(y, m, 0).getDate();
  if (dayNum < 1 || dayNum > dim) { clearOverflowPanels(); return; }

  const occExp = EXPENSES.computeOccurrences(EXPENSES.readMaster(), y, m)
    .filter(o => o.date.getDate() === dayNum)
    .map(o => ({ name: o.name, amount: o.amt }));

  const occXfer = TRANSFERS.readAndExpandForMonth(y, m)
    .filter(t => t.date.getDate() === dayNum)
    .map(t => ({ name: `Transfer: ${t.name}`, amount: t.amount }));

  const items = occExp.concat(occXfer).sort((a, b) => b.amount - a.amount);
  const visN = C.visibleLines || 4;
  if (items.length <= visN) { clearOverflowPanels(); return; }

  const side = (dgi >= 4) ? "AG" : "AB";
  _renderPanel(side, y, m, dayNum, items);
}



 function onEdit(e){
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const name = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1  = e.range.getA1Notation();
  const val = e.value;

  // Make D available everywhere in this handler
  const D = CFG.datedAccountBalance;

  // === Live add new values to lists
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
if (e && e.range && e.range.getSheet().getName() === "_Lists") {
  try { Lists_ApplyValidations(); } catch(err) { Logger.log("Validation update failed: " + err); }
  return;
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

    if (name === CFG.income.master.sheet){
      const M=CFG.income.master;
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colAcct)) return void Lists_AddAccount(v);
    }

    if (name === CFG.expenses.master.sheet){
      const M=CFG.expenses.master;
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colAcct)) return void Lists_AddAccount(v);
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colName)) return void Lists_AddExpenseName(v);
      if (inRows(M.rowStart,M.rowEnd) && inCol(M.colCat))  return void Lists_AddCategory(v);
    }

    try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}
  })();

  // === Savings MONTH table – early, lightweight
  if (name === CFG.singleMonth.sheet){
    const M = CFG.savings.monthly;
    if (row >= M.rowStart && row <= M.rowEnd) {
      try { Savings_RecomputeMonthFromMonthOnly(); } catch(err) { Logger.log("Savings error: " + err); }
    }
  }

  // === Savings MASTER sheet – recompute on edits to key cols
  if (name === CFG.savings.sheet){
    const T = CFG.savings.table;
    if (row >= T.rowStart && row <= T.rowEnd) {
      const editCols = [ _.c(T.colStartAmount), _.c(T.colGoalAmount), _.c(T.colStartDate), _.c(T.colContribution), _.c(T.colFreq) ];
      if (editCols.includes(col)) {
        const ss = SpreadsheetApp.getActive();
        const savingsSh = ss.getSheetByName(CFG.savings.sheet);
        const mainRows = SAVINGS.readMain();

        // Recompute master table (Percent + Finish Date)
        SAVINGS.recomputeMain(savingsSh, mainRows);

        // Update month table
        const mSh = ss.getSheetByName(CFG.singleMonth.sheet);
        if (mSh) {
          const y = _.year();
          const mIdx = _.monthIdx();
          SAVINGS.populateMonth(mSh, y, mIdx, true, mainRows);
          Savings_RecomputeMonthFromMonthOnly();
        }

        // Update tiles
        SAVINGS.recomputeTiles(savingsSh, mainRows);
        return;
      }
    }
  }

  // === DATED ACCOUNT BALANCE ===
  if (name === D.sheet){
    const tCell = sh.getRange(D.targetDate);
    const tRow  = tCell.getRow();
    const tCol  = tCell.getColumn();

    // H2 date changed → refresh but preserve actuals
    if (row === tRow && col === tCol) {
      try {
        DATED_ACCOUNT_BALANCE.refreshSourceData(false);
        DATED_ACCOUNT_BALANCE.updateDatedAccountBalance();
      } catch(_) {}
      return;
    }

    // Any edit in expenses/transfers/income subtables → recompute
    if ((row >= D.expenses.rowStart && row <= D.expenses.rowEnd) ||
        (row >= D.transfers.rowStart && row <= D.transfers.rowEnd) ||
        (row >= D.income.rowStart    && row <= D.income.rowEnd)) {
      try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_) {}
      return;
    }
  }

  // Transfers sheet edit → recompute dated + month
  if (name === CFG.transfersPayments.sheet){
    try {
      DATED_ACCOUNT_BALANCE.refreshSourceData(false);
      DATED_ACCOUNT_BALANCE.updateDatedAccountBalance();
      Monthly_UpdateBankAndCategoryFormatting();
    } catch(_) {}
    return;
  }

  // AEP change in any dated subtable → recompute dated
  if ((row >= D.expenses.rowStart && row <= D.expenses.rowEnd && (col === _.c(D.expenses.colPred) || col === _.c(D.expenses.colAct))) ||
      (row >= D.transfers.rowStart && row <= D.transfers.rowEnd && (col === _.c(D.transfers.colPred) || col === _.c(D.transfers.colAct))) ||
      (row >= D.income.rowStart    && row <= D.income.rowEnd    && (col === _.c(D.income.colPred)    || col === _.c(D.income.colAct)))) {
    try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_) {}
    return;
  }

  // Include checkbox toggled in any dated subtable → recompute dated
  if ((row >= D.expenses.rowStart  && row <= D.expenses.rowEnd  && col === _.c(D.expenses.colInclude)) ||
      (row >= D.transfers.rowStart && row <= D.transfers.rowEnd && col === _.c(D.transfers.colInclude)) ||
      (row >= D.income.rowStart    && row <= D.income.rowEnd    && col === _.c(D.income.colInclude))) {
    try { DATED_ACCOUNT_BALANCE.updateDatedAccountBalance(); } catch(_) {}
    return;
  }

  // Warn dupes on account name edits
  if (name === CFG.singleMonth.sheet){
    const B = CFG.bank;
    if (row>=B.rowStart && row<=B.rowEnd && col===_.c(B.colName)) { try{ WarnDupes_UpdateAll(); }catch(_){ } }
  }
  if (name === D.sheet){
    const A = D.accounts;
    if (row>=A.rowStart && row<=A.rowEnd && col===_.c(A.colName)) { try{ WarnDupes_UpdateAll(); }catch(_){ } }
  }

  // Master expenses change → rebuild month + visuals + dated
  if (name === CFG.expenses.master.sheet){
    const mSh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
    if (mSh){
      const y=_.year(), mIdx=_.monthIdx();
      EXPENSES.populateMonth(mSh, y, mIdx, true, EXPENSES.readMaster());
      Monthly_RenderCalendar();
      Expenses_UpdateSheetTiles();
      DATED_ACCOUNT_BALANCE.refreshSourceData(false);
      DATED_ACCOUNT_BALANCE.updateDatedAccountBalance();
    }
    return;
  }

  // Month edits (expenses table) → calendar + tiles/rollups
  if (name === CFG.singleMonth.sheet){
    const ET = CFG.monthly.expenseTable;
    const inExpenseTable = row >= ET.rowStart && row <= ET.rowEnd;
    if (inExpenseTable) { Monthly_RenderCalendar(); }

    _updateTilesOnly();

    const ss = SpreadsheetApp.getActive();
    const savingsSh = ss.getSheetByName(CFG.savings.sheet);
    if (savingsSh) {
      const main = SAVINGS.readMain();
      SAVINGS.recomputeMain(savingsSh, main);
      SAVINGS.recomputeTiles(savingsSh, main);
    }
    // inside the SingleMonth onEdit handler branch (you already have this block)
SpreadsheetApp.flush();
try { Monthly_UpdateBankAndCategoryFormatting(); } catch(_){}

    updateIncomeTiles();
    Expenses_UpdateSheetTiles();
    return;
  }

  // Income tiles sheet edit → rebuild paychecks + month + tiles + dated
  if (name === CFG.income.tiles.sheet){
    try { PAYCHECKS.rebuild(); } catch(_){}
    const mSh = SpreadsheetApp.getActive().getSheetByName(CFG.singleMonth.sheet);
    if (mSh){
      const y=_.year(), mIdx=_.monthIdx();
      INCOME.populateMonth(mSh, y, mIdx, true, INCOME.readMaster());
    }
    _updateTilesOnly();
    updateIncomeTiles();
    Expenses_UpdateSheetTiles();
    DATED_ACCOUNT_BALANCE.refreshSourceData(false);
    DATED_ACCOUNT_BALANCE.updateDatedAccountBalance();
    return;
  }

  // Expenses tiles “refresh all” button
  if (name === CFG.expenses.tiles.sheet){
    updateAll();
    return;
  }

  // Paychecks sheet edit → propagate
  if (name === CFG.income.paychecks.sheet){
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
