import { useState, useEffect, useMemo, useRef, createContext, useContext } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard,
  BookOpen,
  FileText,
  List,
  BarChart3,
  Plus,
  Trash2,
  CheckCircle2,
  AlertCircle,
  Upload,
  Download,
  Languages,
  Archive,
} from "lucide-react";

/* ---------------------------------------------------------
   기본 계정과목 (원가계산 제외 · 단순 상거래 기준)
--------------------------------------------------------- */
const DEFAULT_ACCOUNTS = [
  { id: "a1", code: "101", name: "현금", nameEn: "Cash", type: "자산" },
  { id: "a2", code: "102", name: "보통예금", nameEn: "Bank Deposits", type: "자산" },
  { id: "a3", code: "108", name: "매출채권", nameEn: "Accounts Receivable", type: "자산" },
  { id: "a4", code: "120", name: "재고자산", nameEn: "Inventory", type: "자산" },
  { id: "a5", code: "150", name: "유형자산", nameEn: "Fixed Assets", type: "자산" },
  { id: "l1", code: "201", name: "매입채무", nameEn: "Accounts Payable", type: "부채" },
  { id: "l2", code: "205", name: "미지급금", nameEn: "Accrued Liabilities", type: "부채" },
  { id: "l3", code: "210", name: "차입금", nameEn: "Borrowings", type: "부채" },
  { id: "e1", code: "301", name: "자본금", nameEn: "Capital Stock", type: "자본" },
  { id: "e2", code: "305", name: "이익잉여금", nameEn: "Retained Earnings", type: "자본" },
  { id: "r1", code: "401", name: "매출", nameEn: "Sales Revenue", type: "수익" },
  { id: "r2", code: "405", name: "기타수익", nameEn: "Other Income", type: "수익" },
  { id: "x1", code: "501", name: "매출원가", nameEn: "Cost of Goods Sold", type: "비용" },
  { id: "x2", code: "510", name: "급여", nameEn: "Salaries", type: "비용" },
  { id: "x3", code: "515", name: "임차료", nameEn: "Rent Expense", type: "비용" },
  { id: "x4", code: "520", name: "소모품비", nameEn: "Supplies Expense", type: "비용" },
  { id: "x5", code: "525", name: "감가상각비", nameEn: "Depreciation Expense", type: "비용" },
  { id: "x6", code: "530", name: "기타비용", nameEn: "Other Expenses", type: "비용" },
];

const TYPE_ORDER = ["자산", "부채", "자본", "수익", "비용"];
const TYPE_COLOR = {
  자산: "#1C2541",
  부채: "#8A3324",
  자본: "#5B4636",
  수익: "#0E5C46",
  비용: "#8A3324",
};
const TYPE_LABEL = {
  자산: { ko: "자산", en: "Assets" },
  부채: { ko: "부채", en: "Liabilities" },
  자본: { ko: "자본", en: "Equity" },
  수익: { ko: "수익", en: "Revenue" },
  비용: { ko: "비용", en: "Expenses" },
};
const normalOf = (type) => (type === "자산" || type === "비용" ? "debit" : "credit");
const todayStr = () => new Date().toISOString().slice(0, 10);
const yearStartStr = () => `${new Date().getFullYear()}-01-01`;

const fmt = (n) => {
  const v = Math.round(n || 0);
  if (v < 0) return `(${Math.abs(v).toLocaleString("en-US")})`;
  return v.toLocaleString("en-US");
};

const uid = () => Math.random().toString(36).slice(2, 10);

function accName(acc, lang) {
  if (!acc) return "?";
  return lang === "en" && acc.nameEn ? acc.nameEn : acc.name;
}
function typeLabel(type, lang) {
  return TYPE_LABEL[type] ? TYPE_LABEL[type][lang] : type;
}
function normalLabel(normal, lang) {
  if (lang === "en") return normal === "debit" ? "Debit" : "Credit";
  return normal === "debit" ? "차변" : "대변";
}

/* ---------------------------------------------------------
   번역 사전
--------------------------------------------------------- */
const TR = {
  appTitle: { ko: "재무 원장", en: "Finance Ledger" },
  appSub: { ko: "재무회계 · 원가 제외", en: "Financial Accounting · No Costing" },
  navDashboard: { ko: "대시보드", en: "Dashboard" },
  navAccounts: { ko: "계정과목", en: "Chart of Accounts" },
  navEntry: { ko: "전표입력", en: "Journal Entry" },
  navLedger: { ko: "원장", en: "Ledger" },
  navStatements: { ko: "재무제표", en: "Statements" },
  sideFooter: { ko: "전표는 항상 차변 = 대변이어야 저장됩니다.", en: "An entry saves only when debits equal credits." },

  dashTitle: { ko: "대시보드", en: "Dashboard" },
  dashSub: { ko: "기준일 {date}", en: "As of {date}" },
  cardAssets: { ko: "총자산", en: "Total Assets" },
  cardLiab: { ko: "총부채", en: "Total Liabilities" },
  cardEquity: { ko: "총자본", en: "Total Equity" },
  cardNI: { ko: "누적 당기순이익", en: "Cumulative Net Income" },
  won: { ko: "원", en: "KRW" },
  recentEntries: { ko: "최근 전표", en: "Recent Entries" },
  noEntriesDash: { ko: "아직 입력된 전표가 없습니다. 좌측 메뉴에서 [전표입력]으로 첫 거래를 기록해보세요.", en: "No entries yet. Use [Journal Entry] on the left to record your first transaction." },
  colDate: { ko: "일자", en: "Date" },
  colMemo: { ko: "적요", en: "Memo" },
  colAccount: { ko: "계정", en: "Account" },
  colDebit: { ko: "차변", en: "Debit" },
  colCredit: { ko: "대변", en: "Credit" },
  colBalance: { ko: "잔액", en: "Balance" },

  acctTitle: { ko: "계정과목 관리", en: "Chart of Accounts" },
  acctSub: { ko: "자산 · 부채 · 자본 · 수익 · 비용 5개 분류", en: "5 categories: Assets, Liabilities, Equity, Revenue, Expenses" },
  fieldCode: { ko: "계정코드", en: "Code" },
  fieldName: { ko: "계정명", en: "Name" },
  fieldType: { ko: "구분", en: "Type" },
  codePlaceholder: { ko: "예: 601", en: "e.g. 601" },
  namePlaceholder: { ko: "예: 광고선전비", en: "e.g. Advertising Expense" },
  addAccount: { ko: "계정 추가", en: "Add Account" },
  errCodeName: { ko: "계정코드와 계정명을 입력하세요.", en: "Enter both a code and a name." },
  errDupCode: { ko: "이미 존재하는 계정코드입니다.", en: "This account code already exists." },
  errInUse: { ko: "이미 전표에서 사용 중인 계정은 삭제할 수 없습니다.", en: "Cannot delete an account already used in entries." },
  normalBalanceNote: { ko: "({normal} 정상잔액)", en: "(Normal balance: {normal})" },

  entryTitle: { ko: "전표입력 (분개)", en: "Journal Entry" },
  entrySub: { ko: "복식부기 원칙: 차변 합계 = 대변 합계", en: "Double-entry rule: total debits = total credits" },
  uploadDesc: { ko: "여러 건의 전표를 한번에 올리려면 엑셀 양식을 내려받아 작성한 뒤 업로드하세요. (같은 전표번호끼리 묶여 하나의 전표가 됩니다)", en: "To upload many entries at once, download the Excel template, fill it in, then upload it. (Rows sharing the same voucher No. are grouped into one entry)" },
  downloadTemplate: { ko: "템플릿 다운로드", en: "Download Template" },
  uploadExcel: { ko: "엑셀 업로드", en: "Upload Excel" },
  importDone: { ko: "전표 {n}건 업로드 완료", en: "{n} entries uploaded" },
  importErrSuffix: { ko: ", {n}건 오류", en: ", {n} errors" },
  importMoreErr: { ko: "외 {n}건...", en: "and {n} more..." },
  fieldDate: { ko: "일자", en: "Date" },
  fieldMemo: { ko: "적요", en: "Memo" },
  memoPlaceholder: { ko: "예: 사무용품 현금 구매", en: "e.g. Office supplies paid in cash" },
  colAccountTitle: { ko: "계정과목", en: "Account" },
  selectAccount: { ko: "계정 선택", en: "Select account" },
  addLine: { ko: "줄 추가", en: "Add line" },
  balancedYes: { ko: "대차 일치", en: "Balanced" },
  balancedNo: { ko: "대차 불일치", en: "Not balanced" },
  saveEntry: { ko: "전표 저장", en: "Save Entry" },
  sumRow: { ko: "합계", en: "Total" },
  errMemo: { ko: "적요를 입력하세요.", en: "Enter a memo." },
  errLines: { ko: "모든 줄에 계정과 금액을 입력하세요.", en: "Enter an account and amount on every line." },
  errBalance: { ko: "차변 합계와 대변 합계가 일치해야 저장할 수 있습니다.", en: "Debit and credit totals must match before saving." },
  enteredEntries: { ko: "입력된 전표 ({n}건)", en: "Entries ({n})" },
  noEntries: { ko: "입력된 전표가 없습니다.", en: "No entries yet." },
  debitShort: { ko: "차", en: "Dr" },
  creditShort: { ko: "대", en: "Cr" },
  fileReadErr: { ko: "파일을 읽는 중 오류가 발생했습니다. 양식을 확인해주세요.", en: "Could not read the file. Please check the template format." },

  ledgerTitle: { ko: "원장", en: "Ledger" },
  ledgerSub: { ko: "계정과목별 거래 내역과 누적 잔액", en: "Transaction history and running balance by account" },
  selectAccountField: { ko: "계정 선택", en: "Select Account" },
  ledgerHeader: { ko: "{name} 원장 ({type} · {normal} 정상잔액)", en: "{name} Ledger ({type} · Normal balance: {normal})" },
  currentBalance: { ko: "현재 잔액: {n}원", en: "Current balance: {n} KRW" },
  noLedgerRows: { ko: "이 계정에 대한 거래 내역이 없습니다.", en: "No transactions for this account." },

  stmtTitle: { ko: "재무제표", en: "Financial Statements" },
  stmtSub: { ko: "손익계산서 · 재무상태표 (원가계산 제외)", en: "Income Statement · Balance Sheet (no cost accounting)" },
  tabIS: { ko: "손익계산서", en: "Income Statement" },
  tabBS: { ko: "재무상태표", en: "Balance Sheet" },
  fieldStart: { ko: "시작일", en: "Start Date" },
  fieldEnd: { ko: "종료일", en: "End Date" },
  fieldAsOf: { ko: "기준일", en: "As of Date" },

  isTitle: { ko: "손익계산서", en: "Income Statement" },
  isRange: { ko: "{start} ~ {end}", en: "{start} ~ {end}" },
  secRevenue: { ko: "수익", en: "Revenue" },
  secExpense: { ko: "비용", en: "Expenses" },
  totalRevenue: { ko: "수익 합계", en: "Total Revenue" },
  totalExpense: { ko: "비용 합계", en: "Total Expenses" },
  netIncome: { ko: "당기순이익", en: "Net Income" },

  bsTitle: { ko: "재무상태표", en: "Balance Sheet" },
  bsAsOf: { ko: "{date} 기준", en: "As of {date}" },
  secAssets: { ko: "자산", en: "Assets" },
  secLiab: { ko: "부채", en: "Liabilities" },
  secEquity: { ko: "자본", en: "Equity" },
  totalAssets: { ko: "자산 총계", en: "Total Assets" },
  totalLiab: { ko: "부채 합계", en: "Total Liabilities" },
  totalEquity: { ko: "자본 합계", en: "Total Equity" },
  niCumulative: { ko: "당기순이익(누계)", en: "Net Income (cumulative)" },
  totalLiabEquity: { ko: "부채와자본 총계", en: "Total Liabilities & Equity" },
  none: { ko: "해당 없음", en: "None" },

  navAssets: { ko: "고정자산대장", en: "Fixed Asset Register" },
  assetsTitle: { ko: "고정자산대장등록", en: "Fixed Asset Register" },
  assetsSub: { ko: "유형자산 개별 등록 및 정액법 감가상각 계산", en: "Register fixed assets and compute straight-line depreciation" },
  assetsNote: { ko: "이 표는 정액법 기준 참고용 계산입니다. 재무제표에 실제 반영하려면 아래 [감가상각비 전표 생성]을 눌러 전표를 저장하세요.", en: "This table is a reference calculation (straight-line method). To reflect it in the statements, click [Post Depreciation Entry] below." },
  fieldAssetName: { ko: "자산명", en: "Asset Name" },
  fieldAcquireDate: { ko: "취득일", en: "Acquisition Date" },
  fieldCost: { ko: "취득가액", en: "Acquisition Cost" },
  fieldUsefulLife: { ko: "내용연수(년)", en: "Useful Life (yrs)" },
  fieldSalvage: { ko: "잔존가치", en: "Salvage Value" },
  assetNamePlaceholder: { ko: "예: 업무용 노트북", en: "e.g. Office Laptop" },
  addAsset: { ko: "자산 등록", en: "Register Asset" },
  errAssetForm: { ko: "자산명, 취득일, 취득가액, 내용연수를 입력하세요.", en: "Enter name, acquisition date, cost, and useful life." },
  colAcquireDate: { ko: "취득일", en: "Acquired" },
  colCost: { ko: "취득가액", en: "Cost" },
  colAccumDep: { ko: "감가상각누계액", en: "Accum. Depreciation" },
  colBookValue: { ko: "장부가액", en: "Book Value" },
  totalRow: { ko: "합계", en: "Total" },
  noAssets: { ko: "등록된 고정자산이 없습니다.", en: "No fixed assets registered yet." },
  postDepEntry: { ko: "당월 감가상각비 전표 생성", en: "Post This Month's Depreciation Entry" },
  postDepDone: { ko: "감가상각비 전표가 생성되었습니다 ({n}원).", en: "Depreciation entry posted ({n} KRW)." },
  postDepZero: { ko: "이번 달 반영할 감가상각비가 없습니다 (전액 상각 완료 또는 자산 없음).", en: "No depreciation to post this month (fully depreciated or no assets)." },
  postDepMemo: { ko: "당월 감가상각비 반영", en: "Monthly depreciation" },
};

function interpolate(str, vars) {
  if (!vars) return str;
  return Object.keys(vars).reduce((s, k) => s.replaceAll(`{${k}}`, vars[k]), str);
}

const LangCtx = createContext({ lang: "ko", t: (k) => k });
function useT() {
  return useContext(LangCtx);
}

/* -------- 엑셀 날짜/문자 값을 YYYY-MM-DD 문자열로 변환 -------- */
function toDateStr(val) {
  if (val === undefined || val === null || val === "") return todayStr();
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  return s;
}

/* -------- 엑셀 워크북 → 전표 배열로 변환 -------- */
function parseVoucherWorkbook(workbook, accounts) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const groups = {};
  const errors = [];

  rows.forEach((row, i) => {
    const rowNo = i + 2;
    const codeVal = String(row["계정코드"] ?? row["Account Code"] ?? "").trim();
    const nameVal = String(row["계정명"] ?? row["Account Name"] ?? "").trim();
    let acc = accounts.find((a) => a.code === codeVal);
    if (!acc && nameVal) acc = accounts.find((a) => a.name === nameVal || a.nameEn === nameVal);
    if (!acc) {
      errors.push(`${rowNo}: ${codeVal || nameVal || "?"}`);
      return;
    }
    const debit = Number(row["차변"] ?? row["Debit"]) || 0;
    const credit = Number(row["대변"] ?? row["Credit"]) || 0;
    if (!debit && !credit) {
      errors.push(`${rowNo}: 0`);
      return;
    }
    const voucherNo = String(row["전표번호"] ?? row["Voucher No"] ?? "").trim();
    const dateVal = row["일자"] ?? row["Date"];
    const memoVal = row["적요"] ?? row["Memo"];
    const key = voucherNo || `${toDateStr(dateVal)}__${memoVal ?? ""}`;
    if (!groups[key]) groups[key] = { date: toDateStr(dateVal), memo: String(memoVal ?? "").trim() || "-", lines: [] };
    groups[key].lines.push({ accountId: acc.id, debit, credit });
  });

  const newEntries = [];
  Object.entries(groups).forEach(([key, g]) => {
    const td = g.lines.reduce((s, l) => s + l.debit, 0);
    const tc = g.lines.reduce((s, l) => s + l.credit, 0);
    if (td !== tc || td === 0) {
      errors.push(`${key}: ${fmt(td)} / ${fmt(tc)}`);
      return;
    }
    newEntries.push({ id: uid(), date: g.date, memo: g.memo, lines: g.lines });
  });

  return { newEntries, errors };
}

function downloadTemplate(lang) {
  const wb = XLSX.utils.book_new();
  const header = lang === "en"
    ? ["Voucher No", "Date", "Memo", "Account Code", "Debit", "Credit"]
    : ["전표번호", "일자", "적요", "계정코드", "차변", "대변"];
  const memo1 = lang === "en" ? "Office supplies paid in cash" : "사무용품 현금 구매";
  const memo2 = lang === "en" ? "Cash sales" : "상품 매출 (현금)";
  const aoa = [
    header,
    ["V001", "2026-01-05", memo1, "520", 30000, ""],
    ["V001", "2026-01-05", memo1, "101", "", 30000],
    ["V002", "2026-01-10", memo2, "101", 500000, ""],
    ["V002", "2026-01-10", memo2, "401", "", 500000],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 10 }, { wch: 12 }, { wch: 24 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, lang === "en" ? "VoucherTemplate" : "전표양식");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = lang === "en" ? "voucher_template.xlsx" : "전표업로드양식.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

/* -------- 정액법 감가상각 계산 (경과 개월 기준) -------- */
function monthsBetween(fromDate, toDate) {
  const f = new Date(fromDate);
  const t = new Date(toDate);
  if (isNaN(f) || isNaN(t) || t < f) return 0;
  let months = (t.getFullYear() - f.getFullYear()) * 12 + (t.getMonth() - f.getMonth());
  if (t.getDate() < f.getDate()) months -= 1;
  return Math.max(0, months);
}

function depreciationInfo(asset, asOfDate) {
  const cost = Number(asset.cost) || 0;
  const salvage = Number(asset.salvage) || 0;
  const life = Number(asset.usefulLifeYears) || 1;
  const depreciable = Math.max(0, cost - salvage);
  const monthlyDep = depreciable / (life * 12);
  const elapsedMonths = monthsBetween(asset.acquireDate, asOfDate);
  const accumulated = Math.min(depreciable, monthlyDep * elapsedMonths);
  return { monthlyDep, accumulated, bookValue: cost - accumulated };
}

export default function FinanceERP() {
  const [accounts, setAccounts] = useState(DEFAULT_ACCOUNTS);
  const [entries, setEntries] = useState([]);
  const [fixedAssets, setFixedAssets] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [lang, setLang] = useState("ko");

  const t = (key, vars) => interpolate((TR[key] ? TR[key][lang] : key) || key, vars);

  /* -------------------- 영속 저장 (브라우저 저장) -------------------- */
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("erp-data");
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          if (parsed.accounts && parsed.accounts.length) {
            // 이전 버전(영문명 없음 / 신규 계정 없음)에 저장된 데이터면 기본 계정과목 기준으로 보완
            const existingCodes = new Set(parsed.accounts.map((a) => a.code));
            const merged = parsed.accounts.map((a) => {
              if (a.nameEn) return a;
              const match = DEFAULT_ACCOUNTS.find((d) => d.code === a.code);
              return match ? { ...a, nameEn: match.nameEn } : a;
            });
            const missingDefaults = DEFAULT_ACCOUNTS.filter((d) => !existingCodes.has(d.code));
            setAccounts([...merged, ...missingDefaults]);
          } else {
            setAccounts(DEFAULT_ACCOUNTS);
          }
          setEntries(parsed.entries || []);
          setFixedAssets(parsed.fixedAssets || []);
        }
        const langRes = await window.storage.get("erp-lang");
        if (langRes && langRes.value) setLang(langRes.value);
      } catch (e) {
        // 저장된 데이터 없음 - 기본값 사용
      }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await window.storage.set("erp-data", JSON.stringify({ accounts, entries, fixedAssets }));
      } catch (e) {
        // 저장 실패 - 무시하고 계속 진행
      }
    })();
  }, [accounts, entries, fixedAssets, loaded]);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await window.storage.set("erp-lang", lang);
      } catch (e) {}
    })();
  }, [lang, loaded]);

  /* -------------------- 계산 헬퍼 -------------------- */
  const accountById = useMemo(() => {
    const m = {};
    accounts.forEach((a) => (m[a.id] = a));
    return m;
  }, [accounts]);

  function balanceOf(accountId, opts = {}) {
    const { start, end } = opts;
    let debit = 0,
      credit = 0;
    entries.forEach((e) => {
      if (start && e.date < start) return;
      if (end && e.date > end) return;
      e.lines.forEach((l) => {
        if (l.accountId === accountId) {
          debit += Number(l.debit) || 0;
          credit += Number(l.credit) || 0;
        }
      });
    });
    const acc = accountById[accountId];
    const normal = acc ? normalOf(acc.type) : "debit";
    const raw = debit - credit;
    return normal === "debit" ? raw : -raw;
  }

  const netIncomeUpTo = (asOf) => {
    let total = 0;
    accounts.forEach((a) => {
      if (a.type === "수익") total += balanceOf(a.id, { end: asOf });
      if (a.type === "비용") total -= balanceOf(a.id, { end: asOf });
    });
    return total;
  };

  const totalsAsOf = (asOf) => {
    let assets = 0,
      liab = 0,
      equityBase = 0;
    accounts.forEach((a) => {
      const b = balanceOf(a.id, { end: asOf });
      if (a.type === "자산") assets += b;
      if (a.type === "부채") liab += b;
      if (a.type === "자본") equityBase += b;
    });
    const ni = netIncomeUpTo(asOf);
    return { assets, liab, equity: equityBase + ni, ni };
  };

  return (
    <LangCtx.Provider value={{ lang, t }}>
      <div
        className="w-full min-h-[720px] flex text-[#1C2541]"
        style={{
          fontFamily: "'IBM Plex Sans KR','Noto Sans KR',sans-serif",
          background: "#F5F6F3",
        }}
      >
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@600;700&family=IBM+Plex+Sans+KR:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
          .serif { font-family: 'Noto Serif KR', serif; }
          .mono-num { font-family: 'IBM Plex Mono','IBM Plex Sans KR',monospace; font-variant-numeric: tabular-nums; }
          .ledger-table td, .ledger-table th { border-bottom: 1px solid #E3E1D8; }
          .ledger-rule { border-left: 2px solid #8A3324; border-right: 2px solid #8A3324; }
          input, select { outline: none; }
          input:focus, select:focus { box-shadow: 0 0 0 2px rgba(28,37,65,0.25); border-radius: 2px; }
        `}</style>

        <aside
          className="w-[210px] shrink-0 flex flex-col py-6"
          style={{ background: "#16233F", color: "#EDECE4" }}
        >
          <div className="px-5 pb-4 mb-2 border-b border-white/10">
            <div className="flex items-center justify-between">
              <div className="serif text-lg font-bold tracking-tight">{t("appTitle")}</div>
              <button
                onClick={() => setLang(lang === "ko" ? "en" : "ko")}
                title={lang === "ko" ? "Switch to English" : "한국어로 전환"}
                className="flex items-center gap-1 text-[11px] font-semibold px-2 py-1 rounded border border-white/20 hover:bg-white/10"
              >
                <Languages size={12} />
                {lang === "ko" ? "EN" : "KO"}
              </button>
            </div>
            <div className="text-[11px] opacity-60 mt-1.5">{t("appSub")}</div>
          </div>
          <nav className="flex flex-col gap-1 px-3">
            <SideItem icon={<LayoutDashboard size={16} />} label={t("navDashboard")} active={tab === "dashboard"} onClick={() => setTab("dashboard")} />
            <SideItem icon={<BookOpen size={16} />} label={t("navAccounts")} active={tab === "accounts"} onClick={() => setTab("accounts")} />
            <SideItem icon={<Archive size={16} />} label={t("navAssets")} active={tab === "assets"} onClick={() => setTab("assets")} />
            <SideItem icon={<FileText size={16} />} label={t("navEntry")} active={tab === "entry"} onClick={() => setTab("entry")} />
            <SideItem icon={<List size={16} />} label={t("navLedger")} active={tab === "ledger"} onClick={() => setTab("ledger")} />
            <SideItem icon={<BarChart3 size={16} />} label={t("navStatements")} active={tab === "statements"} onClick={() => setTab("statements")} />
          </nav>
          <div className="mt-auto px-5 pt-6 text-[11px] opacity-50 leading-relaxed">
            {t("sideFooter")}
          </div>
        </aside>

        <main className="flex-1 p-8 overflow-auto">
          {tab === "dashboard" && <Dashboard accounts={accounts} entries={entries} totalsAsOf={totalsAsOf} accountById={accountById} />}
          {tab === "accounts" && <AccountsTab accounts={accounts} setAccounts={setAccounts} entries={entries} />}
          {tab === "assets" && <FixedAssetsTab accounts={accounts} fixedAssets={fixedAssets} setFixedAssets={setFixedAssets} entries={entries} setEntries={setEntries} />}
          {tab === "entry" && <EntryTab accounts={accounts} entries={entries} setEntries={setEntries} />}
          {tab === "ledger" && <LedgerTab accounts={accounts} entries={entries} balanceOf={balanceOf} />}
          {tab === "statements" && <StatementsTab accounts={accounts} balanceOf={balanceOf} netIncomeUpTo={netIncomeUpTo} totalsAsOf={totalsAsOf} />}
        </main>
      </div>
    </LangCtx.Provider>
  );
}

function SideItem({ icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-2.5 px-3 py-2.5 rounded text-sm transition-colors"
      style={{
        background: active ? "rgba(255,255,255,0.12)" : "transparent",
        color: active ? "#FFFFFF" : "#C9C7BB",
        fontWeight: active ? 600 : 400,
      }}
    >
      {icon}
      {label}
    </button>
  );
}

/* =============================== 대시보드 =============================== */
function Dashboard({ accounts, entries, totalsAsOf, accountById }) {
  const { lang, t } = useT();
  const t0 = totalsAsOf(todayStr());
  const recent = [...entries].sort((a, b) => (a.date < b.date ? 1 : -1)).slice(0, 6);

  return (
    <div>
      <Header title={t("dashTitle")} sub={t("dashSub", { date: todayStr() })} />
      <div className="grid grid-cols-4 gap-4 mb-8">
        <Card label={t("cardAssets")} value={t0.assets} color="#1C2541" unit={t("won")} />
        <Card label={t("cardLiab")} value={t0.liab} color="#8A3324" unit={t("won")} />
        <Card label={t("cardEquity")} value={t0.equity} color="#5B4636" unit={t("won")} />
        <Card label={t("cardNI")} value={t0.ni} color={t0.ni >= 0 ? "#0E5C46" : "#8A3324"} unit={t("won")} />
      </div>

      <div className="bg-white border border-[#E3E1D8] rounded">
        <div className="px-5 py-3 border-b border-[#E3E1D8] serif font-semibold text-sm">{t("recentEntries")}</div>
        {entries.length === 0 ? (
          <EmptyState text={t("noEntriesDash")} />
        ) : (
          <table className="w-full text-sm ledger-table">
            <thead>
              <tr className="text-left text-[#8B8A80]">
                <th className="py-2 px-5 font-medium">{t("colDate")}</th>
                <th className="py-2 px-5 font-medium">{t("colMemo")}</th>
                <th className="py-2 px-5 font-medium">{t("colAccount")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colDebit")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colCredit")}</th>
              </tr>
            </thead>
            <tbody>
              {recent.map((e) =>
                e.lines.map((l, i) => (
                  <tr key={e.id + i}>
                    <td className="py-2 px-5 mono-num text-[#8B8A80]">{i === 0 ? e.date : ""}</td>
                    <td className="py-2 px-5 text-[#8B8A80]">{i === 0 ? e.memo : ""}</td>
                    <td className="py-2 px-5">{accName(accountById[l.accountId], lang)}</td>
                    <td className="py-2 px-5 text-right mono-num">{l.debit ? fmt(l.debit) : ""}</td>
                    <td className="py-2 px-5 text-right mono-num">{l.credit ? fmt(l.credit) : ""}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function Card({ label, value, color, unit }) {
  return (
    <div className="bg-white border border-[#E3E1D8] rounded p-4">
      <div className="text-[12px] text-[#8B8A80] mb-1.5">{label}</div>
      <div className="mono-num text-xl font-semibold" style={{ color }}>
        {fmt(value)} <span className="text-xs font-normal text-[#8B8A80]">{unit}</span>
      </div>
    </div>
  );
}

function Header({ title, sub }) {
  return (
    <div className="mb-6">
      <h1 className="serif text-2xl font-bold">{title}</h1>
      {sub && <div className="text-sm text-[#8B8A80] mt-1">{sub}</div>}
    </div>
  );
}

function EmptyState({ text }) {
  return <div className="p-10 text-center text-sm text-[#8B8A80]">{text}</div>;
}

/* =============================== 계정과목 =============================== */
function AccountsTab({ accounts, setAccounts, entries }) {
  const { lang, t } = useT();
  const [form, setForm] = useState({ code: "", name: "", type: "자산" });
  const [msg, setMsg] = useState("");

  const usedIds = useMemo(() => {
    const s = new Set();
    entries.forEach((e) => e.lines.forEach((l) => s.add(l.accountId)));
    return s;
  }, [entries]);

  const addAccount = () => {
    if (!form.code.trim() || !form.name.trim()) {
      setMsg(t("errCodeName"));
      return;
    }
    if (accounts.some((a) => a.code === form.code.trim())) {
      setMsg(t("errDupCode"));
      return;
    }
    setAccounts([...accounts, { id: uid(), code: form.code.trim(), name: form.name.trim(), type: form.type }]);
    setForm({ code: "", name: "", type: form.type });
    setMsg("");
  };

  const removeAccount = (id) => {
    if (usedIds.has(id)) {
      setMsg(t("errInUse"));
      return;
    }
    setAccounts(accounts.filter((a) => a.id !== id));
    setMsg("");
  };

  return (
    <div>
      <Header title={t("acctTitle")} sub={t("acctSub")} />

      <div className="bg-white border border-[#E3E1D8] rounded p-4 mb-6 flex items-end gap-3">
        <Field label={t("fieldCode")}>
          <input className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-24" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder={t("codePlaceholder")} />
        </Field>
        <Field label={t("fieldName")}>
          <input className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-40" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder={t("namePlaceholder")} />
        </Field>
        <Field label={t("fieldType")}>
          <select className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-28" value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
            {TYPE_ORDER.map((tp) => (
              <option key={tp} value={tp}>{typeLabel(tp, lang)}</option>
            ))}
          </select>
        </Field>
        <button onClick={addAccount} className="flex items-center gap-1.5 bg-[#16233F] text-white text-sm px-3 py-1.5 rounded">
          <Plus size={14} /> {t("addAccount")}
        </button>
      </div>
      {msg && <div className="flex items-center gap-1.5 text-sm text-[#8A3324] mb-4"><AlertCircle size={14} />{msg}</div>}

      <div className="grid grid-cols-2 gap-4">
        {TYPE_ORDER.map((type) => (
          <div key={type} className="bg-white border border-[#E3E1D8] rounded">
            <div className="px-4 py-2.5 border-b border-[#E3E1D8] serif font-semibold text-sm flex items-center gap-2">
              <span className="w-2 h-2 rounded-full" style={{ background: TYPE_COLOR[type] }} />
              {typeLabel(type, lang)} {t("normalBalanceNote", { normal: normalLabel(normalOf(type), lang) })}
            </div>
            <table className="w-full text-sm">
              <tbody>
                {accounts.filter((a) => a.type === type).map((a) => (
                  <tr key={a.id} className="border-b border-[#F0EFE8] last:border-0">
                    <td className="py-2 px-4 mono-num text-[#8B8A80] w-14">{a.code}</td>
                    <td className="py-2 px-4">{accName(a, lang)}</td>
                    <td className="py-2 px-4 w-8 text-right">
                      <button onClick={() => removeAccount(a.id)} className="text-[#B7B5A8] hover:text-[#8A3324]">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div className="flex flex-col gap-1">
      <label className="text-[11px] text-[#8B8A80]">{label}</label>
      {children}
    </div>
  );
}

/* =============================== 고정자산대장 =============================== */
function FixedAssetsTab({ accounts, fixedAssets, setFixedAssets, entries, setEntries }) {
  const { lang, t } = useT();
  const [form, setForm] = useState({ name: "", acquireDate: todayStr(), cost: "", usefulLifeYears: "5", salvage: "0" });
  const [msg, setMsg] = useState("");
  const [postMsg, setPostMsg] = useState(null);

  const fixedAssetAccount = accounts.find((a) => a.code === "150");
  const depExpenseAccount = accounts.find((a) => a.code === "525");

  const addAsset = () => {
    if (!form.name.trim() || !form.acquireDate || !Number(form.cost) || !Number(form.usefulLifeYears)) {
      setMsg(t("errAssetForm"));
      return;
    }
    setFixedAssets([
      ...fixedAssets,
      {
        id: uid(),
        name: form.name.trim(),
        acquireDate: form.acquireDate,
        cost: Number(form.cost),
        usefulLifeYears: Number(form.usefulLifeYears),
        salvage: Number(form.salvage) || 0,
      },
    ]);
    setForm({ name: "", acquireDate: todayStr(), cost: "", usefulLifeYears: "5", salvage: "0" });
    setMsg("");
  };

  const removeAsset = (id) => setFixedAssets(fixedAssets.filter((a) => a.id !== id));

  const rows = fixedAssets.map((a) => ({ ...a, ...depreciationInfo(a, todayStr()) }));
  const totalCost = rows.reduce((s, a) => s + a.cost, 0);
  const totalAccum = rows.reduce((s, a) => s + a.accumulated, 0);
  const totalBook = rows.reduce((s, a) => s + a.bookValue, 0);

  const postDepreciation = () => {
    if (!fixedAssetAccount || !depExpenseAccount) return;
    const now = new Date();
    const monthStart = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
    let monthlyTotal = 0;
    rows.forEach((a) => {
      // 취득 이후이고, 아직 전액 상각되지 않은 자산만 이번 달 상각액 반영
      if (a.acquireDate <= todayStr() && a.accumulated < a.cost - (Number(a.salvage) || 0) + 0.01) {
        monthlyTotal += a.monthlyDep;
      }
    });
    monthlyTotal = Math.round(monthlyTotal);
    if (!monthlyTotal) {
      setPostMsg({ ok: false, text: t("postDepZero") });
      return;
    }
    setEntries([
      ...entries,
      {
        id: uid(),
        date: todayStr(),
        memo: t("postDepMemo"),
        lines: [
          { accountId: depExpenseAccount.id, debit: monthlyTotal, credit: 0 },
          { accountId: fixedAssetAccount.id, debit: 0, credit: monthlyTotal },
        ],
      },
    ]);
    setPostMsg({ ok: true, text: t("postDepDone", { n: fmt(monthlyTotal) }) });
  };

  return (
    <div>
      <Header title={t("assetsTitle")} sub={t("assetsSub")} />

      <div className="bg-white border border-[#E3E1D8] rounded p-4 mb-4 flex items-end gap-3 flex-wrap">
        <Field label={t("fieldAssetName")}>
          <input className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-44" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder={t("assetNamePlaceholder")} />
        </Field>
        <Field label={t("fieldAcquireDate")}>
          <input type="date" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm mono-num" value={form.acquireDate} onChange={(e) => setForm({ ...form, acquireDate: e.target.value })} />
        </Field>
        <Field label={t("fieldCost")}>
          <input type="number" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-32 text-right mono-num" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} placeholder="0" />
        </Field>
        <Field label={t("fieldUsefulLife")}>
          <input type="number" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-20 text-right mono-num" value={form.usefulLifeYears} onChange={(e) => setForm({ ...form, usefulLifeYears: e.target.value })} />
        </Field>
        <Field label={t("fieldSalvage")}>
          <input type="number" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-28 text-right mono-num" value={form.salvage} onChange={(e) => setForm({ ...form, salvage: e.target.value })} placeholder="0" />
        </Field>
        <button onClick={addAsset} className="flex items-center gap-1.5 bg-[#16233F] text-white text-sm px-3 py-1.5 rounded">
          <Plus size={14} /> {t("addAsset")}
        </button>
      </div>
      {msg && <div className="flex items-center gap-1.5 text-sm text-[#8A3324] mb-4"><AlertCircle size={14} />{msg}</div>}

      <div className="text-xs text-[#8B8A80] mb-3">{t("assetsNote")}</div>

      <div className="bg-white border border-[#E3E1D8] rounded mb-4">
        {rows.length === 0 ? (
          <EmptyState text={t("noAssets")} />
        ) : (
          <table className="w-full text-sm ledger-table">
            <thead>
              <tr className="text-left text-[#8B8A80]">
                <th className="py-2 px-5 font-medium">{t("fieldAssetName")}</th>
                <th className="py-2 px-5 font-medium">{t("colAcquireDate")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colCost")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colAccumDep")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colBookValue")}</th>
                <th className="w-8"></th>
              </tr>
            </thead>
            <tbody>
              {rows.map((a) => (
                <tr key={a.id}>
                  <td className="py-2 px-5">{a.name}</td>
                  <td className="py-2 px-5 mono-num text-[#8B8A80]">{a.acquireDate}</td>
                  <td className="py-2 px-5 text-right mono-num">{fmt(a.cost)}</td>
                  <td className="py-2 px-5 text-right mono-num">{fmt(a.accumulated)}</td>
                  <td className="py-2 px-5 text-right mono-num font-medium">{fmt(a.bookValue)}</td>
                  <td className="text-center">
                    <button onClick={() => removeAsset(a.id)} className="text-[#B7B5A8] hover:text-[#8A3324]"><Trash2 size={14} /></button>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="border-t border-[#E3E1D8] font-semibold">
                <td className="py-2 px-5" colSpan={2}>{t("totalRow")}</td>
                <td className="py-2 px-5 text-right mono-num">{fmt(totalCost)}</td>
                <td className="py-2 px-5 text-right mono-num">{fmt(totalAccum)}</td>
                <td className="py-2 px-5 text-right mono-num">{fmt(totalBook)}</td>
                <td></td>
              </tr>
            </tfoot>
          </table>
        )}
      </div>

      {rows.length > 0 && (
        <div className="flex items-center gap-3">
          <button onClick={postDepreciation} className="bg-[#16233F] text-white text-sm px-4 py-2 rounded">{t("postDepEntry")}</button>
          {postMsg && (
            <span className={`flex items-center gap-1.5 text-sm ${postMsg.ok ? "text-[#0E5C46]" : "text-[#8A3324]"}`}>
              {postMsg.ok ? <CheckCircle2 size={14} /> : <AlertCircle size={14} />}
              {postMsg.text}
            </span>
          )}
        </div>
      )}
    </div>
  );
}

/* =============================== 전표입력 =============================== */
function EntryTab({ accounts, entries, setEntries }) {
  const { lang, t } = useT();
  const blankLine = () => ({ id: uid(), accountId: "", debit: "", credit: "" });
  const [date, setDate] = useState(todayStr());
  const [memo, setMemo] = useState("");
  const [lines, setLines] = useState([blankLine(), blankLine()]);
  const [msg, setMsg] = useState("");
  const [importResult, setImportResult] = useState(null);
  const fileInputRef = useRef(null);

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const { newEntries, errors } = parseVoucherWorkbook(wb, accounts);
        if (newEntries.length) setEntries((prev) => [...prev, ...newEntries]);
        setImportResult({ success: newEntries.length, errors });
      } catch (err) {
        setImportResult({ success: 0, errors: [t("fileReadErr")] });
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const totalDebit = lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const totalCredit = lines.reduce((s, l) => s + (Number(l.credit) || 0), 0);
  const balanced = totalDebit > 0 && totalDebit === totalCredit;

  const updateLine = (id, field, value) => {
    setLines(lines.map((l) => (l.id === id ? { ...l, [field]: value, ...(field === "debit" ? { credit: "" } : {}), ...(field === "credit" ? { debit: "" } : {}) } : l)));
  };
  const addLine = () => setLines([...lines, blankLine()]);
  const removeLine = (id) => setLines(lines.length > 2 ? lines.filter((l) => l.id !== id) : lines);

  const save = () => {
    if (!memo.trim()) return setMsg(t("errMemo"));
    if (lines.some((l) => !l.accountId || (!Number(l.debit) && !Number(l.credit)))) return setMsg(t("errLines"));
    if (!balanced) return setMsg(t("errBalance"));
    setEntries([
      ...entries,
      { id: uid(), date, memo: memo.trim(), lines: lines.map((l) => ({ accountId: l.accountId, debit: Number(l.debit) || 0, credit: Number(l.credit) || 0 })) },
    ]);
    setDate(todayStr());
    setMemo("");
    setLines([blankLine(), blankLine()]);
    setMsg("");
  };

  const removeEntry = (id) => setEntries(entries.filter((e) => e.id !== id));

  return (
    <div>
      <Header title={t("entryTitle")} sub={t("entrySub")} />

      <div className="bg-white border border-[#E3E1D8] rounded p-4 mb-4 flex items-center justify-between">
        <div className="text-sm text-[#8B8A80]">{t("uploadDesc")}</div>
        <div className="flex gap-2 shrink-0 ml-4">
          <button onClick={() => downloadTemplate(lang)} className="flex items-center gap-1.5 border border-[#D9D7CB] text-sm px-3 py-1.5 rounded text-[#1C2541]">
            <Download size={14} /> {t("downloadTemplate")}
          </button>
          <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-1.5 bg-[#16233F] text-white text-sm px-3 py-1.5 rounded">
            <Upload size={14} /> {t("uploadExcel")}
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
        </div>
      </div>

      {importResult && (
        <div className={`mb-4 p-3 rounded text-sm border ${importResult.errors.length ? "border-[#E7C6B8] bg-[#FBF3EE]" : "border-[#B9DED0] bg-[#F0F9F5]"}`}>
          <div className="flex items-center gap-1.5 font-medium" style={{ color: importResult.errors.length ? "#8A3324" : "#0E5C46" }}>
            {importResult.errors.length ? <AlertCircle size={14} /> : <CheckCircle2 size={14} />}
            {t("importDone", { n: importResult.success })}{importResult.errors.length ? t("importErrSuffix", { n: importResult.errors.length }) : ""}
          </div>
          {importResult.errors.length > 0 && (
            <ul className="mt-1.5 pl-5 list-disc text-[#8A3324] space-y-0.5">
              {importResult.errors.slice(0, 8).map((er, i) => (<li key={i}>{er}</li>))}
              {importResult.errors.length > 8 && <li>{t("importMoreErr", { n: importResult.errors.length - 8 })}</li>}
            </ul>
          )}
        </div>
      )}

      <div className="bg-white border border-[#E3E1D8] rounded p-5 mb-4">
        <div className="flex gap-3 mb-4">
          <Field label={t("fieldDate")}>
            <input type="date" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm mono-num" value={date} onChange={(e) => setDate(e.target.value)} />
          </Field>
          <Field label={t("fieldMemo")}>
            <input className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-80" value={memo} onChange={(e) => setMemo(e.target.value)} placeholder={t("memoPlaceholder")} />
          </Field>
        </div>

        <table className="w-full text-sm mb-2">
          <thead>
            <tr className="text-left text-[#8B8A80]">
              <th className="py-1.5 font-medium w-1/2">{t("colAccountTitle")}</th>
              <th className="py-1.5 font-medium text-right w-1/5">{t("colDebit")}</th>
              <th className="py-1.5 font-medium text-right w-1/5">{t("colCredit")}</th>
              <th className="w-8"></th>
            </tr>
          </thead>
          <tbody>
            {lines.map((l) => (
              <tr key={l.id}>
                <td className="py-1 pr-3">
                  <select className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-full" value={l.accountId} onChange={(e) => updateLine(l.id, "accountId", e.target.value)}>
                    <option value="">{t("selectAccount")}</option>
                    {TYPE_ORDER.map((type) => (
                      <optgroup key={type} label={typeLabel(type, lang)}>
                        {accounts.filter((a) => a.type === type).map((a) => (
                          <option key={a.id} value={a.id}>{a.code} · {accName(a, lang)}</option>
                        ))}
                      </optgroup>
                    ))}
                  </select>
                </td>
                <td className="py-1 pr-1">
                  <input type="number" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-full text-right mono-num" value={l.debit} onChange={(e) => updateLine(l.id, "debit", e.target.value)} placeholder="0" />
                </td>
                <td className="py-1 pl-1">
                  <input type="number" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm w-full text-right mono-num" value={l.credit} onChange={(e) => updateLine(l.id, "credit", e.target.value)} placeholder="0" />
                </td>
                <td className="text-center">
                  <button onClick={() => removeLine(l.id)} className="text-[#B7B5A8] hover:text-[#8A3324]"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="border-t border-[#E3E1D8] font-semibold">
              <td className="py-2">{t("sumRow")}</td>
              <td className="py-2 text-right mono-num">{fmt(totalDebit)}</td>
              <td className="py-2 text-right mono-num">{fmt(totalCredit)}</td>
              <td></td>
            </tr>
          </tfoot>
        </table>

        <div className="flex items-center justify-between mt-3">
          <button onClick={addLine} className="flex items-center gap-1 text-sm text-[#16233F] font-medium">
            <Plus size={14} /> {t("addLine")}
          </button>
          <div className="flex items-center gap-4">
            {totalDebit > 0 && (
              <span className={`flex items-center gap-1 text-sm ${balanced ? "text-[#0E5C46]" : "text-[#8A3324]"}`}>
                {balanced ? <CheckCircle2 size={14} /> : <AlertCircle size={14} />}
                {balanced ? t("balancedYes") : t("balancedNo")}
              </span>
            )}
            <button onClick={save} className="bg-[#16233F] text-white text-sm px-4 py-2 rounded">{t("saveEntry")}</button>
          </div>
        </div>
        {msg && <div className="flex items-center gap-1.5 text-sm text-[#8A3324] mt-3"><AlertCircle size={14} />{msg}</div>}
      </div>

      <div className="bg-white border border-[#E3E1D8] rounded">
        <div className="px-5 py-3 border-b border-[#E3E1D8] serif font-semibold text-sm">{t("enteredEntries", { n: entries.length })}</div>
        {entries.length === 0 ? (
          <EmptyState text={t("noEntries")} />
        ) : (
          [...entries].sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (
            <div key={e.id} className="px-5 py-3 border-b border-[#F0EFE8] last:border-0 flex items-start justify-between">
              <div>
                <div className="text-sm font-medium mono-num">{e.date} · {e.memo}</div>
                <div className="text-xs text-[#8B8A80] mt-1 flex flex-wrap gap-x-4 gap-y-0.5">
                  {e.lines.map((l, i) => (
                    <span key={i}>
                      {accName(accounts.find((a) => a.id === l.accountId), lang)} {l.debit ? `${t("debitShort")} ${fmt(l.debit)}` : `${t("creditShort")} ${fmt(l.credit)}`}
                    </span>
                  ))}
                </div>
              </div>
              <button onClick={() => removeEntry(e.id)} className="text-[#B7B5A8] hover:text-[#8A3324] mt-0.5"><Trash2 size={14} /></button>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

/* =============================== 원장 =============================== */
function LedgerTab({ accounts, entries, balanceOf }) {
  const { lang, t } = useT();
  const [accountId, setAccountId] = useState(accounts[0]?.id || "");
  const acc = accounts.find((a) => a.id === accountId);
  const normal = acc ? normalOf(acc.type) : "debit";

  const rows = useMemo(() => {
    const list = [];
    [...entries].sort((a, b) => (a.date < b.date ? -1 : 1)).forEach((e) => {
      e.lines.forEach((l) => {
        if (l.accountId === accountId) {
          list.push({ date: e.date, memo: e.memo, debit: l.debit, credit: l.credit });
        }
      });
    });
    let bal = 0;
    return list.map((r) => {
      const change = normal === "debit" ? r.debit - r.credit : r.credit - r.debit;
      bal += change;
      return { ...r, balance: bal };
    });
  }, [entries, accountId, normal]);

  return (
    <div>
      <Header title={t("ledgerTitle")} sub={t("ledgerSub")} />
      <div className="mb-4">
        <Field label={t("selectAccountField")}>
          <select className="border border-[#D9D7CB] rounded px-3 py-1.5 text-sm w-64" value={accountId} onChange={(e) => setAccountId(e.target.value)}>
            {TYPE_ORDER.map((type) => (
              <optgroup key={type} label={typeLabel(type, lang)}>
                {accounts.filter((a) => a.type === type).map((a) => (
                  <option key={a.id} value={a.id}>{a.code} · {accName(a, lang)}</option>
                ))}
              </optgroup>
            ))}
          </select>
        </Field>
      </div>

      <div className="bg-white border border-[#E3E1D8] rounded">
        <div className="px-5 py-3 border-b border-[#E3E1D8] serif font-semibold text-sm flex items-center justify-between">
          <span>{t("ledgerHeader", { name: accName(acc, lang), type: typeLabel(acc?.type, lang), normal: normalLabel(normal, lang) })}</span>
          <span className="mono-num">{t("currentBalance", { n: fmt(balanceOf(accountId)) })}</span>
        </div>
        {rows.length === 0 ? (
          <EmptyState text={t("noLedgerRows")} />
        ) : (
          <table className="w-full text-sm ledger-table">
            <thead>
              <tr className="text-left text-[#8B8A80]">
                <th className="py-2 px-5 font-medium">{t("colDate")}</th>
                <th className="py-2 px-5 font-medium">{t("colMemo")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num ledger-rule">{t("colDebit")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num ledger-rule">{t("colCredit")}</th>
                <th className="py-2 px-5 font-medium text-right mono-num">{t("colBalance")}</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i}>
                  <td className="py-2 px-5 mono-num text-[#8B8A80]">{r.date}</td>
                  <td className="py-2 px-5">{r.memo}</td>
                  <td className="py-2 px-5 text-right mono-num ledger-rule">{r.debit ? fmt(r.debit) : ""}</td>
                  <td className="py-2 px-5 text-right mono-num ledger-rule">{r.credit ? fmt(r.credit) : ""}</td>
                  <td className="py-2 px-5 text-right mono-num font-medium">{fmt(r.balance)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

/* =============================== 재무제표 =============================== */
function StatementsTab({ accounts, balanceOf, netIncomeUpTo, totalsAsOf }) {
  const { t } = useT();
  const [mode, setMode] = useState("is");
  const [start, setStart] = useState(yearStartStr());
  const [end, setEnd] = useState(todayStr());
  const [asOf, setAsOf] = useState(todayStr());

  return (
    <div>
      <Header title={t("stmtTitle")} sub={t("stmtSub")} />
      <div className="flex gap-2 mb-5">
        <TabBtn active={mode === "is"} onClick={() => setMode("is")}>{t("tabIS")}</TabBtn>
        <TabBtn active={mode === "bs"} onClick={() => setMode("bs")}>{t("tabBS")}</TabBtn>
      </div>

      {mode === "is" && (
        <div>
          <div className="flex items-end gap-3 mb-4">
            <Field label={t("fieldStart")}><input type="date" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm mono-num" value={start} onChange={(e) => setStart(e.target.value)} /></Field>
            <Field label={t("fieldEnd")}><input type="date" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm mono-num" value={end} onChange={(e) => setEnd(e.target.value)} /></Field>
          </div>
          <IncomeStatement accounts={accounts} balanceOf={balanceOf} start={start} end={end} />
        </div>
      )}

      {mode === "bs" && (
        <div>
          <div className="flex items-end gap-3 mb-4">
            <Field label={t("fieldAsOf")}><input type="date" className="border border-[#D9D7CB] rounded px-2 py-1.5 text-sm mono-num" value={asOf} onChange={(e) => setAsOf(e.target.value)} /></Field>
          </div>
          <BalanceSheet accounts={accounts} balanceOf={balanceOf} netIncomeUpTo={netIncomeUpTo} asOf={asOf} />
        </div>
      )}
    </div>
  );
}

function TabBtn({ active, onClick, children }) {
  return (
    <button onClick={onClick} className="px-4 py-1.5 rounded text-sm font-medium" style={{ background: active ? "#16233F" : "#fff", color: active ? "#fff" : "#1C2541", border: "1px solid #D9D7CB" }}>
      {children}
    </button>
  );
}

function IncomeStatement({ accounts, balanceOf, start, end }) {
  const { lang, t } = useT();
  const revenues = accounts.filter((a) => a.type === "수익").map((a) => ({ ...a, amount: balanceOf(a.id, { start, end }) }));
  const expenses = accounts.filter((a) => a.type === "비용").map((a) => ({ ...a, amount: balanceOf(a.id, { start, end }) }));
  const totalRev = revenues.reduce((s, a) => s + a.amount, 0);
  const totalExp = expenses.reduce((s, a) => s + a.amount, 0);
  const net = totalRev - totalExp;

  return (
    <div className="bg-white border border-[#E3E1D8] rounded max-w-xl">
      <div className="px-6 py-4 border-b border-[#E3E1D8] text-center">
        <div className="serif font-bold text-base">{t("isTitle")}</div>
        <div className="text-xs text-[#8B8A80] mt-0.5">{t("isRange", { start, end })}</div>
      </div>
      <StmtSection title={t("secRevenue")} rows={revenues} lang={lang} noneLabel={t("none")} />
      <StmtSubtotal label={t("totalRevenue")} value={totalRev} />
      <StmtSection title={t("secExpense")} rows={expenses} lang={lang} noneLabel={t("none")} />
      <StmtSubtotal label={t("totalExpense")} value={totalExp} />
      <div className="px-6 py-4 flex justify-between items-center font-bold text-base border-t-2 border-[#16233F]">
        <span className="serif">{t("netIncome")}</span>
        <span className="mono-num" style={{ color: net >= 0 ? "#0E5C46" : "#8A3324" }}>{fmt(net)} {t("won")}</span>
      </div>
    </div>
  );
}

function BalanceSheet({ accounts, balanceOf, netIncomeUpTo, asOf }) {
  const { lang, t } = useT();
  const assets = accounts.filter((a) => a.type === "자산").map((a) => ({ ...a, amount: balanceOf(a.id, { end: asOf }) }));
  const liabilities = accounts.filter((a) => a.type === "부채").map((a) => ({ ...a, amount: balanceOf(a.id, { end: asOf }) }));
  const equity = accounts.filter((a) => a.type === "자본").map((a) => ({ ...a, amount: balanceOf(a.id, { end: asOf }) }));
  const ni = netIncomeUpTo(asOf);

  const totalAssets = assets.reduce((s, a) => s + a.amount, 0);
  const totalLiab = liabilities.reduce((s, a) => s + a.amount, 0);
  const totalEquity = equity.reduce((s, a) => s + a.amount, 0) + ni;
  const totalLiabEquity = totalLiab + totalEquity;
  const balanced = Math.round(totalAssets) === Math.round(totalLiabEquity);

  return (
    <div className="bg-white border border-[#E3E1D8] rounded max-w-3xl">
      <div className="px-6 py-4 border-b border-[#E3E1D8] text-center relative">
        <div className="serif font-bold text-base">{t("bsTitle")}</div>
        <div className="text-xs text-[#8B8A80] mt-0.5">{t("bsAsOf", { date: asOf })}</div>
        <span className={`absolute right-6 top-4 flex items-center gap-1 text-xs ${balanced ? "text-[#0E5C46]" : "text-[#8A3324]"}`}>
          {balanced ? <CheckCircle2 size={13} /> : <AlertCircle size={13} />}
          {balanced ? t("balancedYes") : t("balancedNo")}
        </span>
      </div>
      <div className="grid grid-cols-2 divide-x divide-[#E3E1D8]">
        <div>
          <StmtSection title={t("secAssets")} rows={assets} lang={lang} noneLabel={t("none")} />
          <StmtSubtotal label={t("totalAssets")} value={totalAssets} />
        </div>
        <div>
          <StmtSection title={t("secLiab")} rows={liabilities} lang={lang} noneLabel={t("none")} />
          <StmtSubtotal label={t("totalLiab")} value={totalLiab} />
          <StmtSection title={t("secEquity")} rows={[...equity, { id: "ni", name: t("niCumulative"), nameEn: t("niCumulative"), amount: ni }]} lang={lang} noneLabel={t("none")} />
          <StmtSubtotal label={t("totalEquity")} value={totalEquity} />
          <div className="px-6 py-3 flex justify-between items-center font-bold text-sm border-t-2 border-[#16233F]">
            <span>{t("totalLiabEquity")}</span>
            <span className="mono-num">{fmt(totalLiabEquity)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function StmtSection({ title, rows, lang, noneLabel }) {
  return (
    <div className="px-6 pt-4">
      <div className="text-xs font-semibold text-[#8B8A80] mb-1.5">{title}</div>
      {rows.length === 0 ? (
        <div className="text-xs text-[#B7B5A8] pb-2">{noneLabel}</div>
      ) : (
        rows.map((r) => (
          <div key={r.id} className="flex justify-between text-sm py-1">
            <span>{accName(r, lang)}</span>
            <span className="mono-num">{fmt(r.amount)}</span>
          </div>
        ))
      )}
    </div>
  );
}

function StmtSubtotal({ label, value }) {
  return (
    <div className="px-6 py-2.5 flex justify-between text-sm font-semibold border-t border-b border-[#F0EFE8] mt-1">
      <span>{label}</span>
      <span className="mono-num">{fmt(value)}</span>
    </div>
  );
}
