import * as XLSX from "xlsx";

export type DataSet = {
  fileName: string;
  kind: "sheet" | "pdf" | "text";
  rows: Record<string, unknown>[];
  columns: string[];
  rawText?: string;
  pdfLines?: string[];
  sheets?: string[];
  pages?: number;
  reportAsOf?: string;
};

export type AnalysisResult = {
  title: string;
  confidence: "عالية" | "متوسطة" | "محدودة";
  conclusion?: { label: string; value: string; detail: string; tone: "ok" | "warn" | "bad" | "info" };
  dueScenarios?: { selectedDays: number; items: { days: number; label: string; value: string; selected: boolean }[] };
  dueSchedule?: { asOf: string; includedThrough: string; cutoffInvoiceDate?: string; cutoffBalance?: string; laterDeductions?: string; nextDueDate?: string; nextDueAmount?: string; nextDueRef?: string };
  aging?: { title: string; asOf: string; basis: string; total: string; buckets: { label: string; value: string; count: number; percent: string; tone: "ok" | "info" | "warn" | "bad" }[] };
  supplierFlow?: {
    title: string;
    note: string;
    source: string;
    equation: string;
    items: { label: string; value: string; hint: string; tone: "ok" | "info" | "warn" | "bad" }[];
  };
  summary: { label: string; value: string; tone?: string }[];
  findings: { title: string; detail: string; tone: "ok" | "warn" | "bad" | "info" }[];
  table?: { headers: string[]; rows: string[][] };
  dueTable?: { headers: string[]; rows: string[][] };
  checks: string[];
  sources: string[];
};

const normalize = (value: unknown) => String(value ?? "").normalize("NFKC").trim().toLowerCase().replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").replace(/[ـً-ٰٟ]/g, "").replace(/\s+/g, " ");
function accountingSearchText(value:unknown){
  const base=normalize(value),arabic=[...base].filter(char=>/[ء-ي]/.test(char)).join("");
  return `${base} ${arabic} ${[...arabic].reverse().join("")}`;
}
const westernDigits = (value:string) => value.replace(/[٠-٩]/g,d=>String("٠١٢٣٤٥٦٧٨٩".indexOf(d))).replace(/[۰-۹]/g,d=>String("۰۱۲۳۴۵۶۷۸۹".indexOf(d))).replace(/٫/g,".").replace(/٬/g,",");
function extractReportAsOf(values:unknown[]){
  const text=westernDigits(values.map(value=>String(value??"")).join(" \n "));
  const labeled=[...text.matchAll(/(?:الى|إلى|حتى|to)\s*(?:تاريخ|date)?\s*[:：-]?\s*(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2})/gi)].at(-1)?.[1];
  return labeled||undefined;
}
const number = (value: unknown) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  let s = westernDigits(String(value ?? "").trim())
    .replace(/ر\.?\s*س\.?|sar|usd|eur|gbp|[$€£]/gi, "")
    .replace(/\s/g, "")
    .replace(/−/g, "-");
  if (/^\(.*\)$/.test(s)) s = `-${s.slice(1, -1)}`;
  if (s.includes(".")) s = s.replace(/,/g, "");
  else if (/^-?\d{1,3}(?:,\d{3})+$/.test(s)) s = s.replace(/,/g, "");
  else if (/^-?\d+,\d{1,2}$/.test(s)) s = s.replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s.replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};
const money = (n: number) => new Intl.NumberFormat("ar-SA", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n) + " ر.س";
const qty = (n: number) => new Intl.NumberFormat("ar-SA", { maximumFractionDigits: 2 }).format(n);

type AgingLot={date:Date;amount:number;ref?:string};
const agingLabels=["0–30 يومًا","31–45 يومًا","46–60 يومًا","61–90 يومًا","91–120 يومًا","أكثر من 120 يومًا"];
const agingTones=["ok","info","info","warn","warn","bad"] as const;
function agingIndex(invoiceAge:number){return invoiceAge<=30?0:invoiceAge<=45?1:invoiceAge<=60?2:invoiceAge<=90?3:invoiceAge<=120?4:5}
function consumeOldest(lots:AgingLot[],amount:number){let remaining=amount;lots.sort((a,b)=>a.date.getTime()-b.date.getTime());for(const lot of lots){if(remaining<=.001)break;const paid=Math.min(lot.amount,remaining);lot.amount-=paid;remaining-=paid}return lots.filter(x=>x.amount>.001)}
function effectiveDueDate(lot:AgingLot,creditDays:number){const date=new Date(lot.date);date.setDate(date.getDate()+creditDays);return date}
function isLotDue(lot:AgingLot,asOf:Date,creditDays:number){return asOf.getTime()>=effectiveDueDate(lot,creditDays).getTime()}
function reconcileLots(lots:AgingLot[],target:number,fallbackDate:Date){const current=lots.reduce((s,x)=>s+x.amount,0),difference=target-current;if(difference>.02)lots.push({date:fallbackDate,amount:difference,ref:"رصيد مرحّل"});else if(difference<-.02)lots=consumeOldest(lots,-difference);return lots}
function agingRanges(creditDays:number){const days=Math.max(1,Math.min(365,Math.round(creditDays))),bounds=[60,90,120].filter(bound=>bound>days),ranges:{label:string;min:number;max:number;tone:"ok"|"info"|"warn"|"bad"}[]=[{label:`0–${days} يومًا · غير مستحق`,min:0,max:days,tone:"ok"}];let start=days+1;for(const bound of bounds){ranges.push({label:`${start}–${bound} يومًا · مستحق`,min:start,max:bound,tone:bound<=60?"info":"warn"});start=bound+1}ranges.push({label:`أكثر من ${Math.max(days,120)} يومًا · مستحق`,min:start,max:Infinity,tone:"bad"});return ranges}
function buildAging(lots:AgingLot[],asOf:Date,basis:string,title="أعمار المستحقات",creditDays=30){const ranges=agingRanges(creditDays),raw=ranges.map(()=>({amount:0,count:0}));for(const lot of lots.filter(x=>x.amount>.001)){const age=Math.max(0,Math.floor((asOf.getTime()-lot.date.getTime())/86400000)),index=ranges.findIndex(range=>age>=range.min&&age<=range.max);raw[index].amount+=lot.amount;raw[index].count++}const total=raw.reduce((sum,item)=>sum+item.amount,0);return{title,asOf:pdfDateFormatter.format(asOf),basis:`${basis} تم تجميع كل الفواتير من 0 إلى ${creditDays} يومًا في خانة واحدة غير مستحقة، ويبدأ الاستحقاق من اليوم ${creditDays+1}.`,total:money(total),buckets:raw.map((bucket,index)=>({label:ranges[index].label,value:money(bucket.amount),count:bucket.count,percent:total?`${(bucket.amount/total*100).toFixed(1)}%`:"0%",tone:ranges[index].tone}))}}

function parseCsv(text: string) {
  const lines: string[][] = []; let row: string[] = []; let cell = ""; let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], next = text[i + 1];
    if (c === '"' && quoted && next === '"') { cell += '"'; i++; }
    else if (c === '"') quoted = !quoted;
    else if ((c === "," || c === ";" || c === "\t") && !quoted) { row.push(cell.trim()); cell = ""; }
    else if ((c === "\n" || c === "\r") && !quoted) { if (c === "\r" && next === "\n") i++; row.push(cell.trim()); if (row.some(Boolean)) lines.push(row); row = []; cell = ""; }
    else cell += c;
  }
  row.push(cell.trim()); if (row.some(Boolean)) lines.push(row);
  const headerIndex = lines.findIndex(r => r.filter(Boolean).length >= 2);
  if (headerIndex < 0) return { rows: [], columns: [] };
  const columns = lines[headerIndex].map((h, i) => h || `عمود ${i + 1}`);
  return { columns, rows: lines.slice(headerIndex + 1).map(r => Object.fromEntries(columns.map((h, i) => [h, r[i] ?? ""]))) };
}

export async function readAccountingFile(file: File): Promise<DataSet> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext === "xlsx" || ext === "xls") {
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const first = wb.SheetNames[0];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[first], { header: 1, defval: "", raw: false });
    const headerWords=/تاريخ|بيان|مرجع|مستند|فاتور|مدين|دائن|رصيد|debit|credit|balance|date|reference|invoice/i;
    const candidates=matrix.slice(0,40).map((row,index)=>{
      const cells=row.map(value=>String(value??"").trim()),filled=cells.filter(Boolean).length;
      const keywordHits=cells.filter(cell=>headerWords.test(normalize(cell))).length;
      return{index,cells,filled,score:keywordHits*10+Math.min(filled,8)};
    }).filter(candidate=>candidate.filled>=2).sort((a,b)=>b.score-a.score||a.index-b.index);
    const headerIndex=candidates[0]?.index??0,rawHeaders=(matrix[headerIndex]||[]).map((value,index)=>String(value||`عمود ${index+1}`).trim());
    const seen=new Map<string,number>(),columns=rawHeaders.map((header,index)=>{const base=header||`عمود ${index+1}`,count=seen.get(base)||0;seen.set(base,count+1);return count?`${base} (${count+1})`:base});
    const rows=matrix.slice(headerIndex+1).filter(row=>row.some(value=>String(value??"").trim())).map(row=>Object.fromEntries(columns.map((column,index)=>[column,row[index]??""]))),reportAsOf=extractReportAsOf(matrix.slice(0,headerIndex).flat());
    return { fileName: file.name, kind: "sheet", rows, columns, sheets: wb.SheetNames, reportAsOf };
  }
  if (ext === "pdf") {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const pages: string[] = [];
    const allLines: string[] = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const content = await (await pdf.getPage(p)).getTextContent();
      const parts = content.items.filter((item) => "str" in item && item.str.trim()).map((item) => {
        const textItem = item as { str: string; transform?: number[] };
        return { str: textItem.str.trim(), x: textItem.transform?.[4] ?? 0, y: textItem.transform?.[5] ?? 0 };
      }).sort((a,b) => b.y-a.y);
      const grouped: { y:number; parts:{str:string;x:number}[] }[] = [];
      for (const part of parts) {
        const line = grouped.find(group => Math.abs(group.y-part.y) <= 2);
        if (line) line.parts.push({str:part.str,x:part.x});
        else grouped.push({y:part.y,parts:[{str:part.str,x:part.x}]});
      }
      const pageLines = grouped.sort((a,b)=>b.y-a.y).map(line=>{
        const joined=line.parts.map(part=>part.str).join("");
        const arabic=(joined.match(/[\u0600-\u06ff]/g)||[]).length;
        const latin=(joined.match(/[A-Za-z]/g)||[]).length;
        const rtl=arabic>latin;
        return line.parts.sort((a,b)=>rtl?b.x-a.x:a.x-b.x).map(part=>part.str).join(" ").replace(/\s+/g," ").trim();
      }).filter(Boolean);
      allLines.push(...pageLines);
      pages.push(pageLines.join("\n"));
    }
    return { fileName: file.name, kind: "pdf", rows: [], columns: [], rawText: pages.join("\n\n"), pdfLines: allLines, pages: pdf.numPages, reportAsOf:extractReportAsOf(allLines.slice(0,80)) };
  }
  const parsed = parseCsv(await file.text());
  return { fileName: file.name, kind: "text", ...parsed };
}

const aliases: Record<string, string[]> = {
  date: ["التاريخ", "تاريخ", "date", "transaction date", "تاريخ الحركه"],
  invoiceDate: ["تاريخ اصدار الفاتوره", "تاريخ الفاتوره", "تاريخ فاتوره", "تاريخ المستند", "تاريخ الوثيقه", "invoice date", "invoice issue date", "document date", "issue date"],
  due: ["تاريخ الاستحقاق", "الاستحقاق", "due date", "due"],
  debit: ["مدين", "debit", "debit amount", "مدين محلي"],
  credit: ["دائن", "credit", "credit amount", "دائن محلي"],
  balance: ["الرصيد", "رصيد", "balance", "running balance", "الرصيد المتبقي"],
  amount: ["المبلغ", "مبلغ", "amount", "القيمه", "قيمه", "اجمالي"],
  ref: ["المرجع", "رقم المرجع", "reference", "ref", "رقم القيد", "رقم المستند", "رقم الفاتوره"],
  name: ["الاسم", "اسم الحساب", "العميل", "المورد", "name", "account name"],
  item: ["الصنف", "اسم الصنف", "كود الصنف", "item", "item name", "sku", "product"],
  opening: ["رصيد اول المده", "اول المده", "opening", "opening balance", "opening qty"],
  purchases: ["المشتريات", "مشتريات", "الوارد", "purchases", "receipts", "received"],
  sales: ["المبيعات", "مبيعات", "الصادر", "sales", "issues", "issued"],
  returnsIn: ["مرتجع مبيعات", "مرتجعات مبيعات", "sales returns", "return in"],
  returnsOut: ["مرتجع مشتريات", "مرتجعات مشتريات", "purchase returns", "return out"],
  actual: ["الرصيد الفعلي", "الجرد الفعلي", "actual", "actual qty", "counted"],
  system: ["رصيد النظام", "الرصيد الحالي", "system", "system qty", "book qty", "الكمية"],
  stockQty: ["الكميه", "كمية المخزون", "الرصيد الكمي", "كمية الجرد", "qty", "quantity", "on hand", "stock qty", "book quantity"],
  account: ["الحساب", "اسم الحساب", "account", "account name", "البيان"],
  category: ["التصنيف", "نوع الحساب", "الفئه", "category", "account type"],
  revenue: ["الايرادات", "ايرادات", "المبيعات", "صافي المبيعات", "revenue", "sales", "income"],
  cost: ["تكلفه المبيعات", "تكلفه البضاعه", "التكلفه", "cost of sales", "cogs", "cost"],
  expense: ["المصروفات", "المصاريف", "مصروف", "expenses", "expense"],
  invoice: ["رقم الفاتوره", "الفاتوره", "invoice number", "invoice no", "invoice", "document no"],
  net: ["الصافي قبل الضريبه", "المبلغ قبل الضريبه", "الصافي", "net amount", "taxable amount", "net"],
  vat: ["ضريبه القيمه المضافه", "الضريبه", "قيمه الضريبه", "vat", "vat amount", "tax amount"],
  total: ["الاجمالي شامل الضريبه", "الاجمالي", "total including vat", "gross total", "total"],
  type: ["النوع", "نوع الحركه", "type", "transaction type", "flow type"],
  assetCost: ["تكلفه الاصل", "تكلفه الاقتناء", "asset cost", "cost"],
  accumDep: ["مجمع الاهلاك", "الاهلاك المتراكم", "accumulated depreciation", "accum dep"],
  gross: ["الاجمالي", "الراتب الاجمالي", "gross salary", "gross"],
  deductions: ["الاستقطاعات", "الخصومات", "deductions", "deduction"],
  netPay: ["صافي الراتب", "الصافي المستحق", "net salary", "net pay"],
};

function column(ds: DataSet, key: string) {
  const target = aliases[key].map(normalize);
  return ds.columns.find(c => target.includes(normalize(c))) || ds.columns.find(c => target.some(a => normalize(c).includes(a) || a.includes(normalize(c))));
}
const value = (row: Record<string, unknown>, col?: string) => col ? row[col] : "";
function parseDate(v: unknown): Date | null {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v ?? "").trim(); if (!s) return null;
  const parts = s.split(/[\/\-.]/).map(Number);
  if (parts.length === 3) { const [a,b,c] = parts; const d = a > 1900 ? new Date(a,b-1,c) : new Date(c,b-1,a); if (!isNaN(d.getTime())) return d; }
  const d = new Date(s); return isNaN(d.getTime()) ? null : d;
}

function standardChecks(ds: DataSet[]) {
  return [
    `تمت قراءة ${ds.reduce((s,d)=>s+d.rows.length,0).toLocaleString("ar-SA")} صف من ${ds.length} ملف.`,
    "تم تجاهل الصفوف الفارغة والقيم غير الرقمية عند الجمع.",
    "يجب مطابقة الرصيد الافتتاحي والختامي مع التقرير الأصلي قبل الاعتماد.",
    "لا يتم إنشاء أو ترحيل أي قيد محاسبي تلقائيًا.",
  ];
}

function generic(ds: DataSet[]): AnalysisResult {
  const rows = ds.reduce((n,d)=>n+d.rows.length,0), pdfs=ds.filter(d=>d.kind==="pdf");
  const findings = ds.map(d => ({ title:d.fileName, detail:d.kind==="pdf"?`تم استخراج النص من ${d.pages} صفحة. يلزم تحديد نوع التقرير للحصول على مطابقة دقيقة.`:`تم التعرف على ${d.rows.length} حركة و${d.columns.length} عمود.`, tone:(d.rows.length||d.rawText)?"ok":"warn" as const }));
  return { title:"فحص الملفات المرفوعة", confidence:pdfs.length?"محدودة":"متوسطة", summary:[{label:"عدد الملفات",value:String(ds.length)},{label:"إجمالي الحركات",value:rows.toLocaleString("ar-SA")},{label:"ملفات PDF",value:String(pdfs.length)},{label:"حالة القراءة",value:"اكتملت"}], findings, checks:standardChecks(ds), sources:ds.map(d=>d.fileName) };
}

const isBankPdfAction = (action:string) => action.includes("بنك") || action.includes("راجحي") || action.includes("شبكه") || /(^|\s)ساب(?=\s|$)/.test(action);

function pdfTitle(action:string){
  const a=normalize(action);
  if(a.includes("مراجعه شامله"))return "مراجعة شاملة للحسابات من PDF";
  if(a.includes("نقطه اختلاف"))return "مقارنة الملفين وتحديد أول اختلاف من PDF";
  if(isBankPdfAction(a))return "مطابقة البنك والشبكات من PDF";
  if(a.includes("مورد"))return "تحليل حساب المورد من ملف PDF";
  if(a.includes("عميل")||a.includes("ديون")||a.includes("استحقاق"))return "تحليل حساب العميل والاستحقاقات من PDF";
  if(a.includes("مكرر"))return "كشف القيود والفواتير المكررة من PDF";
  if(a.includes("مبيعات")&&a.includes("ضريب"))return "مراجعة المبيعات والضريبة من PDF";
  if(a.includes("مشتريات"))return "مراجعة المشتريات من PDF";
  if(a.includes("حركه الاصناف"))return "مقارنة حركة الأصناف من PDF";
  if(a.includes("مخزون")||a.includes("صنف")||a.includes("جرد"))return "تحليل المخزون من PDF";
  if(a.includes("ربح")||a.includes("خسار")||a.includes("شهري"))return "تحليل الربحية الشهرية من PDF";
  if(a.includes("قوائم"))return "تحليل القوائم المالية من PDF";
  if(a.includes("ميزاني"))return "تحليل الميزانية والمركز المالي من PDF";
  if(a.includes("ميزان")||a.includes("قيد"))return "مراجعة ميزان المراجعة من PDF";
  if(a.includes("تدفق"))return "تحليل التدفقات النقدية من PDF";
  if(a.includes("ضريب"))return "مراجعة ضريبة القيمة المضافة من PDF";
  if(a.includes("زك"))return "مراجعة بيانات الزكاة من PDF";
  if(a.includes("اصول")||a.includes("اهلاك"))return "مراجعة الأصول والإهلاك من PDF";
  if(a.includes("رواتب")||a.includes("عهد"))return "مراجعة الرواتب والعهد من PDF";
  if(a.includes("نشاط تجاري"))return "تحليل النشاط التجاري من PDF";
  return `تقرير ${action} من PDF`;
}

function pdfReviewFocus(action:string){
  const a=normalize(action);
  if(a.includes("مورد"))return "التركيز على الحركات والدفعات وآخر رصيد مستحق ظاهر في كشف المورد.";
  if(a.includes("عميل")||a.includes("ديون")||a.includes("استحقاق"))return "التركيز على فواتير العميل والتحصيلات والرصيد، وتقسيم الأعمار 0–30 و31–45 و46–60 و61–90 و91–120 وأكثر من 120 يومًا من تاريخ إصدار الفاتورة، مع توزيع التحصيل على أقدم فاتورة أولًا (FIFO).";
  if(isBankPdfAction(a))return "التركيز على مبالغ الحركات والرصيد، مع ضرورة رفع ملف النظام أيضًا لإتمام المطابقة.";
  if(a.includes("مكرر"))return "التركيز على أرقام المستندات والتواريخ والمبالغ المتكررة؛ يلزم تأكيد الأعمدة قبل اعتبار الحركة مكررة.";
  if(a.includes("مخزون")||a.includes("صنف")||a.includes("جرد"))return "التركيز على كود الصنف والرصيد المخزني والجرد الفعلي والتكلفة؛ يلزم ملف منظم لتحديد العجز والزيادة بدقة.";
  if(a.includes("ربح")||a.includes("خسار")||a.includes("شهري"))return "التركيز على الإيرادات والتكلفة والمصروفات لكل شهر؛ لا تُعلن نتيجة ربح أو خسارة دون وضوح هذه البنود.";
  if(a.includes("ضريب"))return "التركيز على الصافي والضريبة والإجمالي وأرقام الفواتير، مع بقاء الاعتماد النهائي للمختص الضريبي.";
  if(a.includes("زك"))return "التركيز على عناصر القوائم والوعاء الظاهرة، دون احتساب التزام زكوي نهائي من بيانات غير مكتملة.";
  if(a.includes("ميزان")||a.includes("قوائم")||a.includes("تدفق"))return "التركيز على الأرصدة والإجماليات والتصنيفات المحاسبية الظاهرة مع بيان ما يحتاج تأكيدًا.";
  return "يعرض التقرير السطور والتواريخ والمبالغ التي أمكن قراءتها من الملف وفق الاختيار المحدد.";
}
const pdfDatePattern=/\b(?:\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2})\b/g;
const pdfAmountPattern=/(?:\(?-?\d[\d,]*\.\d{1,2}\)?|\(?-?\d{1,3}(?:,\d{3})+\)?)/g;
const pdfDateFormatter=new Intl.DateTimeFormat("ar-SA-u-ca-gregory",{year:"numeric",month:"2-digit",day:"2-digit"});

type PdfCandidate={line:string;dates:{text:string;date:Date}[];amounts:number[]};
type PdfMovement={date:Date;description:string;raw:string;increase:number;decrease:number;balance:number;entryAmount:number;isOpening:boolean;fragmented:boolean};
type PdfReview={fileName:string;lines:number;movements:PdfMovement[];opening:number;closing:number;derivedClosing:number;closingSource:"printed"|"running";increases:number;decreases:number;quality:number;fragmented:number};

type SupplierKind="opening"|"purchase"|"sale"|"payment"|"return"|"discount"|"increase"|"decrease";
function supplierKind(text:string,change:number,isOpening=false):SupplierKind{
  if(isOpening)return "opening";
  const label=accountingSearchText(text);
  if(/اشعار\s*خصم|خصم\s*(?:مورد|مكتسب|اشعار)|discount note/.test(label))return "discount";
  if(/مرتجع|اشعار\s*دائن|purchase return|sales return|credit note/.test(label))return "return";
  if(/فاتوره\s*(?:مبيعات|بيع)|(?:مبيعات|بيع)\s*فاتوره|مبيعات|بيع|sales invoice|\bsale\b/.test(label))return "sale";
  if(/سداد|سند\s*(?:صرف|دفع|قبض)|(?:صرف|دفع|قبض)\s*سند|دفع|تحويل|حواله|شيك|صرف\s*نقدي|payment|paid|bank transfer/.test(label))return "payment";
  if(/فاتوره\s*(?:مشتريات|شراء|مورد)|(?:مشتريات|شراء|مورد)\s*فاتوره|purchase|supplier invoice/.test(label))return "purchase";
  if(change>.001)return "purchase";
  if(change<-.001)return "payment";
  return "increase";
}
const supplierKindLabel:Record<SupplierKind,string>={
  opening:"رصيد أول المدة",
  purchase:"فاتورة مشتريات",
  sale:"فاتورة مبيعات",
  payment:"سداد",
  return:"مرتجع",
  discount:"إشعار خصم",
  increase:"غير مصنف",
  decrease:"غير مصنف",
};

function explicitPdfClosing(lines:string[]){
  const matches=lines.map((line,index)=>({line,index,label:normalize(westernDigits(line)),amounts:(westernDigits(line).match(pdfAmountPattern)||[]).map(number)})).filter(row=>
    row.amounts.length>0 &&
    !/رصيد افتتاحي|اول المده|opening balance/.test(row.label) &&
    /صافي الرصيد|الرصيد (?:الختامي|النهائي|المستحق|الاجمالي|اخر المده|اخر الفتره)|رصيد (?:اخر|نهايه) المده|closing balance|balance due|amount due/.test(row.label)
  );
  const row=matches.at(-1);
  return row?row.amounts.at(-1)??null:null;
}

function strictPdfDate(value:string){
  const parts=value.split(/[\/\-.]/).map(Number);if(parts.length!==3)return null;
  const [a,b,c]=parts;const year=a>1900?a:c,month=b,day=a>1900?c:a;
  if(year<2000||year>2100||month<1||month>12||day<1||day>31)return null;
  const date=new Date(year,month-1,day);
  return date.getFullYear()===year&&date.getMonth()===month-1&&date.getDate()===day?date:null;
}

function choosePdfDate(options:{text:string;date:Date}[],previous:Date|null){
  if(!previous)return options[0].date;
  const forward=options.filter(x=>x.date.getTime()>=previous.getTime()).sort((a,b)=>a.date.getTime()-b.date.getTime());
  return (forward[0]||options[0]).date;
}

function movementMatches(previous:number,current:number,others:number[]){
  const delta=Math.abs(current-previous),tolerance=Math.max(.05,delta*.001);
  if(others.some(x=>Math.abs(Math.abs(x)-delta)<=tolerance))return true;
  for(let i=0;i<others.length;i++)for(let j=i+1;j<others.length;j++)if(Math.abs(Math.abs(others[i]-others[j])-delta)<=tolerance)return true;
  return delta<=tolerance&&others.some(x=>Math.abs(x)<=tolerance);
}

function balanceEdgeScore(rows:PdfCandidate[],edge:"first"|"last"){
  let previous:number|null=null,matched=0,evaluated=0;
  for(const row of rows){const index=edge==="first"?0:row.amounts.length-1,balance=row.amounts[index],others=row.amounts.filter((_,i)=>i!==index);if(previous!==null){evaluated++;if(movementMatches(previous,balance,others))matched++}previous=balance}
  return evaluated?matched/evaluated:0;
}

function balanceEdgeActivity(rows:PdfCandidate[],edge:"first"|"last"){
  const values=rows.map(row=>row.amounts[edge==="first"?0:row.amounts.length-1]);
  const nonZero=values.filter(value=>Math.abs(value)>.001).length;
  const changes=values.slice(1).filter((value,index)=>Math.abs(value-values[index])>.001).length;
  return (nonZero+changes*2)/Math.max(1,values.length*3);
}

function chooseBalanceEdge(rows:PdfCandidate[]):"first"|"last"{
  const firstScore=balanceEdgeScore(rows,"first"),lastScore=balanceEdgeScore(rows,"last");
  const firstActivity=balanceEdgeActivity(rows,"first"),lastActivity=balanceEdgeActivity(rows,"last");
  // بعض كشوف الـ ERP تضع عمودًا صفريًا ثابتًا على طرف الجدول. لا يجوز اعتباره
  // رصيدًا جارياً لمجرد أن اختبار التسلسل تعادل بين الطرفين.
  if(Math.abs(firstScore-lastScore)<=.08&&Math.abs(firstActivity-lastActivity)>.05)return firstActivity>lastActivity?"first":"last";
  return lastScore>firstScore?"last":"first";
}

function cleanPdfDescription(line:string,isOpening:boolean,increase:number,decrease:number){
  const stripped=line.replace(pdfDatePattern," ").replace(pdfAmountPattern," ").replace(/[\u200e\u200f\u202a-\u202e]/g," ").replace(/\b\d+\b/g," ").replace(/[|_:؛]+/g," ").replace(/\s+/g," ").trim();
  const tokens=stripped.split(" ").filter(Boolean),arabic=tokens.filter(x=>/[\u0600-\u06ff]/.test(x)),single=arabic.filter(x=>/^[\u0600-\u06ff]$/.test(x)).length;
  const fragmented=arabic.length>=4&&single/arabic.length>=.45;
  const references=(stripped.match(/\b[A-Za-z][A-Za-z0-9_\/-]{2,}\b/g)||[]).filter(x=>!/^(balance|debit|credit|date)$/i.test(x)).slice(0,2);
  const fallback=isOpening?"رصيد افتتاحي":increase>0||decrease>0?"حركة مالية غير مصنفة":"حركة دون قيمة مالية";
  const reference=references.length?` — مرجع ${references.join(" ")}`:"";
  if(fragmented||!stripped||stripped.length>120)return{text:`${fallback}${reference}`,fragmented};
  return{text:stripped.slice(0,100),fragmented};
}

function pdfDescriptionReference(description:string){return description.match(/ — مرجع .+$/)?.[0]||""}
function pdfDisplayDescription(movement:PdfMovement,customer:boolean){
  if(movement.isOpening)return "رصيد أول المدة";
  const raw=accountingSearchText(movement.raw),reference=pdfDescriptionReference(movement.description);
  if(customer){
    if(/اشعار\s*خصم|خصم\s*(?:عميل|اشعار)|discount note/.test(raw))return `إشعار خصم للعميل${reference}`;
    if(/مرتجع|اشعار\s*دائن|sales return|credit note/.test(raw))return `مرتجع مبيعات/إشعار دائن للعميل${reference}`;
    if(/سداد|سند\s*(?:قبض|صرف|دفع)|(?:قبض|صرف|دفع)\s*سند|تحصيل|قبض|حواله|تحويل|receipt|payment|paid/.test(raw))return `تحصيل من العميل${reference}`;
    if(/فاتوره\s*(?:مبيعات|بيع)|(?:مبيعات|بيع)\s*فاتوره|مبيعات|بيع|sales invoice|invoice/.test(raw))return `فاتورة مبيعات للعميل${reference}`;
    return movement.increase>0?`قيد مدين للعميل${reference}`:movement.decrease>0?`قيد دائن للعميل${reference}`:`حركة عميل غير مصنفة${reference}`;
  }
  if(!/^حركه ماليه غير مصنفه/.test(normalize(movement.description)))return movement.description;
  return movement.increase>0?`حركة مدينة${reference}`:movement.decrease>0?`حركة دائنة${reference}`:`حركة مالية غير مصنفة${reference}`;
}
function pdfCustomerKind(movement:PdfMovement){
  if(movement.isOpening)return "رصيد أول المدة";
  const raw=accountingSearchText(movement.raw);
  if(/اشعار\s*خصم|خصم\s*(?:عميل|اشعار)|discount note/.test(raw))return "إشعار خصم";
  if(/مرتجع|اشعار\s*دائن|return|credit note/.test(raw))return "مرتجع";
  if(/سداد|سند\s*(?:قبض|صرف|دفع)|(?:قبض|صرف|دفع)\s*سند|تحصيل|قبض|حواله|تحويل|receipt|payment|paid/.test(raw))return "سداد";
  if(/فاتوره\s*(?:مبيعات|بيع)|(?:مبيعات|بيع)\s*فاتوره|مبيعات|بيع|sales invoice|invoice/.test(raw))return "فاتورة مبيعات";
  if(movement.increase>.001)return "فاتورة مبيعات";
  if(movement.decrease>.001)return "سداد";
  return "غير مصنف";
}
function pdfGeneralKind(movement:PdfMovement){
  if(movement.isOpening)return "رصيد أول المدة";
  const raw=accountingSearchText(movement.raw);
  if(/اشعار\s*خصم|خصم\s*(?:مكتسب|مسموح|اشعار)|discount note/.test(raw))return "إشعار خصم";
  if(/مرتجع|اشعار\s*دائن|return|credit note/.test(raw))return "مرتجع";
  if(/سداد|سند\s*(?:قبض|صرف|دفع)|(?:قبض|صرف|دفع)\s*سند|تحصيل|قبض|دفع|حواله|تحويل|شيك|receipt|payment|paid|bank transfer/.test(raw))return "سداد";
  if(/فاتوره\s*(?:مشتريات|شراء|مورد)|(?:مشتريات|شراء|مورد)\s*فاتوره|purchase|supplier invoice/.test(raw))return "فاتورة مشتريات";
  if(/فاتوره\s*(?:مبيعات|بيع)|(?:مبيعات|بيع)\s*فاتوره|مبيعات|بيع|sales invoice|\bsale\b/.test(raw))return "فاتورة مبيعات";
  return "غير مصنف";
}

function readPdfReview(ds:DataSet):PdfReview|null{
  const lines=(ds.pdfLines?.length?ds.pdfLines:(ds.rawText||"").split(/\n+/)).map(westernDigits).map(x=>x.trim()).filter(Boolean);
  const candidates=lines.map(line=>{
    const dates=[...line.matchAll(pdfDatePattern)].map(match=>({text:match[0],date:strictPdfDate(match[0])})).filter((x):x is {text:string;date:Date}=>!!x.date);
    const amounts=(line.match(pdfAmountPattern)||[]).map(number);
    return{line,dates,amounts};
  }).filter(row=>row.dates.length>0&&row.amounts.length>=2);
  if(!candidates.length)return null;
  let previousDate:Date|null=null;
  const dated=candidates.map(row=>{const date=choosePdfDate(row.dates,previousDate);previousDate=date;return{...row,date}});
  const edge=chooseBalanceEdge(candidates);
  const quality=balanceEdgeScore(candidates,edge),movements:PdfMovement[]=[];let previousBalance:number|null=null,fragmented=0;
  for(const row of dated){
    const balanceIndex=edge==="first"?0:row.amounts.length-1,balance=row.amounts[balanceIndex],entryAmount=Math.max(0,...row.amounts.filter((_,index)=>index!==balanceIndex).map(x=>Math.abs(x))),first=previousBalance===null,explicitOpening=first&&/رصيد افتتاحي|اول المده|opening balance/.test(normalize(row.line));
    let delta=first?0:balance-previousBalance!;
    const description=cleanPdfDescription(row.line,explicitOpening,Math.max(0,delta),Math.max(0,-delta));
    const isOpening=first&&explicitOpening;
    if(first&&!isOpening){const kind=supplierKind(row.line,1),direction=kind==="sale"||kind==="payment"||kind==="return"?-1:1;delta=entryAmount?entryAmount*direction:balance}
    const increase=Math.max(0,delta),decrease=Math.max(0,-delta);if(description.fragmented)fragmented++;
    movements.push({date:row.date,description:description.text,raw:row.line,increase,decrease,balance,entryAmount,isOpening,fragmented:description.fragmented});previousBalance=balance;
  }
  const opening=movements[0].balance,derivedClosing=movements[movements.length-1].balance,printedCandidate=explicitPdfClosing(lines);
  // تجاهل صفر الترويسة/الإجمالي الكاذب عندما يثبت تسلسل الحركات وجود رصيد فعلي.
  const printedClosing=printedCandidate===0&&Math.abs(derivedClosing)>.02?null:printedCandidate;
  const closing=printedClosing===null?derivedClosing:printedClosing;
  return{fileName:ds.fileName,lines:lines.length,movements,opening,closing,derivedClosing,closingSource:printedClosing===null?"running":"printed",increases:movements.reduce((s,x)=>s+x.increase,0),decreases:movements.reduce((s,x)=>s+x.decrease,0),quality,fragmented};
}

function buildPdfSupplierFlow(review:PdfReview){
  const sign=review.derivedClosing<0?-1:1,first=review.movements[0],firstBalance=first.balance*sign,opening=first.isOpening?firstBalance:firstBalance-(first.increase-first.decrease);
  const totals={purchase:0,sale:0,payment:0,return:0,discount:0,increase:0,decrease:0};let previous=opening;
  for(const movement of review.movements){
    const current=movement.balance*sign,change=movement.isOpening?0:current-previous,kind=supplierKind(`${movement.raw} ${movement.description}`,change,movement.isOpening);
    if(change>0)totals[kind==="purchase"?"purchase":"increase"]+=change;
    else if(change<0){const target=kind==="sale"?"sale":kind==="payment"?"payment":kind==="return"?"return":kind==="discount"?"discount":"decrease";totals[target]+=-change}
    previous=current;
  }
  const calculated=opening+totals.purchase+totals.increase-totals.sale-totals.payment-totals.return-totals.discount-totals.decrease,target=Math.abs(review.closing),unread=target-calculated;
  const items=[
    {label:"رصيد أول المدة",value:money(opening),hint:"الرصيد الذي بدأ به الكشف",tone:"info" as const},
    {label:"مشتريات",value:money(totals.purchase),hint:"فواتير ومستندات مشتريات",tone:"warn" as const},
    {label:"مبيعات",value:money(totals.sale),hint:"فواتير ومستندات مبيعات",tone:"ok" as const},
    {label:"سداد",value:money(totals.payment),hint:"دفعات وتحويلات مالية",tone:"ok" as const},
    {label:"مرتجع",value:money(totals.return),hint:"مرتجعات وخصومات",tone:"ok" as const},
    {label:"إشعار خصم",value:money(totals.discount),hint:"إشعارات خصم مثبتة في الكشف",tone:"ok" as const},
    {label:"غير مصنف",value:money(totals.increase+totals.decrease),hint:"لم يوضح الملف نوع المستند",tone:"info" as const},
    ...(Math.abs(unread)>.02?[{label:"فرق سطور غير مكتملة القراءة",value:money(Math.abs(unread)),hint:unread>0?"الرصيد المطبوع أعلى من مجموع الحركات المقروءة":"الرصيد المطبوع أقل من مجموع الحركات المقروءة",tone:"bad" as const}]:[]),
    {label:"صافي الرصيد حسب الكشف",value:money(target),hint:review.closingSource==="printed"?"من الرصيد الختامي المطبوع في الملف":"من آخر رصيد جارٍ تم التحقق من تسلسله",tone:target?"warn" as const:"ok" as const},
  ];
  return{title:"ملخص حركة حساب المورد",note:"تم تصنيف المستندات إلى فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، وإشعار خصم. الحركة التي لا يمكن قراءة نوعها تظهر غير مصنفة دون تخمين.",source:review.closingSource==="printed"?"تم اعتماد الرصيد الختامي المطبوع في كشف الحساب":"تم احتساب الصافي من آخر رصيد جارٍ في الحركات المقروءة",equation:"رصيد أول المدة + فواتير المشتريات والقيود الدائنة − فواتير المبيعات − السداد − المرتجع − إشعار الخصم − القيود المدينة = صافي الرصيد",items};
}

function pdfSupplierRows(review:PdfReview,multiple:boolean){
  const sign=review.derivedClosing<0?-1:1,first=review.movements[0],firstBalance=first.balance*sign;let previous=first.isOpening?firstBalance:firstBalance-(first.increase-first.decrease);
  return review.movements.slice(-300).map(movement=>{
    const current=movement.balance*sign,change=movement.isOpening?0:current-previous,kind=supplierKind(`${movement.raw} ${movement.description}`,change,movement.isOpening);previous=current;
    const generic=/^حركه ماليه غير مصنفه/.test(normalize(movement.description)),description=generic?`${supplierKindLabel[kind]}${pdfDescriptionReference(movement.description)}`:movement.description;
    const cells=[pdfDateFormatter.format(movement.date),supplierKindLabel[kind],description,change<0?money(-change):"—",change>0?money(change):"—",money(current)];
    return multiple?[review.fileName,...cells]:cells;
  });
}

function buildPdfAging(review:PdfReview,asOf:Date,supplier:boolean,creditDays=30){
  const sign=review.derivedClosing<0?-1:1,first=review.movements[0],firstBalance=first.balance*sign,opening=first.isOpening?firstBalance:firstBalance-(first.increase-first.decrease);let lots:AgingLot[]=opening>0?[{date:first.date,amount:opening,ref:"رصيد افتتاحي"}]:[],previous=opening;
  for(const movement of review.movements){const current=movement.balance*sign,change=movement.isOpening?0:current-previous;if(change>.001)lots.push({date:movement.date,amount:change,ref:movement.description});else if(change<-.001)lots=consumeOldest(lots,-change);previous=current}
  const latest=review.movements[review.movements.length-1];lots=reconcileLots(lots,Math.abs(review.closing),latest.date);
  return buildAging(lots,asOf,"يُعتبر التاريخ الموجود في سطر الفاتورة هو تاريخ إصدارها، وتُضاف إليه مدة الائتمان قبل بدء حساب التأخر، مع توزيع السداد على أقدم فاتورة أولًا (FIFO).",supplier?"أعمار مستحقات المورد بعد مدة الائتمان":"أعمار مستحقات العميل بعد مدة الائتمان",creditDays);
}

function pdfAnalysis(action:string,ds:DataSet[],creditDays=30):AnalysisResult{
  const pages=ds.reduce((s,d)=>s+(d.pages||0),0),text=ds.map(d=>d.rawText||d.pdfLines?.join(" ")||"").join(" "),reviews=ds.map(readPdfReview).filter((x):x is PdfReview=>!!x),movements=reviews.flatMap(x=>x.movements);
  if(!movements.length){const scanned=text.replace(/\s/g,"").length<12;return{title:pdfTitle(action),confidence:"محدودة",summary:[{label:"عدد الصفحات",value:String(pages)},{label:"السطور المقروءة",value:String(ds.reduce((s,d)=>s+(d.pdfLines?.length||0),0))},{label:"الحركات المؤكدة",value:"0",tone:"bad"},{label:"حالة الملف",value:scanned?"PDF مصوّر وغير قابل للقراءة":"لم تظهر أعمدة مالية مؤكدة",tone:"bad"}],findings:[{title:"لم يتم إنشاء نتائج رقمية",detail:scanned?"الصفحات تبدو صورًا ممسوحة. استخدم PDF نصيًا أو Excel.":"وجد النظام نصًا، لكنه لم يجد في السطر نفسه تاريخًا وقيمتين ماليتين على الأقل؛ لذلك لم يخمّن الرصيد.",tone:"bad"},{title:"أفضل صيغة للدقة",detail:"صدّر الكشف من نظام الحسابات بصيغة Excel أو CSV، أو PDF نصي يحتوي أعمدة التاريخ والحركة والرصيد.",tone:"info"}],checks:[`تم فحص ${pages} صفحة.`,"لم يتم اعتماد أرقام غير مؤكدة.","لا يتم إنشاء أو ترحيل أي قيد محاسبي تلقائيًا."],sources:ds.map(d=>d.fileName)};}
  const dates=movements.map(x=>x.date.getTime()),start=new Date(Math.min(...dates)),end=new Date(Math.max(...dates)),single=reviews.length===1?reviews[0]:null,quality=reviews.reduce((s,x)=>s+x.quality,0)/reviews.length,request=normalize(action),supplier=request.includes("مورد"),customer=request.includes("عميل")||request.includes("ديون")||request.includes("استحقاق"),dueLabel=supplier?"صافي رصيد المورد حسب الكشف":customer?"إجمالي المستحق على العميل":"الرصيد الختامي";
  const tableRows=supplier?reviews.flatMap(review=>pdfSupplierRows(review,reviews.length>1)):reviews.flatMap(review=>review.movements.slice(-300).map(row=>{const cells=customer?[pdfDateFormatter.format(row.date),pdfCustomerKind(row),pdfDisplayDescription(row,true),row.increase?money(row.increase):"—",row.decrease?money(row.decrease):"—",money(row.balance)]:[pdfDateFormatter.format(row.date),pdfGeneralKind(row),pdfDisplayDescription(row,false),row.increase?money(row.increase):"—",row.decrease?money(row.decrease):"—",money(row.balance)];return reviews.length>1?[review.fileName,...cells]:cells}));
  const movementSummary=single?(customer?`رصيد أول المدة ${money(single.opening)}، فواتير المبيعات والحركات المدينة ${money(single.increases)}، التحصيلات والمرتجعات الدائنة ${money(single.decreases)}، والرصيد الختامي ${money(single.closing)}.`:`رصيد أول المدة ${money(single.opening)}، إجمالي الحركات المدينة ${money(single.increases)}، إجمالي الحركات الدائنة ${money(single.decreases)}، والرصيد الختامي ${money(single.closing)}.`):`تم فصل حركات ${reviews.length} ملفات، ولكل ملف رصيده المستقل.`;
  const fragmented=reviews.reduce((s,x)=>s+x.fragmented,0),totalLines=reviews.reduce((s,x)=>s+x.lines,0),closingConfirmed=!!single&&(single.closingSource==="printed"||single.quality>=.85),runningBalanceUsable=!!single&&single.closingSource==="running"&&single.movements.length>=5&&single.quality>=.55,closingUsable=closingConfirmed||runningBalanceUsable,closingValue=single&&closingUsable?money(Math.abs(single.closing)):"غير مؤكد من ملف PDF",closingDifference=single?Math.abs(Math.abs(single.closing)-Math.abs(single.derivedClosing)):0;
  const conclusion=single&&(supplier||customer)?{label:dueLabel,value:closingValue,detail:closingConfirmed?(single.closing===0?`لا يوجد رصيد ظاهر حتى ${pdfDateFormatter.format(single.movements[single.movements.length-1].date)}.`:single.closingSource==="printed"?`تم اعتماد الرصيد الختامي المطبوع داخل كشف الحساب حتى ${pdfDateFormatter.format(single.movements[single.movements.length-1].date)}، وليس رقمًا تقديريًا من جمع الحركات.`:`تم احتساب الصافي من آخر رصيد جارٍ حتى ${pdfDateFormatter.format(single.movements[single.movements.length-1].date)} بعد نجاح اختبار تسلسل الحركات بنسبة مرتفعة.`):runningBalanceUsable?`رصيد مبدئي من آخر رصيد جارٍ في حركة بتاريخ ${pdfDateFormatter.format(single.movements[single.movements.length-1].date)}. لم يظهر سطر رصيد ختامي مطبوع؛ راجع القيمة قبل الاعتماد النهائي.`:"لم يظهر رصيد ختامي مطبوع ولم ينجح تسلسل الأعمدة بدرجة كافية؛ لذلك لم يعتمد النظام مبلغًا قد يكون خاطئًا. ارفع Excel للدقة الكاملة.",tone:closingConfirmed?(single.closing===0?"ok" as const:"warn" as const):runningBalanceUsable?"info" as const:"bad" as const}:undefined;
  const reportDate=single?parseDate(ds[0].reportAsOf):null,asOf=reportDate?new Date(reportDate):end;asOf.setHours(0,0,0,0);const aging=single&&(supplier||customer)&&closingUsable?buildPdfAging(single,asOf,supplier,creditDays):undefined,supplierFlow=single&&supplier&&closingUsable?buildPdfSupplierFlow(single):undefined,agingMatches=!!aging&&aging.total===closingValue;
  const addFileHeader=(headers:string[])=>reviews.length>1?["الملف",...headers]:headers;
  const supplierHeaders=addFileHeader(["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"]);
  const customerHeaders=addFileHeader(["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"]);
  const generalHeaders=addFileHeader(["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"]);
  const findings:AnalysisResult["findings"]=[
    {title:"تم تصحيح قراءة التقرير",detail:`تم اعتماد ${movements.length} حركة مرتبطة بتاريخ وقيم مالية، واستبعاد تواريخ العناوين والطباعة من فترة الحساب.`,tone:"ok"},
    ...(supplier&&single?[{title:"مصدر صافي رصيد المورد",detail:closingConfirmed?(single.closingSource==="printed"?`القيمة المعتمدة ${money(Math.abs(single.closing))} مأخوذة من سطر الرصيد الختامي المطبوع في الكشف.`:`القيمة ${money(Math.abs(single.closing))} مأخوذة من آخر رصيد جارٍ بعد تحقق قوي من تسلسل الحركات.`):runningBalanceUsable?`القيمة المبدئية ${money(Math.abs(single.closing))} مأخوذة من آخر رصيد جارٍ في الكشف، وتحتاج مراجعة قبل الاعتماد النهائي.`:"لم يعتمد النظام قيمة المستحق؛ لأن الرصيد الختامي غير مطبوع وتسلسل الأعمدة غير كافٍ.",tone:closingConfirmed?"ok" as const:runningBalanceUsable?"warn" as const:"bad" as const},{title:"تصنيف المشتريات والمبيعات",detail:"يفك النظام الحروف العربية المتقطعة أو المعكوسة أولًا لتحديد فاتورة المشتريات، فاتورة المبيعات، السداد، المرتجع أو إشعار الخصم. عند تعذر اسم المستند تمامًا، يصنف ارتفاع رصيد المورد كمشتريات وانخفاضه كسداد بصورة مبدئية تحتاج مراجعة.",tone:"ok" as const}]:[{title:customer?"نتيجة استحقاق العميل":"ملخص الحركة المحاسبية",detail:single&&customer&&closingUsable?`${dueLabel} هو ${money(Math.abs(single.closing))}. ${movementSummary}`:movementSummary,tone:single&&closingUsable?"warn" as const:"info" as const}]),
    ...(aging?[{title:"مطابقة إجمالي الأعمار مع المديونية",detail:agingMatches?`مجموع فئات الأعمار الست ${aging.total} يساوي ${dueLabel} دون فرق.`:`مجموع الأعمار ${aging.total} لا يساوي ${dueLabel} ${closingValue}؛ يلزم مراجعة الملف.`,tone:agingMatches?"ok" as const:"bad" as const}]:[]),
    ...(supplier&&single&&single.closingSource==="printed"&&closingDifference>.02?[{title:"فرق بين الرصيد المطبوع والحركات المقروءة",detail:`يوجد فرق ${money(closingDifference)}؛ تم اعتماد الرصيد المطبوع في الكشف وإظهار الفرق كبنود لم تكتمل قراءتها بدل تغيير المستحق.`,tone:"warn" as const}]:[]),
    {title:"تنسيق البيان العربي",detail:fragmented?`تم استبدال النص العربي المتقطع في ${fragmented} حركة ببيان محاسبي واضح، مع الاحتفاظ بأي مرجع إنجليزي ظاهر.`:"النصوص المقروءة ظهرت بصورة سليمة.",tone:"ok"},
    {title:"نطاق المراجعة",detail:pdfReviewFocus(action),tone:"info"},
    {title:"سلامة تحديد الرصيد",detail:closingConfirmed?(single?.closingSource==="printed"?"تمت مطابقة النتيجة مع الرصيد الختامي المطبوع داخل الملف.":"تسلسل الأرصدة متوافق مع مبالغ الحركة بدرجة مرتفعة."):runningBalanceUsable?"تم إظهار آخر رصيد جارٍ كقيمة مبدئية لأن عدد الحركات وتسلسلها يسمحان بالتحليل، مع وجوب المراجعة قبل الاعتماد.":"لا توجد ثقة كافية لتحديد الرصيد النهائي من PDF؛ يلزم Excel أو CSV.",tone:closingConfirmed?"ok":"warn"},
  ];
  const accountingRule=supplier?"في كشف المورد تُصنف المستندات إلى فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، وإشعار خصم.":customer?"في كشف العميل تُصنف المستندات إلى فاتورة مبيعات، سداد، مرتجع، وإشعار خصم.":"يظهر نوع المستند: فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف إذا تعذرت قراءة البيان.";
  return{title:supplier?"تحليل مشتريات ومبيعات المورد من PDF":pdfTitle(action),confidence:closingUsable||(!supplier&&!customer&&quality>=.55)?"متوسطة":"محدودة",conclusion,aging,supplierFlow,summary:[{label:"عدد الصفحات",value:String(pages)},{label:"الحركات المحاسبية",value:String(movements.length)},{label:"الفترة الفعلية",value:`${pdfDateFormatter.format(start)} — ${pdfDateFormatter.format(end)}`},{label:dueLabel,value:single?(supplier||customer?closingValue:money(single.closing)):"راجع رصيد كل ملف",tone:single&&closingUsable?(single.closing?"warn":"ok"):"bad"}],findings,table:{headers:supplier?supplierHeaders:customer?customerHeaders:generalHeaders,rows:tableRows},checks:[`تمت قراءة ${pages} صفحة و${totalLines} سطرًا نصيًا.`,"أعمار الديون محسوبة حتى آخر تاريخ حركة في الكشف، وليس حتى تاريخ اليوم.",accountingRule,closingConfirmed?"تمت مطابقة صافي النتيجة مع مصدر الرصيد الموضح في التقرير.":runningBalanceUsable?"تم عرض آخر رصيد جارٍ كقيمة مبدئية تحتاج مراجعة قبل الاعتماد.":"لم يتم اعتماد مبلغ مستحق غير مؤكد.",...(aging?[agingMatches?"مجموع فئات الأعمار يساوي إجمالي المديونية الظاهر.":"مجموع فئات الأعمار لا يساوي إجمالي المديونية؛ يلزم مراجعة الملف."]:[]),movements.length>300?"يظهر أحدث 300 حركة في الجدول حتى تكون آخر العمليات ظاهرة، ويمكن تنزيل النتائج بصيغة CSV.":"تم عرض كل الحركات المؤكدة في الجدول.","لا يتم إنشاء أو ترحيل أي قيد محاسبي تلقائيًا."],sources:ds.map(d=>d.fileName)};
}

function ledgerPolarity(d:DataSet,debit?:string,credit?:string,balance?:string){
  if(!debit||!credit||!balance)return 1;
  let previous:number|null=null,directError=0,reverseError=0,tests=0;
  for(const row of d.rows){const rawBalance=value(row,balance);if(String(rawBalance??"").trim()==="")continue;const current=number(rawBalance),dr=number(value(row,debit)),cr=number(value(row,credit));if(previous!==null&&(Math.abs(dr)+Math.abs(cr)>.001)){const delta=current-previous,direct=dr-cr;directError+=Math.abs(delta-direct);reverseError+=Math.abs(delta+direct);tests++}previous=current}
  return tests&&reverseError+.01<directError?-1:1;
}

function prepareLedgerAging(d:DataSet,supplier:boolean,closing:number,now:Date,creditDays=30){
  const debit=column(d,"debit"),credit=column(d,"credit"),bal=column(d,"balance"),amount=column(d,"amount"),date=column(d,"date"),invoiceDate=column(d,"invoiceDate"),ageDate=invoiceDate||date,ref=column(d,"ref"),sign=closing<0?-1:1,polarity=bal?ledgerPolarity(d,debit,credit,bal):(supplier?-1:1);let lots:AgingLot[]=[],previousBalance:number|null=null,fallbackDate:Date|null=null,runningExposure=0;const ledgerRows:{date:Date;change:number;balance:number;ref:string;isInvoice:boolean;isDeduction:boolean;index:number}[]=[];
  const transactionRows=d.rows.filter(row=>!!parseDate(value(row,ageDate))||(!!bal&&String(value(row,bal)??"").trim()!==""));
  const texts=transactionRows.map(row=>accountingSearchText(Object.values(row).join(" "))),invoicePattern=supplier?/فاتوره.*مشتريات|مشتريات.*اجل|purchase/:/فاتوره.*مبيعات|مبيعات.*اجل|sales invoice/,deductionPattern=supplier?/مردود.*مشتريات|مرتجع.*مشتريات|سند.*صرف|سداد|تحويل|payment/:/مردود.*مبيعات|مرتجع.*مبيعات|سند.*قبض|تحصيل|تحويل|receipt/,hasNamedInvoices=texts.some(text=>invoicePattern.test(text)),hasNamedDeductions=texts.some(text=>deductionPattern.test(text));
  for(const [index,row] of transactionRows.entries()){const issueDate=parseDate(value(row,ageDate)),lotDate=issueDate||now;fallbackDate=issueDate||fallbackDate;const rawBalance=value(row,bal),hasBalance=!!bal&&String(rawBalance??"").trim()!=="";let change=0;if(hasBalance){const current=number(rawBalance)*sign;change=previousBalance===null?polarity*(number(value(row,debit))-number(value(row,credit))):current-previousBalance;previousBalance=current;runningExposure=current}else if(debit||credit){const dr=number(value(row,debit)),cr=number(value(row,credit));change=polarity*(dr-cr);runningExposure+=change}else if(amount){change=number(value(row,amount));runningExposure+=change}const rowRef=String(value(row,ref)||"—"),rowText=texts[index],explicitInvoice=invoicePattern.test(rowText),isInvoice=hasNamedInvoices?explicitInvoice:change>.001,isDeduction=change<-.001&&(hasNamedDeductions?deductionPattern.test(rowText):true);if(issueDate)ledgerRows.push({date:issueDate,change,balance:Math.max(0,runningExposure),ref:rowRef,isInvoice,isDeduction,index});if(change>.001)lots.push({date:lotDate,amount:change,ref:rowRef});else if(change<-.001)lots=consumeOldest(lots,-change)}
  lots=reconcileLots(lots,Math.abs(closing),fallbackDate||now).sort((a,b)=>a.date.getTime()-b.date.getTime());
  const basis=invoiceDate?`الاستحقاق محسوب من تاريخ إصدار كل فاتورة + المدة المختارة. يُؤخذ الرصيد الجاري عند آخر فاتورة حل استحقاقها، ثم تُخصم المردودات والسداد اللاحق لها فقط.`:date?`الاستحقاق محسوب من التاريخ الموجود في سطر الفاتورة + المدة المختارة. يُؤخذ الرصيد عند آخر فاتورة مستحقة ثم تُخصم حركات السداد والمردودات اللاحقة.`:"لم يظهر تاريخ فاتورة صالح؛ لا يمكن اعتماد توزيع الأعمار دون تاريخ الإصدار.";
  const aging=buildAging(lots,now,basis,supplier?"أعمار مستحقات المورد بعد مدة الائتمان":"أعمار مستحقات العميل بعد مدة الائتمان",creditDays);
  const rows=lots.map(lot=>{const invoiceAge=Math.max(0,Math.floor((now.getTime()-lot.date.getTime())/86400000)),dueDate=effectiveDueDate(lot,creditDays),due=isLotDue(lot,now,creditDays),range=agingRanges(creditDays).find(item=>invoiceAge>=item.min&&invoiceAge<=item.max);return[lot.ref||"—",pdfDateFormatter.format(lot.date),pdfDateFormatter.format(dueDate),`تاريخ الفاتورة + ${creditDays} يومًا`,String(invoiceAge),range?.label||`${invoiceAge} يومًا`,due?"مستحق":"غير مستحق",money(lot.amount)]});
  const dueCalculation=(days:number)=>{const cutoff=[...ledgerRows].filter(row=>row.isInvoice&&isLotDue({date:row.date,amount:Math.max(0,row.change),ref:row.ref},now,days)).sort((a,b)=>a.date.getTime()-b.date.getTime()||a.index-b.index).at(-1);if(!cutoff)return{value:0,deductions:0,cutoff:undefined};const deductions=ledgerRows.filter(row=>row.index>cutoff.index&&row.isDeduction).reduce((sum,row)=>sum-row.change,0);return{value:Math.max(0,cutoff.balance-deductions),deductions,cutoff}};
  const dueAt=(days:number)=>dueCalculation(days).value;
  const dueTotal=dueAt(creditDays),nonDueTotal=Math.max(0,Math.abs(closing)-dueTotal),scenarioDays=[...new Set([30,45,60,90,120,creditDays])].sort((a,b)=>a-b);
  const periodName=(days:number)=>days===30?"شهر":days===45?"شهر ونصف":days===60?"شهران":days===90?"3 أشهر":days===120?"4 أشهر":`${days} يومًا`;
  const dueScenarios={selectedDays:creditDays,items:scenarioDays.map(days=>({days,label:`المستحق بعد ${periodName(days)} (${days} يومًا)`,value:money(dueAt(days)),selected:days===creditDays}))};
  const selectedCalculation=dueCalculation(creditDays),dueLots=lots.filter(lot=>isLotDue(lot,now,creditDays)).sort((a,b)=>effectiveDueDate(a,creditDays).getTime()-effectiveDueDate(b,creditDays).getTime());
  const futureLots=lots.filter(lot=>!isLotDue(lot,now,creditDays)).sort((a,b)=>effectiveDueDate(a,creditDays).getTime()-effectiveDueDate(b,creditDays).getTime());
  const lastIncluded=dueLots.at(-1),nextDue=futureLots[0];
  const dueSchedule={asOf:pdfDateFormatter.format(now),includedThrough:selectedCalculation.cutoff?pdfDateFormatter.format(effectiveDueDate({date:selectedCalculation.cutoff.date,amount:0},creditDays)):lastIncluded?pdfDateFormatter.format(effectiveDueDate(lastIncluded,creditDays)):"لا توجد فاتورة مستحقة",...(selectedCalculation.cutoff?{cutoffInvoiceDate:pdfDateFormatter.format(selectedCalculation.cutoff.date),cutoffBalance:money(selectedCalculation.cutoff.balance),laterDeductions:money(selectedCalculation.deductions)}:{}),...(nextDue?{nextDueDate:pdfDateFormatter.format(effectiveDueDate(nextDue,creditDays)),nextDueAmount:money(nextDue.amount),nextDueRef:nextDue.ref||"—"}:{})};
  return{aging,rows,openCount:lots.length,basis,dueTotal,nonDueTotal,dueScenarios,dueSchedule};
}

function prepareSupplierSheet(d:DataSet,closing:number){
  const debit=column(d,"debit"),credit=column(d,"credit"),balance=column(d,"balance"),date=column(d,"invoiceDate")||column(d,"date"),ref=column(d,"ref"),invoice=column(d,"invoice"),type=column(d,"type"),account=column(d,"account"),sign=closing<0?-1:1,polarity=balance?ledgerPolarity(d,debit,credit,balance):-1;
  const totals={purchase:0,sale:0,payment:0,return:0,discount:0,increase:0,decrease:0};let previous:number|null=null,opening=0;
  const rows=d.rows.filter(row=>!!parseDate(value(row,date))||(!!balance&&String(value(row,balance)??"").trim()!=="")).map(row=>{
    const dr=number(value(row,debit)),cr=number(value(row,credit)),rawBalance=value(row,balance),hasBalance=String(rawBalance??"").trim()!=="",current=hasBalance?number(rawBalance)*sign:null;
    const text=[value(row,type),value(row,account),value(row,ref),value(row,invoice)].filter(Boolean).join(" — ")||"—",isOpening=/رصيد افتتاحي|اول المده|opening balance/.test(normalize(text));
    let change=polarity*(dr-cr);
    if(previous===null){if(isOpening){opening=current??0;change=0}else if(current!==null){opening=current-change}else opening=0}
    else if(current!==null)change=current-previous;
    const running=current??(previous??opening)+change,detected=supplierKind(text,change,isOpening),kind=change>.001&&detected==="sale"?"purchase":detected;
    if(change>0)totals[kind==="purchase"?"purchase":"increase"]+=change;
    else if(change<0){const target=kind==="sale"?"sale":kind==="payment"?"payment":kind==="return"?"return":kind==="discount"?"discount":"decrease";totals[target]+=-change}
    previous=running;
    const transactionDate=parseDate(value(row,date));
    return[transactionDate?pdfDateFormatter.format(transactionDate):"—",supplierKindLabel[kind],text,change<0?money(-change):"—",change>0?money(change):"—",money(running)];
  });
  const target=Math.abs(closing),calculated=opening+totals.purchase+totals.increase-totals.sale-totals.payment-totals.return-totals.discount-totals.decrease,unread=target-calculated;
  const flow:NonNullable<AnalysisResult["supplierFlow"]>={title:"ملخص حركة حساب المورد",note:"تم تصنيف المستندات إلى فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، وإشعار خصم طبقًا للبيان الموجود في الملف.",source:balance?"تم اعتماد آخر قيمة غير فارغة في عمود الرصيد":"تم احتساب الصافي من إجمالي الدائن ناقص إجمالي المدين",equation:"رصيد أول المدة + فواتير المشتريات والقيود الدائنة − فواتير المبيعات − السداد − المرتجع − إشعار الخصم − القيود المدينة = صافي الرصيد",items:[
    {label:"رصيد أول المدة",value:money(opening),hint:"الرصيد قبل أول حركة في الملف",tone:"info"},
    {label:"مشتريات",value:money(totals.purchase),hint:"فواتير ومستندات مشتريات",tone:"warn"},
    {label:"مبيعات",value:money(totals.sale),hint:"فواتير ومستندات مبيعات",tone:"ok"},
    {label:"سندات الصرف والسداد",value:money(totals.payment),hint:"تُخصم من رصيد المورد ومن أقدم الفواتير أولًا",tone:"ok"},
    {label:"مرتجع",value:money(totals.return),hint:"مرتجعات وخصومات",tone:"ok"},
    {label:"إشعار خصم",value:money(totals.discount),hint:"إشعارات خصم مثبتة في الكشف",tone:"ok"},
    {label:"غير مصنف",value:money(totals.increase+totals.decrease),hint:"لم يوضح الملف نوع المستند",tone:"info"},
    ...(Math.abs(unread)>.02?[{label:"فرق يحتاج مراجعة",value:money(Math.abs(unread)),hint:"فرق بين تفاصيل الحركات والرصيد الختامي",tone:"bad" as const}]:[]),
    {label:"صافي الرصيد حسب الكشف",value:money(target),hint:balance?"مطابق لآخر رصيد مسجل":"محسوب من المدين والدائن",tone:target?"warn":"ok"},
  ]};
  return{flow,rows,totalIncrease:totals.purchase+totals.increase,totalDecrease:totals.sale+totals.payment+totals.return+totals.discount+totals.decrease,totalPayments:totals.payment};
}

function prepareCustomerSheet(d:DataSet,closing:number){
  const debit=column(d,"debit"),credit=column(d,"credit"),balance=column(d,"balance"),date=column(d,"invoiceDate")||column(d,"date"),ref=column(d,"ref"),invoice=column(d,"invoice"),type=column(d,"type"),account=column(d,"account"),sign=closing<0?-1:1,polarity=balance?ledgerPolarity(d,debit,credit,balance):1;
  const totals={invoice:0,receipt:0,return:0,discount:0,increase:0,decrease:0};let previous:number|null=null,opening=0;
  const rows=d.rows.filter(row=>!!parseDate(value(row,date))||(!!balance&&String(value(row,balance)??"").trim()!=="")).map(row=>{const dr=number(value(row,debit)),cr=number(value(row,credit)),rawBalance=value(row,balance),hasBalance=String(rawBalance??"").trim()!=="",current=hasBalance?number(rawBalance)*sign:null,text=[value(row,type),value(row,account),value(row,ref),value(row,invoice)].filter(Boolean).join(" — ")||"—",label=accountingSearchText(text),isOpening=/رصيد افتتاحي|اول المده|opening balance/.test(label);let change=polarity*(dr-cr);if(previous===null){if(isOpening){opening=current??0;change=0}else if(current!==null)opening=current-change}else if(current!==null)change=current-previous;const running=current??(previous??opening)+change;let kind="غير مصنف";if(isOpening)kind="رصيد أول المدة";else if(/مرتجع.*مبيعات|مردود.*مبيعات|sales return|credit note/.test(label))kind="مردود مبيعات";else if(/سند.*قبض|تحصيل|قبض|receipt|payment|paid|تحويل/.test(label))kind="سند قبض/تحصيل";else if(/اشعار.*خصم|خصم.*عميل|discount note/.test(label))kind="إشعار خصم";else if(/فاتوره.*مبيعات|مبيعات.*فاتوره|sales invoice|invoice/.test(label)||change>.001)kind="فاتورة مبيعات";if(change>0){if(kind==="فاتورة مبيعات")totals.invoice+=change;else totals.increase+=change}else if(change<0){const amount=-change;if(kind==="مردود مبيعات")totals.return+=amount;else if(kind==="سند قبض/تحصيل")totals.receipt+=amount;else if(kind==="إشعار خصم")totals.discount+=amount;else totals.decrease+=amount}previous=running;const transactionDate=parseDate(value(row,date));return[transactionDate?pdfDateFormatter.format(transactionDate):"—",kind,text,change>0?money(change):"—",change<0?money(-change):"—",money(running)]});
  const target=Math.abs(closing),calculated=opening+totals.invoice+totals.increase-totals.receipt-totals.return-totals.discount-totals.decrease,unread=target-calculated;
  const flow:NonNullable<AnalysisResult["supplierFlow"]>={title:"ملخص حركة حساب العميل",note:"تم تصنيف مستندات العميل إلى فاتورة مبيعات، سند قبض/تحصيل، مردود مبيعات، إشعار خصم، أو غير مصنف طبقًا للبيان الموجود في الملف.",source:balance?"تم اعتماد آخر قيمة غير فارغة في عمود الرصيد":"تم احتساب الصافي من إجمالي المدين ناقص إجمالي الدائن",equation:"رصيد أول المدة + فواتير المبيعات والقيود المدينة − سندات القبض والتحصيل − مردود المبيعات − إشعارات الخصم − القيود الدائنة = صافي رصيد العميل",items:[{label:"رصيد أول المدة",value:money(opening),hint:"الرصيد الذي بدأ به كشف العميل",tone:"info"},{label:"فواتير المبيعات",value:money(totals.invoice),hint:"تزيد مديونية العميل",tone:"warn"},{label:"سندات القبض والتحصيل",value:money(totals.receipt),hint:"تُخصم من أقدم الفواتير أولًا",tone:"ok"},{label:"مردود المبيعات",value:money(totals.return),hint:"يُخصم من مديونية العميل",tone:"ok"},{label:"إشعارات الخصم",value:money(totals.discount),hint:"خصومات مثبتة في كشف العميل",tone:"ok"},{label:"غير مصنف",value:money(totals.increase+totals.decrease),hint:"لم يوضح الملف نوع المستند",tone:"info"},...(Math.abs(unread)>.02?[{label:"فرق يحتاج مراجعة",value:money(Math.abs(unread)),hint:"فرق بين تفاصيل الحركات والرصيد الختامي",tone:"bad" as const}]:[]),{label:"صافي رصيد العميل حسب الكشف",value:money(target),hint:balance?"مطابق لآخر رصيد مسجل":"محسوب من المدين والدائن",tone:target?"warn":"ok"}]};
  return{flow,rows,totalReceipts:totals.receipt,totalReturns:totals.return};
}

function ledger(ds: DataSet[], supplier=false,creditDays=30): AnalysisResult {
  const d=ds[0], debit=column(d,"debit"), credit=column(d,"credit"), bal=column(d,"balance"),date=column(d,"date"),invoiceDate=column(d,"invoiceDate"),ageDate=invoiceDate||date;
  const transactionRows=d.rows.filter(row=>!!parseDate(value(row,ageDate))||(!!bal&&String(value(row,bal)??"").trim()!=="")),totalDebit=transactionRows.reduce((s,r)=>s+number(value(r,debit)),0), totalCredit=transactionRows.reduce((s,r)=>s+number(value(r,credit)),0);
  const lastBalanceRow=bal?[...d.rows].reverse().find(row=>String(value(row,bal)??"").trim()!==""):undefined,closing=bal&&lastBalanceRow?number(value(lastBalanceRow,bal)):supplier?totalCredit-totalDebit:totalDebit-totalCredit,lastDate=[...d.rows].reverse().map(r=>parseDate(value(r,ageDate))).find((x):x is Date=>!!x),reportDate=parseDate(d.reportAsOf),asOf=reportDate?new Date(reportDate):lastDate?new Date(lastDate):new Date();asOf.setHours(0,0,0,0);
  const prepared=prepareLedgerAging(d,supplier,closing,asOf,creditDays),supplierSheet=supplier?prepareSupplierSheet(d,closing):undefined,customerSheet=!supplier?prepareCustomerSheet(d,closing):undefined,missing=[debit,credit,ageDate].filter(x=>!x).length,dueLabel=supplier?"الرصيد الفعلي المستحق للمورد":"الرصيد الفعلي المستحق على العميل",dueValue=prepared.dueTotal;
  const summary:AnalysisResult["summary"]=supplier&&supplierSheet?[
    {label:"إجمالي رصيد المورد",value:money(Math.abs(closing))},
    {label:`غير مستحق خلال ${creditDays} يومًا`,value:money(prepared.nonDueTotal),tone:"ok"},
    {label:"سندات الصرف والسداد المخصومة",value:money(supplierSheet.totalPayments),tone:"ok"},
    {label:dueLabel,value:money(dueValue),tone:dueValue?"warn":"ok"},
    {label:"فواتير مفتوحة",value:String(prepared.openCount),tone:prepared.openCount?"warn":"ok"},
  ]:[
    {label:"إجمالي رصيد العميل",value:money(Math.abs(closing))},
    {label:`غير مستحق خلال ${creditDays} يومًا`,value:money(prepared.nonDueTotal),tone:"ok"},
    {label:"سندات القبض والتحصيل المخصومة",value:money(customerSheet?.totalReceipts||0),tone:"ok"},
    {label:"مردود المبيعات المخصوم",value:money(customerSheet?.totalReturns||0),tone:"ok"},
    {label:dueLabel,value:money(dueValue),tone:dueValue?"warn":"ok"},
    {label:"فواتير مفتوحة",value:String(prepared.openCount),tone:prepared.openCount?"warn":"ok"},
  ];
  const findings:AnalysisResult["findings"]=[
    {title:"احتساب المستحق الفعلي",detail:`تم أخذ الرصيد الجاري ${prepared.dueSchedule.cutoffBalance||money(0)} عند آخر فاتورة حل استحقاقها، ثم خصم ${prepared.dueSchedule.laterDeductions||money(0)} من ${supplier?"مردودات المشتريات وسندات الصرف والسداد":"مردودات المبيعات وسندات القبض والتحصيلات"} اللاحقة؛ فكان المستحق ${money(dueValue)}.`,tone:dueValue?"warn":"ok"},
    {title:"مصدر صافي الرصيد",detail:bal?"تم أخذ الصافي من آخر خلية رصيد غير فارغة في الكشف.":"لم يظهر عمود رصيد؛ تم احتساب الصافي من المدين والدائن.",tone:bal?"ok":"warn"},
    {title:"قاعدة تاريخ الاستحقاق",detail:`تم احتساب الاستحقاق لكل فاتورة من تاريخ إصدارها + ${creditDays} يومًا. تدخل الفاتورة في المستحق عند وصول تاريخ استحقاقها أو تجاوزه حتى تاريخ القياس ${prepared.dueSchedule.asOf}، وتبقى الفواتير اللاحقة غير مستحقة.`,tone:"ok"},
    {title:"أساس أعمار المستحقات",detail:prepared.basis,tone:"ok"},
    {title:missing?"أعمدة تحتاج تأكيد":"تم التعرف على الأعمدة",detail:missing?"بعض أعمدة المدين/الدائن/تاريخ إصدار الفاتورة غير واضحة؛ راجع أسماء الأعمدة.":"تم احتساب النتائج من الأعمدة الأصلية دون تعديل.",tone:missing?"warn":"ok"},
  ];
  const dueTable=prepared.rows.length?{headers:["المرجع","تاريخ إصدار الفاتورة","تاريخ الاستحقاق","مصدر الاستحقاق","العمر بالأيام","فئة العمر","الحالة","الرصيد المفتوح"],rows:prepared.rows.slice(-100)}:undefined;
  const accountSheet=supplier?supplierSheet:customerSheet,table=accountSheet?{headers:["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"],rows:accountSheet.rows.slice(-300)}:undefined;
  const checks=[...standardChecks(ds),`قاعدة ثابتة: يُختار آخر سطر فاتورة وصل استحقاقه حتى تاريخ القياس بعد مدة ${creditDays} يومًا، ويُعتمد الرصيد الجاري عنده.`,`بعد فاتورة القطع يُخصم فقط ${supplier?"مردود المشتريات وسندات الصرف والسداد":"مردود المبيعات وسندات القبض والتحصيل"}، ولا تُضاف الفواتير الأحدث غير المستحقة.`,"يُعاد تطبيق القاعدة بصورة مستقلة لكل مدة: 30 و45 و60 و90 و120 يومًا.",...(supplier?["صُنفت مستندات المورد إلى فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف عند تعذر قراءة البيان.","صافي الرصيد يطابق آخر رصيد غير فارغ في الكشف عند توفر عمود الرصيد."]:["في حساب العميل تُصنف المستندات إلى فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف عند تعذر قراءة البيان."])];
  return{title:supplier?"تحليل مشتريات ومبيعات المورد":"تحليل مبيعات وتحصيلات العميل",confidence:missing?"متوسطة":"عالية",conclusion:{label:dueLabel,value:money(dueValue),detail:`آخر فاتورة مستحقة بتاريخ ${prepared.dueSchedule.cutoffInvoiceDate||"—"} كان رصيدها الجاري ${prepared.dueSchedule.cutoffBalance||money(0)}، ثم خُصمت الحركات المحددة اللاحقة بقيمة ${prepared.dueSchedule.laterDeductions||money(0)}. لا يدخل الرصيد غير المستحق، ولا تدخل الفواتير الأحدث التي لم يحل استحقاقها.`,tone:dueValue?"warn":"ok"},dueScenarios:prepared.dueScenarios,dueSchedule:prepared.dueSchedule,aging:prepared.aging,supplierFlow:accountSheet?.flow,summary,findings,dueTable,table,checks,sources:ds.map(x=>x.fileName)};
}

type ReconciliationMovement={date:Date|null;amount:number;ref:string};
function reconciliationData(d:DataSet,supplierStatement:boolean,now:Date){
  if(d.kind==="pdf"){const review=readPdfReview(d);if(!review)return{fileName:d.fileName,rows:[] as ReconciliationMovement[],closing:0,aging:undefined};let previous=review.movements[0].balance;const rows=review.movements.slice(1).map(movement=>{const amount=Math.abs(movement.balance-previous);previous=movement.balance;return{date:movement.date,amount,ref:movement.description}}).filter(x=>x.amount>.001);return{fileName:d.fileName,rows,closing:Math.abs(review.closing),aging:supplierStatement?buildPdfAging(review,now,true):undefined}}
  const date=column(d,"date"),ref=column(d,"ref"),amount=column(d,"amount"),debit=column(d,"debit"),credit=column(d,"credit"),balance=column(d,"balance"),rows=d.rows.map(row=>{const movement=amount?Math.abs(number(value(row,amount))):Math.abs(number(value(row,debit))-number(value(row,credit)));return{date:parseDate(value(row,date)),amount:movement,ref:String(value(row,ref)||"—")}}).filter(x=>x.amount>.001),closing=balance?Math.abs(number(value(d.rows[d.rows.length-1]||{},balance))):Math.abs(d.rows.reduce((sum,row)=>sum+number(value(row,debit))-number(value(row,credit)),0));return{fileName:d.fileName,rows,closing,aging:supplierStatement?prepareLedgerAging(d,false,closing,now).aging:undefined};
}
function referenceKey(value:string){return (westernDigits(value).match(/[A-Za-z0-9][A-Za-z0-9_\/-]{2,}/g)||[]).join("|").toLowerCase()}
function supplierReconciliation(ds:DataSet[]):AnalysisResult{
  const companyHint=(d:DataSet)=>/(erp|ledger|استاذ|حسابنا|الشركه|company|system)/i.test(normalize(d.fileName)),ordered=companyHint(ds[0])&&!companyHint(ds[1])?[ds[1],ds[0]]:[ds[0],ds[1]],now=new Date();now.setHours(0,0,0,0);
  const supplier=reconciliationData(ordered[0],true,now),company=reconciliationData(ordered[1],false,now),used=new Set<number>(),matched:{supplier:ReconciliationMovement;company:ReconciliationMovement}[]=[],supplierOnly:ReconciliationMovement[]=[];
  for(const movement of supplier.rows){const key=referenceKey(movement.ref);let best=-1,bestScore=Infinity;company.rows.forEach((candidate,index)=>{if(used.has(index)||Math.abs(candidate.amount-movement.amount)>.02)return;const candidateKey=referenceKey(candidate.ref),referenceMatch=!!key&&!!candidateKey&&(key.includes(candidateKey)||candidateKey.includes(key)),dateDifference=movement.date&&candidate.date?Math.abs(movement.date.getTime()-candidate.date.getTime())/86400000:0;if(!referenceMatch&&movement.date&&candidate.date&&dateDifference>3)return;const score=(referenceMatch?0:10)+dateDifference;if(score<bestScore){best=index;bestScore=score}});if(best>=0){used.add(best);matched.push({supplier:movement,company:company.rows[best]})}else supplierOnly.push(movement)}
  const companyOnly=company.rows.filter((_,index)=>!used.has(index)),difference=Math.abs(supplier.closing-company.closing),formatDate=(date:Date|null)=>date?pdfDateFormatter.format(date):"—",rows=[...matched.map(x=>["مطابق",formatDate(x.supplier.date),money(x.supplier.amount),formatDate(x.company.date),money(x.company.amount),x.supplier.ref]),...supplierOnly.map(x=>["عند المورد فقط",formatDate(x.date),money(x.amount),"—","—",x.ref]),...companyOnly.map(x=>["عند الشركة فقط","—","—",formatDate(x.date),money(x.amount),x.ref])];
  const unmatchedValue=supplierOnly.reduce((s,x)=>s+x.amount,0)+companyOnly.reduce((s,x)=>s+x.amount,0),issues=supplierOnly.length+companyOnly.length;
  return{title:"مطابقة كشف المورد مع حساب الشركة",confidence:supplier.rows.length&&company.rows.length?"متوسطة":"محدودة",conclusion:{label:"فرق الرصيد بين الكشفين",value:money(difference),detail:`رصيد كشف المورد ${money(supplier.closing)} مقابل رصيد حسابه في الشركة ${money(company.closing)}.`,tone:difference>.02?"bad":"ok"},aging:supplier.aging,summary:[{label:"حركات كشف المورد",value:String(supplier.rows.length)},{label:"حركات حساب الشركة",value:String(company.rows.length)},{label:"حركات مطابقة",value:String(matched.length),tone:"ok"},{label:"فروقات تحتاج مراجعة",value:String(issues),tone:issues?"bad":"ok"}],findings:[{title:"كشف المورد",detail:`تم اعتبار ${supplier.fileName} كشف المورد.`,tone:"info"},{title:"كشف الشركة",detail:`تم اعتبار ${company.fileName} حساب المورد في الشركة.`,tone:"info"},{title:"حركات عند المورد فقط",detail:`${supplierOnly.length} حركة بقيمة ${money(supplierOnly.reduce((s,x)=>s+x.amount,0))}.`,tone:supplierOnly.length?"bad":"ok"},{title:"حركات عند الشركة فقط",detail:`${companyOnly.length} حركة بقيمة ${money(companyOnly.reduce((s,x)=>s+x.amount,0))}.`,tone:companyOnly.length?"warn":"ok"},{title:"إجمالي قيمة الحركات المختلفة",detail:money(unmatchedValue),tone:issues?"warn":"ok"},{title:"طريقة المطابقة",detail:"مطابقة تلقائية بالمبلغ والمرجع عند توفره، مع سماح بفارق تاريخ حتى 3 أيام. يجب مراجعة المستند قبل أي قيد تسوية.",tone:"info"}],table:{headers:["الحالة","تاريخ المورد","مبلغ المورد","تاريخ الشركة","مبلغ الشركة","المرجع/البيان"],rows:rows.slice(0,300)},checks:[`تمت مقارنة ${supplier.rows.length} حركة من كشف المورد مع ${company.rows.length} حركة من حساب الشركة.`,"تمت المطابقة حركة بحركة دون تكرار استخدام الحركة نفسها.","لم يتم إنشاء قيد تسوية تلقائيًا.","ارفع كشف المورد أولًا وكشف حسابه من ERP ثانيًا إذا لم تكن أسماء الملفات واضحة."],sources:[supplier.fileName,company.fileName]};
}

function inventory(ds: DataSet[]): AnalysisResult {
  if(ds.length>=2){const actualHint=(d:DataSet)=>/فعلي|جرد|physical|count/i.test(normalize(d.fileName)),ordered=!actualHint(ds[0])&&actualHint(ds[1])?[ds[1],ds[0]]:[ds[0],ds[1]],actualFile=ordered[0],systemFile=ordered[1],actualItem=column(actualFile,"item"),systemItem=column(systemFile,"item"),actualQty=column(actualFile,"actual")||column(actualFile,"stockQty")||column(actualFile,"system")||column(actualFile,"balance"),systemQty=column(systemFile,"system")||column(systemFile,"stockQty")||column(systemFile,"balance")||column(systemFile,"actual"),cost=column(systemFile,"cost")||systemFile.columns.find(c=>/تكلف|cost|سعر متوسط|average cost/i.test(normalize(c))),nameOf=(row:Record<string,unknown>,col?:string)=>String(value(row,col)||"").trim(),keyOf=(row:Record<string,unknown>,col?:string)=>normalize(nameOf(row,col));
    const actualMap=new Map(actualFile.rows.map(row=>[keyOf(row,actualItem),{name:nameOf(row,actualItem),qty:number(value(row,actualQty))}])),systemMap=new Map(systemFile.rows.map(row=>[keyOf(row,systemItem),{name:nameOf(row,systemItem),qty:number(value(row,systemQty)),cost:number(value(row,cost))}])),keys=new Set([...actualMap.keys(),...systemMap.keys()]);keys.delete("");const rows=[...keys].map(key=>{const a=actualMap.get(key),s=systemMap.get(key),difference=(a?.qty||0)-(s?.qty||0),unitCost=s?.cost||0;return{name:a?.name||s?.name||key,system:s?.qty||0,actual:a?.qty||0,difference,unitCost,value:difference*unitCost,status:!a?"غير موجود في الجرد":!s?"غير موجود في النظام":difference<-.0001?"عجز":difference>.0001?"زيادة":"مطابق"}}).sort((a,b)=>Math.abs(b.value)-Math.abs(a.value)||Math.abs(b.difference)-Math.abs(a.difference)),diffs=rows.filter(x=>Math.abs(x.difference)>.0001),shortage=diffs.filter(x=>x.difference<0),surplus=diffs.filter(x=>x.difference>0),shortageValue=shortage.reduce((sum,x)=>sum+Math.abs(x.value),0),surplusValue=surplus.reduce((sum,x)=>sum+Math.abs(x.value),0);
    return{title:"مطابقة الجرد الفعلي مع رصيد النظام",confidence:actualItem&&systemItem&&actualQty&&systemQty?"عالية":"متوسطة",conclusion:{label:"صافي قيمة فرق الجرد",value:money(surplusValue-shortageValue),detail:`قيمة العجز ${money(shortageValue)} مقابل قيمة الزيادة ${money(surplusValue)}. التحليل لا يعدّل أرصدة المخزون.`,tone:diffs.length?"bad":"ok"},summary:[{label:"الأصناف المفحوصة",value:String(rows.length)},{label:"أصناف مطابقة",value:String(rows.length-diffs.length),tone:"ok"},{label:"أصناف ناقصة",value:String(shortage.length),tone:shortage.length?"bad":"ok"},{label:"أصناف زائدة",value:String(surplus.length),tone:surplus.length?"warn":"ok"},{label:"قيمة العجز",value:money(shortageValue),tone:shortageValue?"bad":"ok"},{label:"قيمة الزيادة",value:money(surplusValue),tone:surplusValue?"warn":"ok"}],findings:[{title:diffs.length?"تم اكتشاف فروقات جرد":"الجرد مطابق للنظام",detail:diffs.length?`يوجد ${diffs.length} صنفًا تختلف كميته الفعلية عن الرصيد المخزني.`:"كل كميات الجرد الفعلي تطابق أرصدة النظام في الملفين.",tone:diffs.length?"bad":"ok"},{title:"أصناف غير موجودة في أحد الملفين",detail:`${rows.filter(x=>/غير موجود/.test(x.status)).length} صنفًا يحتاج مراجعة الكود أو اسم الصنف قبل اعتماد الفرق.`,tone:rows.some(x=>/غير موجود/.test(x.status))?"warn":"ok"},{title:"أساس قيمة الفرق",detail:cost?"قيمة الفرق = فرق الكمية × تكلفة الوحدة المقروءة من تقرير النظام.":"لم يظهر عمود تكلفة؛ تم تحديد فرق الكمية، وتظهر قيمة الفرق صفرًا حتى توفير التكلفة.",tone:cost?"ok":"warn"},{title:"حدود التحليل",detail:"هذه مطابقة قراءة فقط؛ لا يتم تعديل كميات النظام أو إنشاء تسويات مخزنية تلقائيًا.",tone:"info"}],table:{headers:["الصنف","رصيد النظام","الجرد الفعلي","فرق الكمية","تكلفة الوحدة","قيمة الفرق","الحالة"],rows:rows.slice(0,500).map(x=>[x.name,qty(x.system),qty(x.actual),qty(x.difference),money(x.unitCost),money(x.value),x.status])},checks:[...standardChecks(ds),"تمت مطابقة الأصناف بالاسم أو الكود بعد توحيد الكتابة.","يجب مراجعة الأصناف غير الموجودة في أحد الملفين قبل اعتماد فرق الجرد."],sources:ordered.map(x=>x.fileName)};
  }
  const d=ds[0],item=column(d,"item"),opening=column(d,"opening"),purchases=column(d,"purchases"),sales=column(d,"sales"),ri=column(d,"returnsIn"),ro=column(d,"returnsOut"),actual=column(d,"actual")||column(d,"system")||column(d,"stockQty"),unitCost=column(d,"cost")||d.columns.find(c=>/تكلف|cost|سعر متوسط|average cost/i.test(normalize(c)));
  const rows=d.rows.map((r,i)=>{const expected=number(value(r,opening))+number(value(r,purchases))+number(value(r,ri))-number(value(r,sales))-number(value(r,ro)),a=number(value(r,actual)),diff=a-expected,cost=number(value(r,unitCost));return{name:String(value(r,item)||`صف ${i+1}`),expected,actual:a,diff,cost,value:diff*cost,status:diff<-.0001?"عجز":diff>.0001?"زيادة":"مطابق"}}).sort((a,b)=>Math.abs(b.value)-Math.abs(a.value)||Math.abs(b.diff)-Math.abs(a.diff)),diffs=rows.filter(x=>Math.abs(x.diff)>.0001),shortage=diffs.filter(x=>x.diff<0),surplus=diffs.filter(x=>x.diff>0),shortageValue=shortage.reduce((s,x)=>s+Math.abs(x.value),0),surplusValue=surplus.reduce((s,x)=>s+Math.abs(x.value),0);
  return {title:"تحليل حركة الأصناف ومطابقة الجرد",confidence:item&&actual?"عالية":"متوسطة",conclusion:{label:"صافي قيمة فرق الجرد",value:money(surplusValue-shortageValue),detail:`قيمة العجز ${money(shortageValue)} وقيمة الزيادة ${money(surplusValue)}.`,tone:diffs.length?"bad":"ok"},summary:[{label:"الأصناف المفحوصة",value:String(rows.length)},{label:"أصناف مطابقة",value:String(rows.length-diffs.length),tone:"ok"},{label:"أصناف ناقصة",value:String(shortage.length),tone:shortage.length?"bad":"ok"},{label:"أصناف زائدة",value:String(surplus.length),tone:surplus.length?"warn":"ok"},{label:"قيمة العجز",value:money(shortageValue),tone:shortageValue?"bad":"ok"},{label:"قيمة الزيادة",value:money(surplusValue),tone:surplusValue?"warn":"ok"}],findings:[{title:diffs.length?"تم اكتشاف فروقات":"لا توجد فروقات بالحركات المقروءة",detail:diffs.length?`يوجد ${diffs.length} صنفًا يحتاج مراجعة الحركة أو الجرد.`:"الرصيد المتوقع يساوي الرصيد الفعلي لكل الأصناف المقروءة.",tone:diffs.length?"bad":"ok"},{title:"معادلة الفحص",detail:"أول المدة + المشتريات + مرتجع المبيعات − المبيعات − مرتجع المشتريات = الرصيد المخزني المتوقع.",tone:"info"},{title:"قيمة الفروق",detail:unitCost?"تم ضرب فرق الكمية في تكلفة الوحدة لكل صنف.":"أضف عمود تكلفة الوحدة لإظهار الأثر المالي للفروق.",tone:unitCost?"ok":"warn"}],table:{headers:["الصنف","الرصيد المتوقع","الجرد الفعلي","فرق الكمية","تكلفة الوحدة","قيمة الفرق","الحالة"],rows:rows.slice(0,500).map(x=>[x.name,qty(x.expected),qty(x.actual),qty(x.diff),money(x.cost),money(x.value),x.status])},checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function trialBalance(ds: DataSet[]): AnalysisResult {
  const d=ds[0],debit=column(d,"debit"),credit=column(d,"credit"),totalDebit=d.rows.reduce((s,r)=>s+number(value(r,debit)),0),totalCredit=d.rows.reduce((s,r)=>s+number(value(r,credit)),0),diff=totalDebit-totalCredit;
  return {title:"مراجعة ميزان المراجعة",confidence:debit&&credit?"عالية":"محدودة",summary:[{label:"إجمالي المدين",value:money(totalDebit)},{label:"إجمالي الدائن",value:money(totalCredit)},{label:"فرق الميزان",value:money(diff),tone:Math.abs(diff)>.01?"bad":"ok"},{label:"الحسابات",value:String(d.rows.length)}],findings:[{title:Math.abs(diff)<=.01?"الميزان متوازن حسابيًا":"الميزان غير متوازن",detail:Math.abs(diff)<=.01?"إجمالي المدين يساوي إجمالي الدائن. يلزم استكمال الفحوص النوعية.":`يوجد فرق مقداره ${money(Math.abs(diff))}.`,tone:Math.abs(diff)<=.01?"ok":"bad"}],checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function monthlyProfit(ds: DataSet[]): AnalysisResult {
  const d=ds[0],date=column(d,"date"),rev=column(d,"revenue"),cost=column(d,"cost"),expense=column(d,"expense"),account=column(d,"account")||column(d,"name"),category=column(d,"category"),debit=column(d,"debit"),credit=column(d,"credit");
  const months=new Map<string,{date:Date,revenue:number,cost:number,expense:number,rows:number}>(); let unclassified=0;
  for(const r of d.rows){const dt=parseDate(value(r,date));if(!dt){unclassified++;continue}const key=`${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}`;const rec=months.get(key)||{date:new Date(dt.getFullYear(),dt.getMonth(),1),revenue:0,cost:0,expense:0,rows:0};rec.rows++;
    if(rev||cost||expense){rec.revenue+=number(value(r,rev));rec.cost+=number(value(r,cost));rec.expense+=number(value(r,expense));}
    else {const label=normalize(`${value(r,category)} ${value(r,account)}`),dr=number(value(r,debit)),cr=number(value(r,credit));if(/مبيعات|ايراد|revenue|sales|income/.test(label))rec.revenue+=cr-dr;else if(/تكلف|بضاع|cogs|cost of sales/.test(label))rec.cost+=dr-cr;else if(/مصروف|مصاريف|expense|رواتب|ايجار|كهرب|اتصالات|تسويق|بنكيه|اهلاك/.test(label))rec.expense+=dr-cr;else unclassified++;}
    months.set(key,rec);
  }
  const list=[...months.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(([key,x])=>({...x,key,net:x.revenue-x.cost-x.expense,margin:x.revenue?((x.revenue-x.cost-x.expense)/x.revenue)*100:0}));
  const total=list.reduce((s,x)=>({revenue:s.revenue+x.revenue,cost:s.cost+x.cost,expense:s.expense+x.expense,net:s.net+x.net}),{revenue:0,cost:0,expense:0,net:0}); const profitable=list.filter(x=>x.net>=0).length,loss=list.length-profitable;
  const rows=list.map((x,i)=>{const prev=i?list[i-1].net:null,change=prev===null||Math.abs(prev)<.01?"—":`${((x.net-prev)/Math.abs(prev)*100).toFixed(1)}%`;return [x.date.toLocaleDateString("ar-SA",{month:"long",year:"numeric"}),money(x.revenue),money(x.cost),money(x.expense),money(x.net),x.net>=0?"ربح":"خسارة",`${x.margin.toFixed(1)}%`,change]});
  return {title:"تحليل الربحية الشهرية",confidence:date&&(rev||account||category)?"عالية":"متوسطة",summary:[{label:"إجمالي الإيرادات",value:money(total.revenue)},{label:"إجمالي التكاليف والمصروفات",value:money(total.cost+total.expense)},{label:"صافي النتيجة",value:money(total.net),tone:total.net>=0?"ok":"bad"},{label:"أشهر ربح / خسارة",value:`${profitable} / ${loss}`,tone:loss?"warn":"ok"}],findings:[{title:list.length?"تم إعداد المقارنة الشهرية":"تعذر تحديد الشهور",detail:list.length?`تم تحليل ${list.length} شهر وربط كل نتيجة بالحركات المقروءة.`:"تأكد من وجود عمود تاريخ صحيح في التقرير.",tone:list.length?"ok":"bad"},{title:"حركات غير مصنفة",detail:`${unclassified} حركة لم تدخل في الربحية لعدم وضوح تاريخها أو تصنيفها.`,tone:unclassified?"warn":"ok"},{title:"قاعدة الاحتساب",detail:"صافي الشهر = الإيرادات − تكلفة المبيعات − المصروفات. الضريبة لا تُعد إيرادًا أو مصروفًا إلا إذا نص التقرير على معالجتها بخلاف ذلك.",tone:"info"}],table:rows.length?{headers:["الشهر","الإيرادات","التكلفة","المصروفات","الصافي","النتيجة","الهامش","التغير عن السابق"],rows}:undefined,checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function duplicates(ds: DataSet[]): AnalysisResult {
  const d=ds[0],ref=column(d,"ref")||column(d,"invoice"),date=column(d,"date"),amount=column(d,"amount")||column(d,"total"),debit=column(d,"debit"),credit=column(d,"credit"),seen=new Map<string,number>(),hits:{row:number;first:number;key:string;amount:number}[]=[];
  d.rows.forEach((r,i)=>{const a=amount?number(value(r,amount)):number(value(r,debit))-number(value(r,credit));const key=normalize(`${value(r,ref)}|${value(r,date)}|${a.toFixed(2)}`);if(key.replace(/[|0.]/g,"").length<2)return;const first=seen.get(key);if(first!==undefined)hits.push({row:i+2,first:first+2,key,amount:a});else seen.set(key,i)});
  return {title:"كشف القيود والفواتير المكررة",confidence:ref&&amount?"عالية":"متوسطة",summary:[{label:"الحركات المفحوصة",value:String(d.rows.length)},{label:"التكرارات المحتملة",value:String(hits.length),tone:hits.length?"bad":"ok"},{label:"قيمة التكرارات",value:money(hits.reduce((s,x)=>s+Math.abs(x.amount),0)),tone:hits.length?"warn":"ok"},{label:"حالة الفحص",value:hits.length?"تحتاج مراجعة":"لا يوجد تكرار ظاهر"}],findings:[{title:hits.length?"تم رصد تكرارات محتملة":"لا توجد مطابقات مكررة",detail:hits.length?"التطابق مبني على المرجع والتاريخ والمبلغ. راجع المستند قبل الحذف أو العكس.":"لم تتكرر مجموعة المرجع والتاريخ والمبلغ في البيانات المقروءة.",tone:hits.length?"bad":"ok"}],table:hits.length?{headers:["الصف المكرر","الصف الأصلي","المبلغ"],rows:hits.slice(0,40).map(x=>[String(x.row),String(x.first),money(x.amount)])}:undefined,checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function journalAudit(ds:DataSet[]):AnalysisResult{
  const d=ds[0],ref=column(d,"ref"),date=column(d,"date"),account=column(d,"account")||column(d,"name"),debit=column(d,"debit"),credit=column(d,"credit"),description=d.columns.find(c=>/بيان|وصف|شرح|description|memo/i.test(normalize(c)));
  if(!debit||!credit)return {...generic(ds),title:"تحليل قيود اليومية",confidence:"محدودة",findings:[{title:"أعمدة المدين والدائن غير واضحة",detail:"ارفع تقريرًا يحتوي على رقم القيد والتاريخ والحساب والمدين والدائن لإتمام التحليل.",tone:"bad"},...generic(ds).findings]};
  type Issue={risk:"عالية"|"متوسطة"|"منخفضة";kind:string;entry:string;date:string;account:string;debit:number;credit:number;detail:string;row:number};
  const issues:Issue[]=[],groups=new Map<string,{debit:number;credit:number;rows:number[];date:string}>(),seen=new Map<string,number>(),amounts:number[]=[];let totalDebit=0,totalCredit=0;
  d.rows.forEach((row,index)=>{const dr=Math.abs(number(value(row,debit))),cr=Math.abs(number(value(row,credit))),entry=String(value(row,ref)||"").trim(),rawDate=String(value(row,date)||"").trim(),accountName=String(value(row,account)||"").trim(),memo=String(value(row,description)||"").trim(),rowNo=index+2,dt=parseDate(value(row,date));totalDebit+=dr;totalCredit+=cr;if(dr||cr)amounts.push(Math.max(dr,cr));const groupKey=normalize(entry||`صف-${rowNo}`),group=groups.get(groupKey)||{debit:0,credit:0,rows:[],date:rawDate};group.debit+=dr;group.credit+=cr;group.rows.push(rowNo);groups.set(groupKey,group);
    if(!entry||!rawDate||!accountName||!memo)issues.push({risk:"متوسطة",kind:"بيانات ناقصة",entry:entry||"—",date:rawDate||"—",account:accountName||"—",debit:dr,credit:cr,detail:[!entry&&"رقم القيد",!rawDate&&"التاريخ",!accountName&&"الحساب",!memo&&"البيان"].filter(Boolean).join("، ")+" غير متاح",row:rowNo});
    if(rawDate&&!dt)issues.push({risk:"عالية",kind:"تاريخ غير صالح",entry:entry||"—",date:rawDate,account:accountName||"—",debit:dr,credit:cr,detail:"تعذر قراءة تاريخ القيد",row:rowNo});
    if(dt&&dt.getTime()>Date.now()+86400000)issues.push({risk:"عالية",kind:"تاريخ مستقبلي",entry:entry||"—",date:rawDate,account:accountName||"—",debit:dr,credit:cr,detail:"تاريخ القيد بعد تاريخ اليوم",row:rowNo});
    if(dr>.001&&cr>.001)issues.push({risk:"عالية",kind:"مدين ودائن معًا",entry:entry||"—",date:rawDate||"—",account:accountName||"—",debit:dr,credit:cr,detail:"السطر يحتوي قيمة في المدين والدائن",row:rowNo});
    const duplicateKey=normalize(`${entry}|${rawDate}|${accountName}|${dr.toFixed(2)}|${cr.toFixed(2)}`),first=seen.get(duplicateKey);if(first!==undefined&&duplicateKey.replace(/[|0.]/g,"").length>3)issues.push({risk:"عالية",kind:"سطر مكرر",entry:entry||"—",date:rawDate||"—",account:accountName||"—",debit:dr,credit:cr,detail:`مطابق للصف ${first}`,row:rowNo});else seen.set(duplicateKey,rowNo);
  });
  for(const [entry,group] of groups)if(Math.abs(group.debit-group.credit)>.02)issues.push({risk:"عالية",kind:"قيد غير متوازن",entry,date:group.date||"—",account:"عدة حسابات",debit:group.debit,credit:group.credit,detail:`فرق ${money(group.debit-group.credit)} في الصفوف ${group.rows.join("، ")}`,row:group.rows[0]});
  const sorted=[...amounts].sort((a,b)=>a-b),median=sorted[Math.floor(sorted.length/2)]||0,unusualThreshold=Math.max(median*5,sorted.length>10?sorted[Math.floor(sorted.length*.95)]||0:0);if(unusualThreshold>0)d.rows.forEach((row,index)=>{const dr=Math.abs(number(value(row,debit))),cr=Math.abs(number(value(row,credit))),amount=Math.max(dr,cr);if(amount>=unusualThreshold&&amount>median)issues.push({risk:"منخفضة",kind:"مبلغ غير معتاد",entry:String(value(row,ref)||"—"),date:String(value(row,date)||"—"),account:String(value(row,account)||"—"),debit:dr,credit:cr,detail:`المبلغ أعلى بكثير من وسيط الحركات ${money(median)}`,row:index+2})});
  const unique=new Map<string,Issue>();for(const issue of issues)unique.set(`${issue.kind}-${issue.row}`,issue);const list=[...unique.values()].sort((a,b)=>({عالية:3,متوسطة:2,منخفضة:1}[b.risk]-{عالية:3,متوسطة:2,منخفضة:1}[a.risk])),high=list.filter(x=>x.risk==="عالية").length,unbalanced=list.filter(x=>x.kind==="قيد غير متوازن").length,duplicatesCount=list.filter(x=>x.kind==="سطر مكرر").length,missing=list.filter(x=>x.kind==="بيانات ناقصة").length;
  return{title:"تحليل قيود اليومية",confidence:ref&&date&&account?"عالية":"متوسطة",conclusion:{label:list.length?"قيود تحتاج مراجعة":"نتيجة فحص قيود اليومية",value:list.length.toLocaleString("ar-SA"),detail:list.length?`تم رصد ${high} ملاحظة عالية الخطورة. التحليل للقراءة والمراجعة فقط ولا يعدّل أو يرحّل القيود.`:"لم تظهر ملاحظات حسابية أو تكرارات في البيانات المقروءة. لا يتم تعديل القيود.",tone:high?"bad":list.length?"warn":"ok"},summary:[{label:"إجمالي المدين",value:money(totalDebit)},{label:"إجمالي الدائن",value:money(totalCredit)},{label:"فرق الملف",value:money(totalDebit-totalCredit),tone:Math.abs(totalDebit-totalCredit)>.02?"bad":"ok"},{label:"قيود غير متوازنة",value:String(unbalanced),tone:unbalanced?"bad":"ok"},{label:"سطور مكررة",value:String(duplicatesCount),tone:duplicatesCount?"bad":"ok"},{label:"بيانات ناقصة",value:String(missing),tone:missing?"warn":"ok"}],findings:[{title:unbalanced?"تم رصد قيود غير متوازنة":"القيود المجمعة متوازنة",detail:unbalanced?`${unbalanced} قيدًا لا يتساوى فيه إجمالي المدين والدائن.`:"إجمالي المدين يساوي الدائن داخل كل رقم قيد مقروء.",tone:unbalanced?"bad":"ok"},{title:duplicatesCount?"توجد سطور مكررة محتملة":"لا توجد سطور مكررة بالكامل",detail:duplicatesCount?`${duplicatesCount} سطرًا تطابق في رقم القيد والتاريخ والحساب والمدين والدائن.`:"لم يظهر تكرار كامل وفق الحقول المقروءة.",tone:duplicatesCount?"bad":"ok"},{title:"المبالغ غير المعتادة",detail:`تمت مقارنة كل مبلغ بوسيط الحركات ${money(median)}؛ النتائج المعروضة إشارات للمراجعة وليست أخطاء مؤكدة.`,tone:list.some(x=>x.kind==="مبلغ غير معتاد")?"warn":"info"},{title:"حدود التحليل",detail:"لا يتم إدخال أو تعديل أو حذف أو ترحيل أي قيد. يجب الرجوع للمستند المؤيد قبل اعتماد أي ملاحظة.",tone:"info"}],table:list.length?{headers:["الخطورة","الملاحظة","رقم القيد","التاريخ","الحساب","مدين","دائن","التفصيل"],rows:list.slice(0,200).map(x=>[x.risk,x.kind,x.entry,x.date,x.account,money(x.debit),money(x.credit),x.detail])}:undefined,checks:[...standardChecks(ds),"تم تجميع السطور حسب رقم القيد لاختبار توازن كل قيد مستقلًا.","تم فحص التكرار الكامل والبيانات الناقصة والتواريخ والمبالغ غير المعتادة."],sources:ds.map(x=>x.fileName)};
}

function invoices(ds: DataSet[], purchase=false, vatOnly=false): AnalysisResult {
  const d=ds[0],invoice=column(d,"invoice")||column(d,"ref"),net=column(d,"net")||column(d,"amount"),vat=column(d,"vat"),total=column(d,"total"),date=column(d,"date"),seen=new Set<string>(),dupes:string[][]=[],errors:string[][]=[];let netSum=0,vatSum=0,totalSum=0;
  d.rows.forEach((r,i)=>{const n=number(value(r,net)),v=number(value(r,vat)),t=total?number(value(r,total)):n+v;netSum+=n;vatSum+=v;totalSum+=t;const id=normalize(value(r,invoice));if(id){if(seen.has(id))dupes.push([String(value(r,invoice)),String(value(r,date)||"—"),money(t)]);else seen.add(id)}if(total&&Math.abs(n+v-t)>.02)errors.push([String(value(r,invoice)||`صف ${i+2}`),money(n),money(v),money(t),money(n+v-t)])});
  const issues=dupes.length+errors.length, title=vatOnly?"مراجعة ضريبة القيمة المضافة":purchase?"مراجعة فواتير المشتريات":"مراجعة فواتير المبيعات والضريبة";
  return {title,confidence:net&&vat?"عالية":"متوسطة",summary:[{label:purchase?"صافي المشتريات":"صافي المبيعات",value:money(netSum)},{label:"إجمالي الضريبة",value:money(vatSum)},{label:"الإجمالي شامل الضريبة",value:money(totalSum)},{label:"ملاحظات الفحص",value:String(issues),tone:issues?"bad":"ok"}],findings:[{title:"مطابقة المعادلة",detail:errors.length?`${errors.length} فاتورة لا يساوي فيها الصافي + الضريبة الإجمالي.`:"كل الفواتير المقروءة متوازنة حسابيًا.",tone:errors.length?"bad":"ok"},{title:"أرقام الفواتير",detail:dupes.length?`يوجد ${dupes.length} رقم فاتورة مكرر.`:"لا توجد أرقام فواتير مكررة في الملف.",tone:dupes.length?"warn":"ok"},{title:"حدود المراجعة",detail:"هذه مطابقة حسابية للملف وليست إقرارًا ضريبيًا. يجب مطابقة الإجماليات مع الإقرار والمستندات الأصلية.",tone:"info"}],table:errors.length?{headers:["الفاتورة","الصافي","الضريبة","الإجمالي","الفرق"],rows:errors.slice(0,40)}:dupes.length?{headers:["الفاتورة المكررة","التاريخ","الإجمالي"],rows:dupes.slice(0,40)}:undefined,checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function statements(ds: DataSet[], balanceOnly=false): AnalysisResult {
  const d=ds[0],account=column(d,"account")||column(d,"name"),category=column(d,"category"),bal=column(d,"balance"),debit=column(d,"debit"),credit=column(d,"credit");let assets=0,liabilities=0,equity=0,revenue=0,expenses=0,unclassified=0;
  d.rows.forEach(r=>{const label=normalize(`${value(r,category)} ${value(r,account)}`),raw=bal?number(value(r,bal)):number(value(r,debit))-number(value(r,credit));if(/اصل|اصول|نقد|بنك|عميل|مخزون|asset|cash|bank|receivable|inventory/.test(label))assets+=Math.abs(raw);else if(/التزام|خصوم|مورد|دائن|liabil|payable/.test(label))liabilities+=Math.abs(raw);else if(/راس المال|حقوق الملكيه|ارباح مبقاه|equity|capital|retained/.test(label))equity+=Math.abs(raw);else if(/مبيعات|ايراد|revenue|sales|income/.test(label))revenue+=Math.abs(raw);else if(/تكلف|مصروف|مصاريف|رواتب|ايجار|اهلاك|cost|expense/.test(label))expenses+=Math.abs(raw);else unclassified++});const equation=assets-liabilities-equity,net=revenue-expenses;
  return {title:balanceOnly?"تحليل الميزانية والمركز المالي":"تحليل القوائم المالية",confidence:account&&category?"عالية":"متوسطة",summary:[{label:"إجمالي الأصول",value:money(assets)},{label:"الالتزامات وحقوق الملكية",value:money(liabilities+equity)},{label:"فرق معادلة الميزانية",value:money(equation),tone:Math.abs(equation)>.02?"bad":"ok"},{label:"صافي الربح/الخسارة",value:money(net),tone:net>=0?"ok":"bad"}],findings:[{title:Math.abs(equation)<=.02?"معادلة الميزانية متوازنة":"معادلة الميزانية غير متوازنة",detail:`الأصول − الالتزامات − حقوق الملكية = ${money(equation)}.`,tone:Math.abs(equation)<=.02?"ok":"bad"},{title:"حسابات غير مصنفة",detail:`${unclassified} حساب لم يدخل في التحليل الآلي ويحتاج تصنيفًا واضحًا.`,tone:unclassified?"warn":"ok"},{title:"نتيجة النشاط",detail:`الإيرادات ${money(revenue)} والمصروفات والتكاليف ${money(expenses)}.`,tone:net>=0?"ok":"bad"}],checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function cashFlow(ds: DataSet[]): AnalysisResult {
  const d=ds[0],type=column(d,"type")||column(d,"category"),amount=column(d,"amount"),debit=column(d,"debit"),credit=column(d,"credit");let operating=0,investing=0,financing=0,unclassified=0;
  d.rows.forEach(r=>{const label=normalize(value(r,type)),a=amount?number(value(r,amount)):number(value(r,debit))-number(value(r,credit));if(/تشغيل|operat/.test(label))operating+=a;else if(/استثمار|invest/.test(label))investing+=a;else if(/تمويل|financ/.test(label))financing+=a;else unclassified++});const net=operating+investing+financing;
  return {title:"تحليل التدفقات النقدية",confidence:type&&amount?"عالية":"متوسطة",summary:[{label:"التشغيلية",value:money(operating)},{label:"الاستثمارية",value:money(investing)},{label:"التمويلية",value:money(financing)},{label:"صافي التغير النقدي",value:money(net),tone:net>=0?"ok":"warn"}],findings:[{title:"حركات غير مصنفة",detail:`${unclassified} حركة تحتاج تحديد نوع التدفق: تشغيلي أو استثماري أو تمويلي.`,tone:unclassified?"warn":"ok"},{title:"حدود النتيجة",detail:"يجب ربط صافي التغير برصيد النقدية أول وآخر الفترة قبل اعتماد قائمة التدفقات.",tone:"info"}],checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function assets(ds: DataSet[]): AnalysisResult {
  const d=ds[0],name=column(d,"item")||column(d,"account")||column(d,"name"),cost=column(d,"assetCost")||column(d,"amount"),dep=column(d,"accumDep");let totalCost=0,totalDep=0;const rows=d.rows.map((r,i)=>{const c=number(value(r,cost)),p=number(value(r,dep));totalCost+=c;totalDep+=p;return [String(value(r,name)||`أصل ${i+1}`),money(c),money(p),money(c-p),c&&p>c?"راجع":"سليم"]});
  return {title:"مراجعة الأصول والإهلاك",confidence:cost&&dep?"عالية":"متوسطة",summary:[{label:"تكلفة الأصول",value:money(totalCost)},{label:"مجمع الإهلاك",value:money(totalDep)},{label:"صافي القيمة الدفترية",value:money(totalCost-totalDep)},{label:"عدد الأصول",value:String(d.rows.length)}],findings:[{title:"فحص مجمع الإهلاك",detail:rows.some(r=>r[4]==="راجع")?"يوجد أصل تجاوز فيه مجمع الإهلاك تكلفته.":"لم يتجاوز مجمع الإهلاك تكلفة أي أصل مقروء.",tone:rows.some(r=>r[4]==="راجع")?"bad":"ok"}],table:{headers:["الأصل","التكلفة","مجمع الإهلاك","الصافي","الحالة"],rows:rows.slice(0,50)},checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function payroll(ds: DataSet[]): AnalysisResult {
  const d=ds[0],name=column(d,"name"),gross=column(d,"gross"),ded=column(d,"deductions"),net=column(d,"netPay"),errors:string[][]=[];let tg=0,td=0,tn=0;d.rows.forEach((r,i)=>{const g=number(value(r,gross)),x=number(value(r,ded)),n=number(value(r,net));tg+=g;td+=x;tn+=n;if(Math.abs(g-x-n)>.02)errors.push([String(value(r,name)||`صف ${i+2}`),money(g),money(x),money(n),money(g-x-n)])});
  return {title:"مراجعة الرواتب والعهد",confidence:gross&&net?"عالية":"متوسطة",summary:[{label:"إجمالي الرواتب",value:money(tg)},{label:"الاستقطاعات",value:money(td)},{label:"صافي المستحق",value:money(tn)},{label:"فروق الحساب",value:String(errors.length),tone:errors.length?"bad":"ok"}],findings:[{title:"مطابقة صافي الرواتب",detail:errors.length?`${errors.length} موظفًا لا يساوي صافي راتبه الإجمالي ناقص الاستقطاعات.`:"كل الصفوف المقروءة متوازنة حسابيًا.",tone:errors.length?"bad":"ok"}],table:errors.length?{headers:["الموظف","الإجمالي","الاستقطاعات","الصافي","الفرق"],rows:errors}:undefined,checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

function zakat(ds: DataSet[]): AnalysisResult {
  const base=statements(ds,true);return {...base,title:"مراجعة بيانات الزكاة",confidence:"متوسطة",findings:[...base.findings,{title:"مسودة مراجعة فقط",detail:"النظام يفحص التصنيف والاتساق الحسابي ولا يحسب الوعاء أو الالتزام النهائي تلقائيًا؛ يلزم اعتماد مختص وفق بيانات وإقرارات الفترة.",tone:"warn"}]};
}

function bank(ds: DataSet[]): AnalysisResult {
  if(ds.length<2)return {...generic(ds),title:"مطابقة البنك",confidence:"محدودة",findings:[{title:"مطلوب ملفان",detail:"ارفع كشف البنك وتقرير حساب البنك من الـ ERP معًا.",tone:"warn"},...generic(ds).findings]};
  type BankMovement={date:Date|null;amount:number;signed:number;ref:string;description:string;row:number};
  const movements=(d:DataSet)=>{if(d.kind==="pdf"){const review=readPdfReview(d);return review?review.movements.filter(x=>!x.isOpening&&x.entryAmount>.001).map((x,index)=>({date:x.date,amount:x.entryAmount,signed:x.increase-x.decrease,ref:referenceKey(x.description)||"—",description:x.description,row:index+2})):[] as BankMovement[]}const date=column(d,"date"),ref=column(d,"ref"),amount=column(d,"amount"),debit=column(d,"debit"),credit=column(d,"credit");return d.rows.map((row,index)=>{const dr=number(value(row,debit)),cr=number(value(row,credit)),raw=amount?number(value(row,amount)):dr-cr,description=Object.values(row).filter(Boolean).join(" — ");return{date:parseDate(value(row,date)),amount:Math.abs(raw),signed:raw,ref:String(value(row,ref)||referenceKey(description)||"—"),description,row:index+2}}).filter(x=>x.amount>.001)};
  const closingBalance=(d:DataSet,rows:BankMovement[])=>{if(d.kind==="pdf"){const review=readPdfReview(d);return review?.closing??rows.reduce((sum,row)=>sum+row.signed,0)}const balance=column(d,"balance"),last=balance?[...d.rows].reverse().find(row=>String(value(row,balance)??"").trim()!==""):undefined;return last?number(value(last,balance)):rows.reduce((sum,row)=>sum+row.signed,0)};
  const openingBalance=(d:DataSet,rows:BankMovement[],closing:number)=>{const balance=column(d,"balance"),first=balance?d.rows.find(row=>String(value(row,balance)??"").trim()!==""):undefined;if(first&&rows.length)return number(value(first,balance))-rows[0].signed;return closing-rows.reduce((sum,row)=>sum+row.signed,0)};
  const [bankFile,systemFile]=ds,bankRows=movements(bankFile),systemRows=movements(systemFile),used=new Set<number>(),matched:{bank:BankMovement;system:BankMovement}[]=[],bankOnly:BankMovement[]=[];
  for(const movement of bankRows){let best=-1,bestScore=Infinity;systemRows.forEach((candidate,index)=>{if(used.has(index)||Math.abs(candidate.amount-movement.amount)>.02)return;const dateDiff=movement.date&&candidate.date?Math.abs(movement.date.getTime()-candidate.date.getTime())/86400000:0,refA=referenceKey(movement.ref),refB=referenceKey(candidate.ref),refMatch=!!refA&&!!refB&&(refA.includes(refB)||refB.includes(refA));if(movement.date&&candidate.date&&dateDiff>3&&!refMatch)return;const score=dateDiff+(refMatch?0:5);if(score<bestScore){best=index;bestScore=score}});if(best>=0){used.add(best);matched.push({bank:movement,system:systemRows[best]})}else bankOnly.push(movement)}
  const systemOnly=systemRows.filter((_,index)=>!used.has(index)),grouped:{bank:BankMovement[];system:BankMovement[]}[]=[];
  const withinDays=(a:Date|null,b:Date|null,days=3)=>!a||!b||Math.abs(a.getTime()-b.getTime())/86400000<=days;
  const combination=(target:number,candidates:BankMovement[])=>{const pool=candidates.slice(0,18);let answer:BankMovement[]|null=null;const visit=(start:number,chosen:BankMovement[],sum:number)=>{if(answer||chosen.length>5||sum>target+.02)return;if(chosen.length>=2&&Math.abs(sum-target)<=.02){answer=[...chosen];return}for(let i=start;i<pool.length;i++)visit(i+1,[...chosen,pool[i]],sum+pool[i].amount)};visit(0,[],0);return answer};
  let remainingBank=[...bankOnly],remainingSystem=[...systemOnly];for(const bankMovement of [...remainingBank]){const candidates=remainingSystem.filter(systemMovement=>withinDays(bankMovement.date,systemMovement.date)),parts=combination(bankMovement.amount,candidates);if(parts){grouped.push({bank:[bankMovement],system:parts});remainingBank=remainingBank.filter(x=>x.row!==bankMovement.row);const rows=new Set(parts.map(x=>x.row));remainingSystem=remainingSystem.filter(x=>!rows.has(x.row))}}
  for(const systemMovement of [...remainingSystem]){const candidates=remainingBank.filter(bankMovement=>withinDays(systemMovement.date,bankMovement.date)),parts=combination(systemMovement.amount,candidates);if(parts){grouped.push({bank:parts,system:[systemMovement]});remainingSystem=remainingSystem.filter(x=>x.row!==systemMovement.row);const rows=new Set(parts.map(x=>x.row));remainingBank=remainingBank.filter(x=>!rows.has(x.row))}}
  const amountDifferences:{bank:BankMovement;system:BankMovement}[]=[];for(const bankMovement of remainingBank){const bankRef=referenceKey(bankMovement.ref);if(!bankRef)continue;const candidate=remainingSystem.find(systemMovement=>{const systemRef=referenceKey(systemMovement.ref),dateDiff=bankMovement.date&&systemMovement.date?Math.abs(bankMovement.date.getTime()-systemMovement.date.getTime())/86400000:0;return !!systemRef&&(bankRef.includes(systemRef)||systemRef.includes(bankRef))&&dateDiff<=3&&Math.abs(bankMovement.amount-systemMovement.amount)>.02});if(candidate)amountDifferences.push({bank:bankMovement,system:candidate})}
  const differingBank=new Set(amountDifferences.map(x=>x.bank.row)),differingSystem=new Set(amountDifferences.map(x=>x.system.row)),pureBankOnly=remainingBank.filter(x=>!differingBank.has(x.row)),pureSystemOnly=remainingSystem.filter(x=>!differingSystem.has(x.row));
  const duplicateRows=(rows:BankMovement[])=>rows.filter((row,index)=>rows.findIndex(other=>other.row!==row.row&&Math.abs(other.amount-row.amount)<.02&&other.date?.getTime()===row.date?.getTime()&&referenceKey(other.ref)===referenceKey(row.ref))<index),bankDuplicates=duplicateRows(bankRows),systemDuplicates=duplicateRows(systemRows);
  const formatDate=(date:Date|null)=>date?pdfDateFormatter.format(date):"—",differenceRows=[...pureBankOnly.map(x=>({date:x.date,label:"بالبنك فقط",bank:x,system:null as BankMovement|null,diff:x.amount})),...pureSystemOnly.map(x=>({date:x.date,label:"بالنظام فقط",bank:null as BankMovement|null,system:x,diff:-x.amount})),...amountDifferences.map(x=>({date:x.bank.date||x.system.date,label:"مبلغ مختلف",bank:x.bank,system:x.system,diff:x.bank.amount-x.system.amount}))].sort((x,y)=>(x.date?.getTime()??Infinity)-(y.date?.getTime()??Infinity)),firstDifference=differenceRows[0];
  const bankClosing=closingBalance(bankFile,bankRows),systemClosing=closingBalance(systemFile,systemRows),balanceDifference=bankClosing-systemClosing,fees=pureBankOnly.filter(x=>/عموله|رسوم|ضريبه|vat|fee/.test(normalize(x.description))),pendingDeposits=pureSystemOnly.filter(x=>/ايداع|تحويل|شبكه|نقاط بيع|deposit|transfer|pos/.test(normalize(x.description)));
  const suggestion=(status:string,text:string)=>{const normalized=normalize(text);if(/عموله|رسوم|ضريبه|vat|fee/.test(normalized))return"قيد مصروف/عمولة بنكية مقترح للمراجعة";if(/ايداع|تحويل|شبكه|نقاط بيع|deposit|transfer|pos/.test(normalized))return"تحقق من إيداع أو تحويل معلق";if(status==="مبلغ مختلف")return"راجع المبلغ والمرجع قبل التسوية";return status==="بالبنك فقط"?"تحقق من مستند البنك وسجله بالنظام إن كان صحيحًا":"تحقق من حركة معلقة أو تاريخ القيمة"};
  const sum=(rows:BankMovement[])=>rows.reduce((total,row)=>total+row.amount,0),dates=(rows:BankMovement[])=>[...new Set(rows.map(row=>formatDate(row.date)))].join(" + "),refs=(rows:BankMovement[])=>rows.map(row=>row.ref).join(" + ");
  const tableRows=[...matched.map(x=>["مطابق",formatDate(x.bank.date),money(x.bank.amount),formatDate(x.system.date),money(x.system.amount),money(x.bank.amount-x.system.amount),x.bank.ref,"—"]),...grouped.map(x=>["مطابقة مجمعة",dates(x.bank),money(sum(x.bank)),dates(x.system),money(sum(x.system)),money(sum(x.bank)-sum(x.system)),`${refs(x.bank)} ⇄ ${refs(x.system)}`,`${x.bank.length} حركة بنك مقابل ${x.system.length} حركة نظام`]),...differenceRows.map(x=>[x.label,formatDate(x.bank?.date??null),x.bank?money(x.bank.amount):"—",formatDate(x.system?.date??null),x.system?money(x.system.amount):"—",money(x.diff),x.bank?.ref||x.system?.ref||"—",suggestion(x.label,`${x.bank?.description||""} ${x.system?.description||""}`)])];
  return {title:"مطابقة كشف البنك مع النظام",confidence:bankRows.length&&systemRows.length?"عالية":"متوسطة",conclusion:{label:"فرق الرصيد بين البنك والنظام",value:money(balanceDifference),detail:`رصيد البنك ${money(bankClosing)} مقابل رصيد النظام ${money(systemClosing)}. التقرير يقترح أسباب التسوية ولا ينشئ أو يرحّل أي قيد تلقائيًا.`,tone:Math.abs(balanceDifference)>.02?"bad":"ok"},summary:[{label:"الرصيد الافتتاحي بالبنك",value:money(openingBalance(bankFile,bankRows,bankClosing))},{label:"الرصيد الختامي بالبنك",value:money(bankClosing)},{label:"الرصيد الختامي بالنظام",value:money(systemClosing)},{label:"حركات مطابقة",value:String(matched.length),tone:"ok"},{label:"مطابقات مجمعة",value:String(grouped.length),tone:"ok"},{label:"بالبنك فقط",value:String(pureBankOnly.length),tone:pureBankOnly.length?"bad":"ok"},{label:"بالنظام فقط",value:String(pureSystemOnly.length),tone:pureSystemOnly.length?"warn":"ok"},{label:"مبالغ مختلفة",value:String(amountDifferences.length),tone:amountDifferences.length?"bad":"ok"},{label:"حركات مكررة",value:String(bankDuplicates.length+systemDuplicates.length),tone:bankDuplicates.length+systemDuplicates.length?"warn":"ok"}],findings:[{title:firstDifference?"أول نقطة اختلاف":"لا توجد فروقات بالحركات المقروءة",detail:firstDifference?`${formatDate(firstDifference.date)} — ${firstDifference.label} — ${money(Math.abs(firstDifference.diff))}.`:"كل الحركات المقروءة تمت مطابقتها بالمبلغ والتاريخ والمرجع، بما فيها التجميعات.",tone:firstDifference?"bad":"ok"},{title:"المطابقات المجمعة",detail:`${grouped.length} مجموعة تم فيها ربط حركة واحدة بعدة حركات أو العكس دون تكرار الاستخدام.`,tone:grouped.length?"ok":"info"},{title:"عمولات ورسوم بنكية تحتاج مراجعة",detail:`${fees.length} حركة بقيمة ${money(fees.reduce((sum,row)=>sum+row.amount,0))}.`,tone:fees.length?"warn":"ok"},{title:"إيداعات أو تحويلات معلقة",detail:`${pendingDeposits.length} حركة بقيمة ${money(pendingDeposits.reduce((sum,row)=>sum+row.amount,0))}.`,tone:pendingDeposits.length?"warn":"ok"},{title:"الحركات المكررة",detail:`كشف البنك: ${bankDuplicates.length}، النظام: ${systemDuplicates.length}.`,tone:bankDuplicates.length+systemDuplicates.length?"warn":"ok"},{title:"طريقة المطابقة",detail:"مطابقة بالمبلغ والتاريخ والمرجع مع سماح بفارق تاريخ حتى 3 أيام، ثم مطابقة مجمعة من حركتين إلى خمس حركات.",tone:"info"}],table:{headers:["الحالة","تاريخ البنك","مبلغ البنك","تاريخ النظام","مبلغ النظام","الفرق","المرجع","الإجراء المقترح"],rows:tableRows.slice(0,500)},checks:[...standardChecks(ds),"ارفع كشف البنك أولًا وكشف حساب البنك من النظام ثانيًا ولنفس الفترة.","تم فحص المطابقات الفردية والمجمعة والحركات المكررة وأول نقطة اختلاف والعمولات والإيداعات المعلقة.","أي قيد ظاهر هو اقتراح للمراجعة فقط؛ لا يتم إنشاء أو ترحيل قيود تلقائيًا."],sources:[bankFile.fileName,systemFile.fileName]};
}

export function analyzeData(action: string, ds: DataSet[],creditDays=30): AnalysisResult {
  if(!ds.length) throw new Error("يرجى اختيار ملف واحد على الأقل.");
  const request=normalize(action);
  // ملف واحد = تحليل كشف المورد دائمًا. المطابقة لا تبدأ إلا عند اختيارها صراحة
  // ورفع ملفين، حتى لا يتوقف التحليل العادي بطلب ملف ثانٍ.
  if(request.includes("مطابقه كشف المورد")&&ds.length>=2) return supplierReconciliation(ds);
  if(ds.some(d=>d.kind==="pdf") && ds.every(d=>!d.rows.length)) return pdfAnalysis(action,ds,creditDays);
  if(request.includes("مراجعه شامله")&&ds.length===1&&column(ds[0],"debit")&&column(ds[0],"credit")&&column(ds[0],"balance")&&column(ds[0],"date")) return ledger(ds,true,creditDays);
  if(action.includes("قيود اليوميه")||action.includes("قيود اليومية")) return journalAudit(ds);
  if(action.includes("مكرر")) return duplicates(ds);
  if(action.includes("ضريبه القيمه")||action.includes("ضريبة القيمة")) return invoices(ds,false,true);
  if(action.includes("مبيعات")&&action.includes("ضريب")) return invoices(ds,false,false);
  if(action.includes("مشتريات")&&!action.includes("مورد")) return invoices(ds,true,false);
  if(action.includes("مورد")) return ledger(ds,true,creditDays);
  if(action.includes("عميل")||action.includes("ديون")||action.includes("استحقاق")) return ledger(ds,false,creditDays);
  if(action.includes("صنف")||action.includes("مخزون")||action.includes("جرد")) return inventory(ds);
  if(action.includes("ربح")||action.includes("خسار")||action.includes("شهري")) return monthlyProfit(ds);
  if(action.includes("ميزاني")) return statements(ds,true);
  if(action.includes("ميزان")||action.includes("قيد")) return trialBalance(ds);
  if(action.includes("بنك")||action.includes("راجحي")||action.includes("ساب")||action.includes("شبك")) return bank(ds);
  if(action.includes("قوائم")) return statements(ds,false);
  if(action.includes("تدفق")) return cashFlow(ds);
  if(action.includes("زكاه")||action.includes("زكاة")) return zakat(ds);
  if(action.includes("اصول")||action.includes("أصول")||action.includes("اهلاك")||action.includes("إهلاك")) return assets(ds);
  if(action.includes("رواتب")||action.includes("عهد")) return payroll(ds);
  if(action.includes("نشاط تجاري")) return monthlyProfit(ds);
  return generic(ds);
}
