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
function reconcileLots(lots:AgingLot[],target:number,fallbackDate:Date){const current=lots.reduce((s,x)=>s+x.amount,0),difference=target-current;if(difference>.02)lots.push({date:fallbackDate,amount:difference,ref:"رصيد مرحّل"});else if(difference<-.02)lots=consumeOldest(lots,-difference);return lots}
function buildAging(lots:AgingLot[],asOf:Date,basis:string,title="أعمار المستحقات",creditDays=30){const raw=agingLabels.map(()=>({amount:0,count:0})),nonDueThrough=creditDays<=30?0:creditDays<=45?1:2;for(const lot of lots.filter(x=>x.amount>.001)){const invoiceAge=Math.max(0,Math.floor((asOf.getTime()-lot.date.getTime())/86400000)),index=agingIndex(invoiceAge);raw[index].amount+=lot.amount;raw[index].count++}const total=raw.reduce((s,x)=>s+x.amount,0);return{title,asOf:pdfDateFormatter.format(asOf),basis:`${basis} العمر يبقى محسوبًا من تاريخ الفاتورة، ومدة الائتمان ${creditDays} يومًا تحدد فقط حالة مستحقة/غير مستحقة.`,total:money(total),buckets:raw.map((bucket,index)=>({label:`${agingLabels[index]} · ${index<=nonDueThrough?"غير مستحق":"مستحق"}`,value:money(bucket.amount),count:bucket.count,percent:total?`${(bucket.amount/total*100).toFixed(1)}%`:"0%",tone:agingTones[index]}))}}

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
  if(a.includes("مخزون")||a.includes("صنف"))return "تحليل المخزون من PDF";
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
  if(a.includes("مخزون")||a.includes("صنف"))return "التركيز على كود الصنف والوارد والصادر والرصيد؛ يلزم ملف منظم لتحديد العجز بدقة.";
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
  const debit=column(d,"debit"),credit=column(d,"credit"),bal=column(d,"balance"),amount=column(d,"amount"),date=column(d,"date"),invoiceDate=column(d,"invoiceDate"),ageDate=invoiceDate||date,ref=column(d,"ref"),sign=closing<0?-1:1,polarity=ledgerPolarity(d,debit,credit,bal);let lots:AgingLot[]=[],previousBalance:number|null=null,fallbackDate:Date|null=null;
  const transactionRows=d.rows.filter(row=>!!parseDate(value(row,ageDate))||(!!bal&&String(value(row,bal)??"").trim()!==""));
  for(const row of transactionRows){const issueDate=parseDate(value(row,ageDate)),lotDate=issueDate||now;fallbackDate=issueDate||fallbackDate;const rawBalance=value(row,bal),hasBalance=!!bal&&String(rawBalance??"").trim()!=="";let change=0;if(hasBalance){const current=number(rawBalance)*sign;change=previousBalance===null?polarity*(number(value(row,debit))-number(value(row,credit))):current-previousBalance;previousBalance=current}else if(debit||credit){const dr=number(value(row,debit)),cr=number(value(row,credit));change=polarity*(dr-cr)}else if(amount)change=number(value(row,amount));if(change>.001)lots.push({date:lotDate,amount:change,ref:String(value(row,ref)||"—")});else if(change<-.001)lots=consumeOldest(lots,-change)}
  lots=reconcileLots(lots,Math.abs(closing),fallbackDate||now).sort((a,b)=>a.date.getTime()-b.date.getTime());
  const basis=invoiceDate?"العمر محسوب من تاريخ إصدار كل فاتورة حتى آخر تاريخ في الكشف، مع توزيع السداد على أقدم فاتورة أولًا (FIFO). مجموع فئات الأعمار الست يطابق إجمالي المديونية دون فرق.":date?"العمر محسوب من التاريخ الموجود في سطر الفاتورة حتى آخر تاريخ في الكشف، مع توزيع السداد على أقدم فاتورة أولًا (FIFO). مجموع فئات الأعمار الست يطابق إجمالي المديونية دون فرق.":"لم يظهر تاريخ فاتورة صالح؛ لا يمكن اعتماد توزيع الأعمار دون تاريخ الإصدار.";
  const aging=buildAging(lots,now,basis,supplier?"أعمار مستحقات المورد بعد مدة الائتمان":"أعمار مستحقات العميل بعد مدة الائتمان",creditDays);
  const rows=lots.map(lot=>{const invoiceAge=Math.max(0,Math.floor((now.getTime()-lot.date.getTime())/86400000)),due=invoiceAge>creditDays;return[lot.ref||"—",pdfDateFormatter.format(lot.date),String(invoiceAge),`${agingLabels[agingIndex(invoiceAge)]} · ${due?"مستحق":"غير مستحق"}`,money(lot.amount)]});
  return{aging,rows,openCount:lots.length,basis};
}

function prepareSupplierSheet(d:DataSet,closing:number){
  const debit=column(d,"debit"),credit=column(d,"credit"),balance=column(d,"balance"),date=column(d,"invoiceDate")||column(d,"date"),ref=column(d,"ref"),invoice=column(d,"invoice"),type=column(d,"type"),account=column(d,"account"),sign=closing<0?-1:1,polarity=ledgerPolarity(d,debit,credit,balance);
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
    {label:"سداد",value:money(totals.payment),hint:"دفعات وتحويلات مالية",tone:"ok"},
    {label:"مرتجع",value:money(totals.return),hint:"مرتجعات وخصومات",tone:"ok"},
    {label:"إشعار خصم",value:money(totals.discount),hint:"إشعارات خصم مثبتة في الكشف",tone:"ok"},
    {label:"غير مصنف",value:money(totals.increase+totals.decrease),hint:"لم يوضح الملف نوع المستند",tone:"info"},
    ...(Math.abs(unread)>.02?[{label:"فرق يحتاج مراجعة",value:money(Math.abs(unread)),hint:"فرق بين تفاصيل الحركات والرصيد الختامي",tone:"bad" as const}]:[]),
    {label:"صافي الرصيد حسب الكشف",value:money(target),hint:balance?"مطابق لآخر رصيد مسجل":"محسوب من المدين والدائن",tone:target?"warn":"ok"},
  ]};
  return{flow,rows,totalIncrease:totals.purchase+totals.increase,totalDecrease:totals.sale+totals.payment+totals.return+totals.discount+totals.decrease};
}

function ledger(ds: DataSet[], supplier=false,creditDays=30): AnalysisResult {
  const d=ds[0], debit=column(d,"debit"), credit=column(d,"credit"), bal=column(d,"balance"),date=column(d,"date"),invoiceDate=column(d,"invoiceDate"),ageDate=invoiceDate||date;
  const transactionRows=d.rows.filter(row=>!!parseDate(value(row,ageDate))||(!!bal&&String(value(row,bal)??"").trim()!=="")),totalDebit=transactionRows.reduce((s,r)=>s+number(value(r,debit)),0), totalCredit=transactionRows.reduce((s,r)=>s+number(value(r,credit)),0);
  const lastBalanceRow=bal?[...d.rows].reverse().find(row=>String(value(row,bal)??"").trim()!==""):undefined,closing=bal&&lastBalanceRow?number(value(lastBalanceRow,bal)):supplier?totalCredit-totalDebit:totalDebit-totalCredit,lastDate=[...d.rows].reverse().map(r=>parseDate(value(r,ageDate))).find((x):x is Date=>!!x),reportDate=parseDate(d.reportAsOf),asOf=reportDate?new Date(reportDate):lastDate?new Date(lastDate):new Date();asOf.setHours(0,0,0,0);
  const prepared=prepareLedgerAging(d,supplier,closing,asOf,creditDays),supplierSheet=supplier?prepareSupplierSheet(d,closing):undefined,missing=[debit,credit,ageDate].filter(x=>!x).length,dueLabel=supplier?"صافي رصيد المورد حسب الكشف":"إجمالي المستحق على العميل",dueValue=Math.abs(closing);
  const summary:AnalysisResult["summary"]=supplier&&supplierSheet?[
    {label:"إجمالي الدائن",value:money(supplierSheet.totalIncrease)},
    {label:"إجمالي المدين",value:money(supplierSheet.totalDecrease)},
    {label:dueLabel,value:money(dueValue),tone:dueValue?"warn":"ok"},
    {label:"فواتير مفتوحة",value:String(prepared.openCount),tone:prepared.openCount?"warn":"ok"},
  ]:[
    {label:"إجمالي المدين",value:money(totalDebit)},
    {label:"إجمالي الدائن",value:money(totalCredit)},
    {label:dueLabel,value:money(dueValue),tone:dueValue?"warn":"ok"},
    {label:"فواتير مفتوحة",value:String(prepared.openCount),tone:prepared.openCount?"warn":"ok"},
  ];
  const findings:AnalysisResult["findings"]=[
    {title:supplier?"تصنيف نوع المستند":"نتيجة استحقاق العميل",detail:supplier?"يعرض الجدول نوع المستند صراحةً: فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف إذا تعذرت قراءة البيان.":`${dueLabel} هو ${money(dueValue)} وفق آخر رصيد في الملف، مع توزيع التحصيلات على أقدم فاتورة أولًا.`,tone:dueValue?"warn":"ok"},
    {title:"مصدر صافي الرصيد",detail:bal?"تم أخذ الصافي من آخر خلية رصيد غير فارغة في الكشف.":"لم يظهر عمود رصيد؛ تم احتساب الصافي من المدين والدائن.",tone:bal?"ok":"warn"},
    {title:"أساس أعمار المستحقات",detail:prepared.basis,tone:"ok"},
    {title:missing?"أعمدة تحتاج تأكيد":"تم التعرف على الأعمدة",detail:missing?"بعض أعمدة المدين/الدائن/تاريخ إصدار الفاتورة غير واضحة؛ راجع أسماء الأعمدة.":"تم احتساب النتائج من الأعمدة الأصلية دون تعديل.",tone:missing?"warn":"ok"},
  ];
  const table=supplier&&supplierSheet?{headers:["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"],rows:supplierSheet.rows.slice(-300)}:prepared.rows.length?{headers:["المرجع","تاريخ إصدار الفاتورة","العمر بالأيام","فئة العمر","الرصيد المفتوح"],rows:prepared.rows.slice(-100)}:undefined;
  const checks=[...standardChecks(ds),...(supplier?["صُنفت مستندات المورد إلى فاتورة مشتريات، فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف عند تعذر قراءة البيان.","أعمار الرصيد تُحسب من تاريخ إصدار الفاتورة الظاهر في كشف الحساب.","صافي الرصيد يطابق آخر رصيد غير فارغ في الكشف عند توفر عمود الرصيد."]:["في حساب العميل تُصنف المستندات إلى فاتورة مبيعات، سداد، مرتجع، إشعار خصم، أو غير مصنف عند تعذر قراءة البيان.","أعمار ديون العميل تُحسب من تاريخ إصدار الفاتورة الظاهر في كشف الحساب.","تم توزيع التحصيلات على أقدم فاتورة أولًا (FIFO)."])];
  return{title:supplier?"تحليل مشتريات ومبيعات المورد":"تحليل حساب العميل",confidence:missing?"متوسطة":"عالية",conclusion:{label:dueLabel,value:money(dueValue),detail:dueValue?`${bal?"تم اعتماد آخر رصيد غير فارغ في الكشف":"تم احتساب الصافي من المدين والدائن"}${lastDate?` بتاريخ ${pdfDateFormatter.format(lastDate)}`:""}.`:`لا يوجد رصيد ظاهر${lastDate?` حتى ${pdfDateFormatter.format(lastDate)}`:""}.`,tone:dueValue?"warn":"ok"},aging:prepared.aging,supplierFlow:supplierSheet?.flow,summary,findings,table,checks,sources:ds.map(x=>x.fileName)};
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
  const d=ds[0],item=column(d,"item"),opening=column(d,"opening"),purchases=column(d,"purchases"),sales=column(d,"sales"),ri=column(d,"returnsIn"),ro=column(d,"returnsOut"),actual=column(d,"actual")||column(d,"system");
  const diffs=d.rows.map((r,i)=>{const expected=number(value(r,opening))+number(value(r,purchases))+number(value(r,ri))-number(value(r,sales))-number(value(r,ro));const a=number(value(r,actual));return {name:String(value(r,item)||`صف ${i+1}`),expected,actual:a,diff:a-expected}}).filter(x=>Math.abs(x.diff)>.0001).sort((a,b)=>Math.abs(b.diff)-Math.abs(a.diff));
  const shortage=diffs.filter(x=>x.diff<0),surplus=diffs.filter(x=>x.diff>0);
  return {title:"مقارنة حركة الأصناف",confidence:item&&actual?"عالية":"متوسطة",summary:[{label:"الأصناف المفحوصة",value:String(d.rows.length)},{label:"أصناف ناقصة",value:String(shortage.length),tone:shortage.length?"bad":"ok"},{label:"أصناف زائدة",value:String(surplus.length),tone:surplus.length?"warn":"ok"},{label:"إجمالي فرق الكمية",value:qty(diffs.reduce((s,x)=>s+x.diff,0))}],findings:[{title:diffs.length?"تم اكتشاف فروقات":"لا توجد فروقات بالحركات المقروءة",detail:diffs.length?`يوجد ${diffs.length} صنفًا يحتاج مراجعة حركة أو جرد.`:"الرصيد المتوقع يساوي الرصيد الفعلي لكل الأصناف المقروءة.",tone:diffs.length?"bad":"ok"},{title:"معادلة الفحص",detail:"أول المدة + المشتريات + مرتجع المبيعات − المبيعات − مرتجع المشتريات = الرصيد المتوقع.",tone:"info"}],table:diffs.length?{headers:["الصنف","المتوقع","الفعلي","الفرق"],rows:diffs.slice(0,30).map(x=>[x.name,qty(x.expected),qty(x.actual),qty(x.diff)])}:undefined,checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
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
  const [a,b]=ds, ac=column(a,"amount"),bc=column(b,"amount"),ad=column(a,"debit"),acr=column(a,"credit"),bd=column(b,"debit"),bcr=column(b,"credit");
  const amounts=(d:DataSet,m?:string,deb?:string,cre?:string)=>d.rows.map(r=>m?number(value(r,m)):number(value(r,deb))-number(value(r,cre)));
  const aa=amounts(a,ac,ad,acr),bb=amounts(b,bc,bd,bcr),used=new Set<number>(),unmatchedA:number[]=[];
  aa.forEach((v,i)=>{const j=bb.findIndex((x,k)=>!used.has(k)&&Math.abs(Math.abs(x)-Math.abs(v))<.01);if(j>=0)used.add(j);else unmatchedA.push(i)}); const unmatchedB=bb.map((_,i)=>i).filter(i=>!used.has(i));
  const totalA=aa.reduce((s,x)=>s+x,0),totalB=bb.reduce((s,x)=>s+x,0);
  return {title:"مطابقة كشف البنك مع النظام",confidence:(ac||ad)&&(bc||bd)?"عالية":"متوسطة",summary:[{label:"حركات الكشف",value:String(a.rows.length)},{label:"حركات النظام",value:String(b.rows.length)},{label:"المعاملات المطابقة",value:String(used.size),tone:"ok"},{label:"فرق الإجماليات",value:money(totalA-totalB),tone:Math.abs(totalA-totalB)>.01?"bad":"ok"}],findings:[{title:"غير موجودة في النظام",detail:`${unmatchedA.length} حركة من ${a.fileName}`,tone:unmatchedA.length?"bad":"ok"},{title:"غير موجودة في البنك",detail:`${unmatchedB.length} حركة من ${b.fileName}`,tone:unmatchedB.length?"warn":"ok"},{title:"طريقة المطابقة",detail:"تمت المطابقة الأولية بالمبلغ. يُنصح بإضافة المرجع والتاريخ للتأكد من الحركات المتكررة.",tone:"info"}],table:{headers:["المصدر","رقم الصف","المبلغ"],rows:[...unmatchedA.slice(0,15).map(i=>[a.fileName,String(i+2),money(aa[i])]),...unmatchedB.slice(0,15).map(i=>[b.fileName,String(i+2),money(bb[i])])]},checks:standardChecks(ds),sources:ds.map(x=>x.fileName)};
}

export function analyzeData(action: string, ds: DataSet[],creditDays=30): AnalysisResult {
  if(!ds.length) throw new Error("يرجى اختيار ملف واحد على الأقل.");
  const request=normalize(action);
  // ملف واحد = تحليل كشف المورد دائمًا. المطابقة لا تبدأ إلا عند اختيارها صراحة
  // ورفع ملفين، حتى لا يتوقف التحليل العادي بطلب ملف ثانٍ.
  if(request.includes("مطابقه كشف المورد")&&ds.length>=2) return supplierReconciliation(ds);
  if(ds.some(d=>d.kind==="pdf") && ds.every(d=>!d.rows.length)) return pdfAnalysis(action,ds,creditDays);
  if(request.includes("مراجعه شامله")&&ds.length===1&&column(ds[0],"debit")&&column(ds[0],"credit")&&column(ds[0],"balance")&&column(ds[0],"date")) return ledger(ds,true,creditDays);
  if(action.includes("مكرر")) return duplicates(ds);
  if(action.includes("ضريبه القيمه")||action.includes("ضريبة القيمة")) return invoices(ds,false,true);
  if(action.includes("مبيعات")&&action.includes("ضريب")) return invoices(ds,false,false);
  if(action.includes("مشتريات")&&!action.includes("مورد")) return invoices(ds,true,false);
  if(action.includes("مورد")) return ledger(ds,true,creditDays);
  if(action.includes("عميل")||action.includes("ديون")||action.includes("استحقاق")) return ledger(ds,false,creditDays);
  if(action.includes("صنف")||action.includes("مخزون")) return inventory(ds);
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
