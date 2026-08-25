import test from "node:test";
import assert from "node:assert/strict";
import { File } from "node:buffer";
import * as XLSX from "xlsx";
import { analyzeData, readAccountingFile } from "../app/analyzer.ts";

async function csv(name: string, value: string) {
  return readAccountingFile(new File([value], name, { type: "text/csv" }) as unknown as globalThis.File);
}

test("reads accounting tables across differently formatted Excel sheets", async () => {
  const workbook=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([["شركة مزايا"],["تقرير داخلي"]]),"غلاف");
  XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([["كشف الحساب حتى 24/08/2026"],[""],["التاريخ","رقم المرجع","مدين","دائن","الرصيد"],["2026-08-01","A-1",100,0,100],["إجمالي",100,0,0,100]]),"الحركات");
  XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([["التاريخ","رقم المرجع","مدين","دائن","الرصيد"],["2026-08-02","A-2",0,50,50]]),"تكملة");
  const bytes=XLSX.write(workbook,{type:"array",bookType:"xlsx"});
  const data=await readAccountingFile(new File([bytes],"multi-sheet.xlsx") as unknown as globalThis.File);
  assert.deepEqual(data.parsedSheets,["الحركات","تكملة"]);
  assert.equal(data.rows.length,2);
  assert.equal(data.reportAsOf,"24/08/2026");
});

test("calculates monthly profit and loss", async () => {
  const data = await csv("profit.csv", "التاريخ,الايرادات,تكلفة المبيعات,المصروفات\n2026-01-01,100000,60000,20000\n2026-02-01,50000,40000,15000");
  const result = analyzeData("الربحية الشهرية", [data]);
  assert.equal(result.title, "تحليل الربحية الشهرية");
  assert.equal(result.table?.rows.length, 2);
  assert.equal(result.table?.rows[0][5], "ربح");
  assert.equal(result.table?.rows[1][5], "خسارة");
});

test("detects inventory shortage", async () => {
  const data = await csv("inventory.csv", "الصنف,رصيد اول المدة,المشتريات,المبيعات,الرصيد الفعلي\nHP 4303,10,5,4,9\nDell Monitor,8,2,3,7");
  const result = analyzeData("قارن حركة الأصناف وحدد الصنف الناقص", [data]);
  assert.equal(result.title, "مقارنة حركة الأصناف");
  assert.equal(result.table?.rows.length, 1);
  assert.equal(result.table?.rows[0][0], "HP 4303");
  assert.equal(result.table?.rows[0][3], "؜-٢");
});

test("calculates customer balance and overdue invoice", async () => {
  const data = await csv("customer.csv", "رقم الفاتورة,تاريخ الاستحقاق,مدين,دائن,الرصيد\nINV-1,2025-01-01,1000,0,1000\nRCPT-1,2025-01-10,0,250,750");
  const result = analyzeData("تحليل حساب عميل", [data]);
  assert.equal(result.title, "تحليل حساب العميل");
  assert.equal(result.summary[2].value, "٧٥٠٫٠٠ ر.س");
  assert.ok((result.table?.rows.length ?? 0) > 0);
});

test("applies invoice issue date aging automatically to customer balances", async () => {
  const data = await csv("recent-customer.csv", "تاريخ إصدار الفاتورة,رقم الفاتورة,مدين,دائن,الرصيد\n2026-08-01,SALE-NEW,1250,0,1250");
  const result = analyzeData("تحليل حساب عميل", [data]);

  assert.equal(result.conclusion?.label, "إجمالي المستحق على العميل");
  assert.equal(result.conclusion?.value, "١٬٢٥٠٫٠٠ ر.س");
  assert.equal(result.aging?.buckets[0].label, "0–30 يومًا");
  assert.equal(result.aging?.buckets[0].value, "١٬٢٥٠٫٠٠ ر.س");
  assert.equal(result.aging?.buckets[1].value, "٠٫٠٠ ر.س");
  assert.match(result.aging?.basis ?? "", /تاريخ إصدار الفاتورة/);
  assert.ok(result.checks.some(check => /أعمار ديون العميل/.test(check)));
  assert.ok(result.checks.some(check => /FIFO/.test(check)));
});

test("keeps customer aging buckets equal to the exact total debt", async () => {
  const data = await csv("customer-50000.csv", "تاريخ إصدار الفاتورة,رقم الفاتورة,مدين,دائن,الرصيد\n2026-08-01,SALE-1,10000,0,10000\n2026-07-01,SALE-2,10000,0,20000\n2026-06-01,SALE-3,10000,0,30000\n2026-05-01,SALE-4,10000,0,40000\n2026-04-01,SALE-5,10000,0,50000");
  const result = analyzeData("تحليل حساب عميل", [data]);

  assert.equal(result.conclusion?.value, "٥٠٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.aging?.total, result.conclusion?.value);
  assert.deepEqual(result.aging?.buckets.map(bucket => bucket.value), ["١٠٬٠٠٠٫٠٠ ر.س","١٠٬٠٠٠٫٠٠ ر.س","١٠٬٠٠٠٫٠٠ ر.س","١٠٬٠٠٠٫٠٠ ر.س","١٠٬٠٠٠٫٠٠ ر.س"]);
  assert.match(result.aging?.basis ?? "", /يطابق إجمالي المديونية/);
});

test("shows a prominent supplier amount due for spreadsheet reports", async () => {
  const data = await csv("supplier.csv", "التاريخ,رقم الفاتورة,تاريخ الاستحقاق,مدين,دائن,الرصيد\n2026-01-01,PUR-1,2026-01-31,0,1000,1000\n2026-01-15,PAY-1,2026-01-31,250,0,750");
  const result = analyzeData("تحليل حساب مورد", [data]);
  assert.equal(result.summary[2].label, "صافي رصيد المورد حسب الكشف");
  assert.equal(result.conclusion?.value, "٧٥٠٫٠٠ ر.س");
  assert.match(result.conclusion?.detail ?? "", /١٥.*٠١.*٢٠٢٦/);
  assert.equal(result.aging?.buckets.length, 5);
  assert.equal(result.aging?.total, "٧٥٠٫٠٠ ر.س");
  assert.equal(result.aging?.buckets[4].value, "٧٥٠٫٠٠ ر.س");
});

test("uses invoice dates only and shows the nearest upcoming due date", async () => {
  const data = await csv("supplier-next-due.csv", "التاريخ,رقم المرجع,نوع الحركة,مدين,دائن,الرصيد\n2026-06-21,PUR-OLD,فاتورة مشتريات,0,10000,10000\n2026-06-30,PUR-NEXT,فاتورة مشتريات,0,5000,15000\n2026-08-24,PAY-1,سند صرف,1000,0,14000");
  data.reportAsOf = "25/08/2026";
  const result = analyzeData("تحليل حساب مورد", [data], 60);

  assert.equal(result.dueSchedule?.nextDueDate, "٢٩‏/٠٨‏/٢٠٢٦");
  assert.equal(result.dueSchedule?.nextDueDays, 4);
  assert.equal(result.dueSchedule?.nextDueCreditDays, 60);
  assert.equal(result.dueSchedule?.nextDueRef, "PUR-NEXT");
  assert.equal(result.dueSchedule?.nextDueAmount, "١٤٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.dueSchedule?.nextDueInvoicesAmount, "٥٬٠٠٠٫٠٠ ر.س");
  assert.notEqual(result.dueSchedule?.nextDueDate, "٢٤‏/٠٨‏/٢٠٢٦");
});

test("reads an unpaid customer invoice list and applies 30 to 60 day terms", async () => {
  const data = await csv("open-customer-invoices.csv", "رقم السند_1,المستند_2,التأريخ_3,مدين_4\n1,فاتوره مبيعات اجل,2026-06-01,1000\n2,فاتوره مبيعات اجل,2026-08-01,500");
  data.reportAsOf = "25/08/2026";
  const result = analyzeData("تحليل حساب عميل", [data], 60);

  assert.equal(result.summary[0].value, "١٬٥٠٠٫٠٠ ر.س");
  assert.equal(result.conclusion?.value, "١٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.summary[1].value, "٥٠٠٫٠٠ ر.س");
  assert.equal(result.dueSchedule?.nextDueDate, "٣٠‏/٠٩‏/٢٠٢٦");
  assert.equal(result.dueSchedule?.nextDueAmount, "١٬٥٠٠٫٠٠ ر.س");
  assert.equal(result.dueSchedule?.nextDueInvoicesAmount, "٥٠٠٫٠٠ ر.س");
  assert.match(result.findings[1].detail, /قائمة فواتير مفتوحة غير مسددة/);
});

test("places a recent supplier invoice in the 0 to 30 day bucket from its issue date", async () => {
  const data = await csv("recent-supplier.csv", "تاريخ إصدار الفاتورة,رقم الفاتورة,مدين,دائن,الرصيد\n2026-08-01,PUR-NEW,0,1250,1250");
  const result = analyzeData("تحليل حساب مورد", [data]);

  assert.equal(result.aging?.buckets[0].label, "0–30 يومًا");
  assert.equal(result.aging?.buckets[0].value, "١٬٢٥٠٫٠٠ ر.س");
  assert.equal(result.aging?.buckets[1].value, "٠٫٠٠ ر.س");
  assert.match(result.aging?.basis ?? "", /تاريخ إصدار الفاتورة/);
});

test("uses all five non-overlapping aging ranges from invoice issue dates", async () => {
  const data = await csv("supplier-aging.csv", "تاريخ إصدار الفاتورة,رقم الفاتورة,مدين,دائن,الرصيد\n2026-08-01,PUR-1,0,100,100\n2026-07-01,PUR-2,0,100,200\n2026-06-01,PUR-3,0,100,300\n2026-05-01,PUR-4,0,100,400\n2026-04-01,PUR-5,0,100,500");
  const result = analyzeData("تحليل حساب مورد", [data]);

  assert.deepEqual(result.aging?.buckets.map(bucket => bucket.label), ["0–30 يومًا","31–60 يومًا","61–90 يومًا","91–120 يومًا","أكثر من 120 يومًا"]);
  assert.deepEqual(result.aging?.buckets.map(bucket => bucket.value), ["١٠٠٫٠٠ ر.س","١٠٠٫٠٠ ر.س","١٠٠٫٠٠ ر.س","١٠٠٫٠٠ ر.س","١٠٠٫٠٠ ر.س"]);
});

test("automatically reconciles a supplier statement with the company ledger", async () => {
  const supplier = await csv("supplier-statement.csv", "التاريخ,رقم المرجع,مدين,دائن,الرصيد\n2026-01-01,INV-1,1000,0,1000\n2026-01-10,INV-2,500,0,1500\n2026-01-20,INV-X,300,0,1800");
  const company = await csv("company-erp.csv", "التاريخ,رقم المرجع,مدين,دائن,الرصيد\n2026-01-02,INV-1,0,1000,-1000\n2026-01-11,INV-2,0,500,-1500\n2026-01-25,OUR-ONLY,0,200,-1700");
  const result = analyzeData("مطابقة كشف المورد مع حسابنا وأظهر الفروقات", [supplier, company]);
  assert.equal(result.title, "مطابقة كشف المورد مع حساب الشركة");
  assert.equal(result.summary[2].value, "2");
  assert.equal(result.summary[3].value, "2");
  assert.equal(result.conclusion?.value, "١٠٠٫٠٠ ر.س");
  assert.equal(result.aging?.total, "١٬٨٠٠٫٠٠ ر.س");
  assert.ok(result.table?.rows.some(row => row[0] === "عند المورد فقط"));
  assert.ok(result.table?.rows.some(row => row[0] === "عند الشركة فقط"));
});

test("asks for both files before starting supplier reconciliation", async () => {
  const supplier = await csv("supplier.csv", "التاريخ,المبلغ\n2026-01-01,100");
  const result = analyzeData("مطابقة كشف المورد مع حسابنا", [supplier]);
  assert.equal(result.title, "مطابقة كشف المورد مع حساب الشركة");
  assert.equal(result.summary[1].value, "2");
  assert.match(result.findings[0].detail, /كشف المورد أولًا/);
});

test("detects duplicated entries", async () => {
  const data = await csv("journal.csv", "رقم المرجع,التاريخ,المبلغ\nJV-1,2026-01-01,500\nJV-1,2026-01-01,500\nJV-2,2026-01-02,300");
  const result = analyzeData("كشف القيود المكررة", [data]);
  assert.equal(result.title, "كشف القيود والفواتير المكررة");
  assert.equal(result.summary[1].value, "1");
});

test("audits journal entries without changing them", async () => {
  const data = await csv("journal.csv", "رقم القيد,التاريخ,الحساب,البيان,مدين,دائن\nJV-1,2026-08-20,الصندوق,قبض,1000,0\nJV-1,2026-08-20,المبيعات,قبض,0,1000\nJV-2,2026-08-21,مصروف,رسوم,500,0\nJV-2,2026-08-21,مصروف,رسوم,500,0");
  const result = analyzeData("تحليل قيود اليومية", [data]);
  assert.equal(result.title, "تحليل قيود اليومية");
  assert.equal(result.summary.find(item=>item.label==="قيود غير متوازنة")?.value, "1");
  assert.equal(result.summary.find(item=>item.label==="سطور مكررة")?.value, "1");
  assert.match(result.checks.join(" "), /لا يتم إنشاء أو ترحيل/);
});

test("shows customer flow like supplier flow with customer deductions", async () => {
  const data = await csv("customer.csv", "التاريخ,رقم القيد,نوع الحركة,مدين,دائن,الرصيد\n2026-06-01,1,فاتورة مبيعات,10000,0,10000\n2026-07-20,2,سند قبض,0,3000,7000\n2026-07-25,3,مردود مبيعات,0,1000,6000");
  const result = analyzeData("تحليل حساب عميل", [data], 45);
  assert.equal(result.title, "تحليل مبيعات وتحصيلات العميل");
  assert.equal(result.supplierFlow?.title, "ملخص حركة حساب العميل");
  assert.equal(result.summary.find(item=>item.label==="سندات القبض والتحصيل المخصومة")?.value, "٣٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.summary.find(item=>item.label==="مردود المبيعات المخصوم")?.value, "١٬٠٠٠٫٠٠ ر.س");
});

test("compares physical count with system inventory and values differences", async () => {
  const physical = await csv("جرد-فعلي.csv", "الصنف,الكمية\nHP-01,8\nDELL-02,6");
  const system = await csv("مخزون-النظام.csv", "الصنف,رصيد النظام,التكلفة\nHP-01,10,100\nDELL-02,5,200");
  const result = analyzeData("مطابقة الجرد الفعلي مع النظام", [physical, system]);
  assert.equal(result.title, "مطابقة الجرد الفعلي مع رصيد النظام");
  assert.equal(result.summary.find(item=>item.label==="أصناف ناقصة")?.value, "1");
  assert.equal(result.summary.find(item=>item.label==="أصناف زائدة")?.value, "1");
  assert.equal(result.summary.find(item=>item.label==="قيمة العجز")?.value, "٢٠٠٫٠٠ ر.س");
  assert.equal(result.summary.find(item=>item.label==="قيمة الزيادة")?.value, "٢٠٠٫٠٠ ر.س");
});

test("locks the approved PDF due balance rule for suppliers", () => {
  const lines=[
    "10,066.65 0 10,066.65 1/1/2026 قيد افتتاحي",
    "10,532.4 0 465.75 19/1/2026 فاتورة مبيعات",
    "11,515.65 0 983.25 1/4/2026 فاتورة مبيعات",
    "12,159.65 0 644 5/4/2026 فاتورة مبيعات",
    "7,159.65 5,000 0 7/4/2026 سند يومية حوالة",
    "7,619.65 0 460 14/4/2026 فاتورة مبيعات",
    "8,194.65 0 575 15/4/2026 فاتورة مبيعات",
    "3,194.65 5,000 0 10/5/2026 سند يومية حوالة",
    "2,950.85 243.8 0 17/6/2026 فاتورة مرتجع مبيعات",
    "12,030.1 0 9,079.25 12/7/2026 فاتورة مبيعات",
    "17,596.1 0 5,566 20/7/2026 فاتورة مبيعات",
    "14,646.1 2,950 0 21/7/2026 سند يومية حوالة",
    "25,024.85 0 10,378.75 10/8/2026 فاتورة مبيعات",
    "26,284.1 0 1,259.25 10/8/2026 فاتورة مبيعات",
    "26,744.1 0 460 23/8/2026 فاتورة مبيعات",
  ];
  const pdf={fileName:"مورد.pdf",kind:"pdf" as const,rows:[],columns:[],pdfLines:lines,rawText:lines.join("\n"),pages:1,reportAsOf:"24/8/2026"};
  const result45=analyzeData("تحليل حساب مورد",[pdf],45),result30=analyzeData("تحليل حساب مورد",[pdf],30);
  assert.equal(result45.conclusion?.value,"٠٫٨٥ ر.س");
  assert.equal(result30.conclusion?.value,"١٤٬٦٤٦٫١٠ ر.س");
});

test("finds invoice VAT arithmetic error", async () => {
  const data = await csv("sales.csv", "رقم الفاتورة,الصافي قبل الضريبة,الضريبة,الاجمالي شامل الضريبة\nINV-1,1000,150,1150\nINV-2,500,75,600");
  const result = analyzeData("مراجعة المبيعات والضريبة", [data]);
  assert.equal(result.title, "مراجعة فواتير المبيعات والضريبة");
  assert.equal(result.summary[3].value, "1");
});

test("checks balance sheet equation", async () => {
  const data = await csv("balance.csv", "اسم الحساب,نوع الحساب,الرصيد\nالبنك,أصول,1000\nالموردون,التزامات,400\nرأس المال,حقوق الملكية,600");
  const result = analyzeData("تحليل الميزانية", [data]);
  assert.equal(result.title, "تحليل الميزانية والمركز المالي");
  assert.equal(result.summary[2].value, "٠٫٠٠ ر.س");
});

test("creates the selected supplier report from extracted PDF text", () => {
  const result = analyzeData("حلّل كشف حساب المورد وحدد الرصيد المستحق", [{
    fileName: "supplier-statement.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 2,
    pdfLines: [
      "كشف حساب المورد",
      "01/01/2026 فاتورة شراء 1,000.00 0.00 1,000.00",
      "15/01/2026 سداد 0.00 250.00 750.00",
      "الرصيد المستحق 750.00",
    ],
  }]);

  assert.equal(result.title, "تحليل مشتريات ومبيعات المورد من PDF");
  assert.equal(result.summary[3].label, "صافي رصيد المورد حسب الكشف");
  assert.equal(result.summary[3].value, "٧٥٠٫٠٠ ر.س");
  assert.equal(result.conclusion?.label, "صافي رصيد المورد حسب الكشف");
  assert.equal(result.conclusion?.value, "٧٥٠٫٠٠ ر.س");
  assert.match(result.conclusion?.detail ?? "", /١٥.*٠١.*٢٠٢٦/);
  assert.equal(result.aging?.buckets.length, 5);
  assert.equal(result.aging?.total, "٧٥٠٫٠٠ ر.س");
  assert.ok((result.table?.rows.length ?? 0) >= 2);
  assert.equal(result.table?.rows[0][1], "فاتورة مشتريات");
  assert.equal(result.table?.rows[1][1], "سداد");
});

test("uses the invoice date read from a PDF statement for the 0 to 30 day bucket", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "recent-supplier.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "30/07/2026 فاتورة شراء 1,000.00 0.00 1,000.00",
      "الرصيد الختامي 1,000.00",
    ],
  }]);

  assert.equal(result.aging?.buckets[0].value, "١٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.aging?.buckets[1].value, "٠٫٠٠ ر.س");
  assert.match(result.aging?.basis ?? "", /تاريخ إصدار الفاتورة/);
});

test("uses the same invoice aging method for customer PDF statements", () => {
  const result = analyzeData("تحليل حساب عميل", [{
    fileName: "recent-customer.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "30/07/2026 فاتورة مبيعات 1,000.00 0.00 1,000.00",
      "الرصيد الختامي 1,000.00",
    ],
  }]);

  assert.equal(result.conclusion?.label, "إجمالي المستحق على العميل");
  assert.equal(result.conclusion?.value, "١٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.aging?.total, result.conclusion?.value);
  assert.equal(result.aging?.buckets[0].value, "١٬٠٠٠٫٠٠ ر.س");
  assert.equal(result.table?.rows[0][1], "فاتورة مبيعات");
  assert.match(result.findings.map(finding => finding.detail).join(" "), /مجموع الخانات الخمس/);
});

test("uses the printed supplier closing balance when it differs from incomplete PDF movements", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "supplier.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 2,
    pdfLines: [
      "01/01/2026 فاتورة شراء 1,000.00 0.00 1,000.00",
      "10/01/2026 بيع للمورد 200.00 0.00 800.00",
      "الرصيد الختامي 650.00",
    ],
  }]);

  assert.equal(result.conclusion?.value, "٦٥٠٫٠٠ ر.س");
  assert.equal(result.aging?.total, "٦٥٠٫٠٠ ر.س");
  assert.match(result.supplierFlow?.source ?? "", /الرصيد الختامي المطبوع/);
  assert.ok(result.supplierFlow?.items.some(item => item.label === "فرق سطور غير مكتملة القراءة"));
  assert.match(result.findings.map(item => item.title).join(" "), /فرق بين الرصيد المطبوع/);
});

test("does not announce a supplier amount when PDF balance columns are unreliable", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "unclear-supplier.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "01/01/2026 حركة 100.00 500.00",
      "02/01/2026 حركة 40.00 900.00",
    ],
  }]);

  assert.equal(result.conclusion?.value, "غير مؤكد من ملف PDF");
  assert.equal(result.aging, undefined);
  assert.equal(result.supplierFlow, undefined);
});

test("shows the last running balance provisionally when a long PDF sequence is usable", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "running-supplier.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 2,
    pdfLines: [
      "01/01/2026 فاتورة مشتريات 100.00 0.00 100.00",
      "02/01/2026 فاتورة مشتريات 100.00 0.00 200.00",
      "03/01/2026 سند صرف 0.00 50.00 150.00",
      "04/01/2026 فاتورة مشتريات 100.00 0.00 250.00",
      "05/01/2026 حركة غير واضحة 999.00 0.00 300.00",
      "06/01/2026 سند صرف 0.00 25.00 275.00",
    ],
  }]);

  assert.equal(result.conclusion?.value, "٢٧٥٫٠٠ ر.س");
  assert.match(result.conclusion?.detail ?? "", /مبدئي.*آخر رصيد جار/);
  assert.equal(result.aging?.total, "٢٧٥٫٠٠ ر.س");
  assert.ok(result.supplierFlow);
  assert.equal(result.confidence, "متوسطة");
});

test("separates supplier purchases, sales and payments in a clear ledger table", async () => {
  const data = await csv("supplier.csv", "التاريخ,البيان,مدين,دائن,الرصيد\n2026-01-01,فاتورة مشتريات,0,1000,1000\n2026-01-10,بيع للمورد,200,0,800\n2026-01-15,سداد للمورد,300,0,500");
  const result = analyzeData("تحليل حساب مورد", [data]);

  assert.equal(result.title, "تحليل مشتريات ومبيعات المورد");
  assert.deepEqual(result.table?.headers, ["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"]);
  assert.equal(result.table?.rows[0][1], "فاتورة مشتريات");
  assert.equal(result.table?.rows[1][1], "فاتورة مبيعات");
  assert.equal(result.table?.rows[2][1], "سداد");
  assert.doesNotMatch(JSON.stringify(result), /زيادة|نقص/);
  assert.equal(result.conclusion?.value, "٥٠٠٫٠٠ ر.س");
});

test("uses debit and credit terms for fragmented supplier PDF movements", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "supplier-fragmented.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "02/01/2026 SAYED م و ر د 26,634.00 60,006.25",
      "03/01/2026 SAYED م و ر د 2,530.00 62,536.25",
      "07/01/2026 SAYED م و ر د 2,817.50 59,718.75",
      "الرصيد الختامي 59,718.75",
    ],
  }]);

  assert.deepEqual(result.table?.headers, ["التاريخ","نوع المستند","البيان/المرجع","مدين","دائن","الرصيد"]);
  assert.equal(result.table?.rows[1][1], "فاتورة مشتريات");
  assert.equal(result.table?.rows[2][1], "سداد");
  assert.doesNotMatch(JSON.stringify(result), /زيادة|نقص/);
});

test("repairs reversed fragmented Arabic document names before classifying supplier movements", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "fragmented-types.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "01/01/2026 ت ا ي ر ت ش م ة ر و ت ا ف 1,000.00 0.00 1,000.00",
      "02/01/2026 ف ر ص د ن س 0.00 250.00 750.00",
      "الرصيد الختامي 750.00",
    ],
  }]);

  assert.deepEqual(result.table?.rows.map(row => row[1]), ["فاتورة مشتريات", "سداد"]);
});

test("classifies purchases sales payments returns and discount notes in one statement", () => {
  const result = analyzeData("مراجعة شاملة لكل الحسابات", [{
    fileName: "mixed-statement.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: [
      "01/01/2026 فاتورة مشتريات 1,000.00 0.00 1,000.00",
      "02/01/2026 فاتورة مبيعات 0.00 200.00 800.00",
      "03/01/2026 سداد 0.00 300.00 500.00",
      "04/01/2026 مرتجع مشتريات 0.00 100.00 400.00",
      "05/01/2026 إشعار خصم 0.00 50.00 350.00",
      "الرصيد الختامي 350.00",
    ],
  }]);

  assert.deepEqual(result.table?.rows.map(row => row[1]), ["فاتورة مشتريات","فاتورة مبيعات","سداد","مرتجع","إشعار خصم"]);
  assert.equal(result.summary.find(item => item.label === "الرصيد الختامي")?.value, "٣٥٠٫٠٠ ر.س");
});

test("shows a clear message for an unreadable scanned PDF", () => {
  const result = analyzeData("تحليل حساب مورد", [{
    fileName: "scan.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 6,
    rawText: "",
  }]);

  assert.equal(result.title, "تحليل حساب المورد من ملف PDF");
  assert.equal(result.summary[3].value, "PDF مصوّر وغير قابل للقراءة");
  assert.equal(result.confidence, "محدودة");
});

test("keeps every dashboard option on an option-specific PDF report", () => {
  const pdf = [{
    fileName: "report.pdf",
    kind: "pdf" as const,
    rows: [],
    columns: [],
    pages: 1,
    pdfLines: ["01/01/2026 رصيد افتتاحي 1,000.00 1,000.00", "02/01/2026 حركة 250.00 750.00"],
  }];
  const options = [
    "مطابقة كشف البنك", "تحليل حساب عميل", "تحليل حساب مورد",
    "أعمار الديون والاستحقاق", "مطابقة شبكة الراجحي", "مطابقة شبكة ساب",
    "مراجعة ميزان المراجعة", "كشف القيود المكررة", "مراجعة المبيعات والضريبة",
    "مراجعة المشتريات", "تحليل المخزون", "مقارنة حركة الأصناف",
    "الربحية الشهرية", "تحليل القوائم المالية", "تحليل الميزانية",
    "التدفقات النقدية", "مراجعة ضريبة القيمة المضافة", "مراجعة الزكاة",
    "الأصول والإهلاك", "الرواتب والعهد", "تحليل النشاط التجاري",
  ];

  for (const option of options) {
    const result = analyzeData(option, pdf);
    assert.notEqual(result.title, "فحص الملفات المرفوعة", option);
    assert.match(result.title, /PDF/, option);
    assert.ok(result.table?.rows.length, option);
  }
});

test("ignores report dates, selects the transaction date, and repairs fragmented Arabic PDF rows", () => {
  const result = analyzeData("مراجعة شاملة لكل الحسابات", [{
    fileName: "account.pdf",
    kind: "pdf",
    rows: [],
    columns: [],
    pages: 6,
    pdfLines: [
      "الفترة 01-01-2025 إلى 21-08-2026",
      "تاريخ الطباعة 21-08-2026",
      "01-01-2025 م ر د ص ر 43372.25 43372.25",
      "4/3/2025 SAYED 05-01-2025 م و ر د 454.25 43826.50",
      "11-01-2025 م و ر د 7561.25 36265.25",
    ],
  }]);

  assert.equal(result.title, "مراجعة شاملة للحسابات من PDF");
  assert.equal(result.summary[1].value, "3");
  assert.equal(result.summary[3].value, "٣٦٬٢٦٥٫٢٥ ر.س");
  assert.match(result.summary[2].value, /٢٠٢٥/);
  assert.doesNotMatch(result.summary[2].value, /٢٠٢٦/);
  assert.equal(result.table?.headers.join("|"), "التاريخ|نوع المستند|البيان\/المرجع|مدين|دائن|الرصيد");
  assert.equal(result.table?.rows[1][1], "غير مصنف");
  assert.match(result.table?.rows[1][2] ?? "", /حركة مدينة/);
  assert.doesNotMatch(result.table?.rows[1][2] ?? "", /م و ر د/);
  assert.doesNotMatch(JSON.stringify(result), /زيادة|نقص/);
});

test("reconciles bank and ERP by amount date and reference and classifies differences", async () => {
  const bankFile = await csv("bank.csv", "التاريخ,المرجع,البيان,المبلغ\n2026-08-01,TRX-1,تحويل عميل,1000\n2026-08-02,FEE-1,عمولة بنكية,-25\n2026-08-05,TRX-2,سداد مورد,-400");
  const systemFile = await csv("erp.csv", "التاريخ,المرجع,البيان,المبلغ\n2026-08-02,TRX-1,تحويل عميل,1000\n2026-08-05,TRX-2,سداد مورد,-400\n2026-08-06,DEP-1,إيداع معلق,600");
  const result = analyzeData("مطابقة كشف البنك", [bankFile, systemFile]);
  assert.equal(result.title, "مطابقة كشف البنك مع النظام");
  assert.equal(result.summary.find(item=>item.label==="حركات مطابقة")?.value, "2");
  assert.equal(result.summary.find(item=>item.label==="بالبنك فقط")?.value, "1");
  assert.equal(result.summary.find(item=>item.label==="بالنظام فقط")?.value, "1");
  assert.match(result.findings.find(item=>item.title==="أول نقطة اختلاف")?.detail??"", /٠٢.*٠٨.*٢٠٢٦/);
  assert.match(result.table?.rows.find(row=>row[0]==="بالبنك فقط")?.[7]??"", /عمولة بنكية/);
  assert.match(result.checks.join(" "), /لا يتم إنشاء أو ترحيل قيود تلقائيًا/);
});

test("flags same-reference bank movements with different amounts", async () => {
  const bankFile = await csv("bank.csv", "التاريخ,المرجع,المبلغ\n2026-08-01,TRX-9,1000");
  const systemFile = await csv("erp.csv", "التاريخ,المرجع,المبلغ\n2026-08-01,TRX-9,900");
  const result = analyzeData("مطابقة كشف البنك", [bankFile, systemFile]);
  assert.equal(result.summary.find(item=>item.label==="مبالغ مختلفة")?.value, "1");
  assert.equal(result.table?.rows.find(row=>row[0]==="مبلغ مختلف")?.[5], "١٠٠٫٠٠ ر.س");
});

test("matches one bank deposit against several ERP movements", async () => {
  const bankFile = await csv("bank.csv", "التاريخ,المرجع,المبلغ\n2026-08-10,BATCH-1,1000");
  const systemFile = await csv("erp.csv", "التاريخ,المرجع,المبلغ\n2026-08-09,SALE-1,400\n2026-08-10,SALE-2,350\n2026-08-11,SALE-3,250");
  const result = analyzeData("مطابقة كشف البنك", [bankFile, systemFile]);
  assert.equal(result.summary.find(item=>item.label==="مطابقات مجمعة")?.value, "1");
  assert.equal(result.summary.find(item=>item.label==="بالبنك فقط")?.value, "0");
  assert.equal(result.summary.find(item=>item.label==="بالنظام فقط")?.value, "0");
  assert.equal(result.table?.rows.find(row=>row[0]==="مطابقة مجمعة")?.[7], "1 حركة بنك مقابل 3 حركة نظام");
});
