import test from "node:test";
import assert from "node:assert/strict";
import { File } from "node:buffer";
import { analyzeData, readAccountingFile } from "../app/analyzer.ts";

async function csv(name:string,text:string){return readAccountingFile(new File([text],name,{type:"text/csv"}) as unknown as globalThis.File)}

test("locks the 45-day actual-due rule for suppliers",async()=>{
  const data=await csv("supplier.csv","التاريخ,البيان,مدين,دائن\n2026-01-01,فاتورة مشتريات,0,1000\n2026-03-15,فاتورة مشتريات,0,500\n2026-03-20,سند صرف,200,0");
  const result=analyzeData("تحليل حساب مورد",[data],45);
  assert.equal(result.conclusion?.label,"الرصيد الفعلي المستحق للمورد");
  assert.equal(result.conclusion?.value,"٨٠٠٫٠٠ ر.س");
  assert.match(result.conclusion?.detail??"",/لا يدخل الرصيد غير المستحق/);
  assert.equal(result.dueScenarios?.items.find(item=>item.days===45)?.value,"٨٠٠٫٠٠ ر.س");
  assert.deepEqual(result.dueScenarios?.items.map(item=>item.days),[30,45,60,90,120]);
});

test("locks the 45-day actual-due rule for customers",async()=>{
  const data=await csv("customer.csv","التاريخ,البيان,مدين,دائن\n2026-01-01,فاتورة مبيعات,1000,0\n2026-03-15,فاتورة مبيعات,500,0\n2026-03-20,سند قبض,0,200");
  const result=analyzeData("تحليل حساب عميل",[data],45);
  assert.equal(result.conclusion?.label,"الرصيد الفعلي المستحق على العميل");
  assert.equal(result.conclusion?.value,"٨٠٠٫٠٠ ر.س");
  assert.match(result.conclusion?.detail??"",/لا يدخل الرصيد غير المستحق/);
  assert.equal(result.dueScenarios?.items.find(item=>item.days===45)?.value,"٨٠٠٫٠٠ ر.س");
});

test("calculates 30, 45 and 60 days independently from invoice issue dates",async()=>{
  const data=await csv("supplier-terms.csv","التاريخ,تاريخ الاستحقاق,البيان,مدين,دائن\n2026-01-20,2026-12-31,فاتورة مشتريات,0,1000\n2026-02-10,2026-12-31,فاتورة مشتريات,0,1000\n2026-02-25,2026-12-31,فاتورة مشتريات,0,1000\n2026-03-20,,مردود مشتريات,200,0\n2026-03-31,,سند صرف,100,0");
  const result=analyzeData("تحليل حساب مورد",[data],45);
  const scenarios=Object.fromEntries((result.dueScenarios?.items??[]).map(item=>[item.days,item.value]));
  assert.equal(scenarios[30],"٢٬٧٠٠٫٠٠ ر.س");
  assert.equal(scenarios[45],"١٬٧٠٠٫٠٠ ر.س");
  assert.equal(scenarios[60],"٧٠٠٫٠٠ ر.س");
  assert.equal(result.conclusion?.value,"١٬٧٠٠٫٠٠ ر.س");
  assert.equal(result.dueTable?.rows[0][3],"تاريخ الفاتورة + 45 يومًا");
  assert.match(result.findings.map(item=>item.detail).join(" "),/تاريخ إصدارها \+ 45 يومًا/);
});

test("deducts customer returns and receipts before calculating the due balance",async()=>{
  const data=await csv("customer-due.csv","التاريخ,تاريخ الاستحقاق,البيان,مدين,دائن\n2026-01-01,2026-02-01,فاتورة مبيعات,1000,0\n2026-03-01,,مرتجع مبيعات,0,200\n2026-03-20,,سند قبض,0,100");
  const result=analyzeData("تحليل حساب عميل",[data],45);
  assert.equal(result.conclusion?.value,"٧٠٠٫٠٠ ر.س");
});

test("includes the August 19 invoice and excludes the September 6 invoice at the August 24 cutoff",async()=>{
  const data=await csv("supplier-cutoff.csv","التاريخ,البيان,مدين,دائن\n2026-07-05,فاتورة مشتريات,0,1000\n2026-07-23,فاتورة مشتريات,0,1000\n2026-08-24,تاريخ القياس,0,0");
  const result=analyzeData("تحليل حساب مورد",[data],45);
  assert.equal(result.conclusion?.value,"١٬٠٠٠٫٠٠ ر.س");
  assert.match(result.dueSchedule?.includedThrough??"",/١٩.*٠٨.*٢٠٢٦/);
  assert.match(result.dueSchedule?.nextDueDate??"",/٠٦.*٠٩.*٢٠٢٦/);
  assert.equal(result.dueSchedule?.nextDueAmount,"١٬٠٠٠٫٠٠ ر.س");
});

test("counts an invoice on its exact due date",async()=>{
  const data=await csv("supplier-exact-date.csv","التاريخ,البيان,مدين,دائن\n2026-07-10,فاتورة مشتريات,0,900\n2026-08-24,تاريخ القياس,0,0");
  const result=analyzeData("تحليل حساب مورد",[data],45);
  assert.equal(result.conclusion?.value,"٩٠٠٫٠٠ ر.س");
});

test("uses the running balance at the latest due supplier invoice then deducts later purchase returns and payments",async()=>{
  const data=await csv("supplier-running-balance.csv","التاريخ,البيان,مدين,دائن,الرصيد\n2026-07-05,فاتورة مشتريات اجل,0,1127,53421.15\n2026-07-23,فاتورة مشتريات اجل,0,1127,54548.15\n2026-07-25,مردود مشتريات اجل,3191.25,0,51356.90\n2026-08-05,سند صرف,5000,0,46356.90\n2026-08-11,فاتورة مشتريات اجل,0,437,46793.90\n2026-08-15,فاتورة مشتريات اجل,0,2127.50,48921.40\n2026-08-18,سند صرف,10000,0,38921.40\n2026-08-24,تاريخ القياس,0,0,38921.40");
  const result=analyzeData("تحليل حساب مورد",[data],45);
  assert.equal(result.dueSchedule?.cutoffInvoiceDate,"٠٥‏/٠٧‏/٢٠٢٦");
  assert.equal(result.dueSchedule?.cutoffBalance,"٥٣٬٤٢١٫١٥ ر.س");
  assert.equal(result.dueSchedule?.laterDeductions,"١٨٬١٩١٫٢٥ ر.س");
  assert.equal(result.conclusion?.value,"٣٥٬٢٢٩٫٩٠ ر.س");
  assert.notEqual(result.dueScenarios?.items.find(item=>item.days===30)?.value,result.conclusion?.value);
});

test("deducts only later sales returns and receipts for customer accounts",async()=>{
  const data=await csv("customer-running-balance.csv","التاريخ,البيان,مدين,دائن,الرصيد\n2026-07-05,فاتورة مبيعات اجل,53421.15,0,53421.15\n2026-07-23,فاتورة مبيعات اجل,1127,0,54548.15\n2026-07-25,مردود مبيعات اجل,0,3191.25,51356.90\n2026-08-05,سند قبض,0,5000,46356.90\n2026-08-11,فاتورة مبيعات اجل,437,0,46793.90\n2026-08-18,سند قبض,0,10000,36793.90\n2026-08-24,تاريخ القياس,0,0,36793.90");
  const result=analyzeData("تحليل حساب عميل",[data],45);
  assert.equal(result.conclusion?.value,"٣٥٬٢٢٩٫٩٠ ر.س");
});
