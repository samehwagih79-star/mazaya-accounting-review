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
