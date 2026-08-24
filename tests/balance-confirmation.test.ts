import test from "node:test";
import assert from "node:assert/strict";
import {
  amountToArabicWords,
  formatCurrencyAmount,
  parseCurrencyAmount,
  toIsoDate,
} from "../app/balance-confirmation-utils.ts";

test("reads the Arabic-formatted balance returned by the analyzer",()=>{
  assert.equal(parseCurrencyAmount("٥٣٬٤٢١٫١٥ ر.س"),53421.15);
  assert.equal(formatCurrencyAmount(53421.15),"53,421.15");
});

test("writes the balance in Arabic for the confirmation letter",()=>{
  assert.equal(amountToArabicWords(12000),"اثنا عشر ألفًا ريال");
  assert.match(amountToArabicWords(53421.15),/ثلاثة وخمسون ألفًا/);
  assert.match(amountToArabicWords(53421.15),/خمس عشرة هللة/);
});

test("uses the statement date when it is written with Arabic digits",()=>{
  assert.equal(toIsoDate("٢٤/٠٨/٢٠٢٦"),"2026-08-24");
});
