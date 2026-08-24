const arabicIndic = "٠١٢٣٤٥٦٧٨٩";
const easternArabicIndic = "۰۱۲۳۴۵۶۷۸۹";

export function toWesternDigits(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String(arabicIndic.indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String(easternArabicIndic.indexOf(digit)))
    .replace(/٫/g, ".")
    .replace(/٬/g, ",")
    .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "");
}

export function parseCurrencyAmount(value: string) {
  const normalized = toWesternDigits(value)
    .replace(/ر\.?\s*س\.?|SAR/gi, "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? Math.abs(parsed) : 0;
}

export function formatCurrencyAmount(value: number) {
  return new Intl.NumberFormat("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Math.abs(value));
}

const ones = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة"];
const teens: Record<number, string> = {
  10: "عشرة",
  11: "أحد عشر",
  12: "اثنا عشر",
  13: "ثلاثة عشر",
  14: "أربعة عشر",
  15: "خمسة عشر",
  16: "ستة عشر",
  17: "سبعة عشر",
  18: "ثمانية عشر",
  19: "تسعة عشر",
};
const tens: Record<number, string> = {
  20: "عشرون",
  30: "ثلاثون",
  40: "أربعون",
  50: "خمسون",
  60: "ستون",
  70: "سبعون",
  80: "ثمانون",
  90: "تسعون",
};
const feminineOnes = ["", "واحدة", "اثنتان", "ثلاث", "أربع", "خمس", "ست", "سبع", "ثمان", "تسع"];
const feminineTeens: Record<number, string> = {
  10: "عشر",
  11: "إحدى عشرة",
  12: "اثنتا عشرة",
  13: "ثلاث عشرة",
  14: "أربع عشرة",
  15: "خمس عشرة",
  16: "ست عشرة",
  17: "سبع عشرة",
  18: "ثماني عشرة",
  19: "تسع عشرة",
};
const hundreds: Record<number, string> = {
  100: "مائة",
  200: "مائتان",
  300: "ثلاثمائة",
  400: "أربعمائة",
  500: "خمسمائة",
  600: "ستمائة",
  700: "سبعمائة",
  800: "ثمانمائة",
  900: "تسعمائة",
};

function underHundred(value: number) {
  if (value < 10) return ones[value];
  if (value < 20) return teens[value];
  const unit = value % 10;
  const ten = value - unit;
  return unit ? `${ones[unit]} و${tens[ten]}` : tens[ten];
}

function feminineUnderHundred(value: number) {
  if (value < 10) return feminineOnes[value];
  if (value < 20) return feminineTeens[value];
  const unit = value % 10;
  const ten = value - unit;
  return unit ? `${feminineOnes[unit]} و${tens[ten]}` : tens[ten];
}

function underThousand(value: number) {
  if (value < 100) return underHundred(value);
  const hundred = Math.floor(value / 100) * 100;
  const rest = value % 100;
  return rest ? `${hundreds[hundred]} و${underHundred(rest)}` : hundreds[hundred];
}

function scaledGroup(value: number, singular: string, dual: string, plural: string, accusative: string) {
  if (value === 1) return singular;
  if (value === 2) return dual;
  if (value >= 3 && value <= 10) return `${underThousand(value)} ${plural}`;
  return `${underThousand(value)} ${accusative}`;
}

export function integerToArabicWords(input: number) {
  const value = Math.max(0, Math.floor(Math.abs(input)));
  if (value === 0) return "صفر";
  const parts: string[] = [];
  const billions = Math.floor(value / 1_000_000_000);
  const millions = Math.floor((value % 1_000_000_000) / 1_000_000);
  const thousands = Math.floor((value % 1_000_000) / 1_000);
  const remainder = value % 1_000;
  if (billions) parts.push(scaledGroup(billions, "مليار", "ملياران", "مليارات", "مليارًا"));
  if (millions) parts.push(scaledGroup(millions, "مليون", "مليونان", "ملايين", "مليونًا"));
  if (thousands) parts.push(scaledGroup(thousands, "ألف", "ألفان", "آلاف", "ألفًا"));
  if (remainder) parts.push(underThousand(remainder));
  return parts.join(" و");
}

export function amountToArabicWords(input: number) {
  const amount = Math.round(Math.abs(input) * 100) / 100;
  const riyals = Math.floor(amount);
  const halalas = Math.round((amount - riyals) * 100);
  const riyalText = `${integerToArabicWords(riyals)} ريال`;
  return halalas ? `${riyalText} و${feminineUnderHundred(halalas)} هللة` : riyalText;
}

export function todayIso() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export function toIsoDate(value?: string) {
  if (!value) return todayIso();
  const normalized = toWesternDigits(value).trim();
  const parts = normalized.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
  if (!parts) return todayIso();
  const first = Number(parts[1]);
  const second = Number(parts[2]);
  const third = Number(parts[3]);
  const year = first > 1900 ? first : third < 100 ? 2000 + third : third;
  const month = first > 1900 ? second : second;
  const day = first > 1900 ? third : first;
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return todayIso();
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

export function formatLetterDate(value: string) {
  const [year, month, day] = value.split("-");
  return year && month && day ? `${day}/${month}/${year}` : value;
}
