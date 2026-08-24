"use client";
/* eslint-disable @next/next/no-img-element */

import {useEffect,useMemo,useState} from "react";
import type {AnalysisResult} from "./analyzer";
import {
  amountToArabicWords,
  formatCurrencyAmount,
  formatLetterDate,
  parseCurrencyAmount,
  toIsoDate,
} from "./balance-confirmation-utils";

type AccountKind="customer"|"supplier";
type BalanceChoice={id:"closing"|"due";label:string;value:number};

function inferAccountKind(result:AnalysisResult):AccountKind{
  return /مورد/.test(`${result.title} ${result.conclusion?.label||""}`)?"supplier":"customer";
}

function suggestedPartyName(result:AnalysisResult){
  const source=result.sources[0]||"";
  const suggestion=source
    .replace(/\.[^.]+$/g,"")
    .replace(/(?:كشف|حساب|statement|ledger|عميل|مورد|customer|supplier)/gi," ")
    .replace(/[_-]+/g," ")
    .replace(/\s+/g," ")
    .trim();
  return suggestion.length>=3?suggestion:"";
}

function balanceChoices(result:AnalysisResult):BalanceChoice[]{
  const summaryClosing=result.summary.find(item=>/الرصيد الختامي|صافي رصيد المورد حسب الكشف|إجمالي المستحق على العميل/.test(item.label));
  const closingText=result.aging?.total||summaryClosing?.value||"";
  const dueText=result.conclusion?.value||"";
  const closing=parseCurrencyAmount(closingText);
  const due=parseCurrencyAmount(dueText);
  const choices:BalanceChoice[]=[];
  if(/[0-9٠-٩۰-۹]/.test(closingText))choices.push({id:"closing",label:"الرصيد الختامي للمطابقة",value:closing});
  if(/[0-9٠-٩۰-۹]/.test(dueText)&&Math.abs(due-closing)>.009)choices.push({id:"due",label:"الرصيد المستحق الفعلي فقط",value:due});
  if(!choices.length&&/[0-9٠-٩۰-۹]/.test(dueText))choices.push({id:"due",label:"الرصيد الظاهر في نتيجة التحليل",value:due});
  return choices;
}

function BalanceConfirmationDialog({result,onClose}:{result:AnalysisResult;onClose:()=>void}){
  const choices=useMemo(()=>balanceChoices(result),[result]);
  const initialAmount=choices[0]?.value||0;
  const [kind,setKind]=useState<AccountKind>(()=>inferAccountKind(result));
  const [partyName,setPartyName]=useState(()=>suggestedPartyName(result));
  const [letterDate,setLetterDate]=useState(()=>toIsoDate(result.dueSchedule?.asOf||result.aging?.asOf));
  const [amount,setAmount]=useState(()=>formatCurrencyAmount(initialAmount));
  const [amountWords,setAmountWords]=useState(()=>amountToArabicWords(initialAmount));
  const [selectedBalance,setSelectedBalance]=useState(choices[0]?.id||"closing");
  const numericAmount=parseCurrencyAmount(amount);
  const displayName=partyName.trim()||"........................................................";
  const displayDate=formatLetterDate(letterDate);
  const isSupplier=kind==="supplier";
  const titleNature=isSupplier?"دائنة":"مدينة";
  const balanceNature=isSupplier?"رصيد دائن باسمكم لدى الشركة":"رصيد مدين للشركة";

  useEffect(()=>{
    const closeOnEscape=(event:KeyboardEvent)=>{if(event.key==="Escape")onClose()};
    const cleanupPrint=()=>document.body.classList.remove("printing-confirmation");
    window.addEventListener("keydown",closeOnEscape);
    window.addEventListener("afterprint",cleanupPrint);
    return()=>{
      window.removeEventListener("keydown",closeOnEscape);
      window.removeEventListener("afterprint",cleanupPrint);
      cleanupPrint();
    };
  },[onClose]);

  function updateAmount(value:string){
    setAmount(value);
    setAmountWords(amountToArabicWords(parseCurrencyAmount(value)));
  }

  function selectBalance(id:string){
    const choice=choices.find(item=>item.id===id);
    setSelectedBalance(id as "closing"|"due");
    if(choice)updateAmount(formatCurrencyAmount(choice.value));
  }

  function printLetter(){
    const previousTitle=document.title;
    document.title=`مطابقة رصيد - ${partyName.trim()||"عميل أو مورد"}`;
    document.body.classList.add("printing-confirmation");
    window.setTimeout(()=>{
      window.print();
      window.setTimeout(()=>{document.body.classList.remove("printing-confirmation");document.title=previousTitle},300);
    },50);
  }

  return <div className="confirmation-modal" role="dialog" aria-modal="true" aria-labelledby="confirmation-dialog-title">
    <div className="confirmation-dialog">
      <div className="confirmation-toolbar">
        <div><span>خطاب رسمي من نتيجة التحليل</span><h2 id="confirmation-dialog-title">إنشاء خطاب مطابقة الأرصدة</h2><p>راجع اسم الطرف والمبلغ ثم اطبع الخطاب أو اختر حفظ بصيغة PDF.</p></div>
        <div className="confirmation-toolbar-actions"><button className="confirmation-print" onClick={printLetter}>طباعة / حفظ PDF</button><button className="confirmation-close" onClick={onClose} aria-label="إغلاق">إغلاق ×</button></div>
      </div>
      <div className="confirmation-workspace">
        <aside className="confirmation-form">
          <h3>بيانات الخطاب</h3>
          <p>تم جلب التاريخ والرصيد من التقرير. راجعهما قبل الاعتماد.</p>
          <label>نوع الحساب</label>
          <div className="confirmation-kind" role="group" aria-label="نوع الحساب">
            <button className={kind==="customer"?"active":""} onClick={()=>setKind("customer")}>عميل</button>
            <button className={kind==="supplier"?"active":""} onClick={()=>setKind("supplier")}>مورد</button>
          </div>
          <label htmlFor="confirmation-party">اسم العميل أو المورد</label>
          <input id="confirmation-party" value={partyName} onChange={event=>setPartyName(event.target.value)} placeholder="مثال: شركة ميثاق العالمية للموارد البشرية"/>
          <label htmlFor="confirmation-date">تاريخ مطابقة الرصيد</label>
          <input id="confirmation-date" type="date" value={letterDate} onChange={event=>setLetterDate(event.target.value)}/>
          {choices.length>1&&<><label htmlFor="confirmation-source">مبلغ الخطاب</label><select id="confirmation-source" value={selectedBalance} onChange={event=>selectBalance(event.target.value)}>{choices.map(choice=><option key={choice.id} value={choice.id}>{choice.label} - {formatCurrencyAmount(choice.value)} ر.س</option>)}</select></>}
          <label htmlFor="confirmation-amount">الرصيد بالأرقام (ر.س)</label>
          <input id="confirmation-amount" inputMode="decimal" value={amount} onChange={event=>updateAmount(event.target.value)}/>
          <label htmlFor="confirmation-words">الرصيد كتابةً</label>
          <textarea id="confirmation-words" rows={3} value={amountWords} onChange={event=>setAmountWords(event.target.value)}/>
          <div className="confirmation-note"><b>تنبيه:</b> مطابقة الأرصدة تعتمد الرصيد الختامي افتراضيًا، وليس رصيد الفواتير المستحق بعد مدة الائتمان، ويمكنك تغييره من القائمة إذا احتجت.</div>
        </aside>
        <div className="confirmation-preview" aria-label="معاينة خطاب مطابقة الأرصدة">
          <article className="confirmation-sheet" dir="rtl">
            <img className="confirmation-letterhead" src="/mazaya-letterhead.jpg" alt="" aria-hidden="true"/>
            <div className="confirmation-letter-content">
              <div className="letter-date"><b>في:</b> {displayDate}</div>
              <h1>خطاب مطابقة أرصدة ({titleNature})</h1>
              <div className="letter-recipient"><strong>السادة / {displayName}</strong><span>المحترمين</span></div>
              <p className="letter-greeting">تحية طيبة وبعد،،،،،</p>
              <h2>طلب مصادقة حساب:</h2>
              <p>تهديكم الشركة أطيب تحياتها وتتمنى لكم دوام التقدم والازدهار،،،</p>
              <p>بناءً على طلب المراقب المالي للشركة نرجو شاكرين تكرمكم بالمصادقة على <b>{balanceNature}</b> في <b>{displayDate}</b>.</p>
              <p>والبالغ <b>({formatCurrencyAmount(numericAmount)})</b> (فقط <b>{amountWords||amountToArabicWords(numericAmount)}</b> فقط لا غير).</p>
              <p>كما يرجى توقيع قسيمة التأييد والمطابقة على الرصيد أدناه أو الإفادة بأية اختلافات إن وجدت، وإرسالها إلينا على البريد الإلكتروني <b>sales@mazayaksa.com</b> أو مناولة يدوية إلى مندوبنا لديكم.</p>
              <p>وإذا لم يرد إلينا ردكم خلال عشرة أيام من تاريخ التسليم فيعتبر رصيد حسابكم معنا صحيحًا ومصادقًا عليه نهائيًا من قبلكم دون أي إنذار.</p>
              <p className="letter-thanks">شاكرين حسن تعاونكم معنا،،،،،</p>
              <div className="letter-divider"/>
              <section className="confirmation-reply">
                <h3>السادة: شركة مزايا لتقنية المعلومات</h3>
                <p>تحية طيبة وبعد،،،،،</p>
                <p>إن رصيد حسابنا لدى شركة مزايا لتقنية المعلومات</p>
                <p className="reply-balance">البالغ: (............................) فقط (........................................................ ريال لا غير)</p>
                <p>صحيحًا في <b>{displayDate}</b> م.</p>
                <p><b>فيما عدا الاختلافات التالية (المرفقة الأسباب):</b></p>
                <p className="difference-line">1- ................................................................................................................</p>
                <p className="difference-line">2- ................................................................................................................</p>
                <div className="signature-row"><span><b>الاسم:</b> ............................</span><span><b>التوقيع:</b> ............................</span><span><b>الختم:</b> ............................</span></div>
              </section>
            </div>
          </article>
        </div>
      </div>
    </div>
  </div>;
}

export function BalanceConfirmationButton({result}:{result:AnalysisResult}){
  const [open,setOpen]=useState(false);
  const balanceText=result.aging?.total||result.conclusion?.value||"";
  const eligible=/عميل|مورد/.test(`${result.title} ${result.conclusion?.label||""}`)&&/[0-9٠-٩۰-۹]/.test(balanceText)&&!/غير مؤكد/.test(balanceText);
  if(!eligible)return null;
  return <><button className="confirmation-trigger" onClick={()=>setOpen(true)}>خطاب مطابقة الرصيد</button>{open&&<BalanceConfirmationDialog result={result} onClose={()=>setOpen(false)}/>}</>;
}
