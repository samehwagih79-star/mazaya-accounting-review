import {env} from "cloudflare:workers";

const schemaSql=`CREATE TABLE IF NOT EXISTS bank_reconciliations (
  id TEXT PRIMARY KEY,
  bank_name TEXT NOT NULL,
  account_name TEXT NOT NULL,
  period TEXT NOT NULL,
  result_json TEXT NOT NULL,
  created_at INTEGER NOT NULL
)`;

async function database(){const db=env.DB;if(!db)throw new Error("قاعدة بيانات الحفظ غير متاحة");await db.prepare(schemaSql).run();return db}

export async function GET(){try{const db=await database(),rows=await db.prepare("SELECT id, bank_name AS bankName, account_name AS accountName, period, result_json AS resultJson, created_at AS createdAt FROM bank_reconciliations ORDER BY created_at DESC LIMIT 60").all();return Response.json({items:rows.results})}catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر تحميل التقارير"},{status:500})}}

export async function POST(request:Request){try{const body=await request.json() as {bankName?:string;accountName?:string;period?:string;result?:unknown},bankName=String(body.bankName||"").trim(),accountName=String(body.accountName||"").trim(),period=String(body.period||"").trim();if(!bankName||!accountName||!/^[0-9]{4}-[0-9]{2}$/.test(period)||!body.result)return Response.json({error:"أكمل اسم البنك والحساب والفترة"},{status:400});const db=await database(),id=crypto.randomUUID(),createdAt=Date.now(),resultJson=JSON.stringify(body.result);await db.prepare("INSERT INTO bank_reconciliations (id, bank_name, account_name, period, result_json, created_at) VALUES (?, ?, ?, ?, ?, ?)").bind(id,bankName,accountName,period,resultJson,createdAt).run();return Response.json({id,bankName,accountName,period,createdAt})}catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر حفظ التقرير"},{status:500})}}

export async function DELETE(request:Request){try{const id=new URL(request.url).searchParams.get("id");if(!id)return Response.json({error:"معرف التقرير مطلوب"},{status:400});const db=await database();await db.prepare("DELETE FROM bank_reconciliations WHERE id = ?").bind(id).run();return Response.json({ok:true})}catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر حذف التقرير"},{status:500})}}
