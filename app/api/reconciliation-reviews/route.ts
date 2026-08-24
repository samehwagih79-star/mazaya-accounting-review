import {env} from "cloudflare:workers";

const schemaSql=`CREATE TABLE IF NOT EXISTS reconciliation_reviews (
  id TEXT PRIMARY KEY,
  reconciliation_id TEXT NOT NULL,
  row_key TEXT NOT NULL,
  row_json TEXT NOT NULL,
  decision TEXT NOT NULL,
  note TEXT NOT NULL,
  proposed_entry TEXT NOT NULL,
  reviewer TEXT NOT NULL,
  created_at INTEGER NOT NULL
)`;
async function database(){const db=env.DB;if(!db)throw new Error("قاعدة بيانات الحفظ غير متاحة");await db.prepare(schemaSql).run();return db}

export async function GET(request:Request){try{const reconciliationId=new URL(request.url).searchParams.get("reconciliationId");if(!reconciliationId)return Response.json({items:[]});const db=await database(),rows=await db.prepare("SELECT id, reconciliation_id AS reconciliationId, row_key AS rowKey, row_json AS rowJson, decision, note, proposed_entry AS proposedEntry, reviewer, created_at AS createdAt FROM reconciliation_reviews WHERE reconciliation_id = ? ORDER BY created_at DESC").bind(reconciliationId).all();return Response.json({items:rows.results})}catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر تحميل سجل الاعتماد"},{status:500})}}

export async function POST(request:Request){try{const body=await request.json() as {reconciliationId?:string;rowKey?:string;row?:unknown;decision?:string;note?:string;proposedEntry?:string;reviewer?:string},reconciliationId=String(body.reconciliationId||"").trim(),rowKey=String(body.rowKey||"").trim(),decision=String(body.decision||"").trim(),reviewer=String(body.reviewer||"").trim();if(!reconciliationId||!rowKey||!decision||!reviewer||!body.row)return Response.json({error:"أكمل قرار الاعتماد واسم المراجع"},{status:400});const db=await database(),id=crypto.randomUUID(),createdAt=Date.now();await db.prepare("INSERT INTO reconciliation_reviews (id, reconciliation_id, row_key, row_json, decision, note, proposed_entry, reviewer, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)").bind(id,reconciliationId,rowKey,JSON.stringify(body.row),decision,String(body.note||""),String(body.proposedEntry||""),reviewer,createdAt).run();return Response.json({id,createdAt})}catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر حفظ قرار الاعتماد"},{status:500})}}
