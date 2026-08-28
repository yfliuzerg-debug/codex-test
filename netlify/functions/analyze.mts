import type { Context, Config } from "@netlify/functions";
import * as XLSX from "xlsx";

type Row = Record<string, any>;
type SoRow = {channel:string, customerCode:string, customerName:string, segment1:string, segment2:string, sku:string, skuName:string, months:Record<string,number>};
type StockRow = {customerCode:string, customerName:string, sku:string, skuName:string, cases:number};
type TtsRow = {customerName:string, customerCode:string, channel:string, sku:string, start:Date|null, end:Date|null, rate:number, cp:string};
type CpRow = {code:string, start:Date|null, end:Date|null, description:string, mechanism:string, channels:string, balance:number, type:string};

const EXCLUDE_SKUS = new Set(["12187803","12585787","12598186","12611209"]);

const esc = (v:any) => String(v ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const num = (v:any) => { const n=Number(v); return Number.isFinite(n)?n:0; };
const money = (v:number) => new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0}).format(v||0);
const normalize = (s:any) => String(s??"").replace(/\s+/g,"").replace(/有限责任公司$/,"有限公司").toLowerCase();
const monthKey=(d:Date)=>`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}`;
const monthRange=(a:Date|null,b:Date|null)=>{ if(!a||!b)return[] as string[]; const out:string[]=[]; const d=new Date(a.getFullYear(),a.getMonth(),1); const e=new Date(b.getFullYear(),b.getMonth(),1); while(d<=e){out.push(monthKey(d));d.setMonth(d.getMonth()+1);} return out; };
const parseDate=(v:any):Date|null=>{ if(!v)return null; if(v instanceof Date && !Number.isNaN(v.getTime()))return v; if(typeof v==="number"){const o=XLSX.SSF.parse_date_code(v); return o?new Date(o.y,o.m-1,o.d):null;} const d=new Date(String(v)); return Number.isNaN(d.getTime())?null:d; };

function readWorkbook(file: File){return file.arrayBuffer().then(ab=>XLSX.read(Buffer.from(ab),{type:"buffer",cellDates:true,raw:true}));}
function sheetRows(wb:any,name?:string){const sn=name&&wb.SheetNames.includes(name)?name:wb.SheetNames[0]; return XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null,raw:true});}
function findHeader(rows:any[][], required:string[]){return rows.findIndex(r=>required.every(k=>r.some(c=>String(c??"").trim()===k)));}
function rowObj(headers:any[], row:any[]):Row{const o:Row={}; headers.forEach((h,i)=>{if(h!=null&&String(h).trim())o[String(h).trim()]=row[i];}); return o;}

function parseBiSo(wb:any):SoRow[]{
  if(!wb.SheetNames.includes("BI-SO(折扣)"))return[];
  const rows=sheetRows(wb,"BI-SO(折扣)"); const hi=findHeader(rows,["Outlet Channel","Customer Code","SKU"]); if(hi<0)return[];
  const h=rows[hi].map(x=>String(x??"").trim()); const monthCols=h.map((x,i)=>/^20\d{4}$/.test(x)?[x,i]:null).filter(Boolean) as [string,number][];
  const idx=(n:string)=>h.indexOf(n);
  return rows.slice(hi+1).map(r=>({channel:String(r[idx("Outlet Channel")]??""),customerCode:String(r[idx("Customer Code")]??""),customerName:String(r[idx("Customer Name")]??""),segment1:String(r[idx("PH5 Segment1_MKT")]??""),segment2:String(r[idx("PH5 Segment2_Detail")]??""),sku:String(r[idx("SKU")]??""),skuName:String(r[idx("SKU Name CN")]??""),months:Object.fromEntries(monthCols.map(([m,i])=>[m,num(r[i])]))})).filter(x=>x.customerName&&x.sku&&!/汇总/.test(x.customerCode));
}
function parseStock(wb:any):StockRow[]{
  if(!wb.SheetNames.includes("BI-Stock"))return[]; const rows=sheetRows(wb,"BI-Stock"); const hi=findHeader(rows,["Customer Code","Customer Name","SKU","SOH (Good) Case"]); if(hi<0)return[]; const h=rows[hi].map(x=>String(x??"").trim()), idx=(n:string)=>h.indexOf(n);
  return rows.slice(hi+1).map(r=>({customerCode:String(r[idx("Customer Code")]??""),customerName:String(r[idx("Customer Name")]??""),sku:String(r[idx("SKU")]??""),skuName:String(r[idx("SKU Name CN")]??""),cases:num(r[idx("SOH (Good) Case")])})).filter(x=>x.customerName&&x.sku&&x.cases!==0);
}
function parsePrices(wb:any):Map<string,number>{
  const out=new Map<string,number>();
  for(const sn of wb.SheetNames){const rows=sheetRows(wb,sn); const hi=findHeader(rows,["产品编码"]); if(hi<0)continue; const h=rows[hi].map(x=>String(x??"").trim()); const skuI=h.indexOf("产品编码"); let priceI=h.findIndex(x=>x.includes("NPS")&&x.includes("未税箱价")); if(priceI<0)priceI=h.findIndex(x=>x==="Price"); for(const r of rows.slice(hi+1)){const sku=String(r[skuI]??"").trim(); const p=num(r[priceI]); if(sku&&p>0)out.set(sku,p);} }
  return out;
}
function parseTts(wb:any):TtsRow[]{
  const rows=sheetRows(wb,wb.SheetNames.includes("PPCenturyBigTable")?"PPCenturyBigTable":undefined); const hi=findHeader(rows,["PPCode","Distributor Name","Activity Code in Cycle Plan"]); if(hi<0)return[]; const h=rows[hi].map(x=>String(x??"").trim()), idx=(n:string)=>h.indexOf(n);
  return rows.slice(hi+1).map(r=>{let rate=num(r[idx("Disc.Rate")]); if(rate>1)rate/=100; return {customerName:String(r[idx("Distributor Name")]??""),customerCode:String(r[idx("Distributor Code")]??""),channel:String(r[idx("Outlet Channel")]??r[idx("Channel")]??""),sku:String(r[idx("Product Code")]??""),start:parseDate(r[idx("Promotion Start Date")]),end:parseDate(r[idx("Promotion End Date")]),rate,cp:String(r[idx("Activity Code in Cycle Plan")]??"")};}).filter(x=>x.customerName&&x.sku&&x.rate>0);
}
function parseCp(wb:any):CpRow[]{
  const rows=sheetRows(wb,wb.SheetNames.includes("CPCenturyBigTableByCGAmount")?"CPCenturyBigTableByCGAmount":undefined); const hi=findHeader(rows,["Activity Code","Activity Start Date","Balance Amount"]); if(hi<0)return[]; const h=rows[hi].map(x=>String(x??"").trim()), idx=(n:string)=>h.indexOf(n);
  return rows.slice(hi+1).map(r=>({code:String(r[idx("Activity Code")]??""),start:parseDate(r[idx("Activity Start Date")]),end:parseDate(r[idx("Activity End Date")]),description:String(r[idx("Activity Description")]??""),mechanism:String(r[idx("Activity Mechanism")]??""),channels:String(r[idx("SubChannel")]??""),balance:num(r[idx("Balance Amount")]),type:String(r[idx("PromotionType")]??r[idx("Activity Type")]??"")})).filter(x=>x.code&&x.balance>0);
}
function channelAllowed(cp:CpRow,ch:string){const c=cp.channels.toLowerCase(), x=String(ch??"").toLowerCase(); if(!c)return true; if(c.includes(x))return true; const aliases:Record<string,string[]>={"ws":["wholesale"],"wholesale":["wholesale"],"sd1":["sd1","sub-d"],"mm-a":["mm-a"],"mm-b":["mm-b"],"mm-c":["mm-c"],"b2b":["b2b"],"b2c":["b2c"]}; return (aliases[x]||[]).some(a=>c.includes(a));}
function rowEligible(cp:CpRow,r:SoRow){if(EXCLUDE_SKUS.has(r.sku))return false; if(!channelAllowed(cp,r.channel))return false; const t=(cp.description+" "+cp.mechanism).toLowerCase(); const s=(r.segment1+" "+r.segment2+" "+r.skuName).toLowerCase(); const tests:[RegExp,string[]][]=[[/乳铁|免疫/, ["乳铁","immune"]],[/蓝绿/, ["senior blue","senior green","蓝","绿"]],[/中老年/, ["senior"]],[/n3/i,["n3"]],[/全家/, ["family","全家","fcmp"]],[/炼奶/, ["culinary","炼奶","eagle"]],[/学生/, ["child","学生"]]]; for(const [re,keys] of tests){if(re.test(t)&&!keys.some(k=>s.includes(k)))return false;} return true;}
function rateFor(cp:CpRow,r:SoRow):number|null{
  const t=(cp.description+" "+cp.mechanism).replace(/：/g,":");
  const pick=(re:RegExp)=>{const m=t.match(re); return m?Number(m[1])/100:null;};
  const s=(r.segment1+" "+r.segment2+" "+r.skuName).toLowerCase();
  if(s.includes("senior")){const x=pick(/中老年[^%]{0,30}?(\d+(?:\.\d+)?)%/i); if(x!=null)return x;}
  if(s.includes("n3")){const x=pick(/N3[^%]{0,30}?(\d+(?:\.\d+)?)%/i); if(x!=null)return x;}
  if(s.includes("family")||s.includes("全家")||s.includes("fcmp")){const x=pick(/全家[^%]{0,30}?(\d+(?:\.\d+)?)%/i); if(x!=null)return x;}
  const all=[...t.matchAll(/(\d+(?:\.\d+)?)\s*%/g)].map(m=>Number(m[1])/100); const uniq=[...new Set(all)]; return uniq.length===1?uniq[0]:null;
}
function overlapRate(tts:TtsRow[], customer:string, sku:string, month:string, ch:string){let mx=0; for(const p of tts){if(normalize(p.customerName)!==normalize(customer)||p.sku!==sku)continue; if(p.channel&&ch&&!normalize(p.channel).includes(normalize(ch))&&!normalize(ch).includes(normalize(p.channel)))continue; if(!monthRange(p.start,p.end).includes(month))continue; mx=Math.max(mx,p.rate);} return mx;}
function septemberForecast(r:SoRow){const y25=r.months["202509"]||0; const avg=((r.months["202605"]||0)+(r.months["202606"]||0)+(r.months["202607"]||0))/3; return Math.max(y25,(y25+avg)/2);}

function page(title:string,body:string){return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.wrap{max-width:1500px;margin:auto;padding:24px}.top{display:flex;justify-content:space-between;align-items:center}.top a{color:#1769ff;text-decoration:none}.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:16px 0}.card,.panel{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:14px}.card b{font-size:22px;display:block}.card span,.muted{color:#6b7688}.panel{margin:14px 0;overflow:auto}table{border-collapse:collapse;width:100%;min-width:950px}th,td{border-bottom:1px solid #edf0f5;padding:9px 8px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8fafc;font-size:12px}.num{text-align:right}.tag{display:inline-block;padding:2px 7px;border-radius:999px;background:#eef5ff;color:#1769ff;font-size:12px}.warn{background:#fff8e8;border:1px solid #f0d89a;padding:10px 12px;border-radius:9px}.ok{background:#ecfbf3;border:1px solid #bde8cf;padding:10px 12px;border-radius:9px}.btn{display:inline-block;border:0;background:#1769ff;color:#fff;padding:9px 13px;border-radius:8px;text-decoration:none}@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.wrap{padding:14px}}</style></head><body><div class="wrap"><div class="top"><div><h2 style="margin:0">${esc(title)}</h2><div class="muted">Netlify 服务端计算 · 无前端 JavaScript 依赖</div></div><a href="/">重新上传</a></div>${body}</div></body></html>`;}

export default async (req: Request, context: Context) => {
  if(req.method!=="POST")return new Response("Method Not Allowed",{status:405});
  try{
    const fd=await req.formData(); const bi=fd.get("bi"),ttsF=fd.get("tts"),cpF=fd.get("cp"),product=fd.get("product");
    if(!(bi instanceof File)||!(ttsF instanceof File)||!(cpF instanceof File)||!(product instanceof File))return new Response(page("文件不完整",`<div class="warn">必须上传 BI、TTS、CP预算、产品目录四个文件。</div>`),{status:400,headers:{"content-type":"text/html; charset=utf-8"}});
    const [biW,ttsW,cpW,prodW]=await Promise.all([readWorkbook(bi),readWorkbook(ttsF),readWorkbook(cpF),readWorkbook(product)]);
    const so=parseBiSo(biW), stock=parseStock(biW), tts=parseTts(ttsW), cps=parseCp(cpW), prices=parsePrices(prodW);
    if(!prices.size){for(const [k,v] of parsePrices(biW))prices.set(k,v);}
    const focus=String(fd.get("focus")??"").split(/[,，]/).map(s=>s.trim()).filter(Boolean); const minBalance=num(fd.get("minBalance")??500);
    const now=new Date(); const cur=monthKey(now); const weight=now.getDate()<=10?0.10:now.getDate()<=20?0.20:0.30;
    const clientMap=new Map<string,{name:string,so:number,stock:number}>();
    for(const r of so){const k=normalize(r.customerName); const o=clientMap.get(k)||{name:r.customerName,so:0,stock:0}; o.so+=(r.months[cur]||0)*1000; clientMap.set(k,o);}
    for(const r of stock){const k=normalize(r.customerName); const o=clientMap.get(k)||{name:r.customerName,so:0,stock:0}; o.stock+=r.cases*(prices.get(r.sku)||0); clientMap.set(k,o);}
    let clients=[...clientMap.values()]; if(focus.length)clients=clients.filter(c=>focus.some(f=>normalize(c.name).includes(normalize(f)))); else clients.sort((a,b)=>(b.so+b.stock)-(a.so+a.stock)); clients=clients.slice(0,20);
    const activeCps=cps.filter(c=>c.balance>=minBalance).sort((a,b)=>b.balance-a.balance);
    const matrix:any[]=[];
    for(const c of clients){
      for(const cp of activeCps){let grossBasis=0, fee=0, hasRate=false, unknownRate=false, overlap=0; const months=monthRange(cp.start,cp.end);
        for(const r of so){if(normalize(r.customerName)!==normalize(c.name)||!rowEligible(cp,r))continue; for(const m of months){let base=(r.months[m]||0)*1000; if(m===cur){const st=stock.filter(x=>normalize(x.customerName)===normalize(c.name)&&x.sku===r.sku).reduce((a,x)=>a+x.cases*(prices.get(x.sku)||0),0); base+=st*weight;} else if(m==="202609" && !r.months[m]) base=septemberForecast(r)*1000; if(!base)continue; grossBasis+=base; const rr=rateFor(cp,r); if(rr==null){unknownRate=true;continue;} hasRate=true; const ov=overlapRate(tts,c.name,r.sku,m,r.channel); overlap+=base*Math.min(rr,ov); fee+=base*Math.max(0,rr-ov); }}
        matrix.push({client:c.name,cp,basis:grossBasis,fee:hasRate?fee:null,unknownRate,overlap}); }
    }
    const cards=`<div class="cards"><div class="card"><span>BI SO明细</span><b>${money(so.length)}</b></div><div class="card"><span>库存SKU记录</span><b>${money(stock.length)}</b></div><div class="card"><span>TTS折扣计划</span><b>${money(tts.length)}</b></div><div class="card"><span>可用CP</span><b>${money(activeCps.length)}</b></div><div class="card"><span>目录SKU</span><b>${money(prices.size)}</b></div></div>`;
    const note=`<div class="ok">后端解析成功。当前月库存只作用于“剩余承载”估算：${now.getDate()}日对应固定权重 ${(weight*100).toFixed(0)}%；历史已实现SO不改写。同客户/渠道/SKU/月若存在折扣重叠，按已存在计划的最高折扣率扣减。</div>`;
    const ctable=`<div class="panel"><h3>客户快照</h3><table><thead><tr><th>客户</th><th class="num">本月已实现SO</th><th class="num">当前库存NPS估值</th><th class="num">本月库存计入承载</th></tr></thead><tbody>${clients.map(c=>`<tr><td>${esc(c.name)}</td><td class="num">${money(c.so)}</td><td class="num">${money(c.stock)}</td><td class="num">${money(c.stock*weight)}</td></tr>`).join("")}</tbody></table></div>`;
    const cpTable=`<div class="panel"><h3>可用CP</h3><table><thead><tr><th>CP</th><th>活动</th><th>有效期</th><th>机制</th><th>渠道</th><th class="num">余额</th></tr></thead><tbody>${activeCps.map(c=>`<tr><td><b>${esc(c.code)}</b></td><td>${esc(c.description)}</td><td>${esc(c.start?.toLocaleDateString("zh-CN")||"")} → ${esc(c.end?.toLocaleDateString("zh-CN")||"")}</td><td>${esc(c.mechanism)}</td><td>${esc(c.channels)}</td><td class="num"><b>${money(c.balance)}</b></td></tr>`).join("")}</tbody></table></div>`;
    const mtable=`<div class="panel"><h3>客户 × CP 承载候选</h3><div class="muted">SO承载基数按CP有效期和渠道筛选；9月缺实际值时使用既定预测公式。费率无法从CP机制唯一识别时不硬算金额。</div><table><thead><tr><th>客户</th><th>CP</th><th>活动</th><th class="num">CP余额</th><th class="num">可见SO/库存承载基数</th><th class="num">重叠折扣扣减</th><th class="num">估算可承载费用</th><th>状态</th></tr></thead><tbody>${matrix.map(x=>`<tr><td>${esc(x.client)}</td><td>${esc(x.cp.code)}</td><td>${esc(x.cp.description)}</td><td class="num">${money(x.cp.balance)}</td><td class="num">${money(x.basis)}</td><td class="num">${money(x.overlap)}</td><td class="num"><b>${x.fee==null?"—":money(Math.min(x.fee,x.cp.balance))}</b></td><td>${x.fee==null?'<span class="tag">需人工费率/规则</span>':'<span class="tag">可测算</span>'}${x.unknownRate&&x.fee!=null?' <span class="tag">部分产品需确认</span>':''}</td></tr>`).join("")}</tbody></table></div>`;
    const warn=`<div class="warn">当前版本先完成“服务端四文件解析 + CP范围 + 承载候选”。陈列、TT、区域自主等无法仅靠折扣率表达的CP会保留人工判断，不自动硬分；下一阶段再接拍板台账与TTS模板输出。</div>`;
    return new Response(page("费用规划助手｜承载结果",cards+note+ctable+cpTable+mtable+warn),{headers:{"content-type":"text/html; charset=utf-8","cache-control":"no-store"}});
  }catch(e:any){console.error(e); return new Response(page("计算失败",`<div class="warn"><b>服务端已收到请求，但解析失败：</b><br>${esc(e?.message||e)}</div>`),{status:500,headers:{"content-type":"text/html; charset=utf-8","cache-control":"no-store"}});}
};

export const config: Config = { path: "/analyze" };
