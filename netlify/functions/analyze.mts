import type { Context, Config } from "@netlify/functions";
import * as XLSX from "xlsx";

type SoRow = {channel:string; customerCode:string; customerName:string; brandName:string; segment1:string; segment2:string; segment3:string; sku:string; skuName:string; months:Record<string,number>};
type StockRow = {customerCode:string; customerName:string; sku:string; cases:number};
type RateRow = {sku:string; start:Date|null; end:Date|null; rate:number; kind:string};
type TtsRow = {customerName:string; customerCode:string; channel:string; sku:string; start:Date|null; end:Date|null; rate:number; cp:string};
type CpRow = {code:string; start:Date|null; end:Date|null; description:string; mechanism:string; channels:string; mg1:string[]; balance:number; type:string};

const EXCLUDE_SKUS = new Set(["12187803","12585787","12598186","12611209"]);
const TT_ALLOWED = ["天虹","东方"];
const esc=(v:any)=>String(v??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;");
const num=(v:any)=>{const n=Number(v);return Number.isFinite(n)?n:0};
const money=(v:number)=>new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0}).format(v||0);
const normalize=(s:any)=>String(s??"").replace(/&amp;/g,"&").replace(/\s+/g,"").replace(/有限责任公司$/,"有限公司").toLowerCase();
const monthKey=(d:Date)=>`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}`;
const monthRange=(a:Date|null,b:Date|null)=>{if(!a||!b)return[] as string[];const out:string[]=[];const d=new Date(a.getFullYear(),a.getMonth(),1),e=new Date(b.getFullYear(),b.getMonth(),1);while(d<=e){out.push(monthKey(d));d.setMonth(d.getMonth()+1)}return out};
const parseDate=(v:any):Date|null=>{if(!v)return null;if(v instanceof Date&&!Number.isNaN(v.getTime()))return v;if(typeof v==="number"){const o=XLSX.SSF.parse_date_code(v);return o?new Date(o.y,o.m-1,o.d):null}const s=String(v).trim();const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(m)return new Date(+m[3],+m[1]-1,+m[2]);const d=new Date(s);return Number.isNaN(d.getTime())?null:d};
const splitList=(v:any)=>String(v??"").split(/[,，;]/).map(x=>x.trim()).filter(Boolean);
const daysInMonth=(y:number,m:number)=>new Date(y,m,0).getDate();

async function readWorkbook(file:File){const ab=await file.arrayBuffer();return XLSX.read(Buffer.from(ab),{type:"buffer",cellDates:true,raw:true,nodim:true} as any)}
function sheetRows(wb:any,name?:string){const sn=name&&wb.SheetNames.includes(name)?name:wb.SheetNames[0];return XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null,raw:true}) as any[][]}
function findHeader(rows:any[][],required:string[]){return rows.findIndex(r=>required.every(k=>r.some(c=>String(c??"").trim()===k)))}
function idxMap(h:any[]){const o:Record<string,number>={};h.forEach((v,i)=>{if(v!=null)o[String(v).trim()]=i});return o}

function parseBiSo(wb:any):SoRow[]{
  if(!wb.SheetNames.includes("BI-SO(折扣)"))return[];const rows=sheetRows(wb,"BI-SO(折扣)"),hi=findHeader(rows,["Outlet Channel","Customer Code","SKU"]);if(hi<0)return[];const h=rows[hi].map(x=>String(x??"").trim()),ix=idxMap(h);const monthCols=h.map((x,i)=>/^20\d{4}$/.test(x)?[x,i]:null).filter(Boolean) as [string,number][];
  return rows.slice(hi+1).map(r=>({channel:String(r[ix["Outlet Channel"]]??""),customerCode:String(r[ix["Customer Code"]]??""),customerName:String(r[ix["Customer Name"]]??""),brandName:String(r[ix["Brand Name"]]??""),segment1:String(r[ix["PH5 Segment1_MKT"]]??""),segment2:String(r[ix["PH5 Segment2_Detail"]]??""),segment3:String(r[ix["PH5 Segment3_Detail"]]??""),sku:String(r[ix["SKU"]]??""),skuName:String(r[ix["SKU Name CN"]]??""),months:Object.fromEntries(monthCols.map(([m,i])=>[m,num(r[i])]))})).filter(x=>x.customerName&&x.sku&&!/汇总/.test(x.customerCode));
}
function parseStock(wb:any):StockRow[]{
  if(!wb.SheetNames.includes("BI-Stock"))return[];const rows=sheetRows(wb,"BI-Stock"),hi=findHeader(rows,["Customer Code","Customer Name","SKU","SOH (Good) Case"]);if(hi<0)return[];const h=rows[hi].map(x=>String(x??"").trim()),ix=idxMap(h);
  return rows.slice(hi+1).map(r=>({customerCode:String(r[ix["Customer Code"]]??""),customerName:String(r[ix["Customer Name"]]??""),sku:String(r[ix["SKU"]]??""),cases:num(r[ix["SOH (Good) Case"]])})).filter(x=>x.customerName&&x.sku&&x.cases!==0);
}
function parseRates(wb:any):RateRow[]{
  if(!wb.SheetNames.includes("SKURate"))return[];const rows=sheetRows(wb,"SKURate"),hi=findHeader(rows,["SKUCode","StartDate","EndDate","Rate"]);if(hi<0)return[];const h=rows[hi].map(x=>String(x??"").trim()),ix=idxMap(h);
  return rows.slice(hi+1).map(r=>{let rate=num(r[ix["_RateValue"]]??r[ix["Rate"]]);if(typeof r[ix["Rate"]]==="string"&&String(r[ix["Rate"]]).includes("%"))rate=parseFloat(String(r[ix["Rate"]]))/100;return{sku:String(r[ix["SKUCode"]]??""),start:parseDate(r[ix["StartDate"]]),end:parseDate(r[ix["EndDate"]]),rate,kind:String(r[ix["DDDT"]]??"")}}).filter(x=>x.sku&&x.start&&x.end);
}
function parsePrices(wb:any):Map<string,number>{
  const out=new Map<string,number>();for(const sn of wb.SheetNames){const rows=sheetRows(wb,sn),hi=findHeader(rows,["产品编码"]);if(hi<0)continue;const h=rows[hi].map(x=>String(x??"").trim()),skuI=h.indexOf("产品编码");let priceI=h.findIndex(x=>x.includes("NPS")&&x.includes("未税箱价"));if(priceI<0)priceI=h.indexOf("Price");if(priceI<0)continue;for(const r of rows.slice(hi+1)){const sku=String(r[skuI]??"").trim(),p=num(r[priceI]);if(sku&&p>0)out.set(sku,p)}}return out;
}
function parseTts(wb:any):TtsRow[]{
  const rows=sheetRows(wb,wb.SheetNames.includes("PPCenturyBigTable")?"PPCenturyBigTable":undefined),hi=findHeader(rows,["PPCode","Distributor Name","Activity Code in Cycle Plan"]);if(hi<0)return[];const h=rows[hi].map(x=>String(x??"").trim()),ix=idxMap(h);
  return rows.slice(hi+1).map(r=>{let rate=num(r[ix["Disc.Rate"]]);if(rate>1)rate/=100;const rs=parseDate(r[ix["RemarkStartDate"]]),re=parseDate(r[ix["RemarkEndDate"]]);return{customerName:String(r[ix["Distributor Name"]]??""),customerCode:String(r[ix["Distributor Code"]]??""),channel:String(r[ix["Channel"]]??""),sku:String(r[ix["Product Code"]]??""),start:rs||parseDate(r[ix["Promotion Start Date"]]),end:re||parseDate(r[ix["Promotion End Date"]]),rate,cp:String(r[ix["Activity Code in Cycle Plan"]]??"")}}).filter(x=>x.customerName&&x.sku&&x.rate>0&&x.start&&x.end);
}
function parseCp(wb:any):CpRow[]{
  const rows=sheetRows(wb,wb.SheetNames.includes("CPCenturyBigTableByCGAmount")?"CPCenturyBigTableByCGAmount":undefined),hi=findHeader(rows,["Activity Code","Activity Start Date","Balance Amount"]);if(hi<0)return[];const h=rows[hi].map(x=>String(x??"").trim()),ix=idxMap(h);
  return rows.slice(hi+1).map(r=>({code:String(r[ix["Activity Code"]]??""),start:parseDate(r[ix["Activity Start Date"]]),end:parseDate(r[ix["Activity End Date"]]),description:String(r[ix["Activity Description"]]??""),mechanism:String(r[ix["Activity Mechanism"]]??""),channels:String(r[ix["SubChannel"]]??""),mg1:splitList(r[ix["MG1"]]),balance:num(r[ix["Balance Amount"]]),type:String(r[ix["PromotionType"]]??r[ix["Activity Type"]]??"")})).filter(x=>x.code&&x.balance>0);
}
function cpClientAllowed(cp:CpRow,name:string){if(cp.code==="DA202601011TT")return TT_ALLOWED.some(k=>name.includes(k));return true}
function channelAllowed(cp:CpRow,ch:string){const c=normalize(cp.channels),x=normalize(ch);if(!c)return true;if(c.includes(x))return true;const aliases:Record<string,string[]>={"sm":["hyper&super-nonka13","lkaothers","lkagroup"],"hm":["hyper&super-nka13","hyper&super-nonka13"],"ws":["wholesale"],"mm-a":["mm-a"],"mm-b":["mm-b"],"mm-c":["mm-c"],"speciality":["specialty"],"ss_kiosks":["smallstore"]};return(aliases[x]||[]).some(a=>c.includes(normalize(a)))}
function rowEligible(cp:CpRow,r:SoRow){
  if(!cpClientAllowed(cp,r.customerName)||EXCLUDE_SKUS.has(r.sku)||!channelAllowed(cp,r.channel))return false;const mech=cp.mechanism;const seg3=r.segment3;
  const explicit:string[]=[];if(/新西兰/.test(mech))explicit.push("新西兰");if(/全脂375g/.test(mech))explicit.push("全脂375g");if(/甜奶粉/.test(mech))explicit.push("甜奶粉300g");if(explicit.length)return explicit.some(k=>seg3.includes(k));
  if(cp.mg1.length)return cp.mg1.some(x=>normalize(x)===normalize(r.brandName));return true;
}
function skuRateFor(rates:RateRow[],sku:string,month:string):number|null{
  if(!/^20\d{4}$/.test(month))return null;const y=+month.slice(0,4),m=+month.slice(4,6),s=new Date(y,m-1,1),e=new Date(y,m,0);const vals=[...new Set(rates.filter(x=>x.sku===sku&&x.start!<=e&&x.end!>=s&&x.rate>0).map(x=>+x.rate.toFixed(6)))];return vals.length===1?vals[0]:null;
}
function mechanismRate(cp:CpRow,r:SoRow):number|null{
  const t=(cp.description+" "+cp.mechanism).replace(/：/g,":");const s=(r.segment1+" "+r.segment2+" "+r.segment3+" "+r.skuName).toLowerCase();const pick=(re:RegExp)=>{const m=t.match(re);return m?Number(m[1])/100:null};
  for(const [re,key] of [[/中老年[^%]{0,40}?(\d+(?:\.\d+)?)%/i,"senior"],[/N3[^%]{0,40}?(\d+(?:\.\d+)?)%/i,"n3"],[/全家[^%]{0,40}?(\d+(?:\.\d+)?)%/i,"family"]] as any[]){if(s.includes(key)){const v=pick(re);if(v!=null)return v}}
  const all=[...t.matchAll(/(\d+(?:\.\d+)?)\s*%/g)].map(m=>Number(m[1])/100),uniq=[...new Set(all)];return uniq.length===1?uniq[0]:null;
}
function targetRate(cp:CpRow,r:SoRow,rates:RateRow[],month:string){return mechanismRate(cp,r)??skuRateFor(rates,r.sku,month)}
function overlapRate(tts:TtsRow[],customer:string,sku:string,month:string,ch:string){let mx=0;for(const p of tts){if(normalize(p.customerName)!==normalize(customer)||p.sku!==sku)continue;if(p.channel&&ch&&!normalize(p.channel).includes(normalize(ch))&&!normalize(ch).includes(normalize(p.channel)))continue;if(!monthRange(p.start,p.end).includes(month))continue;mx=Math.max(mx,p.rate)}return mx}
function r156(r:SoRow){const ly=r.months["202509"]||0,avg=((r.months["202605"]||0)+(r.months["202606"]||0)+(r.months["202607"]||0))/3;return Math.max(ly,(ly+avg)/2)*1000}
function currentMonthBasis(r:SoRow,stockValue:number,now:Date){const m=monthKey(now),actual=(r.months[m]||0)*1000,done=Math.max(1,now.getDate()-1),total=daysInMonth(now.getFullYear(),now.getMonth()+1),remaining=Math.max(0,total-done),trendRemain=actual/done*remaining,remainingWeeks=remaining/7,inventoryRemain=stockValue/4*remainingWeeks,w=now.getDate()<=10?.10:now.getDate()<=20?.20:.30;return actual+trendRemain*(1-w)+inventoryRemain*w}

function page(title:string,body:string){return`<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.wrap{max-width:1500px;margin:auto;padding:24px}.top{display:flex;justify-content:space-between;align-items:center}.top a{color:#1769ff;text-decoration:none}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin:16px 0}.card,.panel{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:14px}.card b{font-size:22px;display:block}.card span,.muted{color:#6b7688}.panel{margin:14px 0;overflow:auto}table{border-collapse:collapse;width:100%;min-width:1050px}th,td{border-bottom:1px solid #edf0f5;padding:9px 8px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8fafc;font-size:12px}.num{text-align:right}.tag{display:inline-block;padding:2px 7px;border-radius:999px;background:#eef5ff;color:#1769ff;font-size:12px}.warn{background:#fff8e8;border:1px solid #f0d89a;padding:10px 12px;border-radius:9px}.ok{background:#ecfbf3;border:1px solid #bde8cf;padding:10px 12px;border-radius:9px}@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.wrap{padding:14px}}</style></head><body><div class="wrap"><div class="top"><div><h2 style="margin:0">${esc(title)}</h2><div class="muted">Netlify 服务端计算 · 无前端 JavaScript 依赖</div></div><a href="/">重新上传</a></div>${body}</div></body></html>`}

function selfTest(){const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet([["A","B"],[1,2],[3,4]]);(ws as any)["!ref"]="A1:B1";XLSX.utils.book_append_sheet(wb,ws,"T");const buf=XLSX.write(wb,{type:"buffer",bookType:"xlsx"});const reread=XLSX.read(buf,{type:"buffer",nodim:true} as any);const rows=XLSX.utils.sheet_to_json(reread.Sheets.T,{header:1,defval:null}) as any[][];return rows.length>=3}

export default async(req:Request,context:Context)=>{
  const u=new URL(req.url);if(req.method==="GET"&&u.searchParams.get("selftest")==="1")return new Response(JSON.stringify({ok:selfTest(),runtime:"xlsx-nodim",time:new Date().toISOString()}),{headers:{"content-type":"application/json"}});if(req.method!=="POST")return new Response("Method Not Allowed",{status:405});
  try{
    const fd=await req.formData(),bi=fd.get("bi"),ttsF=fd.get("tts"),cpF=fd.get("cp"),product=fd.get("product");if(!(bi instanceof File)||!(ttsF instanceof File)||!(cpF instanceof File)||!(product instanceof File))return new Response(page("文件不完整",`<div class="warn">必须上传 BI、TTS、CP预算、产品目录四个文件。</div>`),{status:400,headers:{"content-type":"text/html; charset=utf-8"}});
    const[biW,ttsW,cpW,prodW]=await Promise.all([readWorkbook(bi),readWorkbook(ttsF),readWorkbook(cpF),readWorkbook(product)]);const so=parseBiSo(biW),stock=parseStock(biW),rates=parseRates(biW),tts=parseTts(ttsW),cps=parseCp(cpW),prices=parsePrices(prodW);if(!prices.size)for(const[k,v]of parsePrices(biW))prices.set(k,v);
    const focus=String(fd.get("focus")??"").split(/[,，]/).map(s=>s.trim()).filter(Boolean),minBalance=num(fd.get("minBalance")??500),now=new Date(),cur=monthKey(now);let clientNames=[...new Set(so.map(x=>x.customerName))];if(focus.length)clientNames=clientNames.filter(n=>focus.some(f=>normalize(n).includes(normalize(f))));else clientNames=clientNames.slice(0,30);const activeCps=cps.filter(c=>c.balance>=minBalance).sort((a,b)=>b.balance-a.balance);
    const stockMap=new Map<string,number>();for(const s of stock){const k=`${normalize(s.customerName)}|${s.sku}`,v=s.cases*(prices.get(s.sku)||0);stockMap.set(k,(stockMap.get(k)||0)+v)}
    const matrix:any[]=[];for(const client of clientNames){for(const cp of activeCps){let hard=0,current=0,soft=0,overlap=0,unknown=0;for(const r of so){if(normalize(r.customerName)!==normalize(client)||!rowEligible(cp,r))continue;for(const m of monthRange(cp.start,cp.end)){const rate=targetRate(cp,r,rates,m);if(rate==null){unknown++;continue}const ov=overlapRate(tts,client,r.sku,m,r.channel),effective=Math.max(0,rate-ov);let base=0,kind="hard";if(m===cur){base=currentMonthBasis(r,stockMap.get(`${normalize(client)}|${r.sku}`)||0,now);kind="current"}else if(m>cur){base=m==="202609"?r156(r):0;kind="soft"}else base=(r.months[m]||0)*1000;if(base<=0)continue;overlap+=base*Math.min(rate,ov);const fee=base*effective;if(kind==="hard")hard+=fee;else if(kind==="current")current+=fee;else soft+=fee}}
      matrix.push({client,cp,hard,current,soft,total:Math.min(cp.balance,hard+current+soft),overlap,unknown})}}
    const cards=`<div class="cards"><div class="card"><span>BI SO行</span><b>${money(so.length)}</b></div><div class="card"><span>库存行</span><b>${money(stock.length)}</b></div><div class="card"><span>SKU Rate</span><b>${money(rates.length)}</b></div><div class="card"><span>TTS折扣行</span><b>${money(tts.length)}</b></div><div class="card"><span>可用CP</span><b>${money(activeCps.length)}</b></div><div class="card"><span>目录SKU</span><b>${money(prices.size)}</b></div></div>`;
    const note=`<div class="ok">四文件已在服务端解析。TTS优先使用“备注计划开始/结束日期”；本月承载=已实现SO + 趋势剩余SO/库存隐含剩余SO加权（库存权重10%/20%/30%，库存按4周）；未来9月使用R-156，库存不直接硬加。</div>`;
    const table=`<div class="panel"><h3>客户 × CP 承载候选</h3><table><thead><tr><th>客户</th><th>CP</th><th>活动</th><th class="num">余额</th><th class="num">历史HARD</th><th class="num">本月预测</th><th class="num">未来SOFT</th><th class="num">重叠折扣扣减</th><th class="num">建议可见上限</th><th>状态</th></tr></thead><tbody>${matrix.map(x=>`<tr><td>${esc(x.client)}</td><td><b>${esc(x.cp.code)}</b></td><td>${esc(x.cp.description)}</td><td class="num">${money(x.cp.balance)}</td><td class="num">${money(x.hard)}</td><td class="num">${money(x.current)}</td><td class="num">${money(x.soft)}</td><td class="num">${money(x.overlap)}</td><td class="num"><b>${money(x.total)}</b></td><td>${x.unknown?`<span class="tag">${x.unknown}个费率桶待确认</span>`:'<span class="tag">可测算</span>'}</td></tr>`).join("")}</tbody></table></div>`;
    const cpTable=`<div class="panel"><h3>CP范围</h3><table><thead><tr><th>CP</th><th>有效期</th><th>机制</th><th>MG1</th><th>渠道</th><th class="num">余额</th></tr></thead><tbody>${activeCps.map(c=>`<tr><td><b>${esc(c.code)}</b></td><td>${esc(c.start?.toLocaleDateString("zh-CN")||"")} → ${esc(c.end?.toLocaleDateString("zh-CN")||"")}</td><td>${esc(c.mechanism)}</td><td>${esc(c.mg1.join(" / "))}</td><td>${esc(c.channels)}</td><td class="num">${money(c.balance)}</td></tr>`).join("")}</tbody></table></div>`;
    return new Response(page("费用规划助手｜承载结果",cards+note+table+cpTable+`<div class="warn">当前页面用于验算承载。拍板台账和自动TTS模板输出尚未接到服务端版；先以真实数据对账通过后再接。</div>`),{headers:{"content-type":"text/html; charset=utf-8","cache-control":"no-store"}});
  }catch(e:any){console.error(e);return new Response(page("计算失败",`<div class="warn"><b>服务端已收到请求，但解析失败：</b><br>${esc(e?.message||e)}</div>`),{status:500,headers:{"content-type":"text/html; charset=utf-8","cache-control":"no-store"}})}
};
export const config:Config={path:"/analyze"};
