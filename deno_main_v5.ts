import { Buffer } from "node:buffer";
import * as XLSX from "xlsx";
import { renderFilterDashboard as renderFilterDashboardV3 } from "./filter_ui_v3.ts";

(globalThis as any).Buffer = Buffer;

const engine = await import("./netlify/functions/_shared/engine.mts");
const { analyzeFiles, BASELINE, money } = engine;

type AnalysisResult = Awaited<ReturnType<typeof analyzeFiles>>;

const COOKIE = "expense_planner_auth";
const SESSION_SECONDS = 8 * 60 * 60;
const MAX_FILTER_EXPORT_BYTES = 12 * 1024 * 1024;

const env = (name: string) => Deno.env.get(name) || "";
const esc = (v: unknown) => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;");

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function sha256(ab: ArrayBuffer) {
  return bytesToHex(new Uint8Array(await crypto.subtle.digest("SHA-256", ab)));
}

async function hmac(value: string) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(env("AUTH_SECRET")),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  return bytesToHex(new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value))));
}

async function createSessionCookie() {
  const exp = Math.floor(Date.now() / 1000) + SESSION_SECONDS;
  return `${COOKIE}=${exp}.${await hmac(String(exp))}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${SESSION_SECONDS}`;
}

function clearSessionCookie() {
  return `${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`;
}

async function isAuthorized(req: Request) {
  if (!env("AUTH_SECRET") || !env("SHARED_PASSWORD")) return false;
  const cookie = req.headers.get("cookie") || "";
  const match = cookie.match(new RegExp(`(?:^|;\\s*)${COOKIE}=([^;]+)`));
  if (!match) return false;
  const [expText, signature] = match[1].split(".");
  const exp = Number(expText);
  if (!Number.isFinite(exp) || exp < Math.floor(Date.now() / 1000) || !signature) return false;
  return signature === await hmac(expText);
}

function passwordMatches(value: string) {
  const expected = env("SHARED_PASSWORD");
  if (!expected || value.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < value.length; i++) diff |= value.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0;
}

function expandRef(ws: any) {
  let maxR = 0, maxC = 0, seen = false;
  for (const k of Object.keys(ws || {})) {
    if (k.startsWith("!")) continue;
    const m = k.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    maxC = Math.max(maxC, XLSX.utils.decode_col(m[1]));
    maxR = Math.max(maxR, Number(m[2]) - 1);
    seen = true;
  }
  if (seen) ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
}

function sheetRows(wb: any, name: string) {
  const ws = wb?.Sheets?.[name];
  if (!ws) return [] as any[][];
  expandRef(ws);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as any[][];
}

function findHeader(rs: any[][], need: string[]) {
  return rs.findIndex((r) => need.every((k) => r.some((c) => String(c ?? "").trim() === k)));
}

function headerMap(h: any[]) {
  const out: Record<string, number> = {};
  h.forEach((v, i) => { if (v != null) out[String(v).trim()] = i; });
  return out;
}

async function makePack(file: File) {
  const ab = await file.arrayBuffer();
  return {
    wb: XLSX.read(Buffer.from(ab), { type: "buffer", cellDates: true, raw: true }),
    name: file.name,
    size: file.size,
    sha: await sha256(ab),
  };
}

function collectSkus(wb: any) {
  const out = new Set<string>();
  for (const sn of ["BI-SO(折扣)", "BI-Stock"]) {
    if (!wb.SheetNames.includes(sn)) continue;
    const rs = sheetRows(wb, sn);
    const hi = findHeader(rs, ["SKU"]);
    if (hi < 0) continue;
    const x = headerMap(rs[hi]);
    const si = x["SKU"];
    for (const r of rs.slice(hi + 1)) {
      const sku = String(r[si] ?? "").trim();
      if (sku && !/汇总/.test(sku)) out.add(sku);
    }
  }
  return [...out];
}

async function prepareBiForAnalysis(file: File) {
  const ab = await file.arrayBuffer();
  const originalWb = XLSX.read(Buffer.from(ab), { type: "buffer", cellDates: true, raw: true });
  if (!originalWb.SheetNames.includes("BI-Stock")) throw new Error("BI 文件缺少 BI-Stock Sheet。");

  const wb = XLSX.read(Buffer.from(ab), { type: "buffer", cellDates: true, raw: true });
  const rs = sheetRows(wb, "BI-Stock");
  const hi = findHeader(rs, ["SKU", "SOH (Good)"]);
  if (hi < 0) throw new Error("BI-Stock 未找到 SKU / SOH (Good) 字段。");
  const h = rs[hi].map((v) => String(v ?? "").trim());
  const amountIdx = h.indexOf("SOH (Good)");
  let caseIdx = h.indexOf("SOH (Good) Case");
  if (caseIdx < 0) {
    caseIdx = h.length;
    rs[hi][caseIdx] = "SOH (Good) Case";
  }
  for (const row of rs.slice(hi + 1)) row[caseIdx] = row[amountIdx];
  wb.Sheets["BI-Stock"] = XLSX.utils.aoa_to_sheet(rs);

  const transformed = XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  const analysisBi = new File([transformed], file.name.replace(/\.xlsm?$/i, ".xlsx"), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const skus = collectSkus(originalWb);
  const pwb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(pwb, XLSX.utils.aoa_to_sheet([["产品编码", "Price"], ...skus.map((sku) => [sku, 1000])]), "Price");
  const parr = XLSX.write(pwb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  const syntheticProduct = new File([parr], "_internal_stock_value.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  return {
    analysisBi,
    syntheticProduct,
    originalPack: { wb: originalWb, name: file.name, size: file.size, sha: await sha256(ab) },
  };
}

function productGroupCode(v: unknown) {
  const s = String(v ?? "").trim();
  return /^[A-Za-z0-9]{2,8}$/.test(s) ? s : "";
}

function attachProductGroupCodes(r: any) {
  const wb = r?.packs?.[0]?.wb;
  if (!wb || !wb.SheetNames?.includes("BI-SO(折扣)")) return;
  const rs = sheetRows(wb, "BI-SO(折扣)");
  const hi = findHeader(rs, ["Brand Name", "SKU"]);
  if (hi < 0) return;
  const h = rs[hi].map((v) => String(v ?? "").trim());
  const x = headerMap(h);
  const brandIdx = x["Brand Name"], skuIdx = x["SKU"];
  if (!Number.isInteger(brandIdx) || !Number.isInteger(skuIdx) || brandIdx <= 0) return;
  const codeIdx = brandIdx - 1;
  const bySku = new Map<string, string>(), byBrand = new Map<string, string>();
  for (const row of rs.slice(hi + 1)) {
    const sku = String(row[skuIdx] ?? "").trim();
    const brand = String(row[brandIdx] ?? "").trim();
    const code = productGroupCode(row[codeIdx]);
    if (!code) continue;
    if (sku && !bySku.has(sku)) bySku.set(sku, code);
    if (brand && !byBrand.has(brand)) byBrand.set(brand, code);
  }
  for (const s of r.so || []) s.productGroupCode = bySku.get(String(s.sku || "")) || byBrand.get(String(s.brand || "")) || "";
}

function loginPage(message = "") {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>费用规划助手</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.box{width:min(430px,calc(100% - 32px));margin:14vh auto;background:#fff;border:1px solid #dbe3ef;border-radius:16px;padding:26px}.muted{color:#6b7688}.warn{background:#fff3e8;border:1px solid #f0cda5;border-radius:8px;padding:9px;margin:12px 0}.field{width:100%;box-sizing:border-box;padding:11px;border:1px solid #cfd8e6;border-radius:9px;margin:10px 0 14px}.btn{width:100%;border:0;background:#1769ff;color:#fff;padding:11px;border-radius:9px;font-weight:600}</style></head><body><div class="box"><h2 style="margin:0 0 6px">费用规划助手</h2><div class="muted">仅限授权同事使用 · Deno Deploy</div>${message ? `<div class="warn">${esc(message)}</div>` : ""}<form action="/auth" method="post"><input class="field" type="password" name="password" autocomplete="current-password" placeholder="共享密码" required autofocus><button class="btn" type="submit">进入</button></form></div></body></html>`;
}

function workspacePage() {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>费用规划助手</title><style>:root{--nav:#0d1b2a;--bg:#f4f6fa;--panel:#fff;--line:#dbe3ef;--text:#142033;--muted:#6b7688;--blue:#1769ff}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.layout{display:grid;grid-template-columns:210px 1fr;min-height:100vh}.side{background:var(--nav);color:#fff;padding:22px 18px}.side h1{font-size:18px;margin:0 0 4px}.side p{font-size:12px;color:#b7c3d5;margin:0 0 30px}.step{padding:10px 12px;border-radius:9px;margin:7px 0;color:#c9d4e5}.step.active{background:#17365c;color:#fff}.main{padding:26px 28px}.head{display:flex;justify-content:space-between;gap:12px}.head h2{font-size:24px;margin:0 0 4px}.head p{color:var(--muted);margin:0 0 20px}.head a{color:#1769ff;text-decoration:none}.notice{padding:12px 14px;border:1px solid #cfe0ff;background:#eef5ff;border-radius:10px;margin-bottom:18px}.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}.file{border:1px dashed #b9c8dd;border-radius:12px;padding:14px}.file b{display:block;margin-bottom:5px}.file small{display:block;color:var(--muted);margin-bottom:10px}.file input{width:100%}.row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.field label{display:block;font-weight:600;margin-bottom:6px}.field input{width:100%;border:1px solid #cfd8e6;border-radius:9px;padding:10px}.actions{display:flex;justify-content:flex-end;margin-top:16px}.btn{border:0;background:var(--blue);color:#fff;padding:11px 18px;border-radius:9px;font-weight:600}.foot{color:var(--muted);font-size:12px;margin-top:14px}@media(max-width:950px){.grid{grid-template-columns:1fr}.layout{grid-template-columns:1fr}.side{display:none}.row{grid-template-columns:1fr}.main{padding:18px}}</style></head><body><div class="layout"><aside class="side"><h1>费用规划助手</h1><p>共享使用 · Deno Deploy</p><div class="step active">1 数据更新</div><div class="step">2 CP范围</div><div class="step">3 承载 / 拍板</div><div class="step">4 计划输出</div></aside><main class="main"><div class="head"><div><h2>数据更新</h2><p>BI + TTS + CP预算 → 服务端解析 → 承载候选 → 结果筛选</p></div><a href="/logout">退出</a></div><div class="notice">产品目录文件已取消。库存预测直接使用 BI-Stock 的 SOH (Good) 千分位库存金额；一次计算完成后筛选只在浏览器本地执行。</div><form class="panel" action="/analyze" method="post" enctype="multipart/form-data"><div class="grid"><div class="file"><b>BI</b><small>BI-SO(折扣)、BI-Stock、SKURate、产品目录；库存金额读取 SOH (Good)。</small><input required type="file" name="bi" accept=".xlsx,.xlsm,.xls"></div><div class="file"><b>TTS 最新下载</b><small>读取同CP既有计划、重叠费率和回溯备注日期。</small><input required type="file" name="tts" accept=".xlsx,.xlsm,.xls"></div><div class="file"><b>CP预算</b><small>读取余额、有效期、MG1、渠道、机制。</small><input required type="file" name="cp" accept=".xlsx,.xlsm,.xls"></div></div><div class="row"><div class="field"><label>重点客户关键词</label><input name="focus" placeholder="可留空；多个客户用逗号分隔"></div><div class="field"><label>最小展示CP余额</label><input name="minBalance" type="number" value="500" min="0" step="100"></div></div><div class="actions"><button class="btn" type="submit">上传并计算</button></div><div class="foot">三份文件完成一次计算后，可反复筛选与导出，无需重复上传。</div></form></main></div></body></html>`;
}

function page(title: string, body: string) {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.w{max-width:1700px;margin:auto;padding:22px}.top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.actions{display:flex;gap:10px}.top a{color:#1769ff;text-decoration:none}.btn{display:inline-block;background:#1769ff;color:#fff!important;border-radius:9px;padding:9px 13px;font-weight:600}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin:16px 0}.card,.p{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:13px}.card b{display:block;font-size:22px}.card span,.muted{color:#6b7688}.p{margin:14px 0;overflow:auto}.ok{background:#ecfbf3;border:1px solid #bde8cf;padding:10px;border-radius:9px}.warn{background:#fff8e8;border:1px solid #f0d89a;padding:10px;border-radius:9px}table{border-collapse:collapse;width:100%;min-width:1300px}th,td{border-bottom:1px solid #edf0f5;padding:8px;text-align:left}th{background:#f8fafc}.num{text-align:right}.bad{color:#b42318;font-weight:600}.good{color:#067647;font-weight:600}@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.top{display:block}.actions{margin-top:10px}}</style></head><body><div class="w"><div class="top"><div><h2 style="margin:0">${esc(title)}</h2><div class="muted">Deno Deploy · 浏览器本地筛选</div></div><div class="actions"><a href="/workspace">重新上传</a><a href="/logout">退出</a></div></div>${body}</div></body></html>`;
}

function filePanel(r: AnalysisResult) {
  const rows = (r.packs || []).map((p: any, i: number) => {
    const exp = BASELINE[i];
    const ok = exp ? p.sha.startsWith(exp.sha) : false;
    return `<tr><td>${esc(exp?.role || ["BI","TTS","CP预算"][i] || "文件")}</td><td>${esc(p.name)}</td><td class="num">${(p.size / 1024).toFixed(1)}</td><td><code>${p.sha.slice(0, 12)}</code></td><td class="${ok ? "good" : "bad"}">${exp ? (ok ? "与验收基线一致" : `不同；基线 ${exp.sha}`) : "-"}</td></tr>`;
  }).join("");
  return `<div class="p"><h3>本次材料指纹</h3><table style="min-width:850px"><thead><tr><th>角色</th><th>文件名</th><th class="num">大小KB</th><th>SHA-256前12位</th><th>基线状态</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function renderFilterDashboardV5(r: any) {
  let html = renderFilterDashboardV3(r);
  html = html.replace("var groupLevel=\"s3\", metric=\"visibleCapacity\", collapsed=false;", "var groupLevel=\"s3\", metric=\"net\", collapsed=false, capacityThreshold=500;");
  html = html.replace("var metrics=[{key:\"visibleCapacity\",label:\"可见承载\"},{key:\"net\",label:\"净承载\"},{key:\"forecastCapacity\",label:\"预测承载\"},{key:\"so\",label:\"SO\"}];", "var metrics=[{key:\"net\",label:\"净承载\"},{key:\"forecastCapacity\",label:\"预测承载\"},{key:\"so\",label:\"SO\"}];");
  html = html.replace('<div class="rf-metrics" id="metricSwitch"></div>', '<div class="rf-metric-wrap"><div class="muted rf-metric-note">净承载：实际SO按有效费率测算并扣除同CP已占用；预测承载：预测SO按有效费率测算并扣除同CP已占用。</div><label class="rf-threshold">可见承载阈值 <input id="capacityThreshold" type="number" min="0" step="100" value="500"> 元</label><div class="rf-metrics" id="metricSwitch"></div></div>');
  html = html.replace("var a=v.actual>=500?v.actual:0,p=v.forecast>=500?v.forecast:0;", "var a=v.actual>=capacityThreshold?v.actual:0,p=v.forecast>=capacityThreshold?v.forecast:0;");
  html = html.replace('if(metric!=="so")out=out.filter(function(g){return g.unknown||ms.some(function(mm){return Number(g.values[mm]||0)!==0})});', 'if(metric!=="so")out=out.filter(function(g){return g.unknown||ms.some(function(mm){return Math.abs(Number(g.values[mm]||0))>=capacityThreshold})});');
  html = html.replace('document.getElementById("kRows").textContent=money(lastRows.length);', 'document.getElementById("kRows").textContent=money(lastPivot.length);');
  html = html.replace('document.getElementById("rfExport").onclick=exportXlsx;', 'var thr=document.getElementById("capacityThreshold");if(thr){capacityThreshold=Math.max(0,Number(thr.value)||0);thr.onchange=function(){capacityThreshold=Math.max(0,Number(thr.value)||0);apply()}}document.getElementById("rfExport").onclick=exportXlsx;');
  html = html.replace('</style>', '.rf-metric-wrap{display:flex;align-items:center;justify-content:flex-end;gap:10px;flex-wrap:wrap}.rf-metric-note{max-width:760px;font-size:12px}.rf-threshold{display:inline-flex;align-items:center;gap:6px;font-weight:600}.rf-threshold input{width:92px;border:1px solid #cfd8e6;border-radius:8px;padding:6px 8px;text-align:right}</style>');
  return html;
}

function planningRows(r: any) {
  const meta = new Map<string, any>();
  for (const s of r.so || []) {
    const sku = String(s.sku || "");
    if (!sku) continue;
    const old = meta.get(sku) || {};
    meta.set(sku, {
      brand: old.brand || String(s.brand || ""),
      s1: old.s1 || String(s.s1 || ""),
      s2: old.s2 || String(s.s2 || ""),
      s3: old.s3 || String(s.s3 || ""),
      productName: old.productName || String(s.skuName || ""),
      productGroupCode: old.productGroupCode || String(s.productGroupCode || ""),
    });
  }
  return (r.exportRows || []).map((x: any) => {
    const sku = String(x.SKU || ""), m = meta.get(sku) || {};
    return {
      customer: String(x["经销商/客户"] || ""), customerCode: String(x["客户编码"] || ""), cp: String(x.CP || ""), activity: String(x["活动"] || ""), channel: String(x["渠道"] || ""),
      brand: String(x["品牌"] || m.brand || ""), s1: String(m.s1 || ""), s2: String(m.s2 || ""), s3: String(m.s3 || ""), productGroupCode: String(m.productGroupCode || ""),
      sku, productName: String(x["产品名称"] || m.productName || ""), month: String(x["月份"] || ""),
      net: x["净承载"] == null ? null : Number(x["净承载"]), forecastCapacity: x["预测承载"] == null ? null : Number(x["预测承载"]),
    };
  });
}

function renderPlanningAddon(r: any) {
  const rows = JSON.stringify(planningRows(r)).replace(/</g, "\\u003c").replace(/-->/g, "--\\>");
  const cur = JSON.stringify(String(r.currentMonth || ""));
  return `<div class="p rf-table rf-plan" id="v5Plan"><div class="rf-bar"><div><h3>计划汇总</h3><div class="muted">按客户 × 活动 × 渠道 × 产品组汇总；历史取净承载，当前/未来取预测承载，并应用可见承载阈值。</div></div><div class="muted" id="v5PlanCount"></div></div><table><thead id="v5PlanHead"></thead><tbody id="v5PlanBody"></tbody></table></div>
<style>.rf-plan table{min-width:1050px}.rf-plan .v5-total td{font-weight:700;background:#f8fafc}.rf-plan .v5-pg{min-width:220px}.rf-detail.v5-collapsed table{display:none}.v5-detail-toggle{border:1px solid #cfd8e6;background:#fff;color:#1769ff;border-radius:8px;padding:6px 10px;cursor:pointer}</style>
<script>(function(){var V=${rows},CUR=${cur};var fmt=new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0});function money(v){return fmt.format(Number(v)||0)}function cmp(a,b){return String(a).localeCompare(String(b),"zh-CN",{numeric:true})}function uniq(a){return Array.from(new Set(a.filter(function(v){return String(v||"")!==""}))).sort(cmp)}function ml(v){return /^20\\d{4}$/.test(v)?Number(v.slice(4))+"月":v}function sel(id){return new Set(Array.from(document.querySelectorAll("#"+id+" input:checked")).map(function(i){return i.value}).filter(function(v){return v&&v!=="on"}))}function monthsSel(){return new Set(Array.from(document.querySelectorAll("#fMonth .rf-month.on span")).map(function(x){return String(x.textContent||"").trim()}))}function lk(){var t=document.querySelector("#groupLevels label.on span"),v=String(t&&t.textContent||"S3").toLowerCase();return v==="brand"?"brand":v==="s1"?"s1":v==="s2"?"s2":"s3"}function threshold(){var x=document.getElementById("capacityThreshold");return Math.max(0,Number(x&&x.value)||0)}function filtered(){var c=sel("fCustomer"),cc=sel("fCode"),cp=sel("fCP"),ch=sel("fChannel"),br=sel("fBrand"),sku=sel("fSKU"),gr=sel("fGroup"),mo=monthsSel(),k=lk();return V.filter(function(r){if(c.size&&!c.has(r.customer))return false;if(cc.size&&!cc.has(r.customerCode))return false;if(cp.size&&!cp.has(r.cp))return false;if(ch.size&&!ch.has(r.channel))return false;if(br.size&&!br.has(r.brand))return false;if(sku.size&&!sku.has(r.sku))return false;if(gr.size&&!gr.has(String(r[k]||"")))return false;if(mo.size&&!mo.has(ml(r.month)))return false;return true})}function value(r){return r.month<CUR?r.net:r.forecastCapacity}function groupText(r){return [String(r.productGroupCode||""),String(r[lk()]||"")].filter(Boolean).join("｜")}function data(){var a=filtered(),ms=uniq(a.map(function(r){return r.month})),m=new Map(),th=threshold();a.forEach(function(r){var v=value(r),unknown=v==null;if(!unknown&&Math.abs(Number(v)||0)<th)return;var g=String(r[lk()]||""),k=[r.customer,r.customerCode,r.cp,r.channel,r.productGroupCode,g].join("\\u241e"),x=m.get(k);if(!x){x={customer:r.customer,customerCode:r.customerCode,activity:[r.cp,r.activity].filter(Boolean).join("｜"),channel:r.channel,group:groupText(r),values:{},unknown:false};m.set(k,x)}if(unknown)x.unknown=true;else x.values[r.month]=(x.values[r.month]||0)+Number(v||0)});var out=Array.from(m.values()).filter(function(x){return x.unknown||ms.some(function(mm){return Number(x.values[mm]||0)!==0})}).sort(function(a,b){return cmp(a.customer,b.customer)||cmp(a.channel,b.channel)||cmp(a.group,b.group)});return{months:ms,rows:out}}function render(){var p=data(),h=document.getElementById("v5PlanHead"),b=document.getElementById("v5PlanBody"),cnt=document.getElementById("v5PlanCount");if(!h||!b)return;h.textContent="";b.textContent="";var hr=document.createElement("tr");["客户","活动","渠道","产品组"].forEach(function(t){var th=document.createElement("th");th.textContent=t;hr.appendChild(th)});p.months.forEach(function(mm){var th=document.createElement("th");th.className="num";th.textContent=ml(mm);hr.appendChild(th)});var th=document.createElement("th");th.className="num";th.textContent="合计";hr.appendChild(th);h.appendChild(hr);var totals={};p.months.forEach(function(mm){totals[mm]=0});var grand=0;p.rows.forEach(function(g){var tr=document.createElement("tr"),td=document.createElement("td"),n=document.createElement("div"),c=document.createElement("span");n.textContent=g.customer;c.className="rf-sub";c.textContent=g.customerCode;td.appendChild(n);td.appendChild(c);tr.appendChild(td);[g.activity,g.channel,g.group].forEach(function(v,i){var x=document.createElement("td");x.textContent=v;if(i===2)x.className="v5-pg";tr.appendChild(x)});var rt=0;p.months.forEach(function(mm){var v=Number(g.values[mm]||0);rt+=v;totals[mm]+=v;var x=document.createElement("td");x.className="num";x.textContent=money(v);tr.appendChild(x)});grand+=rt;var z=document.createElement("td");z.className="num";z.style.fontWeight="700";z.textContent=money(rt);tr.appendChild(z);b.appendChild(tr)});if(p.rows.length){var tr=document.createElement("tr");tr.className="v5-total";var x=document.createElement("td");x.colSpan=4;x.textContent="合计";tr.appendChild(x);p.months.forEach(function(mm){var y=document.createElement("td");y.className="num";y.textContent=money(totals[mm]);tr.appendChild(y)});var y=document.createElement("td");y.className="num";y.textContent=money(grand);tr.appendChild(y);b.appendChild(tr)}else{var tr=document.createElement("tr"),x=document.createElement("td");x.colSpan=5+p.months.length;x.className="rf-empty";x.textContent="当前筛选无可见承载";tr.appendChild(x);b.appendChild(tr)}if(cnt)cnt.textContent=p.rows.length+" 行 · 阈值 "+money(threshold())+" 元"}function codeMap(){var m=new Map();V.forEach(function(r){if(r.sku&&r.productGroupCode&&!m.has(r.sku))m.set(r.sku,r.productGroupCode)});return m}function groupLabels(){var k=lk(),m=new Map();V.forEach(function(r){var g=String(r[k]||""),c=String(r.productGroupCode||"");if(!g||!c)return;if(!m.has(g))m.set(g,new Set());m.get(g).add(c)});document.querySelectorAll("#fGroup .rf-opt").forEach(function(l){var i=l.querySelector("input"),s=l.querySelector("span");if(!i||!s)return;var base=String(i.value||""),cs=m.get(base),pre=cs&&cs.size?Array.from(cs).sort(cmp).join("/")+"｜":"";s.textContent=pre+base})}function detailGroups(){var m=codeMap();document.querySelectorAll("#detailBody tr").forEach(function(tr){var cells=tr.children;if(cells.length<5)return;var sku=String(cells[4].textContent||"").split("｜")[0].trim(),code=m.get(sku)||"",cell=cells[3],base=cell.getAttribute("data-v5-base");if(base==null){base=String(cell.textContent||"");cell.setAttribute("data-v5-base",base)}cell.textContent=[code,base].filter(Boolean).join("｜")})}function toggleDetail(){var d=document.querySelector(".rf-detail");if(!d||d.querySelector(".v5-detail-toggle"))return;d.classList.add("v5-collapsed");var b=document.createElement("button");b.type="button";b.className="v5-detail-toggle";b.textContent="展开 SKU 明细";b.onclick=function(){var on=d.classList.toggle("v5-collapsed");b.textContent=on?"展开 SKU 明细":"收起 SKU 明细"};var bar=d.querySelector(".rf-bar");if(bar)bar.appendChild(b)}function refresh(){groupLabels();detailGroups();render();toggleDetail();var p=document.getElementById("v5Plan"),d=document.querySelector(".rf-detail");if(p&&d&&p.nextElementSibling!==d)d.parentNode.insertBefore(p,d)}var ob=new MutationObserver(function(){refresh()});var s=document.getElementById("summaryBody"),d=document.getElementById("detailBody");if(s)ob.observe(s,{childList:true,subtree:true});if(d)ob.observe(d,{childList:true,subtree:true});var t=document.getElementById("capacityThreshold");if(t)t.addEventListener("change",refresh);setTimeout(refresh,0)})();</script>`;
}

function addSheet(wb: XLSX.WorkBook, name: string, rows: Record<string, unknown>[], widths: number[]) {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 提示: "无数据" }]);
  ws["!cols"] = widths.map((wch) => ({ wch }));
  ws["!autofilter"] = ws["!ref"] ? { ref: ws["!ref"] } : undefined;
  XLSX.utils.book_append_sheet(wb, ws, name);
  return ws;
}

function buildFilteredWorkbook(payload: any) {
  const wb = XLSX.utils.book_new();
  const summary = Array.isArray(payload?.summaryRows) ? payload.summaryRows.slice(0, 10000) : [];
  const detail = Array.isArray(payload?.detailRows) ? payload.detailRows.slice(0, 50000) : [];
  addSheet(wb, "筛选汇总", summary, [30,34,14,14,14,14,14,14,16]);
  const headers = detail.length ? Object.keys(detail[0]) : [];
  addSheet(wb, "筛选明细", detail, headers.length ? headers.map((h) => h === "客户" ? 30 : h === "活动" ? 34 : h === "SKU" ? 38 : h === "产品组" ? 24 : h.includes("月") ? 14 : 16) : [20]);
  addSheet(wb, "筛选说明", [{ 字段: "当前指标", 值: String(payload?.metricLabel || "") }, { 字段: "产品组层级", 值: String(payload?.groupLevel || "") }, { 字段: "可见承载阈值", 值: String(payload?.capacityThreshold || 500) }], [18,30]);
  return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer);
}

function renderResult(r: AnalysisResult) {
  const skuCount = new Set((r.so || []).map((x: any) => String(x.sku || "")).filter(Boolean)).size;
  const cards = `<div class="cards"><div class="card"><span>BI SO行</span><b>${money(r.so.length)}</b></div><div class="card"><span>库存行</span><b>${money(r.stock.length)}</b></div><div class="card"><span>SKU Rate有效/原始</span><b>${money(r.rateInfo.valid.length)} / ${money(r.rateInfo.raw)}</b></div><div class="card"><span>TTS有效折扣行</span><b>${money(r.tts.length)}</b></div><div class="card"><span>可用CP</span><b>${money(r.active.length)}</b></div><div class="card"><span>BI SKU</span><b>${money(skuCount)}</b></div></div>`;
  const note = `<div class="ok">产品目录上传已取消；库存金额直接取 BI-Stock 的 SOH (Good) 千分位金额。可见承载阈值默认 500 元，可在结果页随时调整。</div>`;
  return page("费用规划助手｜承载结果", filePanel(r) + cards + note + renderFilterDashboardV5(r) + renderPlanningAddon(r));
}

async function handler(req: Request) {
  const url = new URL(req.url), path = url.pathname;
  if (path === "/" && req.method === "GET") {
    if (await isAuthorized(req)) return new Response(null, { status: 303, headers: { location: "/workspace" } });
    return new Response(loginPage(), { headers: { "content-type": "text/html;charset=utf-8", "cache-control": "no-store" } });
  }
  if (path === "/auth" && req.method === "POST") {
    const form = await req.formData();
    if (!passwordMatches(String(form.get("password") ?? ""))) return new Response(loginPage("密码不正确"), { status: 401, headers: { "content-type": "text/html;charset=utf-8", "cache-control": "no-store" } });
    return new Response(null, { status: 303, headers: { location: "/workspace", "set-cookie": await createSessionCookie(), "cache-control": "no-store" } });
  }
  if (path === "/logout") return new Response(null, { status: 303, headers: { location: "/", "set-cookie": clearSessionCookie(), "cache-control": "no-store" } });
  if (!(await isAuthorized(req))) return new Response(null, { status: 303, headers: { location: "/" } });
  if (path === "/workspace" && req.method === "GET") return new Response(workspacePage(), { headers: { "content-type": "text/html;charset=utf-8", "cache-control": "no-store" } });

  if (path === "/analyze" && req.method === "POST") {
    try {
      const f = await req.formData();
      const bi = f.get("bi"), tts = f.get("tts"), cp = f.get("cp");
      if (!(bi instanceof File) || !(tts instanceof File) || !(cp instanceof File)) return new Response(page("文件不完整", '<div class="warn">BI、TTS、CP预算三个文件都必须上传。</div>'), { status: 400, headers: { "content-type": "text/html;charset=utf-8" } });
      const [{ analysisBi, syntheticProduct, originalPack }, ttsPack, cpPack] = await Promise.all([prepareBiForAnalysis(bi), makePack(tts), makePack(cp)]);
      const r = await analyzeFiles({ bi: analysisBi, ttsFile: tts, cpFile: cp, product: syntheticProduct, focus: String(f.get("focus") ?? ""), minBalance: Number(f.get("minBalance") ?? 500) });
      (r as any).packs = [originalPack, ttsPack, cpPack];
      attachProductGroupCodes(r);
      return new Response(renderResult(r), { headers: { "content-type": "text/html;charset=utf-8", "cache-control": "no-store" } });
    } catch (e) {
      return new Response(page("计算失败", `<div class="warn">${esc(e instanceof Error ? e.message : String(e))}</div>`), { status: 500, headers: { "content-type": "text/html;charset=utf-8" } });
    }
  }

  if (path === "/export-filtered" && req.method === "POST") {
    try {
      const contentLength = Number(req.headers.get("content-length") || 0);
      if (contentLength > MAX_FILTER_EXPORT_BYTES) return new Response("筛选结果过大，请缩小筛选范围后再导出。", { status: 413, headers: { "content-type": "text/plain;charset=utf-8" } });
      const payload = await req.json();
      const bytes = buildFilteredWorkbook(payload);
      const filename = `费用规划助手_筛选结果_${new Date().toISOString().slice(0,10).replace(/-/g, "")}.xlsx`;
      return new Response(bytes, { headers: { "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`, "cache-control": "no-store" } });
    } catch (e) {
      return new Response(`导出失败：${e instanceof Error ? e.message : String(e)}`, { status: 500, headers: { "content-type": "text/plain;charset=utf-8" } });
    }
  }

  if (path === "/health") return Response.json({ ok: true, platform: "deno-deploy", version: "filter-v5", time: new Date().toISOString() });
  return new Response("Not Found", { status: 404 });
}

Deno.serve(handler);
