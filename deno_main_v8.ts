import { Buffer } from "node:buffer";
import * as XLSX from "xlsx";

(globalThis as any).Buffer = Buffer;

let capturedHandler: ((req: Request) => Response | Promise<Response>) | null = null;
const originalServe = Deno.serve.bind(Deno);

(Deno as any).serve = (handler: (req: Request) => Response | Promise<Response>) => {
  capturedHandler = handler;
  return { finished: Promise.resolve(), shutdown() {} };
};

await import("./deno_main_v7.ts");
(Deno as any).serve = originalServe;

if (!capturedHandler) throw new Error("未捕获 v7 HTTP handler");

const TARGET_BU = new Set(["dairy", "coffee", "confectionery"]);

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

function sheetRows(wb: any, name?: string) {
  const sn = name && wb.SheetNames.includes(name) ? name : wb.SheetNames[0];
  const ws = wb?.Sheets?.[sn];
  if (!ws) return [] as any[][];
  expandRef(ws);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as any[][];
}

function aliasIndex(h: any[], aliases: string[]) {
  const vals = h.map((v) => String(v ?? "").trim());
  for (const a of aliases) {
    const i = vals.indexOf(a);
    if (i >= 0) return i;
  }
  return -1;
}

function buKey(v: unknown) {
  const s = String(v ?? "").trim().toLowerCase();
  if (s.includes("dairy") || s === "奶品") return "dairy";
  if (s.includes("coffee") || s.includes("咖啡")) return "coffee";
  if (s.includes("confection") || s.includes("糖果")) return "confectionery";
  if (s === "rtd" || s.includes("ready to drink")) return "rtd";
  return s;
}

function buLabel(v: unknown) {
  const k = buKey(v);
  if (k === "dairy") return "Dairy";
  if (k === "coffee") return "Coffee";
  if (k === "confectionery") return "Confectionery";
  return "";
}

async function cpBUMap(file: File) {
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(Buffer.from(ab), { type: "buffer", cellDates: true, raw: true });
  const rs = sheetRows(wb, wb.SheetNames.includes("CPCenturyBigTableByCGAmount") ? "CPCenturyBigTableByCGAmount" : undefined);
  let hi = -1;
  for (let i = 0; i < Math.min(12, rs.length); i++) {
    const code = aliasIndex(rs[i], ["Activity Code", "活动编码"]);
    const bu = aliasIndex(rs[i], ["BU", "业务单元"]);
    if (code >= 0 && bu >= 0) { hi = i; break; }
  }
  const out: Record<string, string> = {};
  if (hi < 0) return out;
  const iCode = aliasIndex(rs[hi], ["Activity Code", "活动编码"]);
  const iBU = aliasIndex(rs[hi], ["BU", "业务单元"]);
  for (const row of rs.slice(hi + 1)) {
    const code = String(row[iCode] ?? "").trim();
    const key = buKey(row[iBU]);
    if (!code || !TARGET_BU.has(key)) continue;
    out[code] = buLabel(key);
  }
  return out;
}

function patchHierarchy(html: string, buByCP: Record<string, string>) {
  if (!html.includes('id="rfPanel"')) return html;

  const buJson = JSON.stringify(buByCP).replace(/</g, "\\u003c");
  html = html.replace(
    "</head>",
    `<script>window.__EP_BU_BY_CP=${buJson};</script></head>`,
  );

  const activityCard = '<div class="rf-card"><div class="rf-title"><span>活动</span><span class="muted" id="nCP"></span></div><div class="rf-list" id="fCP"></div></div>';
  const buCard = '<div class="rf-card"><div class="rf-title"><span>BU</span><span class="muted" id="nBU"></span></div><div class="rf-list" id="fBU"></div></div>';
  if (!html.includes('id="fBU"')) html = html.replace(activityCard, buCard + activityCard);

  html = html.replace(
    "活动显示 CP｜活动名称；SKU 显示 SKU编码｜产品名称；产品组按 Brand → S1 → S2 → S3 单层切换。",
    "层级筛选：BU → 活动CP → Brand → SKU → 产品组；后级只显示当前上级范围内的有效项。客户、客户代码、渠道和月份保持独立筛选。",
  );

  html = html.replace(
    'var S={customer:new Set(),customerCode:new Set(),cp:new Set(),channel:new Set(),brand:new Set(),sku:new Set(),group:new Set(),month:new Set()};',
    'var BU_BY_CP=window.__EP_BU_BY_CP||{};R.forEach(function(r){r.bu=String(BU_BY_CP[String(r.cp||"")]||"")});var S={bu:new Set(),customer:new Set(),customerCode:new Set(),cp:new Set(),channel:new Set(),brand:new Set(),sku:new Set(),group:new Set(),month:new Set()};',
  );

  const oldPairs = 'function pairs(key){var m=new Map();R.forEach(function(r){var v=String(r[key]||"");if(!v||m.has(v))return;m.set(v,key==="cp"?r.activityLabel:key==="sku"?r.skuLabel:v)});return Array.from(m.entries()).sort(function(a,b){return cmp(a[1],b[1])})}';
  const newPairs = 'function hierarchyRows(key){return R.filter(function(r){if(key!=="bu"&&S.bu.size&&!S.bu.has(r.bu))return false;if((key==="brand"||key==="sku"||key==="group")&&S.cp.size&&!S.cp.has(r.cp))return false;if((key==="sku"||key==="group")&&S.brand.size&&!S.brand.has(r.brand))return false;if(key==="group"&&S.sku.size&&!S.sku.has(r.sku))return false;return true})}function pairs(key){var m=new Map();hierarchyRows(key).forEach(function(r){var v=String(r[key]||"");if(!v||m.has(v))return;m.set(v,key==="cp"?r.activityLabel:key==="sku"?r.skuLabel:v)});return Array.from(m.entries()).sort(function(a,b){return cmp(a[1],b[1])})}function keepValid(set,vals){var ok=new Set(vals);Array.from(set).forEach(function(v){if(!ok.has(v))set.delete(v)})}function pruneHierarchy(){keepValid(S.cp,pairs("cp").map(function(p){return p[0]}));keepValid(S.brand,pairs("brand").map(function(p){return p[0]}));keepValid(S.sku,pairs("sku").map(function(p){return p[0]}));keepValid(S.group,uniq(hierarchyRows("group").map(function(r){return r[groupLevel]})))}';
  html = html.replace(oldPairs, newPairs);

  html = html.replace(
    'function groups(){var b=document.getElementById("fGroup");b.textContent="";uniq(R.map(function(r){return r[groupLevel]})).forEach(function(v){b.appendChild(opt(v,v,S.group,apply))});',
    'function groups(){var b=document.getElementById("fGroup");b.textContent="";uniq(hierarchyRows("group").map(function(r){return r[groupLevel]})).forEach(function(v){b.appendChild(opt(v,v,S.group,apply))});',
  );

  html = html.replace(
    'groupLevel=d.key;S.group.clear();groupLevels();groups();apply()',
    'groupLevel=d.key;S.group.clear();groupLevels();apply()',
  );

  html = html.replace(
    'function match(r){if(S.customer.size&&!S.customer.has(r.customer))return false;',
    'function match(r){if(S.bu.size&&!S.bu.has(r.bu))return false;if(S.customer.size&&!S.customer.has(r.customer))return false;',
  );

  html = html.replace(
    'var n=S.customer.size+S.customerCode.size+S.cp.size+S.channel.size+S.brand.size+S.sku.size+S.group.size+S.month.size;',
    'var n=S.bu.size+S.customer.size+S.customerCode.size+S.cp.size+S.channel.size+S.brand.size+S.sku.size+S.group.size+S.month.size;',
  );

  html = html.replace(
    'function apply(){lastRows=R.filter(match);',
    'function apply(){pruneHierarchy();lastRows=R.filter(match);',
  );

  html = html.replace(
    'simple("cp","fCP","nCP");simple("channel","fChannel","nChannel");simple("brand","fBrand","nBrand");simple("sku","fSKU","nSKU");groups();',
    'simple("bu","fBU","nBU");simple("cp","fCP","nCP");simple("channel","fChannel","nChannel");simple("brand","fBrand","nBrand");simple("sku","fSKU","nSKU");groups();',
  );

  const oldPlanFiltered = 'function filtered(){var c=sel("fCustomer"),cc=sel("fCode"),cp=sel("fCP"),ch=sel("fChannel"),br=sel("fBrand"),sku=sel("fSKU"),gr=sel("fGroup"),mo=monthsSel(),k=lk();return V.filter(function(r){if(c.size&&!c.has(r.customer))return false;';
  const newPlanFiltered = 'function filtered(){var bu=sel("fBU"),c=sel("fCustomer"),cc=sel("fCode"),cp=sel("fCP"),ch=sel("fChannel"),br=sel("fBrand"),sku=sel("fSKU"),gr=sel("fGroup"),mo=monthsSel(),k=lk(),bm=window.__EP_BU_BY_CP||{};return V.filter(function(r){if(bu.size&&!bu.has(String(bm[String(r.cp||"")]||"")))return false;if(c.size&&!c.has(r.customer))return false;';
  html = html.replace(oldPlanFiltered, newPlanFiltered);

  return html;
}

async function patchResponse(res: Response, buByCP: Record<string, string>) {
  const ct = res.headers.get("content-type") || "";
  if (!ct.includes("text/html")) return res;
  const html = patchHierarchy(await res.text(), buByCP);
  const headers = new Headers(res.headers);
  headers.delete("content-length");
  return new Response(html, { status: res.status, statusText: res.statusText, headers });
}

function cloneFormData(src: FormData) {
  const out = new FormData();
  for (const [k, v] of src.entries()) {
    if (v instanceof File) out.append(k, v, v.name);
    else out.append(k, String(v));
  }
  return out;
}

Deno.serve(async (req: Request) => {
  const url = new URL(req.url);
  if (url.pathname === "/analyze" && req.method === "POST") {
    const form = await req.formData();
    const cp = form.get("cp");
    let map: Record<string, string> = {};
    if (cp instanceof File) {
      try { map = await cpBUMap(cp); } catch { map = {}; }
    }
    const headers = new Headers();
    const cookie = req.headers.get("cookie");
    if (cookie) headers.set("cookie", cookie);
    const forwarded = new Request(req.url, { method: "POST", headers, body: cloneFormData(form) });
    const res = await capturedHandler!(forwarded);
    return patchResponse(res, map);
  }

  return capturedHandler!(req);
});