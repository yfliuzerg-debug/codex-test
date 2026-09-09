import { Buffer } from "node:buffer";
import * as XLSX from "xlsx";
import { renderFilterDashboard } from "./filter_ui.ts";

(globalThis as any).Buffer = Buffer;

const engine = await import("./netlify/functions/_shared/engine.mts");
const { analyzeFiles, BASELINE, money } = engine;

type AnalysisResult = Awaited<ReturnType<typeof analyzeFiles>>;

const COOKIE = "expense_planner_auth";
const SESSION_SECONDS = 8 * 60 * 60;
const DOWNLOAD_TTL_MS = 30 * 60 * 1000;
const CHUNK_SIZE = 60_000;

const env = (name: string) => Deno.env.get(name) || "";
const esc = (v: unknown) => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;");

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function hmac(value: string) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(env("AUTH_SECRET")),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value));
  return bytesToHex(new Uint8Array(sig));
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

function loginPage(message = "") {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>费用规划助手</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.box{width:min(430px,calc(100% - 32px));margin:14vh auto;background:#fff;border:1px solid #dbe3ef;border-radius:16px;padding:26px}.muted{color:#6b7688}.warn{background:#fff3e8;border:1px solid #f0cda5;border-radius:8px;padding:9px;margin:12px 0}.field{width:100%;box-sizing:border-box;padding:11px;border:1px solid #cfd8e6;border-radius:9px;margin:10px 0 14px}.btn{width:100%;border:0;background:#1769ff;color:#fff;padding:11px;border-radius:9px;font-weight:600}</style></head><body><div class="box"><h2 style="margin:0 0 6px">费用规划助手</h2><div class="muted">仅限授权同事使用 · Deno Deploy</div>${message ? `<div class="warn">${esc(message)}</div>` : ""}<form action="/auth" method="post"><input class="field" type="password" name="password" autocomplete="current-password" placeholder="共享密码" required autofocus><button class="btn" type="submit">进入</button></form></div></body></html>`;
}

function workspacePage() {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>费用规划助手</title><style>:root{--nav:#0d1b2a;--bg:#f4f6fa;--panel:#fff;--line:#dbe3ef;--text:#142033;--muted:#6b7688;--blue:#1769ff}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.layout{display:grid;grid-template-columns:210px 1fr;min-height:100vh}.side{background:var(--nav);color:#fff;padding:22px 18px}.side h1{font-size:18px;margin:0 0 4px}.side p{font-size:12px;color:#b7c3d5;margin:0 0 30px}.step{padding:10px 12px;border-radius:9px;margin:7px 0;color:#c9d4e5}.step.active{background:#17365c;color:#fff}.main{padding:26px 28px}.head{display:flex;justify-content:space-between;gap:12px}.head h2{font-size:24px;margin:0 0 4px}.head p{color:var(--muted);margin:0 0 20px}.head a{color:#1769ff;text-decoration:none}.notice{padding:12px 14px;border:1px solid #cfe0ff;background:#eef5ff;border-radius:10px;margin-bottom:18px}.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.file{border:1px dashed #b9c8dd;border-radius:12px;padding:14px}.file b{display:block;margin-bottom:5px}.file small{display:block;color:var(--muted);margin-bottom:10px}.file input{width:100%}.row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.field label{display:block;font-weight:600;margin-bottom:6px}.field input{width:100%;border:1px solid #cfd8e6;border-radius:9px;padding:10px}.actions{display:flex;justify-content:flex-end;margin-top:16px}.btn{border:0;background:var(--blue);color:#fff;padding:11px 18px;border-radius:9px;font-weight:600}.foot{color:var(--muted);font-size:12px;margin-top:14px}@media(max-width:850px){.layout{grid-template-columns:1fr}.side{display:none}.grid,.row{grid-template-columns:1fr}.main{padding:18px}}</style></head><body><div class="layout"><aside class="side"><h1>费用规划助手</h1><p>双人共享 · Deno Deploy</p><div class="step active">1 数据更新</div><div class="step">2 CP范围</div><div class="step">3 承载 / 拍板</div><div class="step">4 计划输出</div></aside><main class="main"><div class="head"><div><h2>数据更新</h2><p>BI + TTS + CP预算 + 产品目录 → 服务端解析 → 承载候选 → 结果筛选</p></div><a href="/logout">退出</a></div><div class="notice">登录有效8小时。原始文件只用于本次请求计算；计算完成后可按客户、活动、月份、渠道、Brand、SKU、产品组反复筛选，无需重新上传。</div><form class="panel" action="/analyze" method="post" enctype="multipart/form-data"><div class="grid"><div class="file"><b>BI</b><small>BI-SO(折扣)、BI-Stock、SKURate、产品目录。</small><input required type="file" name="bi" accept=".xlsx,.xlsm,.xls"></div><div class="file"><b>TTS 最新下载</b><small>读取同CP既有计划、重叠费率和回溯备注日期。</small><input required type="file" name="tts" accept=".xlsx,.xlsm,.xls"></div><div class="file"><b>CP预算</b><small>读取余额、有效期、MG1、渠道、机制。</small><input required type="file" name="cp" accept=".xlsx,.xlsm,.xls"></div><div class="file"><b>产品目录 / 指导价</b><small>读取 SKU、产品描述、NPS未税箱价。</small><input required type="file" name="product" accept=".xlsx,.xlsm,.xls"></div></div><div class="row"><div class="field"><label>重点客户关键词</label><input name="focus" placeholder="可留空；多个客户用逗号分隔"></div><div class="field"><label>最小展示CP余额</label><input name="minBalance" type="number" value="500" min="0" step="100"></div></div><div class="actions"><button class="btn" type="submit">上传并计算</button></div><div class="foot">一次计算完成后，筛选动作只在浏览器本地执行。</div></form></main></div></body></html>`;
}

function page(title: string, body: string, downloadToken = "") {
  const download = downloadToken ? `<a class="btn" href="/download?token=${encodeURIComponent(downloadToken)}">下载分析结果.xlsx</a>` : "";
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>body{margin:0;background:#f4f6fa;color:#142033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.w{max-width:1700px;margin:auto;padding:22px}.top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.actions{display:flex;gap:10px}.top a{color:#1769ff;text-decoration:none}.btn{display:inline-block;background:#1769ff;color:#fff!important;border-radius:9px;padding:9px 13px;font-weight:600}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin:16px 0}.card,.p{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:13px}.card b{display:block;font-size:22px}.card span,.muted{color:#6b7688}.p{margin:14px 0;overflow:auto}.ok{background:#ecfbf3;border:1px solid #bde8cf;padding:10px;border-radius:9px}.warn{background:#fff8e8;border:1px solid #f0d89a;padding:10px;border-radius:9px}table{border-collapse:collapse;width:100%;min-width:1300px}th,td{border-bottom:1px solid #edf0f5;padding:8px;text-align:left}th{background:#f8fafc}.num{text-align:right}.tag{background:#eef5ff;color:#1769ff;border-radius:999px;padding:2px 7px;font-size:12px}.low{color:#98a2b3}.bad{color:#b42318;font-weight:600}.good{color:#067647;font-weight:600}.activity{min-width:220px;max-width:320px}@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.top{display:block}.actions{margin-top:10px}}</style></head><body><div class="w"><div class="top"><div><h2 style="margin:0">${esc(title)}</h2><div class="muted">Deno Deploy · 浏览器本地筛选</div></div><div class="actions">${download}<a href="/workspace">重新上传</a><a href="/logout">退出</a></div></div>${body}</div></body></html>`;
}

function filePanel(r: AnalysisResult) {
  const rows = r.packs.map((p: any, i: number) => {
    const exp = BASELINE[i];
    const ok = p.sha.startsWith(exp.sha);
    return `<tr><td>${esc(exp.role)}</td><td>${esc(p.name)}</td><td class="num">${(p.size / 1024).toFixed(1)}</td><td><code>${p.sha.slice(0, 12)}</code></td><td class="${ok ? "good" : "bad"}">${ok ? "与验收基线一致" : `不同；基线 ${exp.sha}`}</td></tr>`;
  }).join("");
  return `<div class="p"><h3>本次材料指纹</h3><table style="min-width:850px"><thead><tr><th>角色</th><th>文件名</th><th class="num">大小KB</th><th>SHA-256前12位</th><th>基线状态</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function exportPayload(r: AnalysisResult) {
  return {
    version: "2026-09-deno-v2",
    generatedAt: r.generatedAt,
    currentMonth: r.currentMonth,
    overlayRows: r.exportRows,
    summaryRows: r.matrix.map((x: any) => ({ 客户: x.name, CP: x.cp.code, 活动: x.cp.desc, CP余额: x.cp.balance, 历史HARD: x.hard, 本月已实现: x.currentActual, 本月预测上限: x.currentForecast, 未来SOFT: x.soft, 同CP扣减: x.deduct, 可见上限: x.total, 未知费率桶: x.unknown })),
    fingerprintRows: r.packs.map((p: any, i: number) => ({ 角色: BASELINE[i].role, 文件名: p.name, 大小KB: +(p.size / 1024).toFixed(1), SHA256: p.sha, 基线SHA前12位: BASELINE[i].sha, 是否与验收基线一致: p.sha.startsWith(BASELINE[i].sha) ? "是" : "否" })),
    ruleRows: [
      { 规则: "历史HARD", 说明: "按历史实际SO测算；同CP已存在计划按同客户×渠道×产品范围×月份最高折扣率扣减。" },
      { 规则: "重叠折扣", 说明: "历史回溯只扣同CP既有计划，不让其他CP互相吞掉承载。" },
      { 规则: "500元阈值", 说明: "单月净坑位低于500元保留在明细，但不计入可见承载汇总。" },
      { 规则: "本月预测", 说明: "已实现SO保持实际；剩余月份由趋势SO与库存隐含SO加权，库存按4周，权重按月初10%/月中20%/月末30%。" },
      { 规则: "未来SOFT", 说明: "9月采用R-156预测；库存不直接加入未来SOFT。" },
      { 规则: "用途", 说明: "Overlay明细用于区域自主等后续Overlay，不直接替代人工拍板。" },
    ],
  };
}

function addSheet(wb: XLSX.WorkBook, name: string, rows: Record<string, unknown>[], widths: number[]) {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 提示: "无数据" }]);
  ws["!cols"] = widths.map((wch) => ({ wch }));
  ws["!autofilter"] = ws["!ref"] ? { ref: ws["!ref"] } : undefined;
  XLSX.utils.book_append_sheet(wb, ws, name);
  return ws;
}

function buildWorkbook(payload: any) {
  const wb = XLSX.utils.book_new();
  const overlay = Array.isArray(payload.overlayRows) ? payload.overlayRows : [];
  const summary = Array.isArray(payload.summaryRows) ? payload.summaryRows : [];
  const fingerprints = Array.isArray(payload.fingerprintRows) ? payload.fingerprintRows : [];
  const rules = Array.isArray(payload.ruleRows) ? payload.ruleRows : [];
  const ws = addSheet(wb, "Overlay明细", overlay, [16,24,12,15,28,12,14,34,22,16,10,14,14,10,14,14,14,14,14,42]);
  if (overlay.length) {
    const headers = Object.keys(overlay[0]);
    const rateCol = headers.indexOf("费率");
    const moneyNames = new Set(["CP余额","SO","预测SO","毛承载","TTS已占用","净承载","预测承载","计入可见承载"]);
    for (let r = 1; r <= overlay.length; r++) {
      if (rateCol >= 0) { const c = ws[XLSX.utils.encode_cell({ r, c: rateCol })]; if (c && typeof c.v === "number") c.z = "0.00%"; }
      headers.forEach((h, i) => { if (!moneyNames.has(h)) return; const c = ws[XLSX.utils.encode_cell({ r, c: i })]; if (c && typeof c.v === "number") c.z = "#,##0.00"; });
    }
  }
  addSheet(wb, "客户CP汇总", summary, [28,16,30,14,14,14,14,14,14,14,12]);
  addSheet(wb, "材料指纹", fingerprints, [14,42,12,68,20,20]);
  addSheet(wb, "规则说明", rules, [18,80]);
  const array = XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  return new Uint8Array(array);
}

async function saveDownload(bytes: Uint8Array) {
  const kv = await Deno.openKv();
  const token = crypto.randomUUID();
  const parts = Math.ceil(bytes.length / CHUNK_SIZE);
  const filename = `费用规划助手_分析结果_${new Date().toISOString().slice(0,10).replace(/-/g, "")}.xlsx`;
  const expiresAt = Date.now() + DOWNLOAD_TTL_MS;
  await kv.set(["download", token, "meta"], { parts, filename, expiresAt }, { expireIn: DOWNLOAD_TTL_MS });
  for (let i = 0; i < parts; i++) {
    await kv.set(["download", token, "part", i], bytes.slice(i * CHUNK_SIZE, Math.min(bytes.length, (i + 1) * CHUNK_SIZE)), { expireIn: DOWNLOAD_TTL_MS });
  }
  return token;
}

async function takeDownload(token: string) {
  const kv = await Deno.openKv();
  const metaRes = await kv.get<{parts:number;filename:string;expiresAt:number}>(["download", token, "meta"]);
  const meta = metaRes.value;
  if (!meta || meta.expiresAt < Date.now()) return null;
  const parts: Uint8Array[] = [];
  let total = 0;
  for (let i = 0; i < meta.parts; i++) {
    const res = await kv.get<Uint8Array>(["download", token, "part", i]);
    if (!(res.value instanceof Uint8Array)) return null;
    parts.push(res.value); total += res.value.length;
  }
  const out = new Uint8Array(total); let offset = 0;
  for (const part of parts) { out.set(part, offset); offset += part.length; }
  await kv.delete(["download", token, "meta"]);
  for (let i = 0; i < meta.parts; i++) await kv.delete(["download", token, "part", i]);
  return { bytes: out, filename: meta.filename };
}

function renderResult(r: AnalysisResult, token = "", downloadError = "") {
  const cards = `<div class="cards"><div class="card"><span>BI SO行</span><b>${money(r.so.length)}</b></div><div class="card"><span>库存行</span><b>${money(r.stock.length)}</b></div><div class="card"><span>SKU Rate有效/原始</span><b>${money(r.rateInfo.valid.length)} / ${money(r.rateInfo.raw)}</b></div><div class="card"><span>TTS有效折扣行</span><b>${money(r.tts.length)}</b></div><div class="card"><span>可用CP</span><b>${money(r.active.length)}</b></div><div class="card"><span>目录SKU</span><b>${money(r.prices.size)}</b></div></div>`;
  const note = `<div class="ok">一次计算后的筛选在浏览器本地执行；客户→客户编码、活动→CP、SKU→产品名称、产品组→S3/S2/S1/Brand 路径均保留隐藏关联。</div>${downloadError ? `<div class="warn">分析与筛选可正常使用，但下载缓存未建立：${esc(downloadError)}。Deno KV 绑定后即可恢复 xlsx 下载。</div>` : ""}`;
  return page("费用规划助手｜承载结果", filePanel(r) + cards + note + renderFilterDashboard(r), token);
}

async function handler(req: Request) {
  const url = new URL(req.url);
  const path = url.pathname;

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
      const bi = f.get("bi"), tts = f.get("tts"), cp = f.get("cp"), product = f.get("product");
      if (!(bi instanceof File) || !(tts instanceof File) || !(cp instanceof File) || !(product instanceof File)) return new Response(page("文件不完整", '<div class="warn">四个文件都必须上传。</div>'), { status: 400, headers: { "content-type": "text/html;charset=utf-8" } });
      const r = await analyzeFiles({ bi, ttsFile: tts, cpFile: cp, product, focus: String(f.get("focus") ?? ""), minBalance: Number(f.get("minBalance") ?? 500) });
      let token = ""; let downloadError = "";
      try { token = await saveDownload(buildWorkbook(exportPayload(r))); } catch (e) { downloadError = e instanceof Error ? e.message : String(e); }
      return new Response(renderResult(r, token, downloadError), { headers: { "content-type": "text/html;charset=utf-8", "cache-control": "no-store" } });
    } catch (e) {
      return new Response(page("计算失败", `<div class="warn">${esc(e instanceof Error ? e.message : String(e))}</div>`), { status: 500, headers: { "content-type": "text/html;charset=utf-8" } });
    }
  }

  if (path === "/download" && req.method === "GET") {
    const token = url.searchParams.get("token") || "";
    if (!/^[0-9a-f-]{36}$/i.test(token)) return new Response("无效下载令牌", { status: 400 });
    const result = await takeDownload(token);
    if (!result) return new Response("下载结果已失效或已下载，请重新上传四个文件生成。", { status: 410, headers: { "content-type": "text/plain;charset=utf-8" } });
    return new Response(result.bytes, { headers: { "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(result.filename)}`, "cache-control": "no-store" } });
  }

  if (path === "/health") return Response.json({ ok: true, platform: "deno-deploy", version: "filter-v2", time: new Date().toISOString() });
  return new Response("Not Found", { status: 404 });
}

Deno.serve(handler);
