import { Buffer } from "node:buffer";
import * as XLSX from "xlsx";

(globalThis as any).Buffer = Buffer;

const BRAND_TAG = "\u2063__EP3BU__";
const TARGET_BU = new Set(["dairy", "coffee", "confectionery"]);

type Meta = {
  sku: string;
  bu: string;
  brandCode: string;
  brand: string;
  s1: string;
  s2: string;
  s3: string;
  skuName: string;
  aliases: Set<string>;
  inSO: boolean;
};

type FilePatch = {
  transformedSha12: string;
  originalSha12: string;
  transformedKB: string;
  originalKB: string;
};

let capturedHandler: ((req: Request) => Response | Promise<Response>) | null = null;
const originalServe = Deno.serve.bind(Deno);

(Deno as any).serve = (handler: (req: Request) => Response | Promise<Response>) => {
  capturedHandler = handler;
  return { finished: Promise.resolve(), shutdown() {} };
};

await import("./deno_main_v6.ts");
(Deno as any).serve = originalServe;

if (!capturedHandler) throw new Error("未捕获 v6 HTTP handler");

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function sha256(ab: ArrayBuffer) {
  return bytesToHex(new Uint8Array(await crypto.subtle.digest("SHA-256", ab)));
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

function sheetRows(wb: any, name?: string) {
  const sn = name && wb.SheetNames.includes(name) ? name : wb.SheetNames[0];
  const ws = wb?.Sheets?.[sn];
  if (!ws) return [] as any[][];
  expandRef(ws);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as any[][];
}

function rowValue(row: any[], i: number | undefined) {
  return Number.isInteger(i) && (i as number) >= 0 && (i as number) < row.length ? row[i as number] : null;
}

function headerMap(h: any[]) {
  const out: Record<string, number> = {};
  h.forEach((v, i) => { if (v != null && String(v).trim()) out[String(v).trim()] = i; });
  return out;
}

function findHeader(rs: any[][], need: string[]) {
  return rs.findIndex((r) => {
    const a = r.map((v) => String(v ?? "").trim());
    return need.every((k) => a.includes(k));
  });
}

function norm(v: unknown) {
  return String(v ?? "")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, "")
    .replace(/有限责任公司$/, "有限公司")
    .toLowerCase();
}

function buKey(v: unknown) {
  const s = String(v ?? "").trim().toLowerCase();
  if (!s) return "";
  if (s.includes("dairy") || s === "奶品") return "dairy";
  if (s.includes("coffee") || s.includes("咖啡")) return "coffee";
  if (s.includes("confection") || s.includes("糖果")) return "confectionery";
  if (s === "rtd" || s.includes("ready to drink")) return "rtd";
  return s;
}

function list(v: unknown) {
  return String(v ?? "").split(/[,，;]/).map((x) => x.trim()).filter(Boolean);
}

function taggedBrand(brand: string, sku: string) {
  return brand && sku ? `${brand}${BRAND_TAG}${sku}` : brand;
}

function fillMeta(m: Meta, patch: Partial<Omit<Meta, "aliases" | "inSO">>, alias?: string) {
  if (!m.bu && patch.bu) m.bu = patch.bu;
  if (!m.brandCode && patch.brandCode) m.brandCode = patch.brandCode;
  if (!m.brand && patch.brand) m.brand = patch.brand;
  if (!m.s1 && patch.s1) m.s1 = patch.s1;
  if (!m.s2 && patch.s2) m.s2 = patch.s2;
  if (!m.s3 && patch.s3) m.s3 = patch.s3;
  if (!m.skuName && patch.skuName) m.skuName = patch.skuName;
  if (patch.brand) m.aliases.add(patch.brand);
  if (alias) m.aliases.add(alias);
}

function emptyMeta(sku: string): Meta {
  return { sku, bu: "", brandCode: "", brand: "", s1: "", s2: "", s3: "", skuName: "", aliases: new Set(), inSO: false };
}

function buildMeta(wb: any) {
  const meta = new Map<string, Meta>();

  if (wb.SheetNames.includes("产品目录")) {
    const rs = sheetRows(wb, "产品目录");
    const hi = findHeader(rs, ["SKU", "SKU Name CN"]);
    if (hi >= 0) {
      const x = headerMap(rs[hi]);
      for (const r of rs.slice(hi + 1)) {
        const sku = String(rowValue(r, x["SKU"]) ?? "").trim();
        if (!sku) continue;
        const m = meta.get(sku) || emptyMeta(sku);
        fillMeta(m, {
          bu: buKey(rowValue(r, x["BU"])),
          brandCode: String(rowValue(r, x["Brand"]) ?? "").trim(),
          brand: String(rowValue(r, x["Brand Name"]) ?? "").trim(),
          s1: String(rowValue(r, x["PH5 Segment1_MKT"]) ?? "").trim(),
          s2: String(rowValue(r, x["PH5 Segment2_Detail"]) ?? "").trim(),
          s3: String(rowValue(r, x["PH5 Segment3_Detail"]) ?? "").trim(),
          skuName: String(rowValue(r, x["SKU Name CN"]) ?? "").trim(),
        });
        meta.set(sku, m);
      }
    }
  }

  if (wb.SheetNames.includes("BI-Stock")) {
    const rs = sheetRows(wb, "BI-Stock");
    const hi = findHeader(rs, ["Customer Name", "SKU", "SOH (Good)"]);
    if (hi >= 0) {
      const x = headerMap(rs[hi]);
      for (const r of rs.slice(hi + 1)) {
        const sku = String(rowValue(r, x["SKU"]) ?? "").trim();
        if (!sku) continue;
        const m = meta.get(sku) || emptyMeta(sku);
        fillMeta(m, {
          bu: buKey(rowValue(r, x["BU"])),
          brandCode: String(rowValue(r, x["Brand"]) ?? "").trim(),
          brand: String(rowValue(r, x["Brand Name"]) ?? "").trim(),
          s1: String(rowValue(r, x["PH5 Segment1_MKT"]) ?? "").trim(),
          s2: String(rowValue(r, x["PH5 Segment2_Detail"]) ?? "").trim(),
          s3: String(rowValue(r, x["PH5 Segment3_Detail"]) ?? "").trim(),
          skuName: String(rowValue(r, x["SKU Name CN"]) ?? "").trim(),
        });
        meta.set(sku, m);
      }
    }
  }

  if (wb.SheetNames.includes("BI-SO(折扣)")) {
    const rs = sheetRows(wb, "BI-SO(折扣)");
    const hi = findHeader(rs, ["Outlet Channel", "Customer Code", "SKU"]);
    if (hi >= 0) {
      const x = headerMap(rs[hi]);
      for (const r of rs.slice(hi + 1)) {
        const sku = String(rowValue(r, x["SKU"]) ?? "").trim();
        if (!sku) continue;
        const m = meta.get(sku) || emptyMeta(sku);
        m.inSO = true;
        fillMeta(m, {
          brandCode: String(rowValue(r, x["Brand"]) ?? "").trim(),
          brand: String(rowValue(r, x["Brand Name"]) ?? "").trim(),
          s1: String(rowValue(r, x["PH5 Segment1_MKT"]) ?? "").trim(),
          s2: String(rowValue(r, x["PH5 Segment2_Detail"]) ?? "").trim(),
          skuName: String(rowValue(r, x["SKU Name CN"]) ?? "").trim(),
        });
        meta.set(sku, m);
      }
    }
  }

  if (wb.SheetNames.includes("SKURate")) {
    const rs = sheetRows(wb, "SKURate");
    const hi = findHeader(rs, ["SKUCode", "StartDate", "EndDate", "Rate"]);
    if (hi >= 0) {
      const x = headerMap(rs[hi]);
      for (const r of rs.slice(hi + 1)) {
        const sku = String(rowValue(r, x["SKUCode"]) ?? "").trim();
        if (!sku || !meta.has(sku)) continue;
        const m = meta.get(sku)!;
        fillMeta(m, {
          bu: buKey(rowValue(r, x["BUName"])),
          skuName: String(rowValue(r, x["SKUName"]) ?? "").trim(),
        }, String(rowValue(r, x["MG1Name"]) ?? "").trim());
      }
    }
  }

  return meta;
}

function escapeRe(v: string) {
  return v.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function isNegativeToken(text: string, token: string) {
  const t = String(token || "").trim();
  if (t.length < 2) return false;
  const e = escapeRe(t);
  const left = `(?:^|[^A-Za-z0-9])${e}(?=$|[^A-Za-z0-9])`;
  return new RegExp(`${left}.{0,12}(?:无折扣|不支持|不可核销|不可使用|不得核销|不得使用|不可|不得)`, "is").test(text) ||
    new RegExp(`(?:不含|排除).{0,12}${left}`, "is").test(text);
}

function matchesConfectioneryPositive(text: string, m: Meta, pool: Meta[]) {
  if (/pnm[-－]coated/i.test(text)) return norm(m.s1) === "pnm" && norm(m.s2 || m.s3) === "coated";
  if (/ip口味/i.test(text)) return norm(m.s3) === "ip";

  const strong = /限定产品|仅限|只限/.test(text) || /适用于[^。；\n]{0,40}折扣/.test(text);
  if (!strong) return true;

  const labels = new Set<string>();
  for (const p of pool) {
    for (const v of [p.s1, p.s2, p.s3]) {
      const s = String(v || "").trim();
      if (s.length < 2 || isNegativeToken(text, s)) continue;
      if (norm(text).includes(norm(s))) labels.add(norm(s));
    }
  }
  if (!labels.size) return true;
  return [m.s1, m.s2, m.s3].some((v) => labels.has(norm(v)));
}

function scopeAllows(text: string, m: Meta, pool: Meta[]) {
  for (const tok of [m.sku, m.s1, m.s2, m.s3]) if (isNegativeToken(text, tok)) return false;
  if (m.bu === "confectionery") return matchesConfectioneryPositive(text, m, pool);
  return true;
}

function aliasIndex(h: any[], aliases: string[]) {
  const vals = h.map((v) => String(v ?? "").trim());
  for (const a of aliases) {
    const i = vals.indexOf(a);
    if (i >= 0) return i;
  }
  return -1;
}

const CP_ALIASES = {
  code: ["Activity Code", "活动编码"],
  start: ["Activity Start Date", "活动开始日期"],
  end: ["Activity End Date", "活动结束日期"],
  desc: ["Activity Description", "活动描述"],
  mech: ["Activity Mechanism", "活动机制"],
  bu: ["BU", "业务单元"],
  mg1: ["MG1", "促销品牌"],
  subChannel: ["SubChannel", "渠道"],
  balance: ["Balance Amount", "CG可使用预算金额"],
};

function findCPHeader(rs: any[][]) {
  for (let i = 0; i < Math.min(rs.length, 12); i++) {
    const h = rs[i];
    if (aliasIndex(h, CP_ALIASES.code) >= 0 && aliasIndex(h, CP_ALIASES.start) >= 0 && aliasIndex(h, CP_ALIASES.balance) >= 0) return i;
  }
  return -1;
}

async function transformBI(file: File) {
  const originalAB = await file.arrayBuffer();
  const wb = XLSX.read(Buffer.from(originalAB), { type: "buffer", cellDates: true, raw: true });
  const meta = buildMeta(wb);

  if (!wb.SheetNames.includes("BI-SO(折扣)")) throw new Error("BI 文件缺少 BI-SO(折扣) Sheet。");
  const soRows = sheetRows(wb, "BI-SO(折扣)");
  const soHi = findHeader(soRows, ["Outlet Channel", "Customer Code", "SKU"]);
  if (soHi < 0) throw new Error("BI-SO(折扣) 未识别到 Outlet Channel / Customer Code / SKU 表头。");
  const soX = headerMap(soRows[soHi]);
  for (const r of soRows.slice(soHi + 1)) {
    const sku = String(rowValue(r, soX["SKU"]) ?? "").trim();
    const m = meta.get(sku);
    if (!m || !sku) continue;
    const brand = m.brand || String(rowValue(r, soX["Brand Name"]) ?? "").trim();
    if (soX["Brand"] != null && m.brandCode) r[soX["Brand"]] = m.brandCode;
    if (soX["Brand Name"] != null && brand) r[soX["Brand Name"]] = taggedBrand(brand, sku);
    if (soX["PH5 Segment1_MKT"] != null && m.s1) r[soX["PH5 Segment1_MKT"]] = m.s1;
    if (soX["PH5 Segment2_Detail"] != null && m.s2) r[soX["PH5 Segment2_Detail"]] = m.s2;
    if (soX["SKU Name CN"] != null && m.skuName) r[soX["SKU Name CN"]] = m.skuName;
  }
  wb.Sheets["BI-SO(折扣)"] = XLSX.utils.aoa_to_sheet(soRows);

  const dirHeader = ["BU", "Brand", "Brand Name", "PH5 Segment1_MKT", "PH5 Segment2_Detail", "PH5 Segment3_Detail", "SKU", "SKU Name CN"];
  const dirRows: any[][] = [dirHeader];
  for (const m of Array.from(meta.values()).sort((a, b) => a.sku.localeCompare(b.sku, "zh-CN", { numeric: true }))) {
    if (!m.sku || !m.inSO) continue;
    dirRows.push([m.bu.toUpperCase(), m.brandCode, taggedBrand(m.brand, m.sku), m.s1, m.s2, m.s3, m.sku, m.skuName]);
  }
  wb.Sheets["产品目录"] = XLSX.utils.aoa_to_sheet(dirRows);
  if (!wb.SheetNames.includes("产品目录")) wb.SheetNames.push("产品目录");

  const transformedAB = XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  const transformed = new File([transformedAB], file.name.replace(/\.xlsm?$/i, ".xlsx"), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  return {
    file: transformed,
    meta,
    originalSha: await sha256(originalAB),
    transformedSha: await sha256(transformedAB),
    originalSize: originalAB.byteLength,
    transformedSize: transformedAB.byteLength,
  };
}

async function transformCP(file: File, meta: Map<string, Meta>) {
  const originalAB = await file.arrayBuffer();
  const wb = XLSX.read(Buffer.from(originalAB), { type: "buffer", cellDates: true, raw: true });
  const rs = sheetRows(wb, wb.SheetNames.includes("CPCenturyBigTableByCGAmount") ? "CPCenturyBigTableByCGAmount" : undefined);
  const hi = findCPHeader(rs);
  if (hi < 0) throw new Error("CP预算未识别到活动编码 / 开始日期 / 可用余额字段（支持中英文表头）。");
  const h = rs[hi];
  const iCode = aliasIndex(h, CP_ALIASES.code), iStart = aliasIndex(h, CP_ALIASES.start), iEnd = aliasIndex(h, CP_ALIASES.end);
  const iDesc = aliasIndex(h, CP_ALIASES.desc), iMech = aliasIndex(h, CP_ALIASES.mech), iBU = aliasIndex(h, CP_ALIASES.bu);
  const iMG1 = aliasIndex(h, CP_ALIASES.mg1), iSub = aliasIndex(h, CP_ALIASES.subChannel), iBal = aliasIndex(h, CP_ALIASES.balance);

  const out: any[][] = [["Activity Code", "Activity Start Date", "Activity End Date", "Activity Description", "Activity Mechanism", "BU", "MG1", "SubChannel", "Balance Amount"]];
  const currentMeta = Array.from(meta.values()).filter((m) => m.inSO && TARGET_BU.has(m.bu));

  for (const r of rs.slice(hi + 1)) {
    const code = String(rowValue(r, iCode) ?? "").trim();
    if (!code) continue;
    const bu = buKey(rowValue(r, iBU));
    if (!TARGET_BU.has(bu)) continue;
    const desc = String(rowValue(r, iDesc) ?? "").trim();
    const mech = String(rowValue(r, iMech) ?? "").trim();
    const text = `${desc} ${mech}`;
    const originalMG1 = list(rowValue(r, iMG1));
    const mgNorm = new Set(originalMG1.map(norm));
    const basePool = currentMeta.filter((m) => {
      if (m.bu !== bu) return false;
      if (!mgNorm.size) return true;
      return Array.from(m.aliases).some((a) => mgNorm.has(norm(a)));
    });
    const allowed = basePool.filter((m) => scopeAllows(text, m, basePool));
    const syntheticMG1 = Array.from(new Set(allowed.map((m) => taggedBrand(m.brand, m.sku)).filter(Boolean)));

    out.push([
      code,
      rowValue(r, iStart),
      rowValue(r, iEnd),
      desc,
      mech,
      bu === "dairy" ? "Dairy" : bu === "coffee" ? "Coffee" : "Confectionery",
      syntheticMG1.length ? syntheticMG1.join(",") : "__EP_NO_MATCH__",
      String(rowValue(r, iSub) ?? "").trim(),
      rowValue(r, iBal),
    ]);
  }

  const owb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(owb, XLSX.utils.aoa_to_sheet(out), "CPCenturyBigTableByCGAmount");
  const transformedAB = XLSX.write(owb, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  const transformed = new File([transformedAB], file.name.replace(/\.xlsm?$/i, ".xlsx"), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  return {
    file: transformed,
    originalSha: await sha256(originalAB),
    transformedSha: await sha256(transformedAB),
    originalSize: originalAB.byteLength,
    transformedSize: transformedAB.byteLength,
  };
}

function patchFor(originalSha: string, transformedSha: string, originalSize: number, transformedSize: number): FilePatch {
  return {
    originalSha12: originalSha.slice(0, 12),
    transformedSha12: transformedSha.slice(0, 12),
    originalKB: (originalSize / 1024).toFixed(1),
    transformedKB: (transformedSize / 1024).toFixed(1),
  };
}

function patchHTML(html: string, patches: FilePatch[] = []) {
  html = html.replace('name="minBalance" type="number" value="500"', 'name="minBalance" type="number" value="100"');
  html = html.replace(
    "产品目录文件已取消。库存预测直接使用 BI-Stock 的 SOH (Good) 千分位库存金额；一次计算完成后筛选只在浏览器本地执行。",
    "已启用 Dairy / Coffee / Confectionery 3BU 兼容；RTD 不纳入。产品层级优先使用 BI 内置产品目录，并由 BI-Stock / SKURate 补齐缺失映射；库存预测继续使用 SOH (Good) 千分位库存金额。",
  );
  html = html.replace(
    "产品目录上传已取消；库存金额直接取 BI-Stock 的 SOH (Good) 千分位金额。可见承载阈值默认 500 元，可在结果页随时调整。",
    "3BU 兼容已启用（Dairy / Coffee / Confectionery，RTD 排除）；库存金额直接取 BI-Stock 的 SOH (Good) 千分位金额。可见承载阈值仍默认 500 元，可在结果页随时调整。",
  );

  const tagRe = new RegExp(`${BRAND_TAG}[0-9A-Za-z._-]+`, "g");
  html = html.replace(tagRe, "");
  html = html.replace(/\\u2063__EP3BU__[0-9A-Za-z._-]+/g, "");

  for (const p of patches) {
    html = html.replace(
      `<td class="num">${p.transformedKB}</td><td><code>${p.transformedSha12}</code>`,
      `<td class="num">${p.originalKB}</td><td><code>${p.originalSha12}</code>`,
    );
  }
  return html;
}

async function responseWithPatchedHTML(res: Response, patches: FilePatch[] = []) {
  const ct = res.headers.get("content-type") || "";
  if (!ct.includes("text/html")) return res;
  const html = patchHTML(await res.text(), patches);
  const headers = new Headers(res.headers);
  headers.delete("content-length");
  return new Response(html, { status: res.status, statusText: res.statusText, headers });
}

Deno.serve(async (req: Request) => {
  const url = new URL(req.url);
  if (url.pathname === "/analyze" && req.method === "POST") {
    try {
      const f = await req.formData();
      const bi = f.get("bi"), tts = f.get("tts"), cp = f.get("cp");
      if (!(bi instanceof File) || !(tts instanceof File) || !(cp instanceof File)) {
        return capturedHandler!(req);
      }

      const biX = await transformBI(bi);
      const cpX = await transformCP(cp, biX.meta);
      const nf = new FormData();
      nf.append("bi", biX.file, biX.file.name);
      nf.append("tts", tts, tts.name);
      nf.append("cp", cpX.file, cpX.file.name);
      nf.append("focus", String(f.get("focus") ?? ""));
      const minBalance = String(f.get("minBalance") ?? "").trim();
      nf.append("minBalance", minBalance || "100");

      const headers = new Headers();
      const cookie = req.headers.get("cookie");
      if (cookie) headers.set("cookie", cookie);
      const forwarded = new Request(req.url, { method: "POST", headers, body: nf });
      const res = await capturedHandler!(forwarded);
      return responseWithPatchedHTML(res, [
        patchFor(biX.originalSha, biX.transformedSha, biX.originalSize, biX.transformedSize),
        patchFor(cpX.originalSha, cpX.transformedSha, cpX.originalSize, cpX.transformedSize),
      ]);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return new Response(`3BU兼容预处理失败：${msg}`, { status: 500, headers: { "content-type": "text/plain;charset=utf-8" } });
    }
  }

  const res = await capturedHandler!(req);
  return responseWithPatchedHTML(res);
});
