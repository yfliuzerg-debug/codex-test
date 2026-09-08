import * as XLSX from "xlsx";

export type SO = {
  channel: string;
  code: string;
  name: string;
  brand: string;
  s1: string;
  s2: string;
  s3: string;
  sku: string;
  skuName: string;
  months: Record<string, number>;
};

export type Stock = { name: string; sku: string; cases: number };
export type Rate = { sku: string; start: Date; end: Date; rate: number };
export type TTS = {
  name: string;
  channel: string;
  sku: string;
  start: Date;
  end: Date;
  rate: number;
  cp: string;
  remarks: string;
  productRemark: string;
};
export type CP = {
  code: string;
  start: Date;
  end: Date;
  desc: string;
  mech: string;
  channels: string;
  mg1: string[];
  balance: number;
};
export type Pack = { wb: any; name: string; size: number; sha: string };
export type MonthCell = {
  raw: number;
  deduct: number;
  net: number;
  counted: number;
  unknown: number;
  actualRaw: number;
  actualDeduct: number;
  actualNet: number;
  actualCounted: number;
  kind: string;
};

export type ExportRow = {
  CP: string;
  活动: string;
  CP余额: number;
  客户编码: string;
  "经销商/客户": string;
  渠道: string;
  SKU: string;
  产品名称: string;
  产品组: string;
  品牌: string;
  月份: string;
  SO: number;
  预测SO: number;
  费率: number | null;
  毛承载: number | null;
  TTS已占用: number | null;
  净承载: number | null;
  预测承载: number | null;
  计入可见承载: number | null;
  "数据口径/提示": string;
};

export type MatrixRow = {
  name: string;
  cp: CP;
  per: Record<string, MonthCell>;
  hard: number;
  currentActual: number;
  currentForecast: number;
  soft: number;
  deduct: number;
  total: number;
  unknown: number;
};

export const BASELINE = [
  { role: "BI", sha: "4d3b994be5b9" },
  { role: "TTS", sha: "ce4795e9bc02" },
  { role: "CP预算", sha: "3ca0d72da4b8" },
  { role: "产品目录", sha: "8cf1b9ece38e" },
];

const EXCLUDE = new Set(["12187803", "12585787", "12598186", "12611209"]);

export const norm = (v: any) =>
  String(v ?? "")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, "")
    .replace(/有限责任公司$/, "有限公司")
    .toLowerCase();

export const num = (v: any) => {
  const x = Number(v);
  return Number.isFinite(x) ? x : 0;
};

export const money = (v: number) =>
  new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(v || 0);

export const month = (d: Date) =>
  `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}`;

export const labelMonth = (m: string) => `${+m.slice(4)}月`;

const list = (v: any) =>
  String(v ?? "")
    .split(/[,，;]/)
    .map((x) => x.trim())
    .filter(Boolean);

const parseDate = (v: any): Date | null => {
  if (!v) return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    const o = XLSX.SSF.parse_date_code(v);
    return o ? new Date(o.y, o.m - 1, o.d) : null;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
};

export const months = (a: Date, b: Date) => {
  const out: string[] = [];
  const d = new Date(a.getFullYear(), a.getMonth(), 1);
  const e = new Date(b.getFullYear(), b.getMonth(), 1);
  while (d <= e) {
    out.push(month(d));
    d.setMonth(d.getMonth() + 1);
  }
  return out;
};

const daysIn = (y: number, m: number) => new Date(y, m, 0).getDate();
const parseRate = (v: any) => {
  if (typeof v === "string" && v.includes("%")) return parseFloat(v) / 100;
  const x = num(v);
  return x > 1 ? x / 100 : x;
};

export async function pack(file: File): Promise<Pack> {
  const ab = await file.arrayBuffer();
  const dig = await crypto.subtle.digest("SHA-256", ab);
  const sha = Array.from(new Uint8Array(dig))
    .map((x) => x.toString(16).padStart(2, "0"))
    .join("");
  return {
    wb: XLSX.read(Buffer.from(ab), { type: "buffer", cellDates: true, raw: true }),
    name: file.name,
    size: file.size,
    sha,
  };
}

function expandRef(ws: any) {
  let maxR = 0;
  let maxC = 0;
  let seen = false;
  for (const k of Object.keys(ws)) {
    if (k.startsWith("!")) continue;
    const m = k.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    maxC = Math.max(maxC, XLSX.utils.decode_col(m[1]));
    maxR = Math.max(maxR, +m[2] - 1);
    seen = true;
  }
  if (seen) {
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  }
}

function rows(wb: any, sheet?: string) {
  const sn = sheet && wb.SheetNames.includes(sheet) ? sheet : wb.SheetNames[0];
  const ws = wb.Sheets[sn];
  expandRef(ws);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as any[][];
}

function header(rs: any[][], need: string[]) {
  return rs.findIndex((r) => need.every((k) => r.some((c) => String(c ?? "").trim() === k)));
}

function ix(h: any[]) {
  const out: Record<string, number> = {};
  h.forEach((v, i) => {
    if (v != null) out[String(v).trim()] = i;
  });
  return out;
}

function biMeta(wb: any) {
  const out = new Map<string, { brand: string; s3: string }>();
  if (!wb.SheetNames.includes("产品目录")) return out;
  const rs = rows(wb, "产品目录");
  const hi = header(rs, ["SKU", "SKU Name CN"]);
  if (hi < 0) return out;
  const x = ix(rs[hi]);
  for (const r of rs.slice(hi + 1)) {
    const sku = String(r[x["SKU"]] ?? "");
    if (sku) {
      out.set(sku, {
        brand: String(r[x["Brand Name"]] ?? ""),
        s3: String(r[x["PH5 Segment3_Detail"]] ?? ""),
      });
    }
  }
  return out;
}

export function parseSO(wb: any): SO[] {
  if (!wb.SheetNames.includes("BI-SO(折扣)")) return [];
  const meta = biMeta(wb);
  const rs = rows(wb, "BI-SO(折扣)");
  const hi = header(rs, ["Outlet Channel", "Customer Code", "SKU"]);
  if (hi < 0) return [];
  const h = rs[hi].map((v) => String(v ?? "").trim());
  const x = ix(h);
  const mc = h
    .map((v, i) => (/^20\d{4}$/.test(v) ? [v, i] : null))
    .filter(Boolean) as [string, number][];
  return rs
    .slice(hi + 1)
    .map((r) => {
      const sku = String(r[x["SKU"]] ?? "");
      const m = meta.get(sku);
      return {
        channel: String(r[x["Outlet Channel"]] ?? ""),
        code: String(r[x["Customer Code"]] ?? ""),
        name: String(r[x["Customer Name"]] ?? ""),
        brand: m?.brand || String(r[x["Brand Name"]] ?? ""),
        s1: String(r[x["PH5 Segment1_MKT"]] ?? ""),
        s2: String(r[x["PH5 Segment2_Detail"]] ?? ""),
        s3: m?.s3 || "",
        sku,
        skuName: String(r[x["SKU Name CN"]] ?? ""),
        months: Object.fromEntries(mc.map(([mm, i]) => [mm, num(r[i])])),
      };
    })
    .filter((r) => r.name && r.sku && !/汇总/.test(r.code));
}

export function parseStock(wb: any): Stock[] {
  if (!wb.SheetNames.includes("BI-Stock")) return [];
  const rs = rows(wb, "BI-Stock");
  const hi = header(rs, ["Customer Name", "SKU", "SOH (Good) Case"]);
  if (hi < 0) return [];
  const x = ix(rs[hi]);
  return rs
    .slice(hi + 1)
    .map((r) => ({
      name: String(r[x["Customer Name"]] ?? ""),
      sku: String(r[x["SKU"]] ?? ""),
      cases: num(r[x["SOH (Good) Case"]]),
    }))
    .filter((r) => r.name && r.sku && r.cases !== 0);
}

export function parseRates(wb: any) {
  if (!wb.SheetNames.includes("SKURate")) return { raw: 0, valid: [] as Rate[] };
  const rs = rows(wb, "SKURate");
  const hi = header(rs, ["SKUCode", "StartDate", "EndDate", "Rate"]);
  if (hi < 0) return { raw: 0, valid: [] as Rate[] };
  const x = ix(rs[hi]);
  const out: Rate[] = [];
  let raw = 0;
  for (const r of rs.slice(hi + 1)) {
    const sku = String(r[x["SKUCode"]] ?? "");
    if (!sku) continue;
    raw++;
    const start = parseDate(r[x["StartDate"]]);
    const end = parseDate(r[x["EndDate"]]);
    const rate = parseRate(r[x["_RateValue"]] ?? r[x["Rate"]]);
    if (start && end && rate > 0) out.push({ sku, start, end, rate });
  }
  return { raw, valid: out };
}

export function parsePrices(wb: any) {
  const out = new Map<string, number>();
  for (const sn of wb.SheetNames) {
    const rs = rows(wb, sn);
    const hi = header(rs, ["产品编码"]);
    if (hi < 0) continue;
    const h = rs[hi].map((v) => String(v ?? "").trim());
    const x = ix(h);
    const si = x["产品编码"];
    let pi = h.findIndex((v) => v.includes("NPS") && v.includes("未税箱价"));
    if (pi < 0) pi = x["Price"];
    if (pi == null || pi < 0) continue;
    for (const r of rs.slice(hi + 1)) {
      const sku = String(r[si] ?? "");
      const p = num(r[pi]);
      if (sku && p > 0) out.set(sku, p);
    }
  }
  return out;
}

export function parseTTS(wb: any): TTS[] {
  const rs = rows(wb, wb.SheetNames.includes("PPCenturyBigTable") ? "PPCenturyBigTable" : undefined);
  const hi = header(rs, ["PPCode", "Distributor Name", "Activity Code in Cycle Plan"]);
  if (hi < 0) return [];
  const x = ix(rs[hi]);
  const out: TTS[] = [];
  for (const r of rs.slice(hi + 1)) {
    const rate = parseRate(r[x["Disc.Rate"]]);
    const start = parseDate(r[x["RemarkStartDate"]]) || parseDate(r[x["Promotion Start Date"]]);
    const end = parseDate(r[x["RemarkEndDate"]]) || parseDate(r[x["Promotion End Date"]]);
    if (rate > 0 && start && end && r[x["Distributor Name"]] && r[x["Product Code"]]) {
      out.push({
        name: String(r[x["Distributor Name"]]),
        channel: String(r[x["Channel"]] ?? ""),
        sku: String(r[x["Product Code"]]),
        start,
        end,
        rate,
        cp: String(r[x["Activity Code in Cycle Plan"]] ?? ""),
        remarks: String(r[x["Remarks"]] ?? ""),
        productRemark: String(r[x["Product Remark"]] ?? ""),
      });
    }
  }
  return out;
}

export function parseCP(wb: any): CP[] {
  const rs = rows(wb, wb.SheetNames.includes("CPCenturyBigTableByCGAmount") ? "CPCenturyBigTableByCGAmount" : undefined);
  const hi = header(rs, ["Activity Code", "Activity Start Date", "Balance Amount"]);
  if (hi < 0) return [];
  const x = ix(rs[hi]);
  const out: CP[] = [];
  for (const r of rs.slice(hi + 1)) {
    const start = parseDate(r[x["Activity Start Date"]]);
    const end = parseDate(r[x["Activity End Date"]]);
    const balance = num(r[x["Balance Amount"]]);
    if (start && end && balance > 0 && r[x["Activity Code"]]) {
      out.push({
        code: String(r[x["Activity Code"]]),
        start,
        end,
        desc: String(r[x["Activity Description"]] ?? ""),
        mech: String(r[x["Activity Mechanism"]] ?? ""),
        channels: String(r[x["SubChannel"]] ?? ""),
        mg1: list(r[x["MG1"]]),
        balance,
      });
    }
  }
  return out;
}

function clientOK(cp: CP, name: string) {
  return cp.code !== "DA202601011TT" || ["天虹", "东方"].some((k) => name.includes(k));
}

function channelTokens(bi: string) {
  const x = norm(bi);
  const a: Record<string, string[]> = {
    sm: ["hyper&super", "lkagroup", "lkaothers"],
    hm: ["hyper&super"],
    ws: ["wholesale"],
    "mm-a": ["mm-a"],
    "mm-b": ["mm-b"],
    "mm-c": ["mm-c"],
    speciality: ["specialty"],
    ss_kiosks: ["smallstore"],
    sd1: ["sd1", "sub-d"],
    scd: ["scd"],
    cgp: ["cgp"],
    b2b: ["b2b"],
    b2c: ["b2c"],
    cvs: ["cvs"],
  };
  return a[x] || [x];
}

function channelOK(cp: CP, ch: string) {
  const c = norm(cp.channels);
  return !c || channelTokens(ch).some((v) => c.includes(norm(v)));
}

function ttsChannelOK(tch: string, bch: string) {
  const t = norm(tch);
  return !t || channelTokens(bch).some((v) => t.includes(norm(v)));
}

function eligible(cp: CP, r: SO) {
  if (!clientOK(cp, r.name) || EXCLUDE.has(r.sku) || !channelOK(cp, r.channel)) return false;
  const explicit: string[] = [];
  if (/新西兰/.test(cp.mech)) explicit.push("新西兰");
  if (/全脂375g/.test(cp.mech)) explicit.push("全脂375g");
  if (/甜奶粉/.test(cp.mech)) explicit.push("甜奶粉300g");
  if (explicit.length) return explicit.some((v) => r.s3.includes(v));
  if (cp.mg1.length) return cp.mg1.some((v) => norm(v) === norm(r.brand));
  return true;
}

function skuRate(rs: Rate[], sku: string, mm: string) {
  const y = +mm.slice(0, 4);
  const m = +mm.slice(4, 6);
  const s = new Date(y, m - 1, 1);
  const e = new Date(y, m, 0);
  const vals = [
    ...new Set(
      rs
        .filter((r) => r.sku === sku && r.start <= e && r.end >= s && r.rate > 0)
        .map((r) => +r.rate.toFixed(6)),
    ),
  ];
  return vals.length === 1 ? vals[0] : null;
}

function mechRate(cp: CP, r: SO) {
  const t = (cp.desc + " " + cp.mech).replace(/：/g, ":");
  const s = (r.s1 + " " + r.s2 + " " + r.s3 + " " + r.skuName).toLowerCase();
  const tests: [RegExp, string][] = [
    [/中老年[^%]{0,40}?(\d+(?:\.\d+)?)%/i, "senior"],
    [/N3[^%]{0,40}?(\d+(?:\.\d+)?)%/i, "n3"],
    [/全家[^%]{0,40}?(\d+(?:\.\d+)?)%/i, "family"],
  ];
  for (const [q, k] of tests) {
    if (s.includes(k)) {
      const m = t.match(q);
      if (m) return Number(m[1]) / 100;
    }
  }
  const all = [...t.matchAll(/(\d+(?:\.\d+)?)\s*%/g)].map((m) => Number(m[1]) / 100);
  const u = [...new Set(all)];
  return u.length === 1 ? u[0] : null;
}

function cpRateFromTTS(cp: CP, tts: TTS[]) {
  const u = [...new Set(tts.filter((t) => t.cp === cp.code).map((t) => +t.rate.toFixed(6)))];
  return u.length === 1 ? u[0] : null;
}

function targetRate(cp: CP, r: SO, rs: Rate[], tts: TTS[], mm: string) {
  return cpRateFromTTS(cp, tts) ?? mechRate(cp, r) ?? skuRate(rs, r.sku, mm);
}

function ttsScopeOK(p: TTS, r: SO) {
  if (p.sku === r.sku) return true;
  const z = p.remarks + " " + p.productRemark;
  const s = (r.s1 + " " + r.s2 + " " + r.s3 + " " + r.brand + " " + r.skuName).toLowerCase();
  if (/(该BU|奶品).*(所有SKU|全品).*核销/.test(z)) return true;
  if (r.s3 === "新西兰" && /新西兰/.test(z)) return true;
  if (r.s3 === "全脂375g" && /全脂375g/.test(z)) return true;
  if (r.s3 === "甜奶粉300g" && /(300g甜奶粉|甜奶粉)/.test(z)) return true;
  if (/中老年/.test(z) && s.includes("senior")) return true;
  if (/n3/i.test(z) && s.includes("n3")) return true;
  if (/全家/.test(z) && (s.includes("family") || s.includes("fcmp") || s.includes("全家"))) return true;
  return false;
}

function overlapSameCP(tts: TTS[], cp: CP, r: SO, mm: string) {
  let mx = 0;
  for (const p of tts) {
    if (
      p.cp !== cp.code ||
      norm(p.name) !== norm(r.name) ||
      !ttsChannelOK(p.channel, r.channel) ||
      !ttsScopeOK(p, r)
    ) {
      continue;
    }
    if (months(p.start, p.end).includes(mm)) mx = Math.max(mx, p.rate);
  }
  return mx;
}

function r156(r: SO) {
  const ly = r.months["202509"] || 0;
  const avg = ((r.months["202605"] || 0) + (r.months["202606"] || 0) + (r.months["202607"] || 0)) / 3;
  return Math.max(0, Math.max(ly, (ly + avg) / 2) * 1000);
}

function actualBasis(r: SO, mm: string) {
  return Math.max(0, (r.months[mm] || 0) * 1000);
}

function stockShare(so: SO[], r: SO, mm: string) {
  const g = so.filter((x) => norm(x.name) === norm(r.name) && x.sku === r.sku);
  const pos = g.map((x) => actualBasis(x, mm));
  const sum = pos.reduce((a, b) => a + b, 0);
  if (sum > 0) return actualBasis(r, mm) / sum;
  const hist = g.map((x) => Math.max(0, (x.months["202607"] || 0) + (x.months["202606"] || 0) + (x.months["202605"] || 0)));
  const hs = hist.reduce((a, b) => a + b, 0);
  const i = g.indexOf(r);
  return hs > 0 ? hist[i] / hs : 1 / Math.max(1, g.length);
}

function forecastBasis(so: SO[], r: SO, stockValue: number, now: Date) {
  const mm = month(now);
  const actual = actualBasis(r, mm);
  const done = Math.max(1, now.getDate() - 1);
  const remaining = Math.max(0, daysIn(now.getFullYear(), now.getMonth() + 1) - done);
  const trend = (actual / done) * remaining;
  const inv = (stockValue * stockShare(so, r, mm)) / 4 * (remaining / 7);
  const w = now.getDate() <= 10 ? 0.1 : now.getDate() <= 20 ? 0.2 : 0.3;
  return actual + trend * (1 - w) + inv * w;
}

function productGroup(r: SO) {
  return r.s3 || r.s2 || r.s1 || r.brand || "";
}

function hintFor(mm: string, cur: string, unknown: boolean, overlapRate: number, rate: number | null, net: number) {
  const kind = mm < cur ? "历史HARD" : mm === cur ? "本月实际+预测" : "未来SOFT";
  const notes = [kind];
  if (unknown || rate == null) notes.push("费率待确认");
  if (overlapRate > 0) notes.push(`同CP最高重叠费率${(overlapRate * 100).toFixed(2)}%`);
  if (net > 0 && net < 500) notes.push("净坑位<500，不计入可见承载");
  if (rate != null && overlapRate >= rate && rate > 0) notes.push("同CP既有计划已完全覆盖");
  return notes.join("；");
}

export type AnalysisResult = {
  packs: Pack[];
  so: SO[];
  stock: Stock[];
  rateInfo: { raw: number; valid: Rate[] };
  tts: TTS[];
  cps: CP[];
  prices: Map<string, number>;
  active: CP[];
  names: string[];
  matrix: MatrixRow[];
  exportRows: ExportRow[];
  currentMonth: string;
  generatedAt: string;
};

export async function analyzeFiles(args: {
  bi: File;
  ttsFile: File;
  cpFile: File;
  product: File;
  focus?: string;
  minBalance?: number;
  now?: Date;
}): Promise<AnalysisResult> {
  const [bp, tp, cpPack, pp] = await Promise.all([
    pack(args.bi),
    pack(args.ttsFile),
    pack(args.cpFile),
    pack(args.product),
  ]);

  const so = parseSO(bp.wb);
  const stock = parseStock(bp.wb);
  const rateInfo = parseRates(bp.wb);
  const rates = rateInfo.valid;
  const tts = parseTTS(tp.wb);
  const cps = parseCP(cpPack.wb);
  const prices = parsePrices(pp.wb);
  const focus = String(args.focus ?? "")
    .split(/[,，]/)
    .map((v) => v.trim())
    .filter(Boolean);
  const minBalance = num(args.minBalance ?? 500);
  const now = args.now ?? new Date();
  const cur = month(now);
  const active = cps.filter((c) => c.balance >= minBalance).sort((a, b) => b.balance - a.balance);

  const sm = new Map<string, number>();
  for (const s of stock) {
    const k = `${norm(s.name)}|${s.sku}`;
    sm.set(k, (sm.get(k) || 0) + s.cases * (prices.get(s.sku) || 0));
  }

  let names = [...new Set(so.map((r) => r.name))];
  if (focus.length) names = names.filter((v) => focus.some((q) => norm(v).includes(norm(q))));
  else names = names.slice(0, 30);

  const matrix: MatrixRow[] = [];
  const exportRows: ExportRow[] = [];

  for (const name of names) {
    for (const cp of active) {
      const per: Record<string, MonthCell> = {};
      for (const mm of months(cp.start, cp.end)) {
        per[mm] = {
          raw: 0,
          deduct: 0,
          net: 0,
          counted: 0,
          unknown: 0,
          actualRaw: 0,
          actualDeduct: 0,
          actualNet: 0,
          actualCounted: 0,
          kind: mm < cur ? "HARD" : mm === cur ? "CURRENT" : "SOFT",
        };
      }

      for (const r of so) {
        if (norm(r.name) !== norm(name) || !eligible(cp, r)) continue;
        for (const mm of Object.keys(per)) {
          const actual = actualBasis(r, mm);
          let predictedSO = 0;
          if (mm === cur) predictedSO = forecastBasis(so, r, sm.get(`${norm(name)}|${r.sku}`) || 0, now);
          else if (mm > cur && mm === "202609") predictedSO = r156(r);

          const rate = targetRate(cp, r, rates, tts, mm);
          if (rate == null) {
            per[mm].unknown++;
            if (actual > 0 || predictedSO > 0) {
              exportRows.push({
                CP: cp.code,
                活动: cp.desc,
                CP余额: cp.balance,
                客户编码: r.code,
                "经销商/客户": r.name,
                渠道: r.channel,
                SKU: r.sku,
                产品名称: r.skuName,
                产品组: productGroup(r),
                品牌: r.brand,
                月份: mm,
                SO: actual,
                预测SO: predictedSO,
                费率: null,
                毛承载: null,
                TTS已占用: null,
                净承载: null,
                预测承载: null,
                计入可见承载: null,
                "数据口径/提示": hintFor(mm, cur, true, 0, null, 0),
              });
            }
            continue;
          }

          const ov = overlapSameCP(tts, cp, r, mm);
          const effective = Math.max(0, rate - ov);
          const actualGross = actual * rate;
          const actualDeduct = actual * Math.min(rate, ov);
          const actualNet = actual * effective;
          const predictedGross = predictedSO * rate;
          const predictedDeduct = predictedSO * Math.min(rate, ov);
          const predictedNet = predictedSO * effective;

          per[mm].actualRaw += actualGross;
          per[mm].actualDeduct += actualDeduct;
          per[mm].actualNet += actualNet;

          let base = 0;
          if (mm < cur) base = actual;
          else if (mm === cur) base = predictedSO;
          else base = predictedSO;

          if (base > 0) {
            per[mm].raw += predictedGross || actualGross;
            per[mm].deduct += predictedDeduct || actualDeduct;
            per[mm].net += predictedNet || actualNet;
          }

          if (actual > 0 || predictedSO > 0 || actualDeduct > 0 || predictedDeduct > 0) {
            const visibleNet = mm < cur ? actualNet : mm === cur ? predictedNet : predictedNet;
            exportRows.push({
              CP: cp.code,
              活动: cp.desc,
              CP余额: cp.balance,
              客户编码: r.code,
              "经销商/客户": r.name,
              渠道: r.channel,
              SKU: r.sku,
              产品名称: r.skuName,
              产品组: productGroup(r),
              品牌: r.brand,
              月份: mm,
              SO: actual,
              预测SO: predictedSO,
              费率: rate,
              毛承载: actualGross,
              TTS已占用: actualDeduct,
              净承载: actualNet,
              预测承载: predictedNet,
              计入可见承载: visibleNet >= 500 ? visibleNet : 0,
              "数据口径/提示": hintFor(mm, cur, false, ov, rate, visibleNet),
            });
          }
        }
      }

      for (const mm of Object.keys(per)) {
        per[mm].counted = per[mm].net >= 500 ? per[mm].net : 0;
        per[mm].actualCounted = per[mm].actualNet >= 500 ? per[mm].actualNet : 0;
      }

      const hard = Object.entries(per)
        .filter(([m]) => m < cur)
        .reduce((a, [, v]) => a + v.counted, 0);
      const currentActual = per[cur]?.actualCounted || 0;
      const currentForecast = per[cur]?.counted || 0;
      const soft = Object.entries(per)
        .filter(([m]) => m > cur)
        .reduce((a, [, v]) => a + v.counted, 0);
      const deduct = Object.values(per).reduce((a, v) => a + v.deduct, 0);
      const unknown = Object.values(per).reduce((a, v) => a + v.unknown, 0);

      matrix.push({
        name,
        cp,
        per,
        hard,
        currentActual,
        currentForecast,
        soft,
        deduct,
        total: Math.min(cp.balance, hard + currentForecast + soft),
        unknown,
      });
    }
  }

  return {
    packs: [bp, tp, cpPack, pp],
    so,
    stock,
    rateInfo,
    tts,
    cps,
    prices,
    active,
    names,
    matrix,
    exportRows,
    currentMonth: cur,
    generatedAt: now.toISOString(),
  };
}
