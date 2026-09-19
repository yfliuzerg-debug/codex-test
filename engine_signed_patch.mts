import * as base from "./netlify/functions/_shared/engine.mts?base=1";

export * from "./netlify/functions/_shared/engine.mts?base=1";

const EXCLUDE = new Set(["12187803", "12585787", "12598186", "12611209"]);
const daysIn = (y: number, m: number) => new Date(y, m, 0).getDate();

function clientOK(cp: any, name: string) {
  return cp.code !== "DA202601011TT" || ["天虹", "东方"].some((k) => name.includes(k));
}

function channelTokens(bi: string) {
  const x = base.norm(bi);
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

function channelOK(cp: any, ch: string) {
  const c = base.norm(cp.channels);
  return !c || channelTokens(ch).some((v) => c.includes(base.norm(v)));
}

function ttsChannelOK(tch: string, bch: string) {
  const t = base.norm(tch);
  return !t || channelTokens(bch).some((v) => t.includes(base.norm(v)));
}

function eligible(cp: any, r: any) {
  if (!clientOK(cp, r.name) || EXCLUDE.has(r.sku) || !channelOK(cp, r.channel)) return false;
  const explicit: string[] = [];
  if (/新西兰/.test(cp.mech)) explicit.push("新西兰");
  if (/全脂375g/.test(cp.mech)) explicit.push("全脂375g");
  if (/甜奶粉/.test(cp.mech)) explicit.push("甜奶粉300g");
  if (explicit.length) return explicit.some((v) => r.s3.includes(v));
  if (cp.mg1.length) return cp.mg1.some((v: string) => base.norm(v) === base.norm(r.brand));
  return true;
}

function skuRate(rs: any[], sku: string, mm: string) {
  const y = +mm.slice(0, 4);
  const m = +mm.slice(4, 6);
  const s = new Date(y, m - 1, 1);
  const e = new Date(y, m, 0);
  const vals = [...new Set(rs.filter((r) => r.sku === sku && r.start <= e && r.end >= s && r.rate > 0).map((r) => +r.rate.toFixed(6)))];
  return vals.length === 1 ? Number(vals[0]) : null;
}

function mechRate(cp: any, r: any) {
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

function cpRateFromTTS(cp: any, tts: any[]) {
  const u = [...new Set(tts.filter((t) => t.cp === cp.code).map((t) => +t.rate.toFixed(6)))];
  return u.length === 1 ? u[0] : null;
}

function targetRate(cp: any, r: any, rs: any[], tts: any[], mm: string) {
  return cpRateFromTTS(cp, tts) ?? mechRate(cp, r) ?? skuRate(rs, r.sku, mm);
}

function ttsScopeOK(p: any, r: any) {
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

function overlapSameCP(tts: any[], cp: any, r: any, mm: string) {
  let mx = 0;
  for (const p of tts) {
    if (p.cp !== cp.code || base.norm(p.name) !== base.norm(r.name) || !ttsChannelOK(p.channel, r.channel) || !ttsScopeOK(p, r)) continue;
    if (base.months(p.start, p.end).includes(mm)) mx = Math.max(mx, p.rate);
  }
  return mx;
}

function r156(r: any) {
  const ly = r.months["202509"] || 0;
  const avg = ((r.months["202605"] || 0) + (r.months["202606"] || 0) + (r.months["202607"] || 0)) / 3;
  return Math.max(0, Math.max(ly, (ly + avg) / 2) * 1000);
}

function actualBasis(r: any, mm: string) {
  return (r.months[mm] || 0) * 1000;
}

function stockShare(so: any[], r: any, mm: string) {
  const g = so.filter((x) => base.norm(x.name) === base.norm(r.name) && x.sku === r.sku);
  const pos = g.map((x) => Math.max(0, actualBasis(x, mm)));
  const sum = pos.reduce((a, b) => a + b, 0);
  if (sum > 0) return Math.max(0, actualBasis(r, mm)) / sum;
  const hist = g.map((x) => Math.max(0, (x.months["202607"] || 0) + (x.months["202606"] || 0) + (x.months["202605"] || 0)));
  const hs = hist.reduce((a, b) => a + b, 0);
  const i = g.indexOf(r);
  return hs > 0 ? hist[i] / hs : 1 / Math.max(1, g.length);
}

function forecastBasis(so: any[], r: any, stockValue: number, now: Date) {
  const mm = base.month(now);
  const actual = actualBasis(r, mm);
  const done = Math.max(1, now.getDate() - 1);
  const remaining = Math.max(0, daysIn(now.getFullYear(), now.getMonth() + 1) - done);
  const trend = (actual / done) * remaining;
  const inv = (stockValue * stockShare(so, r, mm)) / 4 * (remaining / 7);
  const w = now.getDate() <= 10 ? 0.1 : now.getDate() <= 20 ? 0.2 : 0.3;
  return actual + trend * (1 - w) + inv * w;
}

function productGroup(r: any) {
  return r.s3 || r.s2 || r.s1 || r.brand || "";
}

function hintFor(mm: string, cur: string, unknown: boolean, overlapRate: number, rate: number | null, net: number) {
  const kind = mm < cur ? "历史HARD" : mm === cur ? "本月实际+预测" : "未来SOFT";
  const notes = [kind];
  if (unknown || rate == null) notes.push("费率待确认");
  if (overlapRate > 0) notes.push(`同CP最高重叠费率${(overlapRate * 100).toFixed(2)}%`);
  if (Math.abs(net) > 0 && Math.abs(net) < 500) notes.push("净坑位绝对值<500，不计入可见承载");
  if (rate != null && overlapRate >= rate && rate > 0) notes.push("同CP既有计划已完全覆盖");
  return notes.join("；");
}

export async function analyzeFiles(args: {
  bi: File;
  ttsFile: File;
  cpFile: File;
  product: File;
  focus?: string;
  minBalance?: number;
  now?: Date;
}) {
  const [bp, tp, cpPack, pp] = await Promise.all([
    base.pack(args.bi),
    base.pack(args.ttsFile),
    base.pack(args.cpFile),
    base.pack(args.product),
  ]);

  const so = base.parseSO(bp.wb);
  const stock = base.parseStock(bp.wb);
  const rateInfo = base.parseRates(bp.wb);
  const rates = rateInfo.valid;
  const tts = base.parseTTS(tp.wb);
  const cps = base.parseCP(cpPack.wb);
  const prices = base.parsePrices(pp.wb);
  const focus = String(args.focus ?? "").split(/[,，]/).map((v) => v.trim()).filter(Boolean);
  const minBalance = base.num(args.minBalance ?? 500);
  const now = args.now ?? new Date();
  const cur = base.month(now);
  const active = cps.filter((c) => c.balance >= minBalance).sort((a, b) => b.balance - a.balance);

  const sm = new Map<string, number>();
  for (const s of stock) {
    const k = `${base.norm(s.name)}|${s.sku}`;
    sm.set(k, (sm.get(k) || 0) + s.cases * (prices.get(s.sku) || 0));
  }

  let names = [...new Set(so.map((r) => r.name))];
  if (focus.length) names = names.filter((v) => focus.some((q) => base.norm(v).includes(base.norm(q))));
  else names = names.slice(0, 30);

  const matrix: any[] = [];
  const exportRows: any[] = [];

  for (const name of names) {
    for (const cp of active) {
      const per: Record<string, any> = {};
      for (const mm of base.months(cp.start, cp.end)) {
        per[mm] = { raw: 0, deduct: 0, net: 0, counted: 0, unknown: 0, actualRaw: 0, actualDeduct: 0, actualNet: 0, actualCounted: 0, kind: mm < cur ? "HARD" : mm === cur ? "CURRENT" : "SOFT" };
      }

      for (const r of so) {
        if (base.norm(r.name) !== base.norm(name) || !eligible(cp, r)) continue;
        for (const mm of Object.keys(per)) {
          const actual = actualBasis(r, mm);
          let predictedSO = 0;
          if (mm === cur) predictedSO = forecastBasis(so, r, sm.get(`${base.norm(name)}|${r.sku}`) || 0, now);
          else if (mm > cur && mm === "202609") predictedSO = r156(r);

          const rate = targetRate(cp, r, rates, tts, mm);
          if (rate == null) {
            per[mm].unknown++;
            if (actual !== 0 || predictedSO !== 0) {
              exportRows.push({ CP: cp.code, 活动: cp.desc, CP余额: cp.balance, 客户编码: r.code, "经销商/客户": r.name, 渠道: r.channel, SKU: r.sku, 产品名称: r.skuName, 产品组: productGroup(r), 品牌: r.brand, 月份: mm, SO: actual, 预测SO: predictedSO, 费率: null, 毛承载: null, TTS已占用: null, 净承载: null, 预测承载: null, 计入可见承载: null, "数据口径/提示": hintFor(mm, cur, true, 0, null, 0) });
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

          const basis = mm < cur ? actual : predictedSO;
          if (basis !== 0) {
            per[mm].raw += predictedGross || actualGross;
            per[mm].deduct += predictedDeduct || actualDeduct;
            per[mm].net += predictedNet || actualNet;
          }

          if (actual !== 0 || predictedSO !== 0 || actualDeduct !== 0 || predictedDeduct !== 0) {
            const visibleNet = mm < cur ? actualNet : predictedNet;
            exportRows.push({
              CP: cp.code, 活动: cp.desc, CP余额: cp.balance, 客户编码: r.code, "经销商/客户": r.name, 渠道: r.channel, SKU: r.sku, 产品名称: r.skuName, 产品组: productGroup(r), 品牌: r.brand, 月份: mm,
              SO: actual, 预测SO: predictedSO, 费率: rate, 毛承载: actualGross, TTS已占用: actualDeduct, 净承载: actualNet, 预测承载: predictedNet,
              计入可见承载: Math.abs(visibleNet) >= 500 ? visibleNet : 0,
              "数据口径/提示": hintFor(mm, cur, false, ov, rate, visibleNet),
            });
          }
        }
      }

      for (const mm of Object.keys(per)) {
        per[mm].counted = Math.abs(per[mm].net) >= 500 ? per[mm].net : 0;
        per[mm].actualCounted = Math.abs(per[mm].actualNet) >= 500 ? per[mm].actualNet : 0;
      }

      const hard = Object.entries(per).filter(([m]) => m < cur).reduce((a, [, v]: any) => a + v.counted, 0);
      const currentActual = per[cur]?.actualCounted || 0;
      const currentForecast = per[cur]?.counted || 0;
      const soft = Object.entries(per).filter(([m]) => m > cur).reduce((a, [, v]: any) => a + v.counted, 0);
      const deduct = Object.values(per).reduce((a: number, v: any) => a + v.deduct, 0);
      const unknown = Object.values(per).reduce((a: number, v: any) => a + v.unknown, 0);

      matrix.push({ name, cp, per, hard, currentActual, currentForecast, soft, deduct, total: Math.min(cp.balance, hard + currentForecast + soft), unknown });
    }
  }

  return { packs: [bp, tp, cpPack, pp], so, stock, rateInfo, tts, cps, prices, active, names, matrix, exportRows, currentMonth: cur, generatedAt: now.toISOString() };
}
