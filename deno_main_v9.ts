let capturedHandler: ((req: Request) => Response | Promise<Response>) | null = null;
const originalServe = Deno.serve.bind(Deno);

(Deno as any).serve = (handler: (req: Request) => Response | Promise<Response>) => {
  capturedHandler = handler;
  return { finished: Promise.resolve(), shutdown() {} };
};

await import("./deno_main_v8.ts");
(Deno as any).serve = originalServe;

if (!capturedHandler) throw new Error("未捕获 v8 HTTP handler");

function ttsSuggestion(v: unknown) {
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.floor((n + 1e-9) / 100) * 100;
}

function patchExportPayload(payload: any) {
  const summaryRows = Array.isArray(payload?.summaryRows)
    ? payload.summaryRows.map((r: any) => ({ ...r, TTS建议金额: ttsSuggestion(r?.["可见上限"]) }))
    : [];
  return { ...payload, summaryRows };
}

function patchHTML(html: string) {
  return html.replace(
    "var a=v.actual>=capacityThreshold?v.actual:0,p=v.forecast>=capacityThreshold?v.forecast:0;",
    "var a=Math.abs(v.actual)>=capacityThreshold?v.actual:0,p=Math.abs(v.forecast)>=capacityThreshold?v.forecast:0;",
  );
}

async function patchHTMLResponse(res: Response) {
  const ct = res.headers.get("content-type") || "";
  if (!ct.includes("text/html")) return res;
  const headers = new Headers(res.headers);
  headers.delete("content-length");
  return new Response(patchHTML(await res.text()), { status: res.status, statusText: res.statusText, headers });
}

Deno.serve(async (req: Request) => {
  const url = new URL(req.url);

  if (url.pathname === "/export-filtered" && req.method === "POST") {
    try {
      const payload = patchExportPayload(await req.json());
      const headers = new Headers(req.headers);
      headers.set("content-type", "application/json");
      headers.delete("content-length");
      const forwarded = new Request(req.url, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
      });
      return capturedHandler!(forwarded);
    } catch {
      return capturedHandler!(req);
    }
  }

  return patchHTMLResponse(await capturedHandler!(req));
});
