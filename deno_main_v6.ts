let capturedHandler: ((req: Request) => Response | Promise<Response>) | null = null;
const originalServe = Deno.serve.bind(Deno);

(Deno as any).serve = (handler: (req: Request) => Response | Promise<Response>) => {
  capturedHandler = handler;
  return { finished: Promise.resolve(), shutdown() {} };
};

await import("./deno_main_v5.ts");
(Deno as any).serve = originalServe;

if (!capturedHandler) throw new Error("未捕获 v5 HTTP handler");

Deno.serve(async (req: Request) => {
  const res = await capturedHandler!(req);
  const ct = res.headers.get("content-type") || "";
  if (!ct.includes("text/html")) return res;

  let html = await res.text();
  html = html.replace(
    'cell.textContent=[code,base].filter(Boolean).join("｜")',
    'var next=[code,base].filter(Boolean).join("｜");if(cell.textContent!==next)cell.textContent=next',
  );
  html = html.replace(
    'body:JSON.stringify({summaryRows:exportSummary(),detailRows:exportDetail(),metricLabel:ml,groupLevel:groupLevel})',
    'body:JSON.stringify({summaryRows:exportSummary(),detailRows:exportDetail(),metricLabel:ml,groupLevel:groupLevel,capacityThreshold:capacityThreshold})',
  );

  const headers = new Headers(res.headers);
  headers.delete("content-length");
  return new Response(html, { status: res.status, statusText: res.statusText, headers });
});
