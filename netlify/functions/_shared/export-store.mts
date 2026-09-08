import type { Context } from "@netlify/functions";
import { getDeployStore, getStore } from "@netlify/blobs";

const STORE_NAME = "expense-analysis-exports";

function storeFor(context: Context) {
  if (context.deploy?.context === "production") {
    return getStore(STORE_NAME, { consistency: "strong" });
  }
  return getDeployStore(STORE_NAME);
}

export async function saveExport(context: Context, payload: unknown) {
  const token = crypto.randomUUID();
  await storeFor(context).setJSON(token, payload);
  return token;
}

export async function takeExport(context: Context, token: string) {
  const store = storeFor(context);
  const data = await store.get(token, { type: "json" });
  if (data != null) await store.delete(token);
  return data;
}
