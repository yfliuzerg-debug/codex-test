import type { Context, Config } from "@netlify/functions";
import * as XLSX from "xlsx";

export default async (req: Request, context: Context) => {
  return new Response("Backend source is staged on the netlify-backend branch. Link Netlify to that branch for deployment.", { status: 503 });
};

export const config: Config = { path: "/analyze" };
