import { createHmac, timingSafeEqual } from "node:crypto";

const COOKIE = "expense_planner_auth";
const SESSION_SECONDS = 8 * 60 * 60;

function getEnv(name: string) {
  return Netlify.env.get(name) || "";
}

function safeEqual(a: string, b: string) {
  const aa = Buffer.from(a);
  const bb = Buffer.from(b);
  return aa.length === bb.length && timingSafeEqual(aa, bb);
}

function sign(exp: number) {
  return createHmac("sha256", getEnv("AUTH_SECRET")).update(String(exp)).digest("hex");
}

export function passwordMatches(value: string) {
  const expected = getEnv("SHARED_PASSWORD");
  return !!expected && safeEqual(value, expected);
}

export function createSessionCookie() {
  const exp = Math.floor(Date.now() / 1000) + SESSION_SECONDS;
  const token = `${exp}.${sign(exp)}`;
  return `${COOKIE}=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${SESSION_SECONDS}`;
}

export function clearSessionCookie() {
  return `${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`;
}

export function isAuthorized(req: Request) {
  const secret = getEnv("AUTH_SECRET");
  if (!secret) return false;
  const cookie = req.headers.get("cookie") || "";
  const match = cookie.match(new RegExp(`(?:^|;\\s*)${COOKIE}=([^;]+)`));
  if (!match) return false;
  const [expText, signature] = match[1].split(".");
  const exp = Number(expText);
  if (!Number.isFinite(exp) || exp < Math.floor(Date.now() / 1000) || !signature) return false;
  return safeEqual(signature, sign(exp));
}

export function loginRedirect() {
  return new Response(null, { status: 303, headers: { location: "/" } });
}
